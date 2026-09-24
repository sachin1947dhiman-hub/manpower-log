import io
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from supabase import create_client

# ------------------------------------------------------------
# SETTINGS  (edit these lists any time)
# ------------------------------------------------------------
PART1_STATIONS = ["Dispatch", "Store", "Lab", "Nachos Line", "Peanut Line",
                  "Potato Line", "ETP", "Miscellaneous", "Other"]
PART2_STATIONS = ["Combo", "Tray", "Can", "Pouch Open", "Stickering",
                  "Leakage Check", "Other"]
DEFAULT_UNIT = {"Combo": "Pouch", "Tray": "Tray", "Can": "Can",
                "Pouch Open": "Pouch", "Stickering": "Nos",
                "Leakage Check": "Pouch", "Other": "Nos"}
UNITS = ["Pouch", "Case", "Kg", "Box", "Tray", "Can", "Nos"]
SHIFTS = ["Day", "Night"]

# worker types: (database column, label on screen, shift hours)
TYPES = [("ladies_10", "Ladies 10h", 10),
         ("ladies_12", "Ladies 12h", 12),
         ("gents_12", "Gents 12h", 12)]
COLS = [t[0] for t in TYPES]
LABELS = [t[1] for t in TYPES]
HOURS = {t[0]: t[2] for t in TYPES}

T_ATT = "mp_attendance"
T_ALLOC = "mp_allocation"

st.set_page_config(page_title="Manpower Log", page_icon="👷", layout="centered")


def today_ist():
    return datetime.now(ZoneInfo("Asia/Kolkata")).date()


# ------------------------------------------------------------
# LOGIN
# ------------------------------------------------------------
def login():
    if st.session_state.get("logged_in"):
        return True
    st.title("👷 Manpower Log")
    name = st.text_input("Your name")
    pw = st.text_input("Enter password", type="password")
    if st.button("Login", type="primary", width="stretch"):
        if not name.strip():
            st.error("Enter your name")
        elif pw == st.secrets["APP_PASSWORD"]:
            st.session_state["logged_in"] = True
            st.session_state["user_name"] = name.strip()
            st.rerun()
        else:
            st.error("Wrong password")
    return False


if not login():
    st.stop()


# ------------------------------------------------------------
# DATABASE
# ------------------------------------------------------------
@st.cache_resource
def get_db():
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])


db = get_db()


def load_sheet(d, shift):
    att = (db.table(T_ATT).select("*")
           .eq("entry_date", d.isoformat()).eq("shift", shift)
           .execute().data)
    alloc = (db.table(T_ALLOC).select("*")
             .eq("entry_date", d.isoformat()).eq("shift", shift)
             .order("id").execute().data)
    return (att[0] if att else None), pd.DataFrame(alloc)


def fetch_range(table, start, end):
    rows, step, offset = [], 1000, 0
    while True:
        res = (db.table(table).select("*")
               .gte("entry_date", start.isoformat())
               .lte("entry_date", end.isoformat())
               .order("entry_date").order("id")
               .range(offset, offset + step - 1).execute())
        rows.extend(res.data)
        if len(res.data) < step:
            break
        offset += step
    df = pd.DataFrame(rows)
    if not df.empty:
        df["entry_date"] = pd.to_datetime(df["entry_date"]).dt.date
    return df


# ------------------------------------------------------------
# CALCULATIONS
# ------------------------------------------------------------
def to_num(df, cols):
    for c in cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df


def people(df):
    return df[COLS].sum(axis=1)


def mh_part1(df):
    """Full shift: each type x its own shift hours."""
    return sum(df[c] * HOURS[c] for c in COLS)


def mh_part2(df):
    """People x hours entered."""
    return people(df) * df["hours"]


# ------------------------------------------------------------
# EXCEL
# ------------------------------------------------------------
def make_excel(att, alloc):
    att = to_num(att.copy(), COLS)
    att["Total Present"] = people(att)
    att["Available Man-Hours"] = mh_part1(att)

    alloc = to_num(alloc.copy(), COLS + ["hours", "output_qty"])
    alloc["people"] = people(alloc)
    alloc["man_hours"] = 0.0
    p1 = alloc["part"] == 1
    alloc.loc[p1, "man_hours"] = mh_part1(alloc[p1])
    alloc.loc[~p1, "man_hours"] = mh_part2(alloc[~p1])

    by = alloc.pivot_table(index=["entry_date", "shift"], columns="part",
                           values="man_hours", aggfunc="sum", fill_value=0)
    by = by.rename(columns={1: "Part 1 Man-Hours", 2: "Part 2 Man-Hours"})
    for c in ["Part 1 Man-Hours", "Part 2 Man-Hours"]:
        if c not in by.columns:
            by[c] = 0.0
    daily = att.merge(by.reset_index(), on=["entry_date", "shift"], how="left")
    daily[["Part 1 Man-Hours", "Part 2 Man-Hours"]] = \
        daily[["Part 1 Man-Hours", "Part 2 Man-Hours"]].fillna(0)
    daily["Not Allocated Man-Hours"] = (daily["Available Man-Hours"]
                                        - daily["Part 1 Man-Hours"]
                                        - daily["Part 2 Man-Hours"])
    daily = daily.rename(columns={"entry_date": "Date", "shift": "Shift",
                                  "entered_by": "Saved By",
                                  **dict(zip(COLS, LABELS))})
    daily = daily[["Date", "Shift"] + LABELS +
                  ["Total Present", "Available Man-Hours", "Part 1 Man-Hours",
                   "Part 2 Man-Hours", "Not Allocated Man-Hours", "Saved By"]]

    base = {"entry_date": "Date", "shift": "Shift", "station": "Station",
            "people": "Total People", "man_hours": "Man-Hours",
            **dict(zip(COLS, LABELS))}
    part1 = alloc[p1].rename(columns=base)[
        ["Date", "Shift", "Station"] + LABELS + ["Total People", "Man-Hours"]]

    part2 = alloc[~p1].copy()
    part2["out_per_mh"] = (part2["output_qty"] /
                           part2["man_hours"].where(part2["man_hours"] > 0)).round(2)
    part2 = part2.rename(columns={**base, "hours": "Hours",
                                  "output_qty": "Output", "output_unit": "Unit",
                                  "out_per_mh": "Output per Man-Hour",
                                  "remarks": "Remarks"})[
        ["Date", "Shift", "Station"] + LABELS +
        ["Total People", "Hours", "Man-Hours", "Output", "Unit",
         "Output per Man-Hour", "Remarks"]]

    summ = alloc.copy()
    summ["Part"] = summ["part"].map({1: "Part 1", 2: "Part 2"})
    summ["output_unit"] = summ["output_unit"].fillna("")
    summ = summ.groupby(["Part", "station", "output_unit"]).agg(
        Days=("entry_date", "nunique"), Man_Hours=("man_hours", "sum"),
        Output=("output_qty", "sum")).reset_index()
    summ["Output per Man-Hour"] = (summ["Output"] /
                                   summ["Man_Hours"].where(summ["Man_Hours"] > 0)).round(2)
    summ.loc[summ["Part"] == "Part 1", ["Output", "Output per Man-Hour"]] = None
    summ.columns = ["Part", "Station", "Unit", "Days", "Total Man-Hours",
                    "Total Output", "Output per Man-Hour"]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", date_format="DD-MM-YYYY") as xw:
        daily.to_excel(xw, sheet_name="Daily Summary", index=False)
        part1.to_excel(xw, sheet_name="Part 1 Stations", index=False)
        part2.to_excel(xw, sheet_name="Part 2 Manual", index=False)
        summ.to_excel(xw, sheet_name="Station Summary", index=False)

        # one sheet per date (if 31 days or less)
        if daily["Date"].nunique() <= 31:
            for d in sorted(daily["Date"].unique()):
                name = d.strftime("%d-%m-%Y")
                row = 0
                for title, df in [("Attendance and balance", daily),
                                  ("Part 1 - Stations", part1),
                                  ("Part 2 - Manual work", part2)]:
                    part = df[df["Date"] == d]
                    if name in xw.book.sheetnames:
                        xw.book[name].cell(row=row + 1, column=1, value=title)
                    else:
                        pd.DataFrame().to_excel(xw, sheet_name=name)
                        xw.book[name].cell(row=row + 1, column=1, value=title)
                    part.to_excel(xw, sheet_name=name, index=False, startrow=row + 1)
                    row += len(part) + 4

        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0
                            for c in col) + 2
                ws.column_dimensions[col[0].column_letter].width = min(max(width, 10), 40)
            if not ws.title[0].isdigit():
                ws.freeze_panes = "A2"
    return buf.getvalue()


# ------------------------------------------------------------
# SCREEN
# ------------------------------------------------------------
st.title("👷 Manpower Log")
c_user, c_out = st.columns([3, 1])
c_user.caption(f"👤 Logged in as **{st.session_state['user_name']}**")
if c_out.button("Logout"):
    st.session_state.clear()
    st.rerun()

tab1, tab2 = st.tabs(["📝 Daily Sheet", "📥 Excel"])

# ================= TAB 1: DAILY SHEET =================
with tab1:
    msg = st.session_state.pop("save_msg", None)
    if msg:
        st.success(msg)

    c1, c2 = st.columns(2)
    sel_date = c1.date_input("Date", value=today_ist(), format="DD/MM/YYYY")
    sel_shift = c2.selectbox("Shift", SHIFTS)

    st.session_state.setdefault("ver", 0)
    cache_key = (sel_date.isoformat(), sel_shift, st.session_state["ver"])
    if st.session_state.get("cache_key") != cache_key:
        try:
            st.session_state["cache_data"] = load_sheet(sel_date, sel_shift)
        except Exception as e:
            st.error(f"Could not load data: {e}")
            st.stop()
        st.session_state["cache_key"] = cache_key
    att, alloc = st.session_state["cache_data"]
    k = "_".join(map(str, cache_key))  # widget key prefix

    if att:
        st.info(f"This sheet is already saved (last saved by {att.get('entered_by') or '-'}). "
                "You can change it and save again.")

    # ---------- STEP 1 : PRESENT ----------
    st.subheader("Step 1 · Total present")
    present = {}
    pc = st.columns(3)
    for i, (col, label, hrs) in enumerate(TYPES):
        present[col] = pc[i].number_input(
            f"{label}", min_value=0, step=1,
            value=int(att[col]) if att else 0, key=f"pr_{col}_{k}")
    avail_mh = sum(present[c] * HOURS[c] for c in COLS)
    st.caption(f"Total present: **{sum(present.values())}** people · "
               f"Available: **{avail_mh:g}** man-hours")

    # ---------- STEP 2 : PART 1 ----------
    st.subheader("Step 2 · Part 1 stations")
    st.caption("Full shift at the station. Enter number of people only.")
    p1 = pd.DataFrame({"Station": PART1_STATIONS})
    for c, lab in zip(COLS, LABELS):
        p1[lab] = 0
    if not alloc.empty and (alloc["part"] == 1).any():
        saved1 = alloc[alloc["part"] == 1]
        for _, r in saved1.iterrows():
            if r["station"] not in p1["Station"].values:
                p1.loc[len(p1), "Station"] = r["station"]
            idx = p1.index[p1["Station"] == r["station"]][0]
            for c, lab in zip(COLS, LABELS):
                p1.loc[idx, lab] = int(r[c] or 0)
        p1[LABELS] = p1[LABELS].fillna(0).astype(int)

    p1_edit = st.data_editor(
        p1, hide_index=True, width="stretch", num_rows="fixed",
        key=f"p1_{k}",
        column_config={
            "Station": st.column_config.TextColumn("Station", disabled=True),
            **{lab: st.column_config.NumberColumn(lab, min_value=0, step=1, format="%d")
               for lab in LABELS},
        })
    p1d = p1_edit.rename(columns=dict(zip(LABELS, COLS)))
    p1d = to_num(p1d, COLS)
    p1_mh = float(mh_part1(p1d).sum())
    p1_tot = {c: int(p1d[c].sum()) for c in COLS}
    st.caption("Part 1 people: " + " + ".join(str(p1_tot[c]) for c in COLS) +
               f" = **{sum(p1_tot.values())}** · **{p1_mh:g}** man-hours")

    # ---------- STEP 3 : PART 2 ----------
    st.subheader("Step 3 · Part 2 manual work")
    st.caption("People move between stations, so enter hours worked. "
               "Use the + row at the bottom to add the same station again "
               "with different hours.")
    if not alloc.empty and (alloc["part"] == 2).any():
        s2 = alloc[alloc["part"] == 2]
        p2 = pd.DataFrame({
            "Station": s2["station"],
            **{lab: s2[c].fillna(0).astype(int) for c, lab in zip(COLS, LABELS)},
            "Hours": pd.to_numeric(s2["hours"]).fillna(0).astype(float),
            "Output": pd.to_numeric(s2["output_qty"]).fillna(0).astype(float),
            "Unit": s2["output_unit"],
            "Remarks": s2["remarks"].fillna(""),
        }).reset_index(drop=True)
        missing = [s for s in PART2_STATIONS if s not in p2["Station"].values]
    else:
        p2 = pd.DataFrame(columns=["Station"] + LABELS +
                          ["Hours", "Output", "Unit", "Remarks"])
        missing = PART2_STATIONS
    if missing:
        extra = pd.DataFrame({
            "Station": missing, **{lab: 0 for lab in LABELS},
            "Hours": 0.0, "Output": 0.0,
            "Unit": [DEFAULT_UNIT.get(s, "Nos") for s in missing],
            "Remarks": "",
        })
        p2 = pd.concat([p2, extra], ignore_index=True) if not p2.empty else extra

    p2_edit = st.data_editor(
        p2, hide_index=True, width="stretch", num_rows="dynamic",
        key=f"p2_{k}",
        column_config={
            "Station": st.column_config.SelectboxColumn(
                "Station", options=PART2_STATIONS, required=True, pinned=True),
            **{lab: st.column_config.NumberColumn(lab, min_value=0, step=1, format="%d")
               for lab in LABELS},
            "Hours": st.column_config.NumberColumn("Hours", min_value=0, max_value=12, step=0.5),
            "Output": st.column_config.NumberColumn("Output", min_value=0),
            "Unit": st.column_config.SelectboxColumn("Unit", options=UNITS),
            "Remarks": st.column_config.TextColumn("Remarks"),
        })
    p2d = p2_edit.rename(columns={**dict(zip(LABELS, COLS)),
                                  "Hours": "hours", "Output": "output_qty"})
    p2d = p2d[p2d["Station"].notna()].copy()
    p2d = to_num(p2d, COLS + ["hours", "output_qty"])
    p2_mh = float(mh_part2(p2d).sum())
    st.caption(f"Part 2: **{p2_mh:g}** man-hours")

    # ---------- BALANCE ----------
    st.subheader("Balance check (man-hours)")
    not_alloc = avail_mh - p1_mh - p2_mh
    m = st.columns(4)
    m[0].metric("Available", f"{avail_mh:g}")
    m[1].metric("Part 1", f"{p1_mh:g}")
    m[2].metric("Part 2", f"{p2_mh:g}")
    m[3].metric("Not allocated", f"{not_alloc:g}")
    if abs(not_alloc) < 0.5:
        st.success("✅ Balanced")
    elif not_alloc > 0:
        st.warning(f"⚠️ {not_alloc:g} man-hours not allocated")
    else:
        st.error(f"❌ {-not_alloc:g} man-hours over — someone may be counted twice")
    over = [lab for c, lab in zip(COLS, LABELS) if p1_tot[c] > present[c]]
    if over:
        st.error("Part 1 has more people than present for: " + ", ".join(over))

    # ---------- SAVE ----------
    if st.button("💾 Save sheet", type="primary", width="stretch"):
        errors = []
        if sum(present.values()) == 0:
            errors.append("Step 1: enter the people present")
        for i, r in p2d.iterrows():
            ppl = sum(r[c] for c in COLS)
            if ppl > 0 and r["hours"] <= 0:
                errors.append(f"Part 2 · {r['Station']}: enter hours")
            if ppl == 0 and (r["hours"] > 0 or r["output_qty"] > 0):
                errors.append(f"Part 2 · {r['Station']}: enter people")
        if errors:
            for e in errors:
                st.error(e)
        else:
            user = st.session_state["user_name"]
            day, sh = sel_date.isoformat(), sel_shift
            rows = []
            for _, r in p1d.iterrows():
                if sum(r[c] for c in COLS) > 0:
                    rows.append({"entry_date": day, "shift": sh, "part": 1,
                                 "station": r["Station"],
                                 **{c: int(r[c]) for c in COLS},
                                 "entered_by": user})
            for _, r in p2d.iterrows():
                if sum(r[c] for c in COLS) > 0:
                    unit = r.get("Unit")
                    rem = r.get("Remarks")
                    rows.append({"entry_date": day, "shift": sh, "part": 2,
                                 "station": r["Station"],
                                 **{c: int(r[c]) for c in COLS},
                                 "hours": float(r["hours"]),
                                 "output_qty": float(r["output_qty"]),
                                 "output_unit": None if pd.isna(unit) else unit,
                                 "remarks": None if pd.isna(rem) or not str(rem).strip()
                                 else str(rem).strip(),
                                 "entered_by": user})
            try:
                db.table(T_ALLOC).delete().eq("entry_date", day).eq("shift", sh).execute()
                db.table(T_ATT).delete().eq("entry_date", day).eq("shift", sh).execute()
                db.table(T_ATT).insert({"entry_date": day, "shift": sh,
                                        **{c: int(present[c]) for c in COLS},
                                        "entered_by": user}).execute()
                if rows:
                    db.table(T_ALLOC).insert(rows).execute()
            except Exception as e:
                st.error(f"Could not save: {e}")
            else:
                st.session_state["save_msg"] = (
                    f"✅ Saved {sel_shift} shift for {sel_date.strftime('%d-%m-%Y')}")
                st.session_state["ver"] += 1
                st.rerun()

    if att:
        with st.expander("🗑️ Delete this sheet"):
            sure = st.checkbox("Yes, delete this date and shift completely",
                               key=f"del_{k}")
            if st.button("Delete", disabled=not sure, key=f"delbtn_{k}"):
                day = sel_date.isoformat()
                db.table(T_ALLOC).delete().eq("entry_date", day).eq("shift", sel_shift).execute()
                db.table(T_ATT).delete().eq("entry_date", day).eq("shift", sel_shift).execute()
                st.session_state["save_msg"] = "🗑️ Sheet deleted"
                st.session_state["ver"] += 1
                st.rerun()

# ================= TAB 2: EXCEL =================
with tab2:
    c1, c2 = st.columns(2)
    x_from = c1.date_input("From", today_ist().replace(day=1),
                           format="DD/MM/YYYY", key="x_from")
    x_to = c2.date_input("To", today_ist(), format="DD/MM/YYYY", key="x_to")

    if st.button("📊 Prepare Excel", type="primary", width="stretch"):
        a = fetch_range(T_ATT, x_from, x_to)
        b = fetch_range(T_ALLOC, x_from, x_to)
        if a.empty:
            st.warning("No sheets saved in this period.")
            st.session_state.pop("xlsx", None)
        else:
            if b.empty:
                b = pd.DataFrame(columns=["entry_date", "shift", "part", "station",
                                          *COLS, "hours", "output_qty",
                                          "output_unit", "remarks"])
            st.session_state["xlsx"] = make_excel(a, b)
            st.session_state["xlsx_name"] = (
                f"Manpower_{x_from.strftime('%d-%m-%Y')}_to_{x_to.strftime('%d-%m-%Y')}.xlsx")

    if "xlsx" in st.session_state:
        st.download_button(
            "⬇️ Download Excel", data=st.session_state["xlsx"],
            file_name=st.session_state["xlsx_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch")
        st.caption("Sheets: Daily Summary, Part 1 Stations, Part 2 Manual, "
                   "Station Summary, plus one sheet for each date.")
