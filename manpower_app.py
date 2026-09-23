import io
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from supabase import create_client

# ------------------------------------------------------------
# SETTINGS  (you can edit these lists any time)
# ------------------------------------------------------------
TABLE = "manpower_log"
TASKS = ["Combo", "Peanut Line", "Pouch Open", "Packing",
         "Export Stickering", "Line", "Other"]
UNITS = ["Pouch", "Case", "Kg", "Box", "Carton", "Nos"]

st.set_page_config(page_title="Manpower Log", page_icon="👷", layout="centered")


def today_ist():
    return datetime.now(ZoneInfo("Asia/Kolkata")).date()


# ------------------------------------------------------------
# PASSWORD
# ------------------------------------------------------------
def login():
    if st.session_state.get("logged_in"):
        return True
    st.title("👷 Manpower Log")
    name = st.text_input("Your name")
    pw = st.text_input("Enter password", type="password")
    if st.button("Login", type="primary", use_container_width=True):
        if not name.strip():
            st.error("Please enter your name")
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


def fetch(start, end):
    """Read all rows between two dates."""
    rows, step, offset = [], 1000, 0
    while True:
        res = (db.table(TABLE).select("*")
               .gte("entry_date", start.isoformat())
               .lte("entry_date", end.isoformat())
               .order("entry_date").order("id")
               .range(offset, offset + step - 1)
               .execute())
        rows.extend(res.data)
        if len(res.data) < step:
            break
        offset += step
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["entry_date"] = pd.to_datetime(df["entry_date"]).dt.date
    for c in ["manpower", "hours", "output_qty"]:
        df[c] = pd.to_numeric(df[c])
    return df


def add_calc(df):
    df = df.copy()
    df["man_hours"] = df["manpower"] * df["hours"]
    df["output_per_man_hour"] = (
        df["output_qty"] / df["man_hours"].where(df["man_hours"] > 0)
    ).round(2)
    return df


def clean(v):
    return None if pd.isna(v) else v


def same(a, b):
    if pd.isna(a) and pd.isna(b):
        return True
    return a == b


# ------------------------------------------------------------
# EXCEL
# ------------------------------------------------------------
def make_excel(df):
    df = add_calc(df)
    if "entered_by" not in df.columns:
        df["entered_by"] = None
    nice = df.rename(columns={
        "entry_date": "Date", "task": "Task", "manpower": "Manpower",
        "hours": "Hours", "man_hours": "Man-Hours", "output_qty": "Output Qty",
        "output_unit": "Unit", "output_per_man_hour": "Output per Man-Hour",
        "remarks": "Remarks", "entered_by": "Entered By",
    })[["Date", "Task", "Manpower", "Hours", "Man-Hours", "Output Qty",
        "Unit", "Output per Man-Hour", "Remarks", "Entered By"]]

    daily = df.groupby("entry_date").agg(
        Tasks=("task", "count"),
        Total_Manpower=("manpower", "sum"),
        Total_Man_Hours=("man_hours", "sum"),
    ).reset_index()
    daily.columns = ["Date", "Tasks", "Total Manpower", "Total Man-Hours"]

    task_sum = df.groupby(["task", "output_unit"], dropna=False).agg(
        Days=("entry_date", "nunique"),
        Total_Manpower=("manpower", "sum"),
        Man_Hours=("man_hours", "sum"),
        Output=("output_qty", "sum"),
    ).reset_index()
    task_sum["Output per Man-Hour"] = (
        task_sum["Output"] / task_sum["Man_Hours"].where(task_sum["Man_Hours"] > 0)
    ).round(2)
    task_sum.columns = ["Task", "Unit", "Days", "Total Manpower",
                        "Total Man-Hours", "Total Output", "Output per Man-Hour"]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", date_format="DD-MM-YYYY") as xw:
        daily.to_excel(xw, sheet_name="Date Summary", index=False)
        task_sum.to_excel(xw, sheet_name="Task Summary", index=False)
        nice.to_excel(xw, sheet_name="All Data", index=False)
        # one sheet per date (only if 60 days or less)
        if nice["Date"].nunique() <= 60:
            for d, g in nice.groupby("Date"):
                g.to_excel(xw, sheet_name=d.strftime("%d-%m-%Y"), index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0
                            for c in col) + 2
                ws.column_dimensions[col[0].column_letter].width = min(max(width, 10), 40)
            ws.freeze_panes = "A2"
    return buf.getvalue()


# ------------------------------------------------------------
# SCREENS
# ------------------------------------------------------------
st.title("👷 Manpower Log")
c_user, c_out = st.columns([3, 1])
c_user.caption(f"👤 Logged in as **{st.session_state['user_name']}**")
if c_out.button("Logout"):
    st.session_state.clear()
    st.rerun()
tab1, tab2, tab3 = st.tabs(["➕ New Entry", "✏️ View & Edit", "📥 Excel"])

# ---------------- TAB 1: NEW ENTRY ----------------
with tab1:
    st.session_state.setdefault("form_id", 0)
    fid = st.session_state["form_id"]
    msg = st.session_state.pop("entry_msg", None)
    if msg:
        st.success(msg)

    entry_date = st.date_input("Date", value=today_ist(),
                               format="DD/MM/YYYY", key=f"date_{fid}")
    n = st.number_input("How many tasks on this date?", min_value=1,
                        max_value=20, value=1, step=1, key=f"n_{fid}")

    records, errors = [], []
    for i in range(int(n)):
        with st.container(border=True):
            st.markdown(f"**Task {i + 1}**")
            task = st.selectbox("Task", TASKS, key=f"task_{fid}_{i}")
            if task == "Other":
                task = st.text_input("Write task name",
                                     key=f"other_{fid}_{i}").strip()
            c1, c2 = st.columns(2)
            mp = c1.number_input("Manpower", min_value=0, step=1,
                                 key=f"mp_{fid}_{i}")
            hrs = c2.number_input("Hours", min_value=0.0, step=0.5,
                                  key=f"hrs_{fid}_{i}")
            c3, c4 = st.columns(2)
            qty = c3.number_input("Output qty", min_value=0.0, step=1.0,
                                  key=f"qty_{fid}_{i}")
            unit = c4.selectbox("Unit", UNITS, key=f"unit_{fid}_{i}")
            remarks = st.text_input("Remarks (optional)", key=f"rem_{fid}_{i}")
            if mp and hrs:
                st.caption(f"Man-hours: {mp * hrs:g}")

            if not task:
                errors.append(f"Task {i + 1}: write the task name")
            if mp <= 0:
                errors.append(f"Task {i + 1}: manpower must be more than 0")
            if hrs <= 0:
                errors.append(f"Task {i + 1}: hours must be more than 0")

            records.append({
                "entry_date": entry_date.isoformat(),
                "task": task,
                "manpower": int(mp),
                "hours": float(hrs),
                "output_qty": float(qty),
                "output_unit": unit,
                "remarks": remarks.strip() or None,
                "entered_by": st.session_state["user_name"],
            })

    if st.button("💾 Save all", type="primary", use_container_width=True):
        if errors:
            for e in errors:
                st.error(e)
        else:
            try:
                db.table(TABLE).insert(records).execute()
            except Exception as e:
                st.error(f"Could not save: {e}")
            else:
                st.session_state["entry_msg"] = (
                    f"✅ Saved {len(records)} task(s) for "
                    f"{entry_date.strftime('%d-%m-%Y')}")
                st.session_state["form_id"] += 1
                st.rerun()

# ---------------- TAB 2: VIEW & EDIT ----------------
with tab2:
    st.session_state.setdefault("edit_ver", 0)
    msg = st.session_state.pop("edit_msg", None)
    if msg:
        st.success(msg)

    c1, c2 = st.columns(2)
    v_from = c1.date_input("From", today_ist() - timedelta(days=7),
                           format="DD/MM/YYYY", key="v_from")
    v_to = c2.date_input("To", today_ist(), format="DD/MM/YYYY", key="v_to")

    df = fetch(v_from, v_to)
    if df.empty:
        st.info("No entries in this period.")
    else:
        fields = ["entry_date", "task", "manpower", "hours",
                  "output_qty", "output_unit", "remarks"]
        if "entered_by" not in df.columns:
            df["entered_by"] = None
        view = df[["id"] + fields + ["entered_by"]].copy()
        view["delete"] = False
        st.caption("Tap a cell to change it. Tick 'Delete?' to remove a row. "
                   "Then press Save changes.")
        edited = st.data_editor(
            view, hide_index=True, use_container_width=True, num_rows="fixed",
            key=f"editor_{st.session_state['edit_ver']}",
            column_config={
                "id": st.column_config.NumberColumn("ID", disabled=True),
                "entry_date": st.column_config.DateColumn("Date", format="DD-MM-YYYY"),
                "task": st.column_config.TextColumn("Task"),
                "manpower": st.column_config.NumberColumn("Manpower", min_value=0, step=1),
                "hours": st.column_config.NumberColumn("Hours", min_value=0, step=0.5),
                "output_qty": st.column_config.NumberColumn("Output", min_value=0),
                "output_unit": st.column_config.SelectboxColumn("Unit", options=UNITS),
                "remarks": st.column_config.TextColumn("Remarks"),
                "entered_by": st.column_config.TextColumn("Entered By", disabled=True),
                "delete": st.column_config.CheckboxColumn("Delete?"),
            },
        )

        if st.button("💾 Save changes", type="primary", use_container_width=True):
            orig = view.set_index("id")
            to_delete = [int(x) for x in edited.loc[edited["delete"], "id"]]
            updates, problems = [], []
            for _, row in edited[~edited["delete"]].iterrows():
                o = orig.loc[row["id"]]
                if all(same(row[f], o[f]) for f in fields):
                    continue
                if pd.isna(row["entry_date"]) or not str(row["task"] or "").strip() \
                        or pd.isna(row["manpower"]) or row["manpower"] <= 0 \
                        or pd.isna(row["hours"]) or row["hours"] <= 0:
                    problems.append(int(row["id"]))
                    continue
                updates.append((int(row["id"]), {
                    "entry_date": row["entry_date"].isoformat(),
                    "task": str(row["task"]).strip(),
                    "manpower": int(row["manpower"]),
                    "hours": float(row["hours"]),
                    "output_qty": None if pd.isna(row["output_qty"]) else float(row["output_qty"]),
                    "output_unit": clean(row["output_unit"]),
                    "remarks": clean(row["remarks"]),
                }))
            if problems:
                st.error(f"Check rows with ID {problems}: date, task, manpower "
                         "and hours must be filled (more than 0). Nothing saved.")
            else:
                try:
                    for rid, data in updates:
                        db.table(TABLE).update(data).eq("id", rid).execute()
                    if to_delete:
                        db.table(TABLE).delete().in_("id", to_delete).execute()
                except Exception as e:
                    st.error(f"Could not save: {e}")
                else:
                    st.session_state["edit_msg"] = (
                        f"✅ Updated {len(updates)} row(s), deleted {len(to_delete)} row(s)")
                    st.session_state["edit_ver"] += 1
                    st.rerun()

# ---------------- TAB 3: EXCEL ----------------
with tab3:
    c1, c2 = st.columns(2)
    x_from = c1.date_input("From", today_ist().replace(day=1),
                           format="DD/MM/YYYY", key="x_from")
    x_to = c2.date_input("To", today_ist(), format="DD/MM/YYYY", key="x_to")

    if st.button("📊 Prepare Excel", type="primary", use_container_width=True):
        data = fetch(x_from, x_to)
        if data.empty:
            st.warning("No entries in this period.")
            st.session_state.pop("xlsx", None)
        else:
            st.session_state["xlsx"] = make_excel(data)
            st.session_state["xlsx_name"] = (
                f"Manpower_{x_from.strftime('%d-%m-%Y')}_to_{x_to.strftime('%d-%m-%Y')}.xlsx")

    if "xlsx" in st.session_state:
        st.download_button(
            "⬇️ Download Excel", data=st.session_state["xlsx"],
            file_name=st.session_state["xlsx_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption("Sheets: Date Summary, Task Summary, All Data, "
                   "plus one sheet for each date.")
