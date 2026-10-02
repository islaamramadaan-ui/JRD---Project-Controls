
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="Project Controls Hub",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# Helpers
# -----------------------------
def excel_serial_to_datetime(value):
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.Timestamp(value)
    try:
        value = float(value)
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(value, unit="D")
    except Exception:
        return pd.to_datetime(value, errors="coerce")

def clean_number(series):
    return pd.to_numeric(series, errors="coerce")

def load_p6_excel(file):
    xls = pd.ExcelFile(file)
    required = {"TASK", "TASKPRED"}
    missing = required - set(xls.sheet_names)
    if missing:
        raise ValueError(f"Missing required sheets: {', '.join(sorted(missing))}")

    task = pd.read_excel(xls, sheet_name="TASK", header=0)
    pred = pd.read_excel(xls, sheet_name="TASKPRED", header=0)

    # Primavera exports often include a second human-readable header row.
    if "task_code" in task.columns:
        task = task[task["task_code"].astype(str) != "Activity ID"].copy()
    if "pred_task_id" in pred.columns:
        pred = pred[pred["pred_task_id"].astype(str) != "Predecessor"].copy()

    # Normalize task columns
    task["task_code"] = task["task_code"].astype(str).str.strip()
    task["task_name"] = task["task_name"].astype(str).str.strip()
    task["status_code"] = task["status_code"].astype(str).str.strip()
    task["wbs_id"] = task["wbs_id"].astype(str).str.strip()

    task["duration_days"] = clean_number(task["target_drtn_hr_cnt"])
    task["total_float_days"] = clean_number(task["total_float_hr_cnt"])
    task["start"] = task["start_date"].apply(excel_serial_to_datetime)
    task["finish"] = task["end_date"].apply(excel_serial_to_datetime)

    # Relationships
    pred["pred_task_id"] = pred["pred_task_id"].astype(str).str.strip()
    pred["task_id"] = pred["task_id"].astype(str).str.strip()
    pred["pred_type"] = pred["pred_type"].astype(str).str.strip()
    pred["lag_days"] = clean_number(pred["lag_hr_cnt"])

    return task, pred

def category_from_task(row):
    code = str(row["task_code"]).upper()
    name = str(row["task_name"]).lower()
    wbs = str(row["wbs_id"]).lower()

    if code.startswith("PQ-") or "prequalification" in name:
        return "Prequalification"
    if code.startswith("SD-") or "shop drawing" in name:
        return "Shop Drawings"
    if code.startswith("MS-") or "material approval" in name or "material submittal" in name:
        return "Material Submittals"
    if code.startswith(("PR-", "PO-", "MF-", "DL-")) or any(x in name for x in ["procurement", "purchase order", "manufacturing", "fabrication", "delivery on site"]):
        return "Procurement"
    if code.startswith("AUTH-") or any(x in name for x in ["approval", "noc", "authority"]) and "material approval" not in name:
        return "Authorities"
    if code.startswith("TC-") or any(x in name for x in ["testing", "commissioning"]):
        return "T&C"
    if code.startswith("HO-") or any(x in name for x in ["handover", "taking over", "snag", "o&m", "as-built"]):
        return "Handover"
    if any(x in wbs for x in ["construction", "substructure", "superstructure", "finishes", "external"]):
        return "Construction"
    return "Other"

def level_from_name(name):
    n = str(name).lower()
    checks = [
        ("Basement 02", ["b2", "basement 02", "basement 2"]),
        ("Basement 01", ["b1", "basement 01", "basement 1"]),
        ("Ground Floor", ["ground floor", " gf ", "- gf", "gf -"]),
        ("Level 01", ["level 01", "level 1", "floor 01", "floor 1"]),
        ("Level 02", ["level 02", "level 2", "floor 02", "floor 2"]),
        ("Roof", ["roof"]),
        ("External Works", ["external"]),
        ("Façade", ["facade", "façade"]),
    ]
    padded = f" {n} "
    for label, keys in checks:
        if any(k in padded for k in keys):
            return label
    return "Other"

def render_metric_card(label, value, sub=None):
    st.metric(label, value, sub if sub else None)

# -----------------------------
# Sidebar
# -----------------------------
st.sidebar.title("Project Controls Hub")
st.sidebar.caption("MVP v0.1 · Primavera Excel Intelligence")

uploaded = st.sidebar.file_uploader(
    "Upload Primavera Excel export",
    type=["xlsx"],
    help="Expected sheets: TASK and TASKPRED",
)

near_critical_threshold = st.sidebar.number_input(
    "Near-critical float threshold (days)",
    min_value=1,
    max_value=90,
    value=20,
    step=1,
)

st.sidebar.markdown("---")
st.sidebar.caption("Core rule")
st.sidebar.markdown("**Code calculates. AI interprets.**")

if uploaded is None:
    st.title("Project Controls Hub")
    st.info("Upload the Primavera Excel export to generate the dashboard.")
    st.stop()

try:
    task, pred = load_p6_excel(uploaded)
except Exception as e:
    st.error(f"Could not read the file: {e}")
    st.stop()

task["category"] = task.apply(category_from_task, axis=1)
task["level"] = task["task_name"].apply(level_from_name)

# -----------------------------
# Baseline facts
# -----------------------------
project_start = task["start"].min()
project_finish = task["finish"].max()
duration_calendar = (project_finish.normalize() - project_start.normalize()).days if pd.notna(project_start) and pd.notna(project_finish) else None

activity_count = len(task)
relationship_count = len(pred)
completed = int(task["status_code"].str.contains("Complete", case=False, na=False).sum())
in_progress = int(task["status_code"].str.contains("Progress|Active", case=False, regex=True, na=False).sum())
not_started = int(task["status_code"].str.contains("Not Started", case=False, na=False).sum())

critical = task[task["total_float_days"] <= 0]
near_critical = task[(task["total_float_days"] > 0) & (task["total_float_days"] <= near_critical_threshold)]
negative_float = task[task["total_float_days"] < 0]

pred_counts = pred.groupby("task_id").size()
succ_counts = pred.groupby("pred_task_id").size()
task["has_pred"] = task["task_code"].isin(pred_counts.index)
task["has_succ"] = task["task_code"].isin(succ_counts.index)

# Ignore milestones when flagging open ends
non_milestone = task[task["duration_days"].fillna(0) > 0]
open_start = non_milestone[~non_milestone["has_pred"]]
open_finish = non_milestone[~non_milestone["has_succ"]]

milestones = task[task["duration_days"].fillna(-1) == 0].copy().sort_values("finish")

# -----------------------------
# Header
# -----------------------------
st.title("JRD · Project Controls Dashboard")
st.caption("Baseline intelligence generated directly from Primavera export")

cols = st.columns(5)
with cols[0]:
    render_metric_card("Activities", f"{activity_count:,}")
with cols[1]:
    render_metric_card("Relationships", f"{relationship_count:,}")
with cols[2]:
    render_metric_card("Critical", f"{len(critical):,}")
with cols[3]:
    render_metric_card("Near Critical", f"{len(near_critical):,}")
with cols[4]:
    render_metric_card("Negative Float", f"{len(negative_float):,}")

cols = st.columns(4)
with cols[0]:
    render_metric_card("Baseline Start", project_start.strftime("%d-%b-%Y") if pd.notna(project_start) else "—")
with cols[1]:
    render_metric_card("Baseline Finish", project_finish.strftime("%d-%b-%Y") if pd.notna(project_finish) else "—")
with cols[2]:
    render_metric_card("Calendar Span", f"{duration_calendar:,} days" if duration_calendar is not None else "—")
with cols[3]:
    render_metric_card("Open Ends", f"{len(open_start)+len(open_finish):,}", f"{len(open_start)} no pred · {len(open_finish)} no succ")

tabs = st.tabs([
    "Executive",
    "Milestones",
    "Schedule Health",
    "Engineering",
    "Procurement",
    "Construction",
    "T&C / Handover",
    "Activities",
])

# -----------------------------
# Executive
# -----------------------------
with tabs[0]:
    c1, c2 = st.columns([1.05, 1])

    with c1:
        st.subheader("Activity Status")
        status_counts = task["status_code"].value_counts(dropna=False).reset_index()
        status_counts.columns = ["Status", "Activities"]
        fig = px.bar(status_counts, x="Status", y="Activities", text="Activities")
        fig.update_layout(height=360, xaxis_title="", yaxis_title="Activities")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.subheader("Programme Composition")
        comp = task["category"].value_counts().reset_index()
        comp.columns = ["Module", "Activities"]
        fig = px.pie(comp, names="Module", values="Activities", hole=0.55)
        fig.update_layout(height=360)
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Baseline Timeline")
    module_dates = (
        task.groupby("category")
        .agg(Start=("start", "min"), Finish=("finish", "max"), Activities=("task_code", "count"))
        .reset_index()
    )
    module_dates = module_dates.dropna(subset=["Start", "Finish"])
    fig = px.timeline(module_dates, x_start="Start", x_end="Finish", y="category", hover_data=["Activities"])
    fig.update_yaxes(autorange="reversed", title="")
    fig.update_layout(height=420, xaxis_title="")
    st.plotly_chart(fig, use_container_width=True)

# -----------------------------
# Milestones
# -----------------------------
with tabs[1]:
    st.subheader("Milestone Register")
    key_words = [
        "site possession", "construction start", "substructure completion",
        "superstructure completion", "mep first fix", "mep final fix",
        "testing and commissioning completion", "authority final approval",
        "taking over certificate", "raft foundation completed",
        "basement", "ground floor structural works completed",
        "floor 01 structural works completed", "floor 02 structural works completed",
        "roof structural works completed", "façade works completed",
        "facade works completed", "external works completed", "handover completed"
    ]
    key_ms = milestones[
        milestones["task_name"].str.lower().apply(lambda x: any(k in x for k in key_words))
    ].copy()
    display = key_ms[["task_code", "task_name", "finish", "total_float_days"]].rename(columns={
        "task_code":"Activity ID", "task_name":"Milestone",
        "finish":"Baseline Date", "total_float_days":"Total Float"
    })
    st.dataframe(display, use_container_width=True, hide_index=True)

    if not key_ms.empty:
        fig = px.scatter(
            key_ms,
            x="finish",
            y="task_name",
            size=np.where(key_ms["total_float_days"].fillna(0).abs()+1 > 0,
                          key_ms["total_float_days"].fillna(0).abs()+1, 1),
            hover_data=["task_code", "total_float_days"],
        )
        fig.update_layout(height=max(420, 28*len(key_ms)), xaxis_title="", yaxis_title="")
        st.plotly_chart(fig, use_container_width=True)

# -----------------------------
# Schedule Health
# -----------------------------
with tabs[2]:
    health_cols = st.columns(4)
    health_cols[0].metric("No Predecessor", f"{len(open_start):,}")
    health_cols[1].metric("No Successor", f"{len(open_finish):,}")
    health_cols[2].metric("TF ≤ 0", f"{len(critical):,}")
    health_cols[3].metric(f"0 < TF ≤ {near_critical_threshold}", f"{len(near_critical):,}")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Lowest Float Activities")
        lowest = task.sort_values("total_float_days").head(25)
        st.dataframe(
            lowest[["task_code","task_name","start","finish","total_float_days","category"]],
            use_container_width=True,
            hide_index=True
        )
    with c2:
        st.subheader("Relationship Types")
        rel_types = pred["pred_type"].value_counts().reset_index()
        rel_types.columns = ["Type", "Count"]
        fig = px.bar(rel_types, x="Type", y="Count", text="Count")
        fig.update_layout(height=380, xaxis_title="", yaxis_title="Relationships")
        st.plotly_chart(fig, use_container_width=True)

    with st.expander("Open-start activities"):
        st.dataframe(open_start[["task_code","task_name","start","finish","total_float_days"]], use_container_width=True, hide_index=True)
    with st.expander("Open-finish activities"):
        st.dataframe(open_finish[["task_code","task_name","start","finish","total_float_days"]], use_container_width=True, hide_index=True)

# -----------------------------
# Engineering
# -----------------------------
with tabs[3]:
    eng = task[task["category"].isin(["Prequalification","Shop Drawings","Material Submittals","Authorities"])].copy()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Prequalifications", f"{(eng['category']=='Prequalification').sum():,}")
    c2.metric("Shop Drawings", f"{(eng['category']=='Shop Drawings').sum():,}")
    c3.metric("Material Submittals", f"{(eng['category']=='Material Submittals').sum():,}")
    c4.metric("Authorities", f"{(eng['category']=='Authorities').sum():,}")

    bycat = (
        eng.groupby("category")
        .agg(Start=("start","min"), Finish=("finish","max"), Activities=("task_code","count"),
             MinFloat=("total_float_days","min"))
        .reset_index()
    )
    st.dataframe(bycat, use_container_width=True, hide_index=True)

    st.subheader("Engineering Activities at Risk")
    risk = eng[eng["total_float_days"] <= near_critical_threshold].sort_values("total_float_days")
    st.dataframe(
        risk[["task_code","task_name","category","start","finish","total_float_days"]],
        use_container_width=True,
        hide_index=True
    )

# -----------------------------
# Procurement
# -----------------------------
with tabs[4]:
    proc = task[task["category"] == "Procurement"].copy()
    st.metric("Procurement Activities", f"{len(proc):,}")
    if not proc.empty:
        p1, p2 = st.columns([1.1,1])
        with p1:
            st.subheader("Procurement Timeline")
            sample = proc.sort_values("start").head(80)
            fig = px.timeline(sample, x_start="start", x_end="finish", y="task_name", hover_data=["task_code","total_float_days"])
            fig.update_yaxes(autorange="reversed", title="")
            fig.update_layout(height=850, xaxis_title="")
            st.plotly_chart(fig, use_container_width=True)
        with p2:
            st.subheader("Lowest Procurement Float")
            st.dataframe(
                proc.sort_values("total_float_days")[["task_code","task_name","start","finish","total_float_days"]].head(50),
                use_container_width=True,
                hide_index=True
            )

# -----------------------------
# Construction
# -----------------------------
with tabs[5]:
    cons = task[task["category"] == "Construction"].copy()
    st.metric("Detected Construction Activities", f"{len(cons):,}")
    levels = (
        task[task["level"] != "Other"]
        .groupby("level")
        .agg(Start=("start","min"), Finish=("finish","max"), Activities=("task_code","count"),
             MinFloat=("total_float_days","min"))
        .reset_index()
    )
    st.subheader("Floor / Location Overview")
    st.dataframe(levels, use_container_width=True, hide_index=True)
    if not levels.empty:
        fig = px.timeline(levels, x_start="Start", x_end="Finish", y="level", hover_data=["Activities","MinFloat"])
        fig.update_yaxes(autorange="reversed", title="")
        fig.update_layout(height=420, xaxis_title="")
        st.plotly_chart(fig, use_container_width=True)

# -----------------------------
# T&C / Handover
# -----------------------------
with tabs[6]:
    tch = task[task["category"].isin(["T&C","Handover"])].copy()
    tc = tch[tch["category"]=="T&C"]
    ho = tch[tch["category"]=="Handover"]
    c1,c2 = st.columns(2)
    c1.metric("T&C Activities", f"{len(tc):,}")
    c2.metric("Handover / Closeout", f"{len(ho):,}")
    st.dataframe(
        tch.sort_values("start")[["task_code","task_name","category","start","finish","total_float_days"]],
        use_container_width=True,
        hide_index=True
    )

# -----------------------------
# All Activities
# -----------------------------
with tabs[7]:
    st.subheader("Activity Browser")
    cats = ["All"] + sorted(task["category"].dropna().unique().tolist())
    selected_cat = st.selectbox("Module", cats)
    search = st.text_input("Search Activity ID / Name")
    view = task.copy()
    if selected_cat != "All":
        view = view[view["category"] == selected_cat]
    if search:
        s = search.lower()
        view = view[
            view["task_code"].str.lower().str.contains(s, na=False) |
            view["task_name"].str.lower().str.contains(s, na=False)
        ]
    st.dataframe(
        view[["task_code","task_name","status_code","category","start","finish","duration_days","total_float_days","wbs_id"]],
        use_container_width=True,
        hide_index=True,
        height=650
    )

st.markdown("---")
st.caption("MVP v0.1 · Baseline intelligence only. Next release: baseline vs update comparison, progress curves, manpower, and automated weekly report.")
