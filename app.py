
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(
    page_title="Project Controls Hub",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# Theme / UI
# ============================================================
st.markdown("""
<style>
.block-container {padding-top: 1.4rem; padding-bottom: 3rem;}
[data-testid="stMetric"] {
    border: 1px solid rgba(128,128,128,.22);
    padding: 14px 16px;
    border-radius: 12px;
    background: rgba(128,128,128,.035);
}
.small-note {font-size:.86rem; opacity:.72;}
.status-ok {color:#149954;font-weight:600;}
.status-warn {color:#c57c00;font-weight:600;}
.status-bad {color:#c73b3b;font-weight:600;}
</style>
""", unsafe_allow_html=True)

# ============================================================
# Helpers
# ============================================================
TECH_HEADER_TASK = "Activity ID"
TECH_HEADER_PRED = "Predecessor"

def fmt_date(x):
    if pd.isna(x):
        return "—"
    return pd.Timestamp(x).strftime("%d-%b-%Y")

def to_dt(s):
    return pd.to_datetime(s, errors="coerce")

def to_num(s):
    return pd.to_numeric(s, errors="coerce")

def read_sheet(xls, name):
    df = pd.read_excel(xls, sheet_name=name, header=0)
    return df

def load_schedule(uploaded):
    xls = pd.ExcelFile(uploaded)
    if "TASK" not in xls.sheet_names or "TASKPRED" not in xls.sheet_names:
        raise ValueError("The workbook must contain TASK and TASKPRED sheets.")

    task = read_sheet(xls, "TASK")
    pred = read_sheet(xls, "TASKPRED")

    if "task_code" not in task.columns:
        raise ValueError("TASK sheet does not contain the expected technical column names.")
    if "pred_task_id" not in pred.columns:
        raise ValueError("TASKPRED sheet does not contain the expected technical column names.")

    task = task[task["task_code"].astype(str) != TECH_HEADER_TASK].copy()
    pred = pred[pred["pred_task_id"].astype(str) != TECH_HEADER_PRED].copy()

    for c in ["task_code","task_name","status_code","wbs_id"]:
        if c in task.columns:
            task[c] = task[c].astype(str).str.strip()

    task["duration_days"] = to_num(task.get("target_drtn_hr_cnt"))
    task["total_float_days"] = to_num(task.get("total_float_hr_cnt"))
    task["start"] = to_dt(task.get("start_date"))
    task["finish"] = to_dt(task.get("end_date"))

    # Optional columns if later exports include them
    aliases = {
        "actual_start":["act_start_date","actual_start_date"],
        "actual_finish":["act_end_date","actual_finish_date"],
        "remain_duration":["remain_drtn_hr_cnt","remaining_duration"],
        "physical_pct":["phys_complete_pct","physical_complete_pct"],
        "pct_complete":["complete_pct","task_complete_pct"],
        "budget_cost":["target_cost","budgeted_cost","planned_cost"],
        "actual_cost":["act_cost","actual_cost"],
        "budget_units":["target_qty","budgeted_units","planned_units"],
        "actual_units":["act_qty","actual_units"],
        "remain_units":["remain_qty","remaining_units"],
    }
    for out, candidates in aliases.items():
        task[out] = np.nan
        for c in candidates:
            if c in task.columns:
                task[out] = task[c]
                break
    for c in ["actual_start","actual_finish"]:
        task[c] = to_dt(task[c])
    for c in ["remain_duration","physical_pct","pct_complete","budget_cost","actual_cost","budget_units","actual_units","remain_units"]:
        task[c] = to_num(task[c])

    pred["pred_task_id"] = pred["pred_task_id"].astype(str).str.strip()
    pred["task_id"] = pred["task_id"].astype(str).str.strip()
    pred["pred_type"] = pred["pred_type"].astype(str).str.strip()
    pred["lag_days"] = to_num(pred.get("lag_hr_cnt"))

    # Detect optional assignment / resource sheets
    optional = {}
    for s in xls.sheet_names:
        su = s.upper()
        if su in {"TASKRSRC","RESOURCEASSIGNMENT","RESOURCE ASSIGNMENT","RESASSIGN","ASSIGNMENTS"}:
            optional["assignments"] = read_sheet(xls, s)
        elif su in {"RSRC","RESOURCES","RESOURCE"}:
            optional["resources"] = read_sheet(xls, s)

    return task, pred, optional, xls.sheet_names

def category_from_row(row):
    code = str(row.get("task_code","")).upper()
    name = str(row.get("task_name","")).lower()
    wbs  = str(row.get("wbs_id","")).lower()

    if code.startswith("PQ-") or "prequalification" in name:
        return "Prequalification"
    if code.startswith("SD-") or "shop drawing" in name:
        return "Shop Drawings"
    if code.startswith("MS-") or "material approval" in name or "material submittal" in name:
        return "Material Submittals"
    if code.startswith(("PR-","PO-","MF-","DL-")) or any(k in name for k in ["procurement","purchase order","manufacturing","fabrication","delivery on site"]):
        return "Procurement"
    if code.startswith("AUTH-") or ("noc" in name) or ("authority" in name):
        return "Authorities"
    if code.startswith("TC-") or any(k in name for k in ["testing","commissioning"]):
        return "T&C"
    if code.startswith("HO-") or any(k in name for k in ["handover","taking over","snag","o&m","as-built","close-out","closeout"]):
        return "Handover"
    if any(k in name for k in ["excavation","raft","pile","slab","column","wall","blockwork","plaster","screed","ceiling","tile","paint","façade","facade","external works"]):
        return "Construction"
    if any(k in wbs for k in ["construction","substructure","superstructure","finishes","external"]):
        return "Construction"
    return "Other"

def level_from_name(name):
    n = f" {str(name).lower()} "
    tests = [
        ("Basement 02", [" b2 ", "basement 02", "basement 2"]),
        ("Basement 01", [" b1 ", "basement 01", "basement 1"]),
        ("Ground Floor", ["ground floor", " gf "]),
        ("Level 01", ["level 01","level 1","floor 01","floor 1"]),
        ("Level 02", ["level 02","level 2","floor 02","floor 2"]),
        ("Roof", ["roof"]),
        ("Façade", ["facade","façade"]),
        ("External Works", ["external"]),
    ]
    for label, keys in tests:
        if any(k in n for k in keys):
            return label
    return "Other"

def enrich(task, pred):
    t = task.copy()
    t["category"] = t.apply(category_from_row, axis=1)
    t["level"] = t["task_name"].apply(level_from_name)

    pred_counts = pred.groupby("task_id").size()
    succ_counts = pred.groupby("pred_task_id").size()
    t["has_pred"] = t["task_code"].isin(pred_counts.index)
    t["has_succ"] = t["task_code"].isin(succ_counts.index)
    return t

def project_facts(task):
    return {
        "start": task["start"].min(),
        "finish": task["finish"].max(),
        "activities": len(task),
        "critical": int((task["total_float_days"] <= 0).sum()),
        "negative": int((task["total_float_days"] < 0).sum()),
        "completed": int(task["status_code"].str.contains("Complete", case=False, na=False).sum()),
        "in_progress": int(task["status_code"].str.contains("Progress|Active", case=False, regex=True, na=False).sum()),
        "not_started": int(task["status_code"].str.contains("Not Started", case=False, na=False).sum()),
    }

def key_milestones(task):
    ms = task[task["duration_days"].fillna(-1) == 0].copy()
    keys = [
        "site possession","construction start","substructure completion",
        "superstructure completion","mep first fix","mep final fix",
        "testing and commissioning completion","authority final approval",
        "taking over certificate","mobilization completed","excavation completed",
        "raft foundation completed","basement 02 structural works completed",
        "basement 01 structural works completed","ground floor structural works completed",
        "floor 01 structural works completed","floor 02 structural works completed",
        "roof structural works completed","façade works completed","facade works completed",
        "external works completed","handover completed"
    ]
    return ms[ms["task_name"].str.lower().apply(lambda x: any(k in x for k in keys))].copy()

def compare_schedules(base, curr):
    b = base.set_index("task_code")
    c = curr.set_index("task_code")
    common = b.index.intersection(c.index)
    rows = []
    for code in common:
        br, cr = b.loc[code], c.loc[code]
        fd = np.nan
        if pd.notna(br["finish"]) and pd.notna(cr["finish"]):
            fd = (pd.Timestamp(cr["finish"]).normalize() - pd.Timestamp(br["finish"]).normalize()).days
        rows.append({
            "Activity ID": code,
            "Activity Name": cr["task_name"],
            "Baseline Finish": br["finish"],
            "Current Finish": cr["finish"],
            "Finish Movement (d)": fd,
            "Baseline Float": br["total_float_days"],
            "Current Float": cr["total_float_days"],
            "Float Movement": (cr["total_float_days"] - br["total_float_days"]) if pd.notna(br["total_float_days"]) and pd.notna(cr["total_float_days"]) else np.nan,
            "Baseline Status": br["status_code"],
            "Current Status": cr["status_code"],
        })
    return pd.DataFrame(rows)

def planned_curve_from_activity_spread(task, value_col=None, freq="W"):
    """
    Time-phase values uniformly across activity calendar span.
    Used only when a valid activity-level budget column exists.
    """
    if value_col is None or value_col not in task.columns:
        return pd.DataFrame()
    data = task.dropna(subset=["start","finish",value_col]).copy()
    data = data[data[value_col].fillna(0) != 0]
    if data.empty:
        return pd.DataFrame()

    chunks = []
    for _, r in data.iterrows():
        dates = pd.date_range(r["start"].normalize(), r["finish"].normalize(), freq="D")
        if len(dates) == 0:
            continue
        daily = float(r[value_col]) / len(dates)
        chunks.append(pd.DataFrame({"date":dates,"value":daily}))
    if not chunks:
        return pd.DataFrame()
    out = pd.concat(chunks).groupby("date", as_index=False)["value"].sum()
    out = out.set_index("date").resample(freq)["value"].sum().reset_index()
    out["cumulative"] = out["value"].cumsum()
    total = out["value"].sum()
    out["cumulative_pct"] = np.where(total != 0, out["cumulative"]/total*100, 0)
    return out

def assignment_curve(assignments, basis="cost", freq="W"):
    """
    Flexible optional parser for common Primavera-like assignment columns.
    If dates/values aren't present, returns empty data.
    """
    if assignments is None or assignments.empty:
        return pd.DataFrame()

    cols = {str(c).lower():c for c in assignments.columns}

    def pick(options):
        for o in options:
            if o in cols: return cols[o]
        return None

    s_col = pick(["start_date","start","task_start_date"])
    f_col = pick(["end_date","finish","finish_date","task_finish_date"])
    if basis == "cost":
        v_col = pick(["target_cost","budgeted_cost","planned_cost","budget_cost","remaining_cost"])
    else:
        v_col = pick(["target_qty","budgeted_units","planned_units","budget_units","remaining_units","target_labor_units"])

    if not all([s_col,f_col,v_col]):
        return pd.DataFrame()

    df = assignments[[s_col,f_col,v_col]].copy()
    df.columns = ["start","finish","value"]
    df["start"] = to_dt(df["start"])
    df["finish"] = to_dt(df["finish"])
    df["value"] = to_num(df["value"])
    return planned_curve_from_activity_spread(df.rename(columns={"value":"budget"}), "budget", freq)

def lookahead(task, data_date, weeks=4):
    if pd.isna(data_date):
        return pd.DataFrame()
    end = pd.Timestamp(data_date) + pd.Timedelta(weeks=weeks)
    t = task[
        ((task["start"] >= data_date) & (task["start"] <= end)) |
        ((task["finish"] >= data_date) & (task["finish"] <= end)) |
        ((task["start"] <= data_date) & (task["finish"] >= data_date))
    ].copy()
    return t.sort_values(["start","finish"])

# ============================================================
# Sidebar Inputs
# ============================================================
st.sidebar.title("Project Controls Hub")
st.sidebar.caption("v0.2 · Baseline + Update Intelligence")

baseline_file = st.sidebar.file_uploader("1. Approved Baseline", type=["xlsx"], key="baseline")
update_file = st.sidebar.file_uploader("2. Current Update (optional)", type=["xlsx"], key="update")
supp_file = st.sidebar.file_uploader(
    "3. Resource / Cost export (optional)",
    type=["xlsx"],
    key="supp",
    help="Use this when cost or manpower/resource assignments are exported separately."
)

near_tf = st.sidebar.number_input("Near-critical threshold (days)", 1, 90, 20, 1)
look_weeks = st.sidebar.selectbox("Lookahead period", [2,4,6,8], index=1)

st.sidebar.markdown("---")
st.sidebar.markdown("**Calculation principle**")
st.sidebar.caption("Code calculates · AI interprets")

if baseline_file is None:
    st.title("Project Controls Hub")
    st.info("Upload the Approved Baseline Primavera Excel export to start.")
    st.stop()

try:
    base_raw, base_pred, base_opt, base_sheets = load_schedule(baseline_file)
    base = enrich(base_raw, base_pred)
except Exception as e:
    st.error(f"Baseline could not be read: {e}")
    st.stop()

curr = None
curr_pred = None
curr_opt = {}
if update_file is not None:
    try:
        curr_raw, curr_pred, curr_opt, curr_sheets = load_schedule(update_file)
        curr = enrich(curr_raw, curr_pred)
    except Exception as e:
        st.error(f"Current update could not be read: {e}")

# Optional supplemental workbook
supp_assign = None
supp_resources = None
supp_sheets = []
if supp_file is not None:
    try:
        sx = pd.ExcelFile(supp_file)
        supp_sheets = sx.sheet_names
        for s in sx.sheet_names:
            su = s.upper()
            if su in {"TASKRSRC","RESOURCEASSIGNMENT","RESOURCE ASSIGNMENT","RESASSIGN","ASSIGNMENTS"}:
                supp_assign = pd.read_excel(sx, sheet_name=s)
            if su in {"RSRC","RESOURCES","RESOURCE"}:
                supp_resources = pd.read_excel(sx, sheet_name=s)
    except Exception as e:
        st.warning(f"Supplemental workbook could not be interpreted: {e}")

assignments = supp_assign if supp_assign is not None else base_opt.get("assignments")

bf = project_facts(base)
cf = project_facts(curr) if curr is not None else None

# ============================================================
# Top Header
# ============================================================
st.title("JRD · Project Controls Hub")
st.caption("Professional project-controls dashboard generated from Primavera exports")

c = st.columns(6)
c[0].metric("Activities", f"{bf['activities']:,}")
c[1].metric("Relationships", f"{len(base_pred):,}")
c[2].metric("Critical", f"{bf['critical']:,}")
c[3].metric(f"Near Critical ≤{near_tf}d", f"{int(((base.total_float_days>0)&(base.total_float_days<=near_tf)).sum()):,}")
c[4].metric("Negative Float", f"{bf['negative']:,}")
c[5].metric("Baseline Finish", fmt_date(bf["finish"]))

# ============================================================
# Data Readiness
# ============================================================
with st.expander("Data readiness & source validation", expanded=False):
    checks = []
    checks.append(("Schedule Activities", True, "TASK sheet detected"))
    checks.append(("Relationships", True, "TASKPRED sheet detected"))
    checks.append(("Current Update", curr is not None, "Upload a current update to enable variance/comparison"))
    has_cost_task = base["budget_cost"].notna().any()
    has_units_task = base["budget_units"].notna().any()
    has_assign = assignments is not None
    checks.append(("Cost Loading", has_cost_task or has_assign, "Needed for Cost S-Curve / Cash Flow"))
    checks.append(("Resource / MHR Loading", has_units_task or has_assign, "Needed for MHR S-Curve / Manpower"))
    for name, ok, note in checks:
        icon = "✅" if ok else "⚠️"
        st.write(f"{icon} **{name}** — {note}")
    st.caption("The application does not fabricate missing project-control values.")

tabs = st.tabs([
    "Executive",
    "Baseline vs Update",
    "Milestones",
    "Schedule Health",
    "Progress Curves",
    "Cash Flow",
    "Manpower",
    "Engineering",
    "Procurement",
    "Construction",
    "Lookahead",
    "T&C / Handover",
    "Report Center",
])

# ============================================================
# Executive
# ============================================================
with tabs[0]:
    st.subheader("Executive Overview")
    row = st.columns(4)
    row[0].metric("Baseline Start", fmt_date(bf["start"]))
    row[1].metric("Baseline Finish", fmt_date(bf["finish"]))
    row[2].metric("Current Forecast", fmt_date(cf["finish"]) if cf else "Upload Update")
    delta_finish = None
    if cf and pd.notna(bf["finish"]) and pd.notna(cf["finish"]):
        delta_finish = (pd.Timestamp(cf["finish"]).normalize()-pd.Timestamp(bf["finish"]).normalize()).days
    row[3].metric("Forecast Movement", f"{delta_finish:+d} d" if delta_finish is not None else "—")

    c1,c2 = st.columns(2)
    with c1:
        status = base["status_code"].value_counts().reset_index()
        status.columns=["Status","Activities"]
        fig=px.bar(status,x="Status",y="Activities",text="Activities",title="Baseline Activity Status")
        fig.update_layout(height=360,xaxis_title="",yaxis_title="Activities")
        st.plotly_chart(fig,use_container_width=True)
    with c2:
        comp=base["category"].value_counts().reset_index()
        comp.columns=["Module","Activities"]
        fig=px.pie(comp,names="Module",values="Activities",hole=.55,title="Programme Composition")
        fig.update_layout(height=360)
        st.plotly_chart(fig,use_container_width=True)

# ============================================================
# Comparison
# ============================================================
with tabs[1]:
    if curr is None:
        st.info("Upload the Current Update to activate Baseline vs Update comparison.")
    else:
        comp = compare_schedules(base, curr)
        moved = comp[comp["Finish Movement (d)"].fillna(0) != 0]
        newcrit = comp[(comp["Baseline Float"]>0) & (comp["Current Float"]<=0)]
        lostcrit = comp[(comp["Baseline Float"]<=0) & (comp["Current Float"]>0)]
        negnew = comp[(comp["Baseline Float"]>=0) & (comp["Current Float"]<0)]

        m=st.columns(5)
        m[0].metric("Forecast Finish", fmt_date(cf["finish"]), f"{delta_finish:+d} d" if delta_finish is not None else None)
        m[1].metric("Moved Activities", f"{len(moved):,}")
        m[2].metric("New Critical", f"{len(newcrit):,}")
        m[3].metric("Lost Critical", f"{len(lostcrit):,}")
        m[4].metric("New Negative Float", f"{len(negnew):,}")

        st.subheader("Largest Finish Movements")
        st.dataframe(
            comp.sort_values("Finish Movement (d)", key=lambda x:x.abs(), ascending=False).head(50),
            use_container_width=True, hide_index=True
        )

        c1,c2=st.columns(2)
        with c1:
            st.subheader("New Critical Activities")
            st.dataframe(newcrit.head(100), use_container_width=True, hide_index=True)
        with c2:
            st.subheader("Lost Critical Activities")
            st.dataframe(lostcrit.head(100), use_container_width=True, hide_index=True)

# ============================================================
# Milestones
# ============================================================
with tabs[2]:
    bms = key_milestones(base)
    if curr is None:
        out=bms[["task_code","task_name","finish","total_float_days"]].rename(columns={
            "task_code":"Activity ID","task_name":"Milestone","finish":"Baseline Date","total_float_days":"Total Float"
        })
    else:
        cms = key_milestones(curr).set_index("task_code")
        rows=[]
        for _,r in bms.iterrows():
            code=r["task_code"]
            cr=cms.loc[code] if code in cms.index else None
            current_finish=cr["finish"] if cr is not None else pd.NaT
            move=(pd.Timestamp(current_finish).normalize()-pd.Timestamp(r["finish"]).normalize()).days if pd.notna(current_finish) and pd.notna(r["finish"]) else np.nan
            rows.append({
                "Activity ID":code,"Milestone":r["task_name"],
                "Baseline Date":r["finish"],"Current Date":current_finish,
                "Movement (d)":move,
                "Current Float": cr["total_float_days"] if cr is not None else np.nan
            })
        out=pd.DataFrame(rows)
    st.dataframe(out,use_container_width=True,hide_index=True)

# ============================================================
# Schedule Health
# ============================================================
with tabs[3]:
    pred_counts=base_pred.groupby("task_id").size()
    succ_counts=base_pred.groupby("pred_task_id").size()
    non_ms=base[base["duration_days"].fillna(0)>0]
    open_start=non_ms[~non_ms["task_code"].isin(pred_counts.index)]
    open_finish=non_ms[~non_ms["task_code"].isin(succ_counts.index)]
    low=base.sort_values("total_float_days").head(50)

    c=st.columns(4)
    c[0].metric("No Predecessor", len(open_start))
    c[1].metric("No Successor", len(open_finish))
    c[2].metric("TF ≤ 0", int((base.total_float_days<=0).sum()))
    c[3].metric(f"0 < TF ≤ {near_tf}", int(((base.total_float_days>0)&(base.total_float_days<=near_tf)).sum()))

    st.subheader("Lowest Float Activities")
    st.dataframe(low[["task_code","task_name","start","finish","total_float_days","category"]],use_container_width=True,hide_index=True)

# ============================================================
# Progress Curves
# ============================================================
with tabs[4]:
    st.subheader("S-Curves")
    st.caption("Cost and MHR are kept as separate progress bases.")

    freq_label=st.radio("Period",["Weekly","Monthly"],horizontal=True,key="curvefreq")
    freq="W" if freq_label=="Weekly" else "MS"

    c1,c2=st.columns(2)
    with c1:
        st.markdown("### Cost S-Curve")
        curve = planned_curve_from_activity_spread(base,"budget_cost",freq) if base["budget_cost"].notna().any() else assignment_curve(assignments,"cost",freq)
        if curve.empty:
            st.warning("Cost-loading data is not present in the current JRD TASK/TASKPRED workbook. Upload Resource/Cost Assignments to activate this chart.")
        else:
            fig=go.Figure()
            fig.add_trace(go.Scatter(x=curve["date"],y=curve["cumulative_pct"],mode="lines",name="Baseline Planned %"))
            fig.update_layout(height=420,yaxis_title="Cumulative %",xaxis_title="")
            st.plotly_chart(fig,use_container_width=True)

    with c2:
        st.markdown("### MHR / Unit S-Curve")
        curve = planned_curve_from_activity_spread(base,"budget_units",freq) if base["budget_units"].notna().any() else assignment_curve(assignments,"units",freq)
        if curve.empty:
            st.warning("Labor-unit / manpower loading is not present in the current workbook. Upload Resource Assignments to activate this chart.")
        else:
            fig=go.Figure()
            fig.add_trace(go.Scatter(x=curve["date"],y=curve["cumulative_pct"],mode="lines",name="Baseline Planned %"))
            fig.update_layout(height=420,yaxis_title="Cumulative %",xaxis_title="")
            st.plotly_chart(fig,use_container_width=True)

# ============================================================
# Cash Flow
# ============================================================
with tabs[5]:
    st.subheader("Cash Flow")
    curve = planned_curve_from_activity_spread(base,"budget_cost","MS") if base["budget_cost"].notna().any() else assignment_curve(assignments,"cost","MS")
    if curve.empty:
        st.warning("Cash Flow requires Cost Loading / Resource Assignments from Primavera.")
        st.markdown("""
**Expected output after cost data is uploaded:**
- Monthly Planned Cost
- Monthly Actual Cost
- Current Forecast
- Cumulative Planned
- Cumulative Actual
- Cumulative Forecast
- Monthly and cumulative cash-flow chart
""")
    else:
        fig=go.Figure()
        fig.add_bar(x=curve["date"],y=curve["value"],name="Monthly Planned")
        fig.add_trace(go.Scatter(x=curve["date"],y=curve["cumulative"],mode="lines+markers",name="Cumulative Planned",yaxis="y2"))
        fig.update_layout(height=500,yaxis_title="Monthly Cost",yaxis2=dict(title="Cumulative Cost",overlaying="y",side="right"),xaxis_title="")
        st.plotly_chart(fig,use_container_width=True)
        st.dataframe(curve,use_container_width=True,hide_index=True)

# ============================================================
# Manpower
# ============================================================
with tabs[6]:
    st.subheader("Manpower & Labor Loading")
    curve = planned_curve_from_activity_spread(base,"budget_units","W") if base["budget_units"].notna().any() else assignment_curve(assignments,"units","W")
    if curve.empty:
        st.warning("Resource Assignment data is required for planned manpower / labor-unit charts.")
        st.markdown("""
Once Resource Assignments are uploaded this module will show:
- Weekly planned labor units
- Monthly planned labor units
- Actual labor units
- Staff / Civil / MEP / Finishes / Subcontractor breakdown where resource codes permit
- Peak manpower
- Planned vs actual histogram
""")
    else:
        fig=px.bar(curve,x="date",y="value",title="Weekly Planned Labor Units")
        fig.update_layout(height=470,xaxis_title="",yaxis_title="Labor Units")
        st.plotly_chart(fig,use_container_width=True)

# ============================================================
# Engineering
# ============================================================
with tabs[7]:
    eng=base[base["category"].isin(["Prequalification","Shop Drawings","Material Submittals","Authorities"])].copy()
    cards=st.columns(4)
    for i,cat in enumerate(["Prequalification","Shop Drawings","Material Submittals","Authorities"]):
        cards[i].metric(cat,int((eng.category==cat).sum()))

    st.subheader("Engineering At-Risk Register")
    risk=eng[eng["total_float_days"]<=near_tf].sort_values("total_float_days")
    st.dataframe(risk[["task_code","task_name","category","start","finish","total_float_days"]],use_container_width=True,hide_index=True)

# ============================================================
# Procurement
# ============================================================
with tabs[8]:
    proc=base[base["category"]=="Procurement"].copy()
    st.metric("Procurement Activities",len(proc))
    c1,c2=st.columns([1,1])
    with c1:
        st.subheader("Lowest Float Procurement")
        st.dataframe(proc.sort_values("total_float_days")[["task_code","task_name","start","finish","total_float_days"]].head(60),use_container_width=True,hide_index=True)
    with c2:
        st.subheader("Long-Lead / Delivery Timeline")
        dl=proc[proc["task_code"].str.startswith(("PO-","MF-","DL-"),na=False)].sort_values("start").head(60)
        if not dl.empty:
            fig=px.timeline(dl,x_start="start",x_end="finish",y="task_name",hover_data=["task_code","total_float_days"])
            fig.update_yaxes(autorange="reversed",title="")
            fig.update_layout(height=900,xaxis_title="")
            st.plotly_chart(fig,use_container_width=True)

# ============================================================
# Construction
# ============================================================
with tabs[9]:
    levels=(
        base[base["level"]!="Other"]
        .groupby("level")
        .agg(Start=("start","min"),Finish=("finish","max"),Activities=("task_code","count"),MinFloat=("total_float_days","min"))
        .reset_index()
    )
    st.subheader("Floor / Location Control")
    st.dataframe(levels,use_container_width=True,hide_index=True)
    if not levels.empty:
        fig=px.timeline(levels,x_start="Start",x_end="Finish",y="level",hover_data=["Activities","MinFloat"])
        fig.update_yaxes(autorange="reversed",title="")
        fig.update_layout(height=440,xaxis_title="")
        st.plotly_chart(fig,use_container_width=True)

# ============================================================
# Lookahead
# ============================================================
with tabs[10]:
    data_date = curr["start"].min() if curr is not None else base["start"].min()
    # Prefer latest actual information later; for baseline-only, user-selected project start is shown transparently.
    st.subheader(f"{look_weeks}-Week Lookahead")
    st.caption(f"Reference date currently used: {fmt_date(data_date)}")
    la = lookahead(curr if curr is not None else base, data_date, look_weeks)
    st.dataframe(la[["task_code","task_name","category","start","finish","total_float_days","status_code"]],use_container_width=True,hide_index=True)

# ============================================================
# T&C / Handover
# ============================================================
with tabs[11]:
    tch=base[base["category"].isin(["T&C","Handover"])].copy()
    c=st.columns(2)
    c[0].metric("T&C",int((tch.category=="T&C").sum()))
    c[1].metric("Handover / Closeout",int((tch.category=="Handover").sum()))
    st.dataframe(tch.sort_values("start")[["task_code","task_name","category","start","finish","total_float_days"]],use_container_width=True,hide_index=True)

# ============================================================
# Report Center
# ============================================================
with tabs[12]:
    st.subheader("Report Center")
    st.markdown("""
### Planned report outputs
**Weekly Progress Report**
- Executive Summary
- Overall Progress
- Cost S-Curve
- MHR / Unit S-Curve
- Cash Flow
- Manpower Histogram
- Milestones
- Schedule Analysis
- Engineering
- Procurement
- Construction
- Lookahead
- T&C / Handover
- Risks & Actions
- Progress Photos

**Management Dashboard**
- 1–3 page concise executive version

**Update Comparison Report**
- Baseline vs Current Update
- Milestone movement
- Critical-path movement
- New negative float
- Engineering / Procurement slippage

The report generator will use company branding and will not mix cost-weighted and MHR-weighted progress.
""")
    if curr is not None:
        comp = compare_schedules(base,curr)
        csv = comp.to_csv(index=False).encode("utf-8-sig")
        st.download_button("Download Update Comparison CSV",csv,"JRD_Update_Comparison.csv","text/csv")

st.markdown("---")
st.caption("JRD Project Controls Hub · v0.2 · Deterministic calculations with transparent data readiness")
