
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, date
from pathlib import Path
from io import BytesIO
import base64, math, os

import plotly.express as px
import plotly.graph_objects as go

# PDF/PPT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

st.set_page_config(
    page_title="JRD Project Controls Hub",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------------------------------------------------
# Brand
# ------------------------------------------------------------------
BRAND_GREEN = "#07795C"
BRAND_SILVER = "#D1D3D4"
BRAND_BLACK = "#121212"
BRAND_YELLOW = "#FFC400"
BRAND_OFFWHITE = "#F7F7F5"
BRAND_RED = "#D62828"

ROOT = Path(__file__).parent if "__file__" in globals() else Path(".")
ASSETS = ROOT / "assets"
ALTA_LOGO = ASSETS / "alta_logo.png"
DDDC_LOGO = ASSETS / "dddc_logo.png"
DAR_LOGO = ASSETS / "dar_logo.png"
COVER_IMG = ASSETS / "jrd_cover.png"

st.markdown(f"""
<style>
html, body, [class*="css"] {{ font-family: Poppins, Arial, sans-serif; }}
.block-container {{padding-top: 1.15rem; padding-bottom: 3rem;}}
[data-testid="stSidebar"] {{background: {BRAND_BLACK};}}
[data-testid="stSidebar"] * {{color: {BRAND_OFFWHITE};}}
div[data-testid="stMetric"] {{
  border: 1px solid rgba(209,211,212,.35);
  border-left: 5px solid {BRAND_GREEN};
  padding: 12px 14px; border-radius: 10px;
  background: rgba(247,247,245,.04);
}}
.kpi-title {{font-size:.74rem; opacity:.72; text-transform:uppercase; letter-spacing:.04em;}}
.kpi-big {{font-size:1.65rem; font-weight:700;}}
.section-title {{
  margin-top: 10px; padding: 8px 12px; border-radius: 6px;
  background:{BRAND_GREEN}; color:white; font-weight:700;
}}
.neg {{color:{BRAND_RED}; font-weight:700;}}
.pos {{color:{BRAND_GREEN}; font-weight:700;}}
.warn {{color:{BRAND_YELLOW}; font-weight:700;}}
.cover-card {{
  position:relative; border-radius:16px; overflow:hidden; margin-bottom:14px;
  border:1px solid rgba(209,211,212,.25);
}}
.cover-overlay {{
  position:absolute; inset:0; background:linear-gradient(90deg,rgba(18,18,18,.88),rgba(18,18,18,.22));
  padding:38px; display:flex; flex-direction:column; justify-content:flex-end;
}}
.cover-title {{font-size:2.4rem; font-weight:800; color:white; line-height:1.0;}}
.cover-sub {{color:{BRAND_SILVER}; font-size:1rem; margin-top:8px;}}
.small {{font-size:.82rem; opacity:.72;}}
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
def img_b64(path):
    if not path.exists():
        return ""
    return base64.b64encode(path.read_bytes()).decode()

def fmt_date(x):
    if pd.isna(x): return "—"
    return pd.Timestamp(x).strftime("%d-%b-%Y")

def to_dt(s): return pd.to_datetime(s, errors="coerce")
def to_num(s): return pd.to_numeric(s, errors="coerce")

def load_schedule(uploaded):
    xls = pd.ExcelFile(uploaded)
    if "TASK" not in xls.sheet_names or "TASKPRED" not in xls.sheet_names:
        raise ValueError("Workbook must contain TASK and TASKPRED sheets.")
    task = pd.read_excel(xls, sheet_name="TASK")
    pred = pd.read_excel(xls, sheet_name="TASKPRED")
    if "task_code" not in task.columns or "pred_task_id" not in pred.columns:
        raise ValueError("Expected Primavera technical column names were not found.")

    task = task[task["task_code"].astype(str).str.strip() != "Activity ID"].copy()
    pred = pred[pred["pred_task_id"].astype(str).str.strip() != "Predecessor"].copy()

    for c in ["task_code","task_name","status_code","wbs_id"]:
        if c in task.columns:
            task[c] = task[c].astype(str).str.strip()

    task["duration_days"] = to_num(task.get("target_drtn_hr_cnt"))
    task["total_float_days"] = to_num(task.get("total_float_hr_cnt"))
    task["start"] = to_dt(task.get("start_date"))
    task["finish"] = to_dt(task.get("end_date"))

    aliases = {
        "actual_start":["act_start_date","actual_start_date"],
        "actual_finish":["act_end_date","actual_finish_date"],
        "remain_duration":["remain_drtn_hr_cnt","remaining_duration"],
        "physical_pct":["phys_complete_pct","physical_complete_pct"],
        "pct_complete":["complete_pct","task_complete_pct"],
        "budget_cost":["target_cost","budgeted_cost","planned_cost","budget_cost"],
        "actual_cost":["act_cost","actual_cost"],
        "budget_units":["target_qty","budgeted_units","planned_units","budget_units","target_labor_units"],
        "actual_units":["act_qty","actual_units","actual_labor_units"],
        "remain_units":["remain_qty","remaining_units","remain_labor_units"],
    }
    for out, candidates in aliases.items():
        task[out] = np.nan
        for c in candidates:
            if c in task.columns:
                task[out] = task[c]; break
    for c in ["actual_start","actual_finish"]:
        task[c] = to_dt(task[c])
    for c in ["remain_duration","physical_pct","pct_complete","budget_cost","actual_cost","budget_units","actual_units","remain_units"]:
        task[c] = to_num(task[c])

    pred["pred_task_id"] = pred["pred_task_id"].astype(str).str.strip()
    pred["task_id"] = pred["task_id"].astype(str).str.strip()
    pred["pred_type"] = pred["pred_type"].astype(str).str.strip()
    pred["lag_days"] = to_num(pred.get("lag_hr_cnt"))

    optional = {}
    for s in xls.sheet_names:
        su = s.upper().replace("_"," ").strip()
        if su in {"TASKRSRC","RESOURCEASSIGNMENT","RESOURCE ASSIGNMENT","RESASSIGN","ASSIGNMENTS"}:
            optional["assignments"] = pd.read_excel(xls, sheet_name=s)
        elif su in {"RSRC","RESOURCES","RESOURCE"}:
            optional["resources"] = pd.read_excel(xls, sheet_name=s)
    return task, pred, optional

def category_from_row(row):
    code = str(row.get("task_code","")).upper()
    name = str(row.get("task_name","")).lower()
    wbs = str(row.get("wbs_id","")).lower()
    if code.startswith("PQ-") or "prequalification" in name: return "Prequalification"
    if code.startswith("SD-") or "shop drawing" in name: return "Shop Drawings"
    if code.startswith("MS-") or "material submittal" in name or "material approval" in name: return "Material Submittals"
    if code.startswith(("PR-","PO-","MF-","DL-")) or any(k in name for k in ["procurement","purchase order","lpo","manufacturing","fabrication","delivery on site"]): return "Procurement"
    if code.startswith("AUTH-") or "authority" in name or " noc" in f" {name}": return "Authorities"
    if code.startswith("TC-") or any(k in name for k in ["testing","commissioning"]): return "T&C"
    if code.startswith("HO-") or any(k in name for k in ["handover","taking over","snag","o&m","as-built","close-out","closeout"]): return "Handover"
    if any(k in name for k in ["excavation","raft","pile","slab","column","wall","blockwork","plaster","screed","ceiling","tile","paint","façade","facade","external works"]): return "Construction"
    if any(k in wbs for k in ["construction","substructure","superstructure","finishes","external"]): return "Construction"
    return "Other"

def package_name(row):
    name = str(row.get("task_name",""))
    # remove common prefixes / action words to group major packages
    cleaned = name
    for token in ["Prequalification","Shop Drawing","Shop Drawings","Material Submittal","Material Approval",
                  "Submit","Submission","Approval","Procurement","Purchase Order","LPO","Manufacturing","Fabrication",
                  "Delivery on Site","Delivery"]:
        cleaned = cleaned.replace(token,"")
    cleaned = cleaned.replace(" - "," ").replace(":"," ").strip(" -")
    return cleaned[:90] if cleaned else name[:90]

def enrich(task, pred):
    t = task.copy()
    t["category"] = t.apply(category_from_row, axis=1)
    t["package"] = t.apply(package_name, axis=1)
    pc = pred.groupby("task_id").size()
    sc = pred.groupby("pred_task_id").size()
    t["has_pred"] = t["task_code"].isin(pc.index)
    t["has_succ"] = t["task_code"].isin(sc.index)
    return t

def project_facts(task):
    return {
        "start": task["start"].min(),
        "finish": task["finish"].max(),
        "activities": len(task),
        "critical": int((task["total_float_days"] <= 0).sum()),
        "negative": int((task["total_float_days"] < 0).sum()),
    }

def planned_curve(task, value_col, freq="W"):
    data = task.dropna(subset=["start","finish",value_col]).copy()
    data = data[data[value_col].fillna(0) != 0]
    if data.empty: return pd.DataFrame()
    parts=[]
    for _,r in data.iterrows():
        dates=pd.date_range(r["start"].normalize(),r["finish"].normalize(),freq="D")
        if len(dates)==0: continue
        daily=float(r[value_col])/len(dates)
        parts.append(pd.DataFrame({"date":dates,"planned":daily}))
    if not parts: return pd.DataFrame()
    out=pd.concat(parts).groupby("date",as_index=False)["planned"].sum()
    out=out.set_index("date").resample(freq)["planned"].sum().reset_index()
    out["cum_planned"]=out["planned"].cumsum()
    total=out["planned"].sum()
    out["cum_planned_pct"]=np.where(total, out["cum_planned"]/total*100,0)
    return out

def assignment_to_activity(assignments):
    if assignments is None or assignments.empty:
        return pd.DataFrame()
    cols={str(c).lower():c for c in assignments.columns}
    def pick(options):
        for o in options:
            if o in cols: return cols[o]
        return None
    act=pick(["task_code","activity id","activity_id","task_id"])
    s=pick(["start_date","start","task_start_date"])
    f=pick(["end_date","finish","finish_date","task_finish_date"])
    pc=pick(["target_cost","budgeted_cost","planned_cost","budget_cost"])
    ac=pick(["act_cost","actual_cost"])
    pu=pick(["target_qty","budgeted_units","planned_units","budget_units","target_labor_units"])
    au=pick(["act_qty","actual_units","actual_labor_units"])
    rn=pick(["resource_name","resource name","rsrc_name"])
    rid=pick(["resource_id","resource id","rsrc_id"])
    out=pd.DataFrame()
    if act: out["activity_id"]=assignments[act].astype(str).str.strip()
    if s: out["start"]=to_dt(assignments[s])
    if f: out["finish"]=to_dt(assignments[f])
    if pc: out["planned_cost"]=to_num(assignments[pc])
    if ac: out["actual_cost"]=to_num(assignments[ac])
    if pu: out["planned_units"]=to_num(assignments[pu])
    if au: out["actual_units"]=to_num(assignments[au])
    if rn: out["resource_name"]=assignments[rn].astype(str)
    if rid: out["resource_id"]=assignments[rid].astype(str)
    return out

def timephase_assign(assign_df, val_col, data_date, freq="W"):
    if assign_df is None or assign_df.empty or val_col not in assign_df.columns:
        return pd.DataFrame()
    d=assign_df.dropna(subset=["start","finish",val_col]).copy()
    d=d[d[val_col].fillna(0)!=0]
    if d.empty:return pd.DataFrame()
    rows=[]
    for _,r in d.iterrows():
        dates=pd.date_range(r["start"].normalize(),r["finish"].normalize(),freq="D")
        if len(dates)==0:continue
        daily=float(r[val_col])/len(dates)
        tmp=pd.DataFrame({"date":dates,val_col:daily})
        rows.append(tmp)
    out=pd.concat(rows).groupby("date",as_index=False)[val_col].sum()
    out=out.set_index("date").resample(freq)[val_col].sum().reset_index()
    return out

def merge_plan_actual(plan, actual, freq, planned_name="planned", actual_name="actual"):
    if plan.empty and actual.empty: return pd.DataFrame()
    out = pd.merge(plan, actual, on="date", how="outer").sort_values("date").fillna(0)
    out["cum_planned"] = out.get(planned_name,0).cumsum()
    out["cum_actual"] = out.get(actual_name,0).cumsum()
    pt=out.get(planned_name,0).sum(); at=out.get(actual_name,0).sum()
    out["cum_planned_pct"]=np.where(pt, out["cum_planned"]/pt*100,0)
    out["cum_actual_pct"]=np.where(pt, out["cum_actual"]/pt*100,0)
    return out

def style_float(v):
    if pd.isna(v): return ""
    return f"color: {BRAND_RED}; font-weight:700;" if v < 0 else ""

def date_window(task, dd, weeks):
    end=dd+pd.Timedelta(weeks=weeks)
    return task[
        ((task["start"]>=dd)&(task["start"]<=end)) |
        ((task["finish"]>=dd)&(task["finish"]<=end)) |
        ((task["start"]<=dd)&(task["finish"]>=dd))
    ].sort_values(["start","finish"])

def status_from_float(tf):
    if pd.isna(tf): return "Unknown"
    if tf < 0: return "DELAY / NEGATIVE FLOAT"
    if tf <= 7: return "Critical"
    if tf <= 20: return "High Risk"
    return "Normal"

def procurement_action(row, dd):
    tf=row.get("total_float_days",np.nan)
    name=str(row.get("task_name","")).lower()
    finish=row.get("finish",pd.NaT)
    status=str(row.get("status_code","")).lower()
    completed=("complete" in status and "not" not in status)
    if completed: return "Completed"
    if pd.notna(tf) and tf<0:
        if "lpo" in name or "purchase order" in name or str(row.get("task_code","")).upper().startswith("PO-"):
            return "IMMEDIATE LPO REQUIRED"
        return "CRITICAL / DELAY"
    if pd.notna(finish) and finish < dd:
        return "OVERDUE"
    if "delivery" in name and pd.notna(finish) and finish <= dd+pd.Timedelta(days=14):
        return "DELIVERY DUE"
    if "lpo" in name or "purchase order" in name or str(row.get("task_code","")).upper().startswith("PO-"):
        if pd.notna(finish) and finish <= dd+pd.Timedelta(days=14):
            return "LPO DUE"
    return "Monitor"

def fig_to_png(fig):
    try:
        return fig.to_image(format="png", scale=2)
    except Exception:
        return None

# ------------------------------------------------------------------
# Sidebar inputs
# ------------------------------------------------------------------
st.sidebar.title("Project Controls Hub")
st.sidebar.caption("v0.3 FINAL · JRD Executive Controls")

baseline_file = st.sidebar.file_uploader("1. Approved Baseline", type=["xlsx"], key="baseline")
update_file = st.sidebar.file_uploader("2. Current Update", type=["xlsx"], key="update")
resource_file = st.sidebar.file_uploader("3. Resource / Cost Export", type=["xlsx"], key="resource")

today_default = date.today()
data_date_input = st.sidebar.date_input("Data Date", value=today_default, format="DD/MM/YYYY")
data_date = pd.Timestamp(data_date_input)

lookahead_weeks = st.sidebar.selectbox("Lookahead Horizon", [2,4,6,8], index=1)
hours_per_day = st.sidebar.number_input("Working Hours / Day", 1.0, 24.0, 8.0, .5)
days_per_week = st.sidebar.number_input("Working Days / Week", 1.0, 7.0, 6.0, .5)
near_tf = st.sidebar.number_input("Near-Critical Threshold (days)", 1, 90, 20, 1)

st.sidebar.markdown("---")
st.sidebar.caption("Code calculates · AI interprets")

# Progress photos
progress_photos = st.sidebar.file_uploader(
    "4. Progress Photos",
    type=["png","jpg","jpeg"],
    accept_multiple_files=True
)

# ------------------------------------------------------------------
# Cover
# ------------------------------------------------------------------
if COVER_IMG.exists():
    cb64=img_b64(COVER_IMG)
    st.markdown(f"""
    <div class="cover-card" style="height:360px;background:url('data:image/png;base64,{cb64}') center/cover;">
      <div class="cover-overlay">
        <div class="cover-title">JUMEIRAH RETAIL DEVELOPMENT</div>
        <div class="cover-sub">Owner / Employer: ALTA &nbsp;&nbsp; | &nbsp;&nbsp; Contractor: DDDC &nbsp;&nbsp; | &nbsp;&nbsp; Consultant: DAR</div>
      </div>
    </div>
    """, unsafe_allow_html=True)
else:
    st.title("JUMEIRAH RETAIL DEVELOPMENT")
    st.caption("Owner / Employer: ALTA | Contractor: DDDC | Consultant: DAR")

if baseline_file is None:
    st.info("Upload the Approved Baseline Primavera Excel export to start.")
    st.stop()

try:
    base_raw, base_pred, base_opt = load_schedule(baseline_file)
    base = enrich(base_raw, base_pred)
except Exception as e:
    st.error(f"Baseline could not be read: {e}")
    st.stop()

curr=None; curr_pred=None; curr_opt={}
if update_file:
    try:
        cr, cp, co = load_schedule(update_file)
        curr=enrich(cr,cp); curr_pred=cp; curr_opt=co
    except Exception as e:
        st.error(f"Current update could not be read: {e}")

# Resource/cost
assignments=None
if resource_file:
    try:
        x=pd.ExcelFile(resource_file)
        for s in x.sheet_names:
            su=s.upper().replace("_"," ").strip()
            if su in {"TASKRSRC","RESOURCEASSIGNMENT","RESOURCE ASSIGNMENT","RESASSIGN","ASSIGNMENTS"}:
                assignments=pd.read_excel(x,sheet_name=s); break
        if assignments is None:
            # best effort: first sheet
            assignments=pd.read_excel(x,sheet_name=x.sheet_names[0])
    except Exception as e:
        st.warning(f"Resource/Cost workbook could not be read: {e}")
if assignments is None:
    assignments = curr_opt.get("assignments") if curr_opt.get("assignments") is not None else base_opt.get("assignments")
assign = assignment_to_activity(assignments)

bf=project_facts(base)
cf=project_facts(curr) if curr is not None else None
forecast_finish = cf["finish"] if cf else bf["finish"]
forecast_move = (pd.Timestamp(forecast_finish).normalize()-pd.Timestamp(bf["finish"]).normalize()).days if pd.notna(forecast_finish) and pd.notna(bf["finish"]) else np.nan
duration_days = (bf["finish"].normalize()-bf["start"].normalize()).days if pd.notna(bf["start"]) and pd.notna(bf["finish"]) else np.nan
elapsed_days = max(0,(data_date.normalize()-bf["start"].normalize()).days) if pd.notna(bf["start"]) else np.nan
remaining_days = max(0,(forecast_finish.normalize()-data_date.normalize()).days) if pd.notna(forecast_finish) else np.nan
elapsed_pct = elapsed_days/duration_days*100 if duration_days and duration_days>0 else np.nan

# Build resource curves
cost_curve=pd.DataFrame(); unit_curve=pd.DataFrame()
if not assign.empty:
    pc=timephase_assign(assign,"planned_cost",data_date,"MS") if "planned_cost" in assign.columns else pd.DataFrame()
    ac=timephase_assign(assign,"actual_cost",data_date,"MS") if "actual_cost" in assign.columns else pd.DataFrame()
    if not pc.empty: pc=pc.rename(columns={"planned_cost":"planned"})
    if not ac.empty: ac=ac.rename(columns={"actual_cost":"actual"})
    cost_curve=merge_plan_actual(pc,ac,"MS","planned","actual") if (not pc.empty or not ac.empty) else pd.DataFrame()

    pu=timephase_assign(assign,"planned_units",data_date,"W") if "planned_units" in assign.columns else pd.DataFrame()
    au=timephase_assign(assign,"actual_units",data_date,"W") if "actual_units" in assign.columns else pd.DataFrame()
    if not pu.empty: pu=pu.rename(columns={"planned_units":"planned"})
    if not au.empty: au=au.rename(columns={"actual_units":"actual"})
    unit_curve=merge_plan_actual(pu,au,"W","planned","actual") if (not pu.empty or not au.empty) else pd.DataFrame()

# Executive progress KPIs
def progress_from_curve(curve):
    if curve.empty: return (np.nan,np.nan,np.nan)
    c=curve[curve["date"]<=data_date]
    if c.empty: return (0.0,0.0,0.0)
    p=float(c["cum_planned_pct"].iloc[-1]) if "cum_planned_pct" in c else np.nan
    a=float(c["cum_actual_pct"].iloc[-1]) if "cum_actual_pct" in c else np.nan
    return p,a,a-p if pd.notna(p) and pd.notna(a) else np.nan
cost_plan,cost_actual,cost_var=progress_from_curve(cost_curve)
mhr_plan,mhr_actual,mhr_var=progress_from_curve(unit_curve)

tabs = st.tabs([
    "Executive Dashboard","Engineering","Procurement","Progress Curves",
    "Cash Flow","Manpower","Milestones","Schedule Health","4-Week Lookahead",
    "Progress Photos","T&C / Handover","Report Center"
])

# ------------------------------------------------------------------
# Executive Dashboard
# ------------------------------------------------------------------
with tabs[0]:
    st.markdown('<div class="section-title">Executive Project Summary</div>', unsafe_allow_html=True)
    a,b = st.columns([1,2])
    with a:
        summary = pd.DataFrame({
            "Description":["Contract Commencement","Contract Completion","Forecast Completion","Total Duration (days)",
                           "Time Elapsed (days)","Remaining Time (days)","Time Elapsed (%)","Data Date","Schedule Status"],
            "Status":[fmt_date(bf["start"]),fmt_date(bf["finish"]),fmt_date(forecast_finish),
                      int(duration_days) if pd.notna(duration_days) else None,
                      int(elapsed_days) if pd.notna(elapsed_days) else None,
                      int(remaining_days) if pd.notna(remaining_days) else None,
                      f"{elapsed_pct:.1f}%" if pd.notna(elapsed_pct) else "—",
                      fmt_date(data_date),
                      "Behind Programme" if pd.notna(forecast_move) and forecast_move>0 else "On / Ahead"]
        })
        st.dataframe(summary, use_container_width=True, hide_index=True)
    with b:
        c=st.columns(6)
        c[0].metric("Forecast Movement", f"{forecast_move:+.0f} d" if pd.notna(forecast_move) else "—")
        c[1].metric("Critical Activities", f"{bf['critical']:,}")
        c[2].metric("Negative Float", f"{bf['negative']:,}")
        c[3].metric("MHR Planned", f"{mhr_plan:.2f}%" if pd.notna(mhr_plan) else "N/A")
        c[4].metric("MHR Actual", f"{mhr_actual:.2f}%" if pd.notna(mhr_actual) else "N/A")
        c[5].metric("Cost Actual", f"{cost_actual:.2f}%" if pd.notna(cost_actual) else "N/A")

        # Progress summary tables
        ptab = pd.DataFrame([
            ["MHR",mhr_plan,mhr_actual,mhr_var],
            ["Cost",cost_plan,cost_actual,cost_var],
        ], columns=["Weightage Basis","Cumulative Plan %","Cumulative Actual %","Variance %"])
        st.dataframe(
            ptab.style.format({"Cumulative Plan %":"{:.2f}%","Cumulative Actual %":"{:.2f}%","Variance %":"{:.2f}%"}).map(
                lambda v: f"color:{BRAND_RED};font-weight:700" if isinstance(v,(float,np.floating)) and v<0 else "",
                subset=["Variance %"]
            ),
            use_container_width=True, hide_index=True
        )

    c1,c2=st.columns(2)
    with c1:
        # Duration chart
        dd=pd.DataFrame({"Metric":["Original Duration","At Completion Duration","Time Elapsed"],
                         "Days":[duration_days, duration_days+(forecast_move if pd.notna(forecast_move) else 0), elapsed_days]})
        fig=px.bar(dd,x="Days",y="Metric",orientation="h",text="Days",title="Project Duration")
        fig.update_layout(height=320,xaxis_title="Days",yaxis_title="")
        st.plotly_chart(fig,use_container_width=True)
    with c2:
        # Management alerts
        active = curr if curr is not None else base
        engrisk=active[active["category"].isin(["Prequalification","Shop Drawings","Material Submittals"]) & (active["total_float_days"]<=near_tf)]
        procrisk=active[(active["category"]=="Procurement") & (active["total_float_days"]<=near_tf)]
        alerts=[
            f"{int((active.total_float_days<0).sum())} activities with negative float",
            f"{len(engrisk)} engineering items at/near critical",
            f"{len(procrisk)} procurement items at/near critical",
            f"Forecast completion movement: {forecast_move:+.0f} days" if pd.notna(forecast_move) else "Forecast movement unavailable",
            f"{lookahead_weeks}-week lookahead starts from {fmt_date(data_date)}",
        ]
        st.markdown("### Management Alerts")
        for x in alerts:
            st.write("•",x)

# ------------------------------------------------------------------
# Engineering
# ------------------------------------------------------------------
with tabs[1]:
    st.markdown('<div class="section-title">Engineering Control · Major Package Wise</div>', unsafe_allow_html=True)
    active=curr if curr is not None else base
    eng=active[active["category"].isin(["Prequalification","Shop Drawings","Material Submittals"])].copy()
    k=st.columns(5)
    k[0].metric("Engineering Items",len(eng))
    k[1].metric("PQ",int((eng.category=="Prequalification").sum()))
    k[2].metric("Shop Drawings",int((eng.category=="Shop Drawings").sum()))
    k[3].metric("Material Submittals",int((eng.category=="Material Submittals").sum()))
    k[4].metric("Negative Float",int((eng.total_float_days<0).sum()))

    matrix = eng.pivot_table(
        index="package", columns="category", values="finish", aggfunc="max"
    ).reset_index()
    st.subheader("Package Matrix")
    st.dataframe(matrix,use_container_width=True,hide_index=True)

    eng["Risk / Status"]=eng["total_float_days"].apply(status_from_float)
    eng["Days to Due"]=(eng["finish"]-data_date).dt.days
    st.subheader("Engineering Action Register")
    show=eng[["package","task_code","task_name","category","finish","total_float_days","Days to Due","Risk / Status"]].sort_values(
        ["total_float_days","finish"],na_position="last"
    )
    styled=show.style.format({"total_float_days":"{:.0f} d","Days to Due":"{:.0f} d"}).map(style_float,subset=["total_float_days"])
    st.dataframe(styled,use_container_width=True,hide_index=True)

# ------------------------------------------------------------------
# Procurement
# ------------------------------------------------------------------
with tabs[2]:
    st.markdown('<div class="section-title">Procurement Action Dashboard</div>', unsafe_allow_html=True)
    active=curr if curr is not None else base
    proc=active[active["category"]=="Procurement"].copy()
    proc["Action"]=proc.apply(procurement_action,axis=1,dd=data_date)
    proc["Days to Finish"]=(proc["finish"]-data_date).dt.days
    priority_order={"IMMEDIATE LPO REQUIRED":0,"CRITICAL / DELAY":1,"OVERDUE":2,"LPO DUE":3,"DELIVERY DUE":4,"Monitor":5,"Completed":6}
    proc["Priority"]=proc["Action"].map(priority_order).fillna(9)
    k=st.columns(6)
    labels=["IMMEDIATE LPO REQUIRED","CRITICAL / DELAY","OVERDUE","LPO DUE","DELIVERY DUE","Completed"]
    for i,lab in enumerate(labels):
        k[i].metric(lab.title(),int((proc["Action"]==lab).sum()))
    st.subheader("Priority Action Register")
    pshow=proc.sort_values(["Priority","total_float_days","finish"])[
        ["task_code","task_name","finish","total_float_days","Days to Finish","Action"]
    ]
    st.dataframe(
        pshow.style.format({"total_float_days":"{:.0f} d","Days to Finish":"{:.0f} d"}).map(style_float,subset=["total_float_days"]),
        use_container_width=True, hide_index=True
    )

# ------------------------------------------------------------------
# Progress Curves
# ------------------------------------------------------------------
with tabs[3]:
    st.markdown('<div class="section-title">Progress Curves</div>', unsafe_allow_html=True)
    c1,c2=st.columns(2)
    with c1:
        st.subheader("MHR / Unit S-Curve")
        if unit_curve.empty:
            st.warning("Resource Assignment units are required.")
        else:
            fig=go.Figure()
            fig.add_trace(go.Scatter(x=unit_curve.date,y=unit_curve.cum_planned_pct,name="Planned",mode="lines"))
            fig.add_trace(go.Scatter(x=unit_curve.date,y=unit_curve.cum_actual_pct,name="Actual",mode="lines"))
            fig.update_layout(height=420,yaxis_title="Cumulative %",xaxis_title="")
            st.plotly_chart(fig,use_container_width=True)
    with c2:
        st.subheader("Cost S-Curve")
        if cost_curve.empty:
            st.warning("Cost loading is required.")
        else:
            fig=go.Figure()
            fig.add_trace(go.Scatter(x=cost_curve.date,y=cost_curve.cum_planned_pct,name="Planned",mode="lines"))
            fig.add_trace(go.Scatter(x=cost_curve.date,y=cost_curve.cum_actual_pct,name="Actual",mode="lines"))
            fig.update_layout(height=420,yaxis_title="Cumulative %",xaxis_title="")
            st.plotly_chart(fig,use_container_width=True)

# ------------------------------------------------------------------
# Cash Flow
# ------------------------------------------------------------------
with tabs[4]:
    st.markdown('<div class="section-title">Cash Flow · Planned vs Actual</div>', unsafe_allow_html=True)
    if cost_curve.empty:
        st.warning("Planned and Actual Cost fields are required from Resource / Cost Assignments.")
    else:
        fig=go.Figure()
        fig.add_bar(x=cost_curve.date,y=cost_curve.planned,name="Monthly Planned")
        fig.add_bar(x=cost_curve.date,y=cost_curve.actual,name="Monthly Actual")
        fig.add_trace(go.Scatter(x=cost_curve.date,y=cost_curve.cum_planned,name="Cumulative Planned",yaxis="y2",mode="lines+markers"))
        fig.add_trace(go.Scatter(x=cost_curve.date,y=cost_curve.cum_actual,name="Cumulative Actual",yaxis="y2",mode="lines+markers"))
        fig.update_layout(barmode="group",height=500,yaxis_title="Monthly Cost",
                          yaxis2=dict(title="Cumulative Cost",overlaying="y",side="right"),xaxis_title="")
        st.plotly_chart(fig,use_container_width=True)
        st.dataframe(cost_curve,use_container_width=True,hide_index=True)

# ------------------------------------------------------------------
# Manpower
# ------------------------------------------------------------------
with tabs[5]:
    st.markdown('<div class="section-title">Manhours & Manpower</div>', unsafe_allow_html=True)
    if unit_curve.empty:
        st.warning("Planned/Actual Labor Units are required.")
    else:
        unit_curve["planned_manpower"]=unit_curve["planned"]/(hours_per_day*days_per_week)
        unit_curve["actual_manpower"]=unit_curve["actual"]/(hours_per_day*days_per_week)
        c1,c2=st.columns(2)
        with c1:
            fig=px.bar(unit_curve,x="date",y=["planned","actual"],barmode="group",title="Weekly Planned vs Actual Manhours")
            fig.update_layout(height=430,xaxis_title="",yaxis_title="Manhours / Labor Units")
            st.plotly_chart(fig,use_container_width=True)
        with c2:
            fig=px.bar(unit_curve,x="date",y=["planned_manpower","actual_manpower"],barmode="group",title="Weekly Planned vs Actual Manpower")
            fig.update_layout(height=430,xaxis_title="",yaxis_title="Headcount")
            st.plotly_chart(fig,use_container_width=True)
        st.caption(f"Headcount conversion: weekly labor hours ÷ ({hours_per_day:g} h/day × {days_per_week:g} days/week).")

# ------------------------------------------------------------------
# Milestones
# ------------------------------------------------------------------
with tabs[6]:
    st.markdown('<div class="section-title">Key Milestones</div>', unsafe_allow_html=True)
    active=curr if curr is not None else base
    ms=active[active["duration_days"].fillna(-1)==0].copy()
    ms["Status"]=ms["total_float_days"].apply(status_from_float)
    st.dataframe(
        ms[["task_code","task_name","finish","total_float_days","Status"]].sort_values("finish")
          .style.format({"total_float_days":"{:.0f} d"}).map(style_float,subset=["total_float_days"]),
        use_container_width=True,hide_index=True
    )

# ------------------------------------------------------------------
# Schedule health
# ------------------------------------------------------------------
with tabs[7]:
    st.markdown('<div class="section-title">Schedule Health</div>', unsafe_allow_html=True)
    active=curr if curr is not None else base
    pred=curr_pred if curr_pred is not None else base_pred
    pc=pred.groupby("task_id").size(); sc=pred.groupby("pred_task_id").size()
    nonms=active[active["duration_days"].fillna(0)>0]
    op=nonms[~nonms.task_code.isin(pc.index)]
    os=nonms[~nonms.task_code.isin(sc.index)]
    c=st.columns(5)
    c[0].metric("No Predecessor",len(op))
    c[1].metric("No Successor",len(os))
    c[2].metric("Critical / Zero Float",int((active.total_float_days<=0).sum()))
    c[3].metric("Negative Float",int((active.total_float_days<0).sum()))
    c[4].metric(f"Near Critical ≤ {near_tf}d",int(((active.total_float_days>0)&(active.total_float_days<=near_tf)).sum()))
    low=active.sort_values("total_float_days").head(100)
    st.dataframe(
        low[["task_code","task_name","finish","total_float_days","category"]]
           .style.format({"total_float_days":"{:.0f} d"}).map(style_float,subset=["total_float_days"]),
        use_container_width=True,hide_index=True
    )

# ------------------------------------------------------------------
# Lookahead
# ------------------------------------------------------------------
with tabs[8]:
    st.markdown(f'<div class="section-title">{lookahead_weeks}-Week Lookahead from Data Date</div>', unsafe_allow_html=True)
    active=curr if curr is not None else base
    la=date_window(active,data_date,lookahead_weeks)
    st.write(f"**Data Date:** {fmt_date(data_date)}  |  **Window End:** {fmt_date(data_date+pd.Timedelta(weeks=lookahead_weeks))}")
    st.dataframe(
        la[["task_code","task_name","category","start","finish","total_float_days","status_code"]]
          .style.format({"total_float_days":"{:.0f} d"}).map(style_float,subset=["total_float_days"]),
        use_container_width=True,hide_index=True
    )

# ------------------------------------------------------------------
# Photos
# ------------------------------------------------------------------
with tabs[9]:
    st.markdown('<div class="section-title">Progress Photos</div>', unsafe_allow_html=True)
    if not progress_photos:
        st.info("Upload weekly site progress photos from the sidebar.")
    else:
        cols=st.columns(3)
        for i,p in enumerate(progress_photos):
            with cols[i%3]:
                st.image(p,use_container_width=True)
                st.caption(p.name)

# ------------------------------------------------------------------
# T&C
# ------------------------------------------------------------------
with tabs[10]:
    st.markdown('<div class="section-title">T&C / Handover</div>', unsafe_allow_html=True)
    active=curr if curr is not None else base
    tch=active[active.category.isin(["T&C","Handover"])].sort_values("start")
    st.dataframe(
        tch[["task_code","task_name","category","start","finish","total_float_days"]]
          .style.format({"total_float_days":"{:.0f} d"}).map(style_float,subset=["total_float_days"]),
        use_container_width=True,hide_index=True
    )

# ------------------------------------------------------------------
# Report generators
# ------------------------------------------------------------------
def make_pdf():
    buf=BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=landscape(A4),rightMargin=12*mm,leftMargin=12*mm,topMargin=12*mm,bottomMargin=12*mm)
    styles=getSampleStyleSheet()
    title=ParagraphStyle("title",parent=styles["Title"],fontName="Helvetica-Bold",fontSize=24,textColor=colors.HexColor(BRAND_GREEN),alignment=TA_CENTER,spaceAfter=8)
    h=ParagraphStyle("h",parent=styles["Heading2"],fontName="Helvetica-Bold",fontSize=14,textColor=colors.HexColor(BRAND_GREEN))
    body=ParagraphStyle("body",parent=styles["BodyText"],fontSize=8.5)
    story=[]
    if COVER_IMG.exists():
        story.append(RLImage(str(COVER_IMG),width=255*mm,height=120*mm))
    story.append(Spacer(1,6))
    story.append(Paragraph("JUMEIRAH RETAIL DEVELOPMENT",title))
    story.append(Paragraph("Owner / Employer: ALTA &nbsp;&nbsp; | &nbsp;&nbsp; Contractor: DDDC &nbsp;&nbsp; | &nbsp;&nbsp; Consultant: DAR",body))
    story.append(Paragraph(f"Data Date: {fmt_date(data_date)}",body))
    story.append(PageBreak())

    story.append(Paragraph("Executive Summary",h))
    data=[
        ["Contract Commencement",fmt_date(bf["start"]),"Contract Completion",fmt_date(bf["finish"])],
        ["Forecast Completion",fmt_date(forecast_finish),"Forecast Movement",f"{forecast_move:+.0f} d" if pd.notna(forecast_move) else "—"],
        ["Time Elapsed",f"{elapsed_days:.0f} d","Remaining Time",f"{remaining_days:.0f} d"],
        ["Critical Activities",str(bf["critical"]),"Negative Float",str(bf["negative"])],
        ["MHR Planned",f"{mhr_plan:.2f}%" if pd.notna(mhr_plan) else "N/A","MHR Actual",f"{mhr_actual:.2f}%" if pd.notna(mhr_actual) else "N/A"],
        ["Cost Planned",f"{cost_plan:.2f}%" if pd.notna(cost_plan) else "N/A","Cost Actual",f"{cost_actual:.2f}%" if pd.notna(cost_actual) else "N/A"],
    ]
    t=Table(data,colWidths=[55*mm,35*mm,55*mm,35*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#F7F7F5")),
        ("GRID",(0,0),(-1,-1),.35,colors.HexColor("#D1D3D4")),
        ("FONTNAME",(0,0),(-1,-1),"Helvetica"),
        ("FONTSIZE",(0,0),(-1,-1),8),
        ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
        ("FONTNAME",(2,0),(2,-1),"Helvetica-Bold"),
    ]))
    story.append(t); story.append(Spacer(1,8))

    # Negative float table
    story.append(Paragraph("Schedule Risk / Negative Float",h))
    active=curr if curr is not None else base
    neg=active[active.total_float_days<0].sort_values("total_float_days").head(30)
    rows=[["Activity ID","Activity Name","Finish","Total Float"]]
    for _,r in neg.iterrows():
        rows.append([r.task_code, str(r.task_name)[:70], fmt_date(r.finish), f"{r.total_float_days:.0f} d"])
    tt=Table(rows,colWidths=[32*mm,125*mm,30*mm,25*mm])
    tt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor(BRAND_GREEN)),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),7.5),
        ("GRID",(0,0),(-1,-1),.3,colors.HexColor("#D1D3D4")),
        ("TEXTCOLOR",(-1,1),(-1,-1),colors.HexColor(BRAND_RED)),
    ]))
    story.append(tt); story.append(PageBreak())

    # Engineering
    story.append(Paragraph("Engineering Control",h))
    eng=(curr if curr is not None else base)
    eng=eng[eng.category.isin(["Prequalification","Shop Drawings","Material Submittals"])].sort_values(["total_float_days","finish"]).head(35)
    rows=[["Package","Stage","Finish","TF","Status"]]
    for _,r in eng.iterrows():
        rows.append([str(r.package)[:55],r.category,fmt_date(r.finish),f"{r.total_float_days:.0f} d" if pd.notna(r.total_float_days) else "—",status_from_float(r.total_float_days)])
    tt=Table(rows,colWidths=[100*mm,40*mm,28*mm,20*mm,35*mm])
    tt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor(BRAND_GREEN)),("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),7),
        ("GRID",(0,0),(-1,-1),.3,colors.HexColor("#D1D3D4")),
    ]))
    story.append(tt); story.append(PageBreak())

    # Procurement
    story.append(Paragraph("Procurement Action Register",h))
    proc=(curr if curr is not None else base)
    proc=proc[proc.category=="Procurement"].copy()
    proc["Action"]=proc.apply(procurement_action,axis=1,dd=data_date)
    proc=proc.sort_values(["total_float_days","finish"]).head(35)
    rows=[["Activity ID","Procurement Item","Finish","TF","Action"]]
    for _,r in proc.iterrows():
        rows.append([r.task_code,str(r.task_name)[:70],fmt_date(r.finish),f"{r.total_float_days:.0f} d" if pd.notna(r.total_float_days) else "—",r.Action])
    tt=Table(rows,colWidths=[28*mm,120*mm,28*mm,18*mm,45*mm])
    tt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor(BRAND_GREEN)),("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),7),
        ("GRID",(0,0),(-1,-1),.3,colors.HexColor("#D1D3D4")),
    ]))
    story.append(tt); story.append(PageBreak())

    # Lookahead
    story.append(Paragraph(f"{lookahead_weeks}-Week Lookahead",h))
    la=date_window(curr if curr is not None else base,data_date,lookahead_weeks).head(45)
    rows=[["Activity ID","Activity Name","Start","Finish","TF"]]
    for _,r in la.iterrows():
        rows.append([r.task_code,str(r.task_name)[:78],fmt_date(r.start),fmt_date(r.finish),f"{r.total_float_days:.0f} d" if pd.notna(r.total_float_days) else "—"])
    tt=Table(rows,colWidths=[30*mm,125*mm,28*mm,28*mm,18*mm])
    tt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor(BRAND_GREEN)),("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),7),
        ("GRID",(0,0),(-1,-1),.3,colors.HexColor("#D1D3D4")),
    ]))
    story.append(tt)

    if progress_photos:
        story.append(PageBreak()); story.append(Paragraph("Progress Photos",h))
        for p in progress_photos[:6]:
            pdata=BytesIO(p.getvalue())
            try:
                story.append(RLImage(pdata,width=110*mm,height=70*mm))
                story.append(Spacer(1,4))
            except: pass

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()

def add_title(slide, text):
    box=slide.shapes.add_textbox(Inches(.5),Inches(.25),Inches(12.3),Inches(.5))
    p=box.text_frame.paragraphs[0]; p.text=text; p.font.size=Pt(24); p.font.bold=True; p.font.color.rgb=RGBColor(7,121,92)

def make_pptx():
    prs=Presentation()
    prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)

    # cover
    slide=prs.slides.add_slide(prs.slide_layouts[6])
    if COVER_IMG.exists():
        slide.shapes.add_picture(str(COVER_IMG),0,0,width=prs.slide_width,height=prs.slide_height)
    overlay=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,Inches(6.3),prs.slide_height)
    overlay.fill.solid(); overlay.fill.fore_color.rgb=RGBColor(18,18,18); overlay.fill.transparency=18; overlay.line.fill.background()
    tb=slide.shapes.add_textbox(Inches(.6),Inches(4.6),Inches(5.4),Inches(1.4))
    tf=tb.text_frame; p=tf.paragraphs[0]; p.text="JUMEIRAH RETAIL\nDEVELOPMENT"; p.font.size=Pt(30); p.font.bold=True; p.font.color.rgb=RGBColor(247,247,245)
    p2=tf.add_paragraph(); p2.text=f"Data Date: {fmt_date(data_date)}"; p2.font.size=Pt(14); p2.font.color.rgb=RGBColor(209,211,212)

    # executive
    slide=prs.slides.add_slide(prs.slide_layouts[6]); add_title(slide,"Executive Project Summary")
    metrics=[
        ("Forecast Completion",fmt_date(forecast_finish)),
        ("Forecast Movement",f"{forecast_move:+.0f} d" if pd.notna(forecast_move) else "—"),
        ("Negative Float",str(bf["negative"])),
        ("MHR Actual",f"{mhr_actual:.2f}%" if pd.notna(mhr_actual) else "N/A"),
        ("Cost Actual",f"{cost_actual:.2f}%" if pd.notna(cost_actual) else "N/A"),
    ]
    x=.6
    for title,val in metrics:
        sh=slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(1.2),Inches(2.25),Inches(1.25))
        sh.fill.solid(); sh.fill.fore_color.rgb=RGBColor(247,247,245); sh.line.color.rgb=RGBColor(209,211,212)
        tf=sh.text_frame; tf.clear()
        p=tf.paragraphs[0]; p.text=title; p.font.size=Pt(11); p.font.color.rgb=RGBColor(18,18,18)
        q=tf.add_paragraph(); q.text=val; q.font.size=Pt(24); q.font.bold=True; q.font.color.rgb=RGBColor(214,40,40) if ("Movement" in title and pd.notna(forecast_move) and forecast_move>0) or title=="Negative Float" else RGBColor(7,121,92)
        x+=2.45

    # Engineering
    slide=prs.slides.add_slide(prs.slide_layouts[6]); add_title(slide,"Engineering · Major Package Wise")
    tb=slide.shapes.add_textbox(Inches(.6),Inches(1.1),Inches(12),Inches(5.8))
    tf=tb.text_frame; tf.word_wrap=True
    eng=(curr if curr is not None else base)
    eng=eng[eng.category.isin(["Prequalification","Shop Drawings","Material Submittals"])].sort_values(["total_float_days","finish"]).head(12)
    for i,(_,r) in enumerate(eng.iterrows()):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.text=f"{r.package[:45]}  |  {r.category}  |  Finish {fmt_date(r.finish)}  |  TF {r.total_float_days:.0f} d"
        p.font.size=Pt(13); p.font.color.rgb=RGBColor(214,40,40) if pd.notna(r.total_float_days) and r.total_float_days<0 else RGBColor(18,18,18)

    # Procurement
    slide=prs.slides.add_slide(prs.slide_layouts[6]); add_title(slide,"Procurement · Priority Actions")
    tb=slide.shapes.add_textbox(Inches(.6),Inches(1.1),Inches(12),Inches(5.8))
    tf=tb.text_frame
    proc=(curr if curr is not None else base)
    proc=proc[proc.category=="Procurement"].copy(); proc["Action"]=proc.apply(procurement_action,axis=1,dd=data_date)
    proc=proc.sort_values(["total_float_days","finish"]).head(12)
    for i,(_,r) in enumerate(proc.iterrows()):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.text=f"{r.task_code} · {r.task_name[:60]} · {r.Action} · TF {r.total_float_days:.0f} d"
        p.font.size=Pt(13); p.font.color.rgb=RGBColor(214,40,40) if pd.notna(r.total_float_days) and r.total_float_days<0 else RGBColor(18,18,18)

    # Lookahead
    slide=prs.slides.add_slide(prs.slide_layouts[6]); add_title(slide,f"{lookahead_weeks}-Week Lookahead · From {fmt_date(data_date)}")
    tb=slide.shapes.add_textbox(Inches(.6),Inches(1.1),Inches(12),Inches(5.8))
    tf=tb.text_frame
    la=date_window(curr if curr is not None else base,data_date,lookahead_weeks).head(15)
    for i,(_,r) in enumerate(la.iterrows()):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.text=f"{r.task_code} · {r.task_name[:70]} · {fmt_date(r.start)} → {fmt_date(r.finish)} · TF {r.total_float_days:.0f} d"
        p.font.size=Pt(12); p.font.color.rgb=RGBColor(214,40,40) if pd.notna(r.total_float_days) and r.total_float_days<0 else RGBColor(18,18,18)

    buf=BytesIO(); prs.save(buf); buf.seek(0); return buf.getvalue()

with tabs[11]:
    st.markdown('<div class="section-title">Report Center</div>', unsafe_allow_html=True)
    st.write("Generate management outputs from the same data and Data Date shown in the dashboard.")
    c1,c2=st.columns(2)
    with c1:
        try:
            pdf=make_pdf()
            st.download_button("Generate Weekly PDF Report",pdf,
                               f"JRD_Weekly_Report_{data_date.strftime('%Y%m%d')}.pdf",
                               "application/pdf",use_container_width=True)
        except Exception as e:
            st.error(f"PDF generation error: {e}")
    with c2:
        try:
            ppt=make_pptx()
            st.download_button("Generate Management Presentation",ppt,
                               f"JRD_Management_Presentation_{data_date.strftime('%Y%m%d')}.pptx",
                               "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                               use_container_width=True)
        except Exception as e:
            st.error(f"PowerPoint generation error: {e}")

st.markdown("---")
st.caption("JRD Project Controls Hub · v0.3 FINAL · DDDC branded · Negative float shown in red as delay indication")
