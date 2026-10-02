
# JRD Project Controls Hub — MVP v0.1

This is the first working prototype of the Project Controls reporting platform.

## What it does

Upload a Primavera P6 Excel export containing:

- `TASK`
- `TASKPRED`

The application automatically generates:

- Project overview
- Activity / relationship counts
- Critical and near-critical activity counts
- Open-end checks
- Milestone register
- Engineering classification
- Procurement dashboard
- Floor/location summary
- T&C and handover register
- Searchable activity browser

## Run locally

1. Install Python 3.11+.
2. Open a terminal in this folder.
3. Run:

```bash
pip install -r requirements.txt
streamlit run app.py
```

4. Open the URL shown by Streamlit.
5. Upload the Primavera Excel export.

## Current scope

This version is baseline-only.

The next version will add:

- Baseline vs weekly update
- Milestone movement
- New / lost critical activities
- Planned vs actual progress
- S-curves
- Manpower
- Engineering / procurement risk to construction
- AI executive summary
- Weekly report generation

## Input assumptions

The Primavera Excel export should use the standard technical column names seen in the JRD export, including:

### TASK
`task_code`, `status_code`, `wbs_id`, `task_name`, `target_drtn_hr_cnt`,
`start_date`, `end_date`, `total_float_hr_cnt`

### TASKPRED
`pred_task_id`, `task_id`, `pred_type`, `lag_hr_cnt`
