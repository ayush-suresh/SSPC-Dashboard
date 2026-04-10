# SSPC Quality Dashboard

## Run Locally (2 minutes)

### 1. Install Python (one-time)
Download from https://python.org — get Python 3.11 or newer

### 2. Install dependencies (one-time)
```
pip install -r requirements.txt
```

### 3. Run the app
```
streamlit run app.py
```
Opens automatically at http://localhost:8501

---

## Deploy to Streamlit Cloud (free, permanent URL)

1. Create a free account at https://streamlit.io
2. Push this folder to a GitHub repo
3. Click "New app" → connect your GitHub repo → select `app.py`
4. Done — you get a permanent URL like `yourteam.streamlit.app`

Anyone with the link can use it — no install needed.

---

## What this sample contains

- All months Jan 2025 → Mar 2026 hardcoded
- Full interactive dashboard:
  - Month selector (sidebar)
  - KPI cards with vs. prev month deltas
  - Cost breakdown (AuST vs CP, good vs scrap)
  - Costed yield by product
  - Top 10 failure modes bar chart
  - Defect detail table
  - Leak valve vs bond breakdown (Mar 2026)
  - Leak Rate trend (6 months)
  - Destroyed Tip Rate trend (6 months)
  - CoPQ/Part all-time trend (highlighted bar)
  - Leak summary KPIs

## Next step: file upload automation

The next version will let users upload the AuST + CP DOR .xlsm files
directly in the browser. The app will extract, compute, and save to a 
database automatically — no Excel needed.
