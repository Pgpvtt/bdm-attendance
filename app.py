# -*- coding: utf-8 -*-
"""
ACPL BDM Attendance — Web App
Upload the WhatsApp group export (exported WITH media) -> pick a date / month / range
-> live dashboard + download the standard "APPLE BDM VISIT" Excel for exactly that period.
"""
import io, datetime
from datetime import date, timedelta
import streamlit as st
import pandas as pd
from engine import (parse_events, filter_events, build_workbook, reconstruct,
                    store_count, store_coverage, anomalies, hm, _tmin, DEFAULT_ORDER)

st.set_page_config(page_title="ACPL BDM Attendance", page_icon="📍", layout="wide")
st.markdown("""<style>.block-container{padding-top:1.4rem;} h1{color:#1F3864;}
[data-testid="stMetricValue"]{color:#1F3864;}</style>""", unsafe_allow_html=True)

st.title("📍 ACPL — BDM Field Visit Attendance")
st.caption("Upload the WhatsApp group export (**Export chat → Attach Media**), choose a period, "
           "and download the ready-made Excel — same format, colours and formulas, nothing to add.")

with st.sidebar:
    st.header("How to use")
    st.markdown("1. WhatsApp → group → **Export chat → Attach Media**\n"
                "2. Upload the `.zip` below\n3. Pick the date / month / range\n"
                "4. Check the dashboard → **Download Excel**")
    st.divider()
    st.subheader("Tracked BDMs"); st.write(", ".join(DEFAULT_ORDER))

up = st.file_uploader("Upload WhatsApp export (.zip or .txt)", type=["zip", "txt"])
if not up:
    st.info("⬆️ Upload the exported chat to begin."); st.stop()

@st.cache_data(show_spinner="Reading export…")
def _load(data: bytes):
    return parse_events(data)

try:
    ev_all, unmapped, all_dates = _load(up.getvalue())
except Exception as e:
    st.error(f"Could not read the export: {e}"); st.stop()
if not all_dates:
    st.warning("No captioned photos found. Make sure you exported **With Media**."); st.stop()

min_d = date.fromisoformat(all_dates[0]); max_d = date.fromisoformat(all_dates[-1])

# ---------------- PERIOD PICKER ----------------
st.markdown("#### 🗓️ Choose period")
pc1, pc2 = st.columns([1, 2])
preset = pc1.selectbox("Quick pick", ["All data", "Latest day", "This month", "Last month",
                                      "Last 7 days", "Custom date / range"])
if preset == "All data":
    s, e = min_d, max_d
elif preset == "Latest day":
    s = e = max_d
elif preset == "This month":
    s, e = max_d.replace(day=1), max_d
elif preset == "Last month":
    e = max_d.replace(day=1) - timedelta(days=1); s = e.replace(day=1)
elif preset == "Last 7 days":
    s, e = max_d - timedelta(days=6), max_d
else:
    rng = pc2.date_input("Pick a single date or a range", (min_d, max_d),
                         min_value=min_d, max_value=max_d, format="DD-MM-YYYY")
    if isinstance(rng, (tuple, list)):
        s, e = (rng[0], rng[-1]) if len(rng) >= 2 else (rng[0], rng[0])
    else:
        s = e = rng
s, e = max(s, min_d), min(e, max_d)
if preset != "Custom date / range":
    pc2.info(f"Showing **{s.strftime('%d-%b-%Y')}**" + ("" if s == e else f"  →  **{e.strftime('%d-%b-%Y')}**"))

# ---------------- BUILD FOR SELECTED PERIOD ----------------
ev = filter_events(ev_all, s.isoformat(), e.isoformat())
sel_dates = sorted({d for b in ev for d in ev[b]})
if not sel_dates:
    st.warning("No visits in that period. Pick another date/range."); st.stop()
wb, (a_dates, latest, per) = build_workbook(ev, DEFAULT_ORDER)
dtl = datetime.datetime.strptime(latest, "%Y-%m-%d")

# ---------------- DOWNLOAD ----------------
buf = io.BytesIO(); wb.save(buf); buf.seek(0)
fname = (f"APPLE BDM VISIT - {dtl.strftime('%d-%m-%Y')}.xlsx" if s == e
         else f"APPLE BDM VISIT - {s.strftime('%d-%m')} to {e.strftime('%d-%m-%Y')}.xlsx")
d1, d2 = st.columns([3, 1])
d1.success(f"✅ Period **{s.strftime('%d-%b')} → {e.strftime('%d-%b-%Y')}**  ·  {len(sel_dates)} day(s)  ·  {len(DEFAULT_ORDER)} BDMs")
d2.download_button("⬇️ Download Excel", buf, file_name=fname,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True)

# ---------------- TODAY / LATEST ----------------
st.subheader(f"▶ Latest day in selection — {dtl.strftime('%d-%b-%Y (%A)')}")
rows = []
for b in DEFAULT_ORDER:
    v = reconstruct(ev.get(b, {}).get(latest, [])); reals = [x for x in v if not x["office"]]
    if reals:
        ins = [_tmin(x["tin"]) for x in reals]; outs = [_tmin(x["tout"]) for x in reals if x["tout"]]
        fm = sum(max(0, _tmin(x["tout"]) - _tmin(x["tin"])) for x in reals if x["tout"])
        rows.append([b, "🟢 Working", len(reals), hm(min(ins)), hm(max(outs)) if outs else "—", hm(fm)])
    else:
        rows.append([b, "🔴 On Leave", 0, "—", "—", "—"])
tdf = pd.DataFrame(rows, columns=["BDM", "Status", "Stores", "First In", "Last Out", "Field Time"])
mc = st.columns(4)
mc[0].metric("Stores (that day)", int(tdf["Stores"].sum()))
mc[1].metric("Working", int((tdf["Status"] == "🟢 Working").sum()))
mc[2].metric("On leave", int((tdf["Status"] == "🔴 On Leave").sum()))
mc[3].metric("Days in selection", len(sel_dates))
st.dataframe(tdf, hide_index=True, use_container_width=True)

# ---------------- PERIOD SUMMARY ----------------
st.subheader("▶ Period summary (selected range)")
prow = [[b, per[b]["days_worked"], per[b]["leave"], per[b]["total"], per[b]["avg_stores"],
         hm(per[b]["field"]), hm(per[b]["avg_store_min"]), hm(per[b]["avg_login"])] for b in DEFAULT_ORDER]
pdf = pd.DataFrame(prow, columns=["BDM", "Days Worked", "Days Leave", "Total Stores",
                                  "Avg Stores/Day", "Total Field Time", "Avg Time/Store", "Avg First-In"])
st.dataframe(pdf, hide_index=True, use_container_width=True)

# ---------------- DAILY MATRIX ----------------
st.subheader("▶ Daily store visits")
mat = []
for d in sel_dates[-31:]:
    dd = datetime.datetime.strptime(d, "%Y-%m-%d")
    r = [dd.strftime("%d-%b %a")] + [per[b]["per_day"].get(d, "") for b in DEFAULT_ORDER]
    r.append(sum(x for x in r[1:] if isinstance(x, int)))
    mat.append(r)
st.dataframe(pd.DataFrame(mat, columns=["Date"] + DEFAULT_ORDER + ["Team"]),
             hide_index=True, use_container_width=True)

# ---------------- STORE COVERAGE ----------------
st.subheader("▶ Store coverage")
cov = store_coverage(ev, DEFAULT_ORDER)
if cov:
    cc1, cc2, cc3 = st.columns(3)
    cc1.metric("Unique stores", len(cov))
    cc2.metric("Total visits", sum(c["visits"] for c in cov))
    cc3.metric("Visited only once", sum(1 for c in cov if c["visits"] == 1))
    cdf = pd.DataFrame([[i + 1, c["store"], c["visits"], c["bdms"],
                         datetime.datetime.strptime(c["last"], "%Y-%m-%d").strftime("%d-%b")]
                        for i, c in enumerate(cov)],
                       columns=["#", "Store", "Visits", "Visited by", "Last Visit"])
    tab1, tab2 = st.tabs([f"🔝 Most visited", "🔻 Least visited (coverage gaps)"])
    tab1.dataframe(cdf.head(20), hide_index=True, use_container_width=True)
    tab2.dataframe(cdf.tail(20).iloc[::-1], hide_index=True, use_container_width=True)

# ---------------- DATA QUALITY ----------------
st.subheader("▶ Data quality & anomalies")
issues, quality = anomalies(ev, DEFAULT_ORDER)
qdf = pd.DataFrame([[q["bdm"], q["visits"], q["missing_exit"], f"{q['pct']}%"] for q in quality],
                   columns=["BDM", "Visits", "Missing 'left' photo", "Missing %"])
qc1, qc2 = st.columns([1, 1])
qc1.markdown("**Missing exit ('left') photo rate**")
qc1.dataframe(qdf, hide_index=True, use_container_width=True)
qc2.markdown(f"**Specific anomalies to review ({len(issues)})**")
if issues:
    idf = pd.DataFrame([[datetime.datetime.strptime(d, "%Y-%m-%d").strftime("%d-%b"), b, s, iss]
                        for d, b, s, iss in issues], columns=["Date", "BDM", "Store", "Issue"])
    qc2.dataframe(idf, hide_index=True, use_container_width=True, height=240)
else:
    qc2.success("No anomalies in this period. 👍")

if unmapped:
    with st.expander("ℹ️ Other posters in the group (not tracked)"):
        st.dataframe(pd.DataFrame(sorted(unmapped.items(), key=lambda x: -x[1]),
                                  columns=["Sender", "Photos"]), hide_index=True, use_container_width=True)
