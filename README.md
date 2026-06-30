# ACPL — BDM Field Visit Attendance

Turns the **WhatsApp BDM attendance group** export into the standard *APPLE BDM VISIT* Excel
(one sheet per BDM, in the exact format the team uses) plus a **dashboard** — automatically.

The BDMs caption each store photo with the store name, and tag a *"left"* photo when they leave.
This app reads those captions + timestamps and rebuilds each day's visits, timings and totals.

## How a user uses it (daily)
1. In WhatsApp open the group **"ACPL BDM's ATTENDANCE"** → **Export chat → Attach Media**
   (must be *With Media* — "Without Media" strips the store-name captions).
2. Open the app link → upload the `.zip`.
3. Review the dashboard (Today / Period / last-14-days).
4. Click **Download Excel** for the records.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
Opens at http://localhost:8501

## Deploy (free public link)
1. Push this folder to a GitHub repo.
2. Go to https://share.streamlit.io → **New app** → pick the repo → main file `app.py` → Deploy.
3. Share the link. To update: edit + push to GitHub (auto-redeploys). To pause: stop the app from the Streamlit dashboard.

## Configuration
- `engine.py` → `DEFAULT_MAP` maps WhatsApp sender → BDM sheet. Add an unsaved phone number
  here if a BDM posts from a second number.
- `DEFAULT_ORDER` is the sheet order / tracked BDM list.

## How it works (logic)
- Each store visit = an **entry** photo (caption = store name) + an **exit** photo (caption contains *"left"*).
  `TIME IN` = entry, `TIME OUT` = exit, `STORE TIME = OUT − IN`. No exit photo → `NO UPDATE`.
- A caption containing *"office"* is shown as an OFFICE row but **not** counted in the store total.
- Duplicate same-minute photos (some BDMs send twice) are de-duplicated.
- Validated against the manually-maintained June file: matches within ±1 store on ~98% of working days.
