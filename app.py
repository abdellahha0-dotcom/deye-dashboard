#!/usr/bin/env python3
"""
Deye Solar Dashboard  --  Streamlit web app
===========================================
A password-protected website that:
  - lets you pick a date range and which stations,
  - runs the Deye Cloud pull when you click a button,
  - shows KPIs and charts (production vs consumption),
  - offers the data as a downloadable Excel file.

Secrets (Deye credentials + the site password) are NOT stored in this file.
They live in Streamlit's secrets manager. See the deploy notes.

Run locally:   streamlit run app.py
"""

import hashlib
import io
import time
from datetime import date, timedelta

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Deye Solar Dashboard", page_icon="☀️", layout="wide")

BASE_URLS = {
    "EU": "https://eu1-developer.deyecloud.com/v1.0",
    "US": "https://us1-developer.deyecloud.com/v1.0",
}


# ---------------------------------------------------------------- helpers
def cfg(key, default=None):
    """Read a value from Streamlit secrets."""
    try:
        return st.secrets[key]
    except Exception:
        return default


def base_url():
    return BASE_URLS[cfg("DEYE_REGION", "EU").upper()]


def sha256(text):
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def safe_sheet_name(name, used):
    clean = "".join(c for c in str(name) if c not in '[]:*?/\\')[:28] or "Station"
    cand, i = clean, 1
    while cand in used:
        cand = f"{clean[:25]}_{i}"; i += 1
    used.add(cand)
    return cand


# ---------------------------------------------------------------- deye api
@st.cache_data(ttl=2400, show_spinner=False)
def get_headers():
    """Log in and return auth headers. Cached ~40 min."""
    r = requests.post(
        f"{base_url()}/account/token?appId={cfg('DEYE_APP_ID')}",
        json={"appSecret": cfg("DEYE_APP_SECRET"), "email": cfg("DEYE_EMAIL"),
              "password": sha256(cfg("DEYE_PASSWORD")), "companyId": str(cfg("DEYE_COMPANY_ID", "0"))},
        headers={"Content-Type": "application/json"}, timeout=30)
    r.raise_for_status()
    data = r.json()
    token = data.get("accessToken") or data.get("token")
    if not token:
        raise RuntimeError(f"Login failed: {data.get('code')} {data.get('msg')}")
    return {"Content-Type": "application/json", "Authorization": "bearer " + token}


@st.cache_data(ttl=1800, show_spinner=False)
def load_stations():
    headers = get_headers()
    stations, page, size = [], 1, 50
    while True:
        d = requests.post(f"{base_url()}/station/list",
                          json={"page": page, "size": size}, headers=headers, timeout=30).json()
        batch = d.get("stationList") or d.get("list") or d.get("records") or []
        for s in batch:
            stations.append({"id": s.get("id") or s.get("stationId"),
                             "name": s.get("name") or s.get("stationName") or f"Station {s.get('id')}"})
        total = d.get("total", len(stations))
        if len(batch) < size or len(stations) >= total or not batch:
            break
        page += 1
    return stations


def get_day(headers, station_id, day_str):
    payload = requests.post(
        f"{base_url()}/station/history",
        json={"stationId": int(station_id), "granularity": 1, "startAt": day_str, "endAt": day_str},
        headers=headers, timeout=30).json()
    items = payload.get("stationDataItems", []) if isinstance(payload, dict) else []
    if not items:
        return pd.DataFrame(columns=["time", "production", "consumption"])
    return pd.DataFrame([{
        "time":        pd.to_datetime(it.get("timeStamp"), unit="s"),
        "production":  pd.to_numeric(it.get("generationPower"),  errors="coerce"),
        "consumption": pd.to_numeric(it.get("consumptionPower"), errors="coerce"),
    } for it in items]).dropna(subset=["time"])


def to_hourly(df):
    if df.empty:
        return df
    df = df.set_index("time").sort_index()
    g = df.resample("1h").agg(production_kw=("production", "mean"),
                              consumption_kw=("consumption", "mean"),
                              frames=("production", "count")).reset_index()
    g = g.rename(columns={"time": "timestamp"})
    g["production_kwh_est"] = g["production_kw"]
    g["consumption_kwh_est"] = g["consumption_kw"]
    g["date"] = g["timestamp"].dt.date
    g["hour"] = g["timestamp"].dt.hour
    return g


def pull(headers, chosen, days, progress):
    rows, steps, done = [], len(chosen) * len(days), 0
    for st_ in chosen:
        for d in days:
            day_str = d.strftime("%Y-%m-%d")
            try:
                hourly = to_hourly(get_day(headers, st_["id"], day_str))
                if not hourly.empty:
                    hourly.insert(0, "station_id", st_["id"])
                    hourly.insert(1, "station_name", st_["name"])
                    rows.append(hourly)
            except Exception as e:
                st.warning(f"{st_['name']} {day_str}: {e}")
            done += 1
            progress.progress(done / steps, text=f"Fetching… {st_['name']} {day_str}")
            time.sleep(0.35)
    if not rows:
        return pd.DataFrame()
    combined = pd.concat(rows, ignore_index=True)
    return combined[["station_id", "station_name", "timestamp", "date", "hour",
                     "production_kw", "consumption_kw",
                     "production_kwh_est", "consumption_kwh_est", "frames"]]


def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        df.to_excel(xl, sheet_name="All (per hour)", index=False)
        used = {"All (per hour)"}
        for _, grp in df.groupby("station_id"):
            grp.to_excel(xl, sheet_name=safe_sheet_name(grp["station_name"].iloc[0], used), index=False)
    return buf.getvalue()


# ---------------------------------------------------------------- login gate
def logged_in():
    if st.session_state.get("authed"):
        return True
    st.title("🔒 Deye Solar Dashboard")
    st.caption("Please log in to continue.")
    pw = st.text_input("Password", type="password")
    if st.button("Log in", type="primary"):
        if cfg("APP_PASSWORD") and pw == cfg("APP_PASSWORD"):
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


# ---------------------------------------------------------------- main app
def main():
    st.title("☀️ Deye Solar — Production & Consumption")

    try:
        stations = load_stations()
    except Exception as e:
        st.error(f"Could not connect to Deye. Check the secrets are set correctly.\n\n{e}")
        st.stop()

    with st.sidebar:
        st.header("Report settings")
        today = date.today()
        start = st.date_input("Start date", today - timedelta(days=2))
        end = st.date_input("End date", today)
        names = [s["name"] for s in stations]
        picked = st.multiselect("Stations", names, default=names)
        run = st.button("▶ Run report", type="primary", use_container_width=True)
        st.divider()
        if st.button("Log out", use_container_width=True):
            st.session_state.clear()
            st.rerun()

    if run:
        if start > end:
            st.error("Start date is after end date."); st.stop()
        chosen = [s for s in stations if s["name"] in picked]
        if not chosen:
            st.error("Pick at least one station."); st.stop()
        days = [start + timedelta(days=n) for n in range((end - start).days + 1)]
        headers = get_headers()
        bar = st.progress(0.0, text="Starting…")
        data = pull(headers, chosen, days, bar)
        bar.empty()
        st.session_state["data"] = data

    data = st.session_state.get("data")
    if data is None:
        st.info("Set your dates and stations on the left, then click **Run report**.")
        st.stop()
    if data.empty:
        st.warning("No data returned for that selection. Try a different date range.")
        st.stop()

    # ---- KPIs
    prod_total = data["production_kwh_est"].sum()
    cons_total = data["consumption_kwh_est"].sum()
    c1, c2, c3 = st.columns(3)
    c1.metric("Est. production (kWh)", f"{prod_total:,.0f}")
    c2.metric("Est. consumption (kWh)", f"{cons_total:,.0f}")
    c3.metric("Stations in report", data["station_name"].nunique())

    st.caption("‘Est.’ = estimated from the intraday power curve (avg power × 1h). "
               "For exact billing kWh use Deye’s daily totals endpoint.")

    # ---- Time series: fleet production vs consumption
    st.subheader("Power over time (all selected stations)")
    ts = (data.groupby("timestamp")[["production_kw", "consumption_kw"]].sum()
          .rename(columns={"production_kw": "Production (kW)", "consumption_kw": "Consumption (kW)"}))
    st.line_chart(ts)

    # ---- Energy per station
    st.subheader("Energy by station (kWh, estimated)")
    per = (data.groupby("station_name")[["production_kwh_est", "consumption_kwh_est"]].sum()
           .rename(columns={"production_kwh_est": "Production", "consumption_kwh_est": "Consumption"}))
    st.bar_chart(per)

    # ---- Optional drill-down
    with st.expander("Look at one station in detail"):
        pick = st.selectbox("Station", sorted(data["station_name"].unique()))
        one = data[data["station_name"] == pick].set_index("timestamp")
        st.line_chart(one[["production_kw", "consumption_kw"]]
                      .rename(columns={"production_kw": "Production (kW)",
                                       "consumption_kw": "Consumption (kW)"}))

    # ---- Table + download
    st.subheader("Data")
    st.dataframe(data, use_container_width=True, height=300)
    fname = f"deye_hourly_{start}_to_{end}.xlsx"
    st.download_button("⬇ Download Excel", data=to_excel_bytes(data),
                       file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")


if __name__ == "__main__":
    if logged_in():
        main()
