"""
GB Power Market Fundamentals Model
====================================
Pulls live data from NESO and Elexon APIs, runs the merit order
dispatch model, and writes a formatted Excel output file.

Data sources:
  - NESO: 2-14 day ahead demand, transmission wind, embedded wind/solar
  - Elexon FOU2T14D: nuclear, biomass and interconnector availability 2-14 days ahead
  - ENTSO-E: day-ahead prices for the interconnector counterparty bidding
    zones, used to shape import prices across the day
  - ECB: EUR/GBP reference rate

Run:
    export ENTSOE_API_TOKEN=...   # free from transparency.entsoe.eu
    python run_model.py

Without ENTSOE_API_TOKEN the model still runs, but falls back to flat
interconnector import prices, which materially flattens the GB price.

Requires:
    pip install requests pandas openpyxl

Output:
    gb_power_model_YYYYMMDD_HHMM.xlsx
"""

import json
import os
import re
import sys
import xml.etree.ElementTree as ET

import requests
import pandas as pd
from datetime import datetime, timedelta, date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference


# ===========================================================================
# ASSUMPTIONS — edit these
# ===========================================================================

GAS_PRICE_P_THERM      = 80.0    # p/therm (NBP)
ETS_PRICE_GBP_TONNE    = 45.0    # £/tonne CO2 — UK ETS
CPS_RATE_GBP_TONNE     = 18.0    # £/tonne CO2 — Carbon Price Support

# Emissions factors (gross CV basis)
ETS_EMISSIONS          = 0.18422  # UKGHGI 2025 median LDZ, tCO2/MWh_th
CPS_EMISSIONS          = 0.18259  # BEIS condensed set 2025, tCO2/MWh_th

# Gas fleet — 8 tranches using Argus published efficiency breakpoints
# (30, 34, 36, 38, 40, 42, 46, 49, 55, 58, 60pc)
# Capacities from NESO 2026/27 Final ALF data mapped to Argus efficiency bands
# CCGT VOM: £5.00/MWh from DESNZ 2025 generation costs summary table (all gas plant, 2024 prices)
# (efficiency, capacity_gw, vom_gbp_mwh)
GAS_TRANCHES = [
    (0.60, 3.9, 5.00),  # Newest CCGT: Pembroke (2012), Marchwood (2009), Fawley (2009)
    (0.58, 2.8, 5.00),  # Recent CCGT: Carrington (2016), Keadby II (2022)
    (0.55, 5.2, 5.00),  # Modern F-class: Staythorpe, West Burton B, Grain, Langage (2010-13)
    (0.49, 6.3, 5.00),  # Mid-vintage 2000s: Saltend, Peterhead, Seabank, Spalding etc.
    (0.46, 6.2, 5.00),  # Late 90s: South Humber Bank, Didcot B, Damhead Creek etc.
    (0.42, 2.7, 5.00),  # Oldest CCGT: Connah's Quay, Medway, Little Barford (1994-96)
    (0.38, 2.5, 5.00),  # OCGT standard peakers
    (0.34, 0.5, 5.00),  # OCGT older/smaller peakers
]
NUCLEAR_GW             = 6.4    # Active installed capacity (fallback if FOU2T14D unavailable)
                                 # Heysham1 1.2GW + Heysham2 1.3GW + Hartlepool 1.2GW
                                 # + Sizewell B 1.3GW + Torness 1.4GW (Elexon BM Units 2025)
                                 # Dungeness B + Hinkley B decommissioned (0 MW in Elexon data)
# Hydro: NPSHYD (run-of-river/reservoir) + PS (pumped storage)
# Registered from Elexon BM Units 2025: NPSHYD ~1.0 GW, PS ~3.1 GW
HYDRO_AVG_GW           = 4.1   # fallback if FOU2T14D unavailable
# Interconnector tranches — each link connects to a different market with its
# own price. Capacities are registered maximums; live availability comes from
# Elexon FOU2T14D.
#
# Import prices are built as  level x intraday shape  (see build_ic_price_shapes):
#   shape — per settlement period, from a trailing window of ENTSO-E day-ahead
#           prices for the counterparty bidding zone, split weekday/weekend.
#           Continental shape is severe now: midday collapses toward (or below)
#           zero on solar, evening peaks near 2x the daily mean.
#   level — a forward baseload price per link when one is supplied (see
#           FORWARD_LEVELS_FILE), otherwise the trailing realised mean.
#           Day-ahead prices do not exist for a 3-14 day ahead window, so the
#           trailing mean is only a proxy — a forward curve is the right input.
#
# France sums three codes (INTFR=IFA1, INTIFA2, INTELEC=ElecLink).
# Ireland sums three (INTIRL=Moyle, INTEW=EWIC, INTGRNL=Greenlink).
# Format: (capacity_gw, label, fou2t14d_fuel_types, entsoe_eic)
INTERCONNECTORS = [
    (1.4, "Interconnector NO", ["INTNSL"], "10YNO-2--------T"),   # North Sea Link Norway (NO2)
    (1.5, "Interconnector DK", ["INTVKL"], "10YDK-1--------W"),   # VikingLink Denmark (DK1)
    (4.1, "Interconnector FR", ["INTFR", "INTIFA2", "INTELEC"], "10YFR-RTE------C"), # IFA1 2.0 + IFA2 1.1 + ElecLink 1.0 GW
    (1.0, "Interconnector BE", ["INTNEM"], "10YBE----------2"),   # Nemo Belgium
    (1.2, "Interconnector NL", ["INTNED"], "10YNL----------L"),   # BritNed Netherlands
    (1.3, "Interconnector IE", ["INTIRL", "INTEW", "INTGRNL"], "10Y1001A1001A59C"),  # Moyle 0.4 + EWIC 0.5 + Greenlink 0.5 GW (SEM)
]

# Forward price inputs for the forecast window, read at run time from
# FORWARD_LEVELS_FILE. It carries three things:
#
#   levels           — one BASELOAD price per interconnector link. Not peak:
#                      the intraday shape is applied on top of this level, so a
#                      peak assessment would count the daily shape twice.
#                      Quoted in "currency" (EUR default, converted at the ECB
#                      rate). Links omitted or null fall back to the trailing
#                      ENTSO-E realised mean.
#   nbp_p_therm      — NBP gas, p/therm. Either one number for the window or a
#                      {date: price} curve; a curve is preferred because NBP
#                      moves day to day and gas sets the GB price in most
#                      periods. Always GBP, never converted.
#   uk_ets_gbp_tonne — UKA carbon, £/tonne. One number: carbon is an annual
#                      compliance instrument with no meaningful shape.
#
# Neither gas nor carbon takes an intraday shape. A gas day is a single
# product, so a plant's fuel cost is flat across all 48 settlement periods,
# and carbon has no intraday dimension at all.
#
# Argus (and any other) forward assessments are LICENSED DATA. This repository
# is public, so those numbers must not be committed to it — the file is
# gitignored. See forward_levels.example.json for the format. Anything absent
# falls back to the hardcoded assumptions above.
FORWARD_LEVELS_FILE    = "forward_levels.json"
# Fallback import prices (£/MWh) used only when ENTSO-E is unavailable — flat,
# and therefore a poor approximation of a shaped continental price.
IC_FALLBACK_PRICE = {
    "Interconnector NO": 87.0, "Interconnector DK": 88.0,
    "Interconnector FR": 80.0, "Interconnector BE": 95.0,
    "Interconnector NL": 91.0, "Interconnector IE": 130.0,
}
# Trailing window (days) used to estimate interconnector price level and shape
IC_SHAPE_DAYS          = 30
CCGT_VOM               = 5.00   # DESNZ 2025 generation costs, all gas plant (£/MWh, 2024 prices)
OCGT_VOM               = 5.0
# Nuclear and hydro VOM: no public DESNZ source — commercially confidential.
# Values are assumptions only. These plant never set the price in this model.
NUCLEAR_VOM            = 4.0   # ASSUMPTION — no published source
HYDRO_VOM              = 3.0   # ASSUMPTION — no published source

# Biomass: Drax (2.58 GW) + Lynemouth (0.42 GW) are CfD/RO supported.
# Bid near zero — CfD makes wholesale price largely irrelevant to dispatch decision.
BIOMASS_GW             = 3.0
BIOMASS_VOM            = 0.5   # Near-zero bid due to CfD/RO support

# System costs adder removed — this model produces the wholesale market clearing
# price (SRMC of marginal plant). BSUoS/TNUoS/DUoS are non-commodity charges
# recovered separately and are not part of wholesale price formation.

# Model start: day 3 from today (NESO 2-14 day ahead data)
MODEL_START_DAY        = 3
MODEL_END_DAY          = 14


# ===========================================================================
# API DATA PULL
# ===========================================================================

# NESO Data Portal (CKAN datastore)
NESO_API               = "https://api.neso.energy/api/3/action/datastore_search"
DEMAND_2_14_RESOURCE   = "7c0411cd-2714-4bb5-a408-adb065edf34d"  # 2-14 day ahead demand
WIND_14DA_RESOURCE     = "93c3048e-1dab-4057-a2a9-417540583929"  # 14 day ahead transmission wind
EMBEDDED_RESOURCE      = "db6c038f-98af-4570-ab60-24d71ebd0ae5"  # embedded wind + solar

# Elexon Insights API
ELEXON_API             = "https://data.elexon.co.uk/bmrs/api/v1"

# ENTSO-E Transparency Platform — day-ahead prices for the counterparty
# bidding zones. Needs a free registration token, read from the environment;
# never hardcode it here.
ENTSOE_API             = "https://web-api.tp.entsoe.eu/api"
ENTSOE_TOKEN_ENV       = "ENTSOE_API_TOKEN"
ENTSOE_RESOLUTION      = {"PT15M": 15, "PT30M": 30, "PT60M": 60, "PT1H": 60}

# ECB daily reference rates — EUR/GBP for converting continental prices
ECB_FX_API             = ("https://data-api.ecb.europa.eu/service/data/EXR/"
                          "D.GBP.EUR.SP00.A")
EUR_GBP_FALLBACK       = 0.855


def fetch_eur_gbp() -> float:
    """Latest ECB EUR->GBP reference rate, falling back to a fixed rate."""
    try:
        r = requests.get(ECB_FX_API,
                         params={"lastNObservations": 1, "format": "csvdata"},
                         timeout=30)
        r.raise_for_status()
        rows = r.text.strip().splitlines()
        cols = rows[0].split(",")
        rate = float(rows[1].split(",")[cols.index("OBS_VALUE")])
        print(f"  EUR/GBP (ECB): {rate}")
        return rate
    except Exception as e:
        print(f"  ECB FX lookup failed ({e}); using {EUR_GBP_FALLBACK}")
        return EUR_GBP_FALLBACK


def fetch_entsoe_da_prices(eic: str, start: date, end: date,
                           token: str) -> pd.Series:
    """
    Day-ahead prices (EUR/MWh) for one bidding zone, as a UTC-indexed Series.

    The A44 document carries one or more TimeSeries whose resolution varies by
    zone (PT15M for the CWE zones, PT60M for SEM). Points use curveType A03,
    where a price holds until the next stated position, so positions can be
    sparse and must be forward-filled onto the full grid — reading the points
    in listed order would misalign the series.
    """
    r = requests.get(ENTSOE_API, params={
        "securityToken": token,
        "documentType":  "A44",
        "in_Domain":     eic,
        "out_Domain":    eic,
        "periodStart":   start.strftime("%Y%m%d%H%M"),
        "periodEnd":     end.strftime("%Y%m%d%H%M"),
    }, timeout=90)
    if r.status_code != 200:
        # Deliberately not raise_for_status: its message embeds the request
        # URL, which carries the security token.
        raise RuntimeError(f"ENTSO-E returned HTTP {r.status_code}")
    root = ET.fromstring(r.text)
    ns = {"n": root.tag.split("}")[0].strip("{")}
    chunks = []
    for ts in root.findall("n:TimeSeries", ns):
        for per in ts.findall("n:Period", ns):
            res = ENTSOE_RESOLUTION[per.find("n:resolution", ns).text]
            t0  = pd.Timestamp(per.find("n:timeInterval/n:start", ns).text)
            t1  = pd.Timestamp(per.find("n:timeInterval/n:end", ns).text)
            n   = int((t1 - t0).total_seconds() // 60 // res)
            grid = pd.Series(index=range(1, n + 1), dtype=float)
            for p in per.findall("n:Point", ns):
                grid[int(p.find("n:position", ns).text)] = \
                    float(p.find("n:price.amount", ns).text)
            grid = grid.ffill()
            chunks.append(pd.Series(
                grid.values,
                index=pd.date_range(t0, periods=n, freq=f"{res}min", tz="UTC")))
    if not chunks:
        return pd.Series(dtype=float)
    s = pd.concat(chunks).sort_index()
    return s[~s.index.duplicated(keep="last")]


def load_forward_inputs(fx: float) -> dict:
    """
    Read forward assessments from FORWARD_LEVELS_FILE.

    Returns {"ic_levels": {label: gbp_mwh},
             "nbp":       pd.Series indexed by date (p/therm) or None,
             "ets":       float or None}

    The file is gitignored so licensed assessments stay out of this public
    repository. A missing or malformed file is not fatal — every caller falls
    back to the hardcoded assumptions.
    """
    empty = {"ic_levels": {}, "nbp": None, "ets": None}
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        FORWARD_LEVELS_FILE)
    if not os.path.exists(path):
        print(f"  No {FORWARD_LEVELS_FILE} — using hardcoded gas/carbon and "
              f"trailing ENTSO-E interconnector levels")
        return empty
    try:
        with open(path) as fh:
            blob = json.load(fh)
        ccy = str(blob.get("currency", "EUR")).upper()
        if ccy not in ("EUR", "GBP"):
            raise ValueError(f"currency must be EUR or GBP, got {ccy!r}")

        # --- interconnector baseload levels ---
        levels  = blob.get("levels") or {}
        known   = {label for _, label, _, _ in INTERCONNECTORS}
        unknown = set(levels) - known
        if unknown:
            print(f"  {FORWARD_LEVELS_FILE}: ignoring unknown links "
                  f"{sorted(unknown)}")
        rate = fx if ccy == "EUR" else 1.0
        ic_levels = {k: float(v) * rate for k, v in levels.items()
                     if k in known and v is not None}

        # --- NBP: a scalar or a {date: p/therm} curve. Always GBP. ---
        nbp_raw = blob.get("nbp_p_therm")
        nbp = None
        if isinstance(nbp_raw, dict) and nbp_raw:
            nbp = pd.Series({pd.to_datetime(k).date(): float(v)
                             for k, v in nbp_raw.items() if v is not None})
            # An untouched template is all nulls — treat that as "not supplied"
            nbp = nbp.sort_index() if not nbp.empty else None
        elif isinstance(nbp_raw, (int, float)):
            nbp = float(nbp_raw)

        ets = blob.get("uk_ets_gbp_tonne")
        ets = float(ets) if isinstance(ets, (int, float)) else None

        bits = [f"{len(ic_levels)} link level(s) in {ccy}"]
        if isinstance(nbp, pd.Series):
            bits.append(f"NBP curve {len(nbp)}d "
                        f"{nbp.min():.1f}-{nbp.max():.1f} p/therm")
        elif nbp is not None:
            bits.append(f"NBP flat {nbp:.1f} p/therm")
        if ets is not None:
            bits.append(f"UK ETS £{ets:.2f}/t")
        print(f"  {FORWARD_LEVELS_FILE}: " + ", ".join(bits))
        return {"ic_levels": ic_levels, "nbp": nbp, "ets": ets}
    except Exception as e:
        print(f"  {FORWARD_LEVELS_FILE} unreadable ({e}); "
              f"falling back to hardcoded assumptions")
        return empty


def resolve_gas_curve(nbp, model_dates) -> pd.Series:
    """
    A p/therm price for every modelled date.

    Accepts a scalar, a {date: price} Series, or None. A curve is reindexed
    onto the modelled dates and filled forward then backward, so a curve that
    is short at either end holds its nearest price rather than snapping back
    to the static assumption mid-window.
    """
    dates = sorted(set(model_dates))
    if isinstance(nbp, pd.Series) and not nbp.empty:
        s = nbp.reindex(dates)
        missing = int(s.isna().sum())
        s = s.ffill().bfill()
        if s.isna().any():
            s = s.fillna(GAS_PRICE_P_THERM)
        if missing == len(dates):
            # Almost always a curve left over from an earlier run: the window
            # rolls daily, so yesterday's dates stop overlapping. Nothing is
            # carried across a disjoint window — fall back to the assumption
            # rather than extrapolate a stale price over the whole run.
            print(f"  WARNING: NBP curve covers {nbp.index.min()}.."
                  f"{nbp.index.max()} but the model runs {dates[0]}..{dates[-1]}"
                  f" — no overlap, so every day falls back to the hardcoded "
                  f"{GAS_PRICE_P_THERM} p/therm. Regenerate with "
                  f"--init-forward.")
        elif missing:
            print(f"  NBP curve missing {missing}/{len(dates)} modelled "
                  f"date(s) — filled from nearest quoted day")
        return s
    level = nbp if isinstance(nbp, (int, float)) else GAS_PRICE_P_THERM
    return pd.Series(level, index=dates)


def build_ic_price_shapes(fx: float, forward: dict) -> dict:
    """
    Build an import price for every interconnector and settlement period.

    Returns {label: {(is_weekend, sp): price_gbp_mwh}}. The level is the
    forward baseload price for the link where one is supplied, else the
    trailing realised mean; the shape is the trailing mean ratio to daily
    mean, by settlement period, split weekday/weekend because continental
    weekend middays sit far lower.

    Any zone that cannot be fetched falls back to a flat IC_FALLBACK_PRICE.
    """
    token = os.environ.get(ENTSOE_TOKEN_ENV)
    flat = {lab: {(w, sp): IC_FALLBACK_PRICE[lab]
                  for w in (False, True) for sp in range(1, 49)}
            for _, lab, _, _ in INTERCONNECTORS}
    if not token:
        print(f"  {ENTSOE_TOKEN_ENV} not set — using flat fallback import prices")
        return flat, {lab: "flat fallback" for lab in flat}

    end     = date.today()
    start   = end - timedelta(days=IC_SHAPE_DAYS)
    shapes, sources = {}, {}
    for cap_gw, label, _, eic in INTERCONNECTORS:
        fwd_level = forward.get(label)
        try:
            print(f"  Fetching ENTSO-E day-ahead {label} ({eic})...")
            s = fetch_entsoe_da_prices(eic, start, end, token)
            if s.empty:
                raise ValueError("no price points returned")
            loc = s.tz_convert("Europe/London")
            # Hourly zones (SEM) upsample by holding the price across the hour
            hh = loc.resample("30min").mean().ffill() * fx
            sp = hh.index.hour * 2 + (hh.index.minute >= 30).astype(int) + 1
            daily = hh.groupby(hh.index.date).transform("mean")
            ratio = hh / daily
            wknd  = hh.index.dayofweek >= 5
            level = fwd_level if fwd_level is not None else float(hh.mean())
            by_sp = {}
            for is_wknd in (False, True):
                mask = wknd if is_wknd else ~wknd
                grp  = ratio[mask].groupby(sp[mask]).mean()
                for p in range(1, 49):
                    by_sp[(is_wknd, p)] = level * float(grp.get(p, 1.0))
            shapes[label] = by_sp
            src = ("forward curve" if fwd_level is not None
                   else f"{IC_SHAPE_DAYS}d ENTSO-E trailing mean")
            sources[label] = src
            print(f"    level £{level:.2f}/MWh ({src}), "
                  f"{IC_SHAPE_DAYS}d shape, {len(hh)} half-hours")
        except Exception as e:
            # Scrub the token in case a lower layer put the URL in the message
            msg = str(e).replace(token, "***")
            print(f"    {label} failed ({msg}); "
                  f"using flat £{IC_FALLBACK_PRICE[label]:.2f}/MWh")
            shapes[label] = flat[label]
            sources[label] = "flat fallback"
    return shapes, sources


def fetch_neso(resource_id: str, limit: int = 5000) -> pd.DataFrame:
    """Fetch all records from a NESO CKAN datastore resource."""
    print(f"  Fetching resource {resource_id[:8]}...")
    params = {"resource_id": resource_id, "limit": limit}
    r = requests.get(NESO_API, params=params, timeout=30)
    r.raise_for_status()
    data = r.json()
    records = data["result"]["records"]
    # If more records exist, paginate
    total = data["result"]["total"]
    offset = limit
    while offset < total:
        params["offset"] = offset
        r = requests.get(NESO_API, params=params, timeout=30)
        records.extend(r.json()["result"]["records"])
        offset += limit
    return pd.DataFrame(records)


def fetch_fou2t14d(fuel_types, start: date, end: date) -> pd.Series:
    """
    Fetch FOU2T14D (Forecast Availability 2-14 Days Ahead) from Elexon Insights API.
    FOU2T14D is a DAILY dataset — one outputUsable figure per fuel type per day.
    fuel_types: str or list of str — multiple types are fetched and summed per day
    (e.g. France = INTFR + INTIFA2 + INTELEC, Ireland = INTIRL + INTEW + INTGRNL).
    Returns a date-indexed Series of outputUsable_mw (daily totals).
    Falls back to empty Series on any error — callers broadcast across all SPs in the day.
    """
    if isinstance(fuel_types, str):
        fuel_types = [fuel_types]
    frames = []
    for fuel_type in fuel_types:
        try:
            print(f"  Fetching FOU2T14D {fuel_type}...")
            r = requests.get(
                f"{ELEXON_API}/datasets/FOU2T14D",
                params={
                    "settlementDateFrom": start.isoformat(),
                    "settlementDateTo":   end.isoformat(),
                    "fuelType":           fuel_type,
                    "format":             "json",
                },
                timeout=30,
            )
            r.raise_for_status()
            data = r.json().get("data", [])
            if not data:
                print(f"  FOU2T14D {fuel_type}: no data returned")
                continue
            df = pd.DataFrame(data)
            # The query params are settlementDateFrom/To, but the response
            # dates the forecast with forecastDate.
            df["date"] = pd.to_datetime(df["forecastDate"]).dt.date
            df["outputUsable_mw"] = pd.to_numeric(df["outputUsable"], errors="coerce")
            # Keep most recent publication per day
            df = (df.sort_values("publishTime")
                    .groupby("date", as_index=False)
                    .last())
            df = df[(df["date"] >= start) & (df["date"] <= end)]
            print(f"  FOU2T14D {fuel_type}: {len(df)} days")
            frames.append(df.set_index("date")["outputUsable_mw"])
        except Exception as e:
            print(f"  FOU2T14D {fuel_type} failed: {e}")
    if not frames:
        return pd.Series(dtype=float)
    # Sum across fuel types (e.g. INTFR + INTIFA2 + INTELEC for France)
    combined = pd.concat(frames, axis=1).sum(axis=1)
    combined.index = pd.Index([d for d in combined.index])
    return combined


def load_forecasts() -> pd.DataFrame:
    """
    Pull wind, embedded and demand forecasts from NESO APIs.
    Returns DataFrame indexed by (date, settlement_period) with:
        wind_gw, solar_gw, demand_gw
    """
    print("Pulling data from NESO APIs...")

    start = date.today() + timedelta(days=MODEL_START_DAY)
    end   = date.today() + timedelta(days=MODEL_END_DAY)

    # --- Transmission wind ---
    wind_raw = fetch_neso(WIND_14DA_RESOURCE)
    wind_raw["date"] = pd.to_datetime(wind_raw["Date"]).dt.date
    wind_raw["sp"]   = wind_raw["Settlement_Period"].astype(int)
    wind = wind_raw[
        (wind_raw["date"] >= start) & (wind_raw["date"] <= end)
    ][["date", "sp", "Wind_Forecast"]].copy()
    wind["Wind_Forecast"] = pd.to_numeric(wind["Wind_Forecast"])

    # --- Embedded wind + solar ---
    emb_raw = fetch_neso(EMBEDDED_RESOURCE)
    emb_raw["date"] = pd.to_datetime(emb_raw["SETTLEMENT_DATE"]).dt.date
    emb_raw["sp"]   = emb_raw["SETTLEMENT_PERIOD"].astype(int)
    emb = emb_raw[
        (emb_raw["date"] >= start) & (emb_raw["date"] <= end)
    ][["date", "sp", "EMBEDDED_WIND_FORECAST", "EMBEDDED_SOLAR_FORECAST"]].copy()
    emb["EMBEDDED_WIND_FORECAST"]  = pd.to_numeric(emb["EMBEDDED_WIND_FORECAST"])
    emb["EMBEDDED_SOLAR_FORECAST"] = pd.to_numeric(emb["EMBEDDED_SOLAR_FORECAST"])

    # --- 2-14 day demand ---
    # This dataset has no settlement period column: it carries the local
    # settlement date (DATE) plus a period-ending clock time (CTIME) as an
    # HHMM integer — 30 (00:30, SP 1) through 2400 (24:00, SP 48).
    dem_raw = fetch_neso(DEMAND_2_14_RESOURCE)
    dem_raw["date"] = pd.to_datetime(dem_raw["DATE"]).dt.date
    ctime = pd.to_numeric(dem_raw["CTIME"])
    dem_raw["sp"] = (((ctime // 100) * 60 + (ctime % 100)) // 30).astype(int)
    dem_raw["demand_mw"] = pd.to_numeric(dem_raw["NATIONALDEMAND"])
    dem = dem_raw[
        (dem_raw["date"] >= start) & (dem_raw["date"] <= end)
    ][["date", "sp", "demand_mw"]].copy()

    # --- Merge ---
    df = wind.merge(emb,  on=["date", "sp"], how="inner")
    df = df.merge(dem, on=["date", "sp"], how="inner")

    # Gross basis: embedded wind and solar appear on the supply side, so
    # demand must be gross of them too. NESO NATIONALDEMAND is transmission
    # system demand — embedded generation is already netted off it — so add
    # the embedded forecasts back to avoid counting them twice.
    df["wind_gw"]   = (df["Wind_Forecast"] + df["EMBEDDED_WIND_FORECAST"]) / 1000
    df["solar_gw"]  = df["EMBEDDED_SOLAR_FORECAST"] / 1000
    df["demand_gw"] = (df["demand_mw"]
                       + df["EMBEDDED_WIND_FORECAST"]
                       + df["EMBEDDED_SOLAR_FORECAST"]) / 1000

    df = df.sort_values(["date", "sp"]).set_index(["date", "sp"])
    print(f"  Loaded {len(df)} settlement periods: "
          f"{df.index.get_level_values('date').min()} to "
          f"{df.index.get_level_values('date').max()}")

    # --- Nuclear availability (FOU2T14D) ---
    nuc = fetch_fou2t14d("NUCLEAR", start, end)
    if not nuc.empty:
        # Broadcast daily value across all 48 SPs in each day
        nuc_gw = nuc / 1000
        df["nuclear_gw"] = df.index.get_level_values("date").map(nuc_gw)
        # Fill any gaps with installed capacity
        df["nuclear_gw"] = df["nuclear_gw"].fillna(NUCLEAR_GW)
    else:
        print(f"  Nuclear: using fixed {NUCLEAR_GW} GW (no FOU2T14D data)")
        df["nuclear_gw"] = NUCLEAR_GW

    # --- Biomass availability (FOU2T14D) ---
    bio = fetch_fou2t14d("BIOMASS", start, end)
    if not bio.empty:
        bio_gw = bio / 1000
        df["biomass_gw"] = df.index.get_level_values("date").map(bio_gw)
        df["biomass_gw"] = df["biomass_gw"].fillna(BIOMASS_GW)
    else:
        print(f"  Biomass: using fixed {BIOMASS_GW} GW (no FOU2T14D data)")
        df["biomass_gw"] = BIOMASS_GW


    # --- Hydro availability (FOU2T14D: NPSHYD + PS) ---
    hydro = fetch_fou2t14d(["NPSHYD", "PS"], start, end)
    if not hydro.empty:
        hydro_gw = hydro / 1000
        df["hydro_gw"] = df.index.get_level_values("date").map(hydro_gw)
        df["hydro_gw"] = df["hydro_gw"].fillna(HYDRO_AVG_GW)
    else:
        print(f"  Hydro: using fixed {HYDRO_AVG_GW} GW (no FOU2T14D data)")
        df["hydro_gw"] = HYDRO_AVG_GW
    # --- Interconnector availability (FOU2T14D) ---
    print("Fetching interconnector availability from Elexon FOU2T14D...")
    ic_cols = []
    for cap_gw, label, fuel_codes, _ in INTERCONNECTORS:
        col = f"ic_{label.lower().replace(' ', '_')}"
        ic_cols.append(col)
        ic = fetch_fou2t14d(fuel_codes, start, end)
        if not ic.empty:
            ic_gw = ic / 1000
            df[col] = df.index.get_level_values("date").map(ic_gw)
            df[col] = df[col].fillna(cap_gw)
        else:
            print(f"  {label}: using fixed {cap_gw} GW (no FOU2T14D data)")
            df[col] = cap_gw

    return df[["wind_gw", "solar_gw", "demand_gw",
               "hydro_gw", "nuclear_gw", "biomass_gw"] + ic_cols]


# ===========================================================================
# MARGINAL COST & MERIT ORDER
# ===========================================================================

def gas_mc(efficiency, vom, gas_p_therm=None, ets_price=None):
    """
    Short run marginal cost of a gas tranche, £/MWh.

    gas_p_therm and ets_price default to the hardcoded assumptions; the model
    passes the day's forward values when forward_levels.json supplies them.
    Both are flat within a day — a gas day is one product and carbon has no
    intraday dimension — so this is constant across settlement periods.
    """
    gas = GAS_PRICE_P_THERM    if gas_p_therm is None else gas_p_therm
    ets_p = ETS_PRICE_GBP_TONNE if ets_price   is None else ets_price
    gbp_per_mwh_th = (gas / 100) / 0.02931
    fuel   = gbp_per_mwh_th / efficiency
    ets    = (ETS_EMISSIONS / efficiency) * ets_p
    cps    = (CPS_EMISSIONS / efficiency) * CPS_RATE_GBP_TONNE
    return fuel + ets + cps + vom


def build_stack(ic_prices: dict = None, gas_p_therm=None, ets_price=None):
    """
    Merit order for one settlement period.

    ic_prices maps interconnector label -> import price for this period. The
    stack must be rebuilt per period because continental prices are shaped
    across the day, which moves each link's position in the merit order.
    gas_p_therm and ets_price set the gas stack for the day.
    """
    stack = [
        ("Wind",    None,                              0.00,          "renewable"),
        ("Solar",   None,                              0.50,          "renewable"),
        ("Hydro",   None, HYDRO_VOM,  "hydro"),  # capacity per-SP from FOU2T14D
        ("Nuclear", None, NUCLEAR_VOM,  "nuclear"),  # capacity set per-SP from FOU2T14D
        ("Biomass", None, BIOMASS_VOM,  "biomass"),   # capacity set per-SP from FOU2T14D
    ]
    for cap_gw, label, _, _ in INTERCONNECTORS:
        mc = (ic_prices or {}).get(label, IC_FALLBACK_PRICE[label])
        stack.append((label, None, mc, "import"))  # capacity per-SP from FOU2T14D
    for i, (eff, cap_gw, vom) in enumerate(GAS_TRANCHES, 1):
        mc = gas_mc(eff, vom, gas_p_therm, ets_price)
        label = f"CCGT T{i} ({eff:.0%})" if eff >= 0.42 else f"OCGT T{i} ({eff:.0%})"
        stack.append((label, cap_gw, mc, "gas"))
    return sorted(stack, key=lambda x: x[2])


# ===========================================================================
# DISPATCH
# ===========================================================================

def dispatch(demand_gw, stack, wind_gw, solar_gw,
             hydro_gw, nuclear_gw, biomass_gw, ic_avail: dict):
    period_output = {
        "Wind":    wind_gw,
        "Solar":   solar_gw,
        "Hydro":   hydro_gw,
        "Nuclear": nuclear_gw,
        "Biomass": biomass_gw,
    }
    period_output.update(ic_avail)  # per-SP interconnector availability
    remaining = demand_gw
    marginal  = None
    for name, cap, mc, ptype in stack:
        if remaining <= 0:
            break
        avail = period_output.get(name, cap)
        gen   = min(avail, remaining)
        if gen > 0:
            remaining -= gen
            marginal   = (name, ptype, mc)
    price = marginal[2] if marginal else 0.0
    return marginal[0] if marginal else "Unserved", round(price, 2)


def run_model(forecasts, ic_shapes, gas_curve, ets_price):
    records = []
    for (dt, sp), row in forecasts.iterrows():
        gas_today = float(gas_curve.get(dt, GAS_PRICE_P_THERM))
        ts = datetime.combine(dt, datetime.min.time()) + timedelta(minutes=30*(sp-1))
        ic_avail = {
            label: row[f"ic_{label.lower().replace(' ', '_')}"]
            for _, label, _, _ in INTERCONNECTORS
        }
        key = (dt.weekday() >= 5, int(sp))
        ic_prices = {label: ic_shapes[label][key]
                     for _, label, _, _ in INTERCONNECTORS}
        stack = build_stack(ic_prices, gas_today, ets_price)
        plant, price = dispatch(
            row["demand_gw"], stack,
            row["wind_gw"], row["solar_gw"],
            row["hydro_gw"], row["nuclear_gw"], row["biomass_gw"],
            ic_avail,
        )
        records.append({
            "Datetime":          ts,
            "Date":              dt,
            "Settlement Period": sp,
            "Demand (GW)":       round(row["demand_gw"],  2),
            "Wind (GW)":         round(row["wind_gw"],    2),
            "Solar (GW)":        round(row["solar_gw"],   2),
            "Hydro (GW)":        round(row["hydro_gw"],   2),
            "Nuclear (GW)":      round(row["nuclear_gw"], 2),
            "Biomass (GW)":      round(row["biomass_gw"], 2),
            "Gas (p/therm)":     round(gas_today, 2),
            "Marginal Plant":    plant,
            "Price (£/MWh)":     price,
        })
    return pd.DataFrame(records)


# ===========================================================================
# EXCEL OUTPUT
# ===========================================================================

# Style constants
DARK_BLUE  = "1F3864"
MID_BLUE   = "2E75B6"
LIGHT_BLUE = "BDD7EE"
PALE_BLUE  = "DEEAF1"
WHITE      = "FFFFFF"
YELLOW     = "FFFF00"
GREEN      = "00B050"
AMBER      = "FFC000"
RED        = "FF0000"
GREY       = "F2F2F2"

def hdr_font(bold=True, white=True, size=10):
    return Font(name="Arial", bold=bold, size=size,
                color=WHITE if white else "000000")

def cell_font(bold=False, size=10, color="000000"):
    return Font(name="Arial", bold=bold, size=size, color=color)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def thin_border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=False)

def fmt_cell(ws, row, col, value, font=None, fill_=None, number_format=None,
             border=None, alignment=None):
    c = ws.cell(row=row, column=col, value=value)
    if font:        c.font        = font
    if fill_:       c.fill        = fill_
    if number_format: c.number_format = number_format
    if border:      c.border      = border
    if alignment:   c.alignment   = alignment
    return c


def write_assumptions_sheet(wb, ic_shapes, ic_level_source, gas_curve, ets_price):
    ws = wb.create_sheet("Assumptions")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 40

    # Title
    ws.merge_cells("A1:C1")
    c = ws["A1"]
    c.value = "GB Power Market Model — Assumptions"
    c.font  = Font(name="Arial", bold=True, size=14, color=WHITE)
    c.fill  = fill(DARK_BLUE)
    c.alignment = Alignment(horizontal="left", vertical="center",
                             indent=1)
    ws.row_dimensions[1].height = 28

    def section(row, title):
        ws.merge_cells(f"A{row}:C{row}")
        c = ws.cell(row=row, column=1, value=title)
        c.font  = Font(name="Arial", bold=True, size=10, color=WHITE)
        c.fill  = fill(MID_BLUE)
        c.alignment = Alignment(indent=1, vertical="center")
        ws.row_dimensions[row].height = 18

    def row_(ws, r, label, value, unit, source="", blue_val=True):
        ws.cell(r, 1, label).font = cell_font()
        vc = ws.cell(r, 2, value)
        vc.font = Font(name="Arial", size=10,
                       color="0000FF" if blue_val else "000000")
        vc.alignment = Alignment(horizontal="center")
        ws.cell(r, 3, unit if not source else f"{unit}   |   {source}").font = \
            Font(name="Arial", size=9, color="595959")
        for col in range(1, 4):
            ws.cell(r, col).border = thin_border()
            ws.cell(r, col).fill   = fill(PALE_BLUE if r % 2 == 0 else WHITE)

    r = 3
    section(r, "Gas & Carbon Prices"); r += 1
    _g = gas_curve if gas_curve is not None else pd.Series(dtype=float)
    if len(_g) and _g.nunique() > 1:
        row_(ws, r, "Gas price (NBP)", f"{_g.mean():.2f} avg",
             f"p/therm  |  daily curve {_g.min():.2f}-{_g.max():.2f} "
             f"over {len(_g)} days, from {FORWARD_LEVELS_FILE}"); r += 1
    else:
        _lvl = float(_g.iloc[0]) if len(_g) else GAS_PRICE_P_THERM
        _src = (f"from {FORWARD_LEVELS_FILE}" if _lvl != GAS_PRICE_P_THERM
                else "hardcoded assumption")
        row_(ws, r, "Gas price (NBP)", _lvl, f"p/therm  |  flat, {_src}"); r += 1
    row_(ws, r, "UK ETS carbon price", ets_price,
         "£/tonne CO2  |  no intraday or daily shape — annual compliance "
         "instrument"); r += 1
    row_(ws, r, "Carbon Price Support",   CPS_RATE_GBP_TONNE,  "£/tonne CO2  |  Frozen at £18 since April 2016"); r += 1

    r += 1
    section(r, "Emissions Factors (Gross CV)"); r += 1
    row_(ws, r, "ETS emissions factor", ETS_EMISSIONS,
         "tCO2/MWh_th",
         "UKGHGI 2025, median of 15 LDZ regions, CEF(g) col", False); r += 1
    row_(ws, r, "CPS emissions factor", CPS_EMISSIONS,
         "tCO2/MWh_th",
         "BEIS GHG Conversion Factors 2025, natural gas gross CV, CO2 only", False); r += 1

    r += 1
    section(r, f"Gas Fleet — {len(GAS_TRANCHES)} Efficiency Tranches"); r += 1
    _gas_avg = float(_g.mean()) if len(_g) else GAS_PRICE_P_THERM
    row_(ws, r, "Source", "NESO 2026/27 Final ALF data, Argus efficiency bands",
         f"SRMC shown at the window-average {_gas_avg:.2f} p/therm", "", False); r += 1
    for i, (eff, cap, vom) in enumerate(GAS_TRANCHES, 1):
        mc   = gas_mc(eff, vom, _gas_avg, ets_price)
        kind = "CCGT" if eff >= 0.42 else "OCGT"
        row_(ws, r,
             f"  T{i} — {kind} {eff:.0%} efficiency",
             f"{cap:.1f} GW",
             f"SRMC: £{mc:.2f}/MWh   |   VOM: £{vom:.2f}/MWh", "", False); r += 1

    r += 1
    section(r, "Other Plant"); r += 1
    row_(ws, r, "Nuclear", f"{NUCLEAR_GW} GW installed",
         "Availability per SP from Elexon FOU2T14D — falls back to installed capacity", "", False); r += 1
    row_(ws, r, "Biomass", f"{BIOMASS_GW} GW installed (Drax + Lynemouth)",
         "Availability per SP from Elexon FOU2T14D — CfD/RO supported, near-zero bid", "", False); r += 1
    row_(ws, r, "Hydro (NPSHYD+PS)", f"{HYDRO_AVG_GW} GW registered",
         "Availability per SP from Elexon FOU2T14D (NPSHYD + PS)", "", False); r += 1
    for cap_gw, label, fuel_codes, eic in INTERCONNECTORS:
        prices = list((ic_shapes or {}).get(label, {}).values())
        if prices:
            lvl  = sum(prices) / len(prices)
            src  = (ic_level_source or {}).get(label, "trailing mean")
            note = (f"Import £{lvl:.2f}/MWh mean, "
                    f"£{min(prices):.2f}–£{max(prices):.2f} shaped by SP  |  "
                    f"level: {src}  |  {eic}")
        else:
            note = f"Import price: £{IC_FALLBACK_PRICE[label]:.2f}/MWh flat"
        row_(ws, r, label, f"{cap_gw} GW registered", note, "", False); r += 1

    r += 1
    section(r, "Variable O&M"); r += 1
    row_(ws, r, "Gas (CCGT & OCGT)", CCGT_VOM,
         "£/MWh  |  DESNZ 2025 generation costs, all gas plant (2024 prices)"); r += 1
    row_(ws, r, "Nuclear", NUCLEAR_VOM,
         "£/MWh  |  ASSUMPTION — no published source, never sets the price"); r += 1
    row_(ws, r, "Hydro", HYDRO_VOM,
         "£/MWh  |  ASSUMPTION — no published source, never sets the price"); r += 1
    row_(ws, r, "Biomass", BIOMASS_VOM,
         "£/MWh  |  Near-zero bid due to CfD/RO support"); r += 1

    r += 1
    section(r, "Price Formation"); r += 1
    row_(ws, r, "Output price", "SRMC of marginal plant",
         "Wholesale market clearing price", "", False); r += 1
    row_(ws, r, "System costs adder", "Not applied",
         "BSUoS/TNUoS/DUoS are non-commodity charges recovered separately "
         "and are not part of wholesale price formation", "", False); r += 1
    row_(ws, r, "Demand basis", "Gross of embedded generation",
         "NESO NATIONALDEMAND is transmission demand (embedded generation "
         "already netted off); embedded wind and solar are added back so "
         "they are not double-counted against the supply stack", "", False); r += 1

    r += 1
    section(r, "Data Sources"); r += 1
    row_(ws, r, "Transmission wind forecast",
         "NESO Data Portal", "14 day ahead, half-hourly", "", False); r += 1
    row_(ws, r, "Embedded wind & solar",
         "NESO Data Portal", "2-14 day ahead, half-hourly", "", False); r += 1
    row_(ws, r, "Demand forecast",
         "NESO Data Portal", "2-14 days ahead half hourly (NATIONALDEMAND)", "", False); r += 1
    row_(ws, r, "Plant & interconnector availability",
         "Elexon FOU2T14D", "2-14 day ahead, daily outputUsable", "", False); r += 1
    row_(ws, r, "Interconnector import prices",
         "ENTSO-E Transparency", f"Day-ahead, trailing {IC_SHAPE_DAYS}d, "
         "shaped by settlement period (weekday/weekend)", "", False); r += 1
    row_(ws, r, "EUR/GBP", "ECB daily reference rate",
         "Applied to all continental day-ahead prices", "", False); r += 1


def write_results_sheet(wb, df):
    ws = wb.create_sheet("Half-Hourly Results")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    cols = ["Datetime", "Date", "Settlement Period", "Demand (GW)",
            "Wind (GW)", "Solar (GW)", "Hydro (GW)", "Nuclear (GW)",
            "Biomass (GW)", "Marginal Plant", "Price (£/MWh)"]
    widths = [20, 12, 8, 12, 10, 10, 10, 11, 12, 22, 14]

    for i, (col, w) in enumerate(zip(cols, widths), 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Title row
    ws.merge_cells(f"A1:{get_column_letter(len(cols))}1")
    c = ws["A1"]
    c.value = "GB Power Market Model — Half-Hourly Results"
    c.font  = Font(name="Arial", bold=True, size=12, color=WHITE)
    c.fill  = fill(DARK_BLUE)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 24

    # Header row
    for i, col in enumerate(cols, 1):
        c = ws.cell(row=2, column=i, value=col)
        c.font      = Font(name="Arial", bold=True, size=10, color=WHITE)
        c.fill      = fill(MID_BLUE)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = thin_border()
    ws.row_dimensions[2].height = 16

    # Plant colour map
    plant_colours = {
        "Wind":          "92D050",   # green
        "Solar":         "FFC000",   # amber
        "Hydro":         "00B0F0",   # light blue
        "Nuclear":       "7030A0",   # purple
        "Biomass":           "C65911",  # dark orange
        "Interconnector NO": "1F77B4",
        "Interconnector DK": "17BECF",
        "Interconnector FR": "2CA02C",
        "Interconnector BE": "9467BD",
        "Interconnector NL": "8C564B",
        "Interconnector IE": "E377C2",
    }

    for r_idx, (_, row) in enumerate(df.iterrows(), 3):
        row_fill = WHITE if r_idx % 2 == 1 else PALE_BLUE
        plant = str(row["Marginal Plant"])

        # Determine plant colour for price cell — gas tranches shade from
        # cheapest (green) to most expensive (dark red) by tranche number
        if "CCGT" in plant or "OCGT" in plant:
            m = re.search(r"T(\d+)", plant)
            t = int(m.group(1)) if m else 3
            colours_gas = ["00B050","70AD47","A9D18E","FFC000","FF9900",
                           "FF7C00","FF0000","C00000"]
            pc = colours_gas[min(t-1, len(colours_gas)-1)]
        else:
            pc = plant_colours.get(plant, "BFBFBF")

        values = [
            row["Datetime"].strftime("%Y-%m-%d %H:%M"),
            str(row["Date"]),
            row["Settlement Period"],
            row["Demand (GW)"],
            row["Wind (GW)"],
            row["Solar (GW)"],
            row["Hydro (GW)"],
            row["Nuclear (GW)"],
            row["Biomass (GW)"],
            plant,
            row["Price (£/MWh)"],
        ]
        fmts = [None, None, "0", "0.00", "0.00", "0.00", "0.00", "0.00",
                "0.00", None, "£#,##0.00"]

        for c_idx, (val, fmt) in enumerate(zip(values, fmts), 1):
            c = ws.cell(row=r_idx, column=c_idx, value=val)
            c.font   = cell_font(size=9)
            c.fill   = fill(row_fill)
            c.border = thin_border()
            c.alignment = Alignment(horizontal="center" if c_idx > 1 else "left",
                                    vertical="center")
            if fmt:
                c.number_format = fmt

        # Colour price cell by marginal plant
        pc_cell = ws.cell(row=r_idx, column=11)
        pc_cell.fill = fill(pc)
        pc_cell.font = Font(name="Arial", size=9, bold=True, color=WHITE)


def write_daily_summary_sheet(wb, df):
    ws = wb.create_sheet("Daily Summary")
    ws.sheet_view.showGridLines = False

    daily = df.groupby("Date").agg(
        Avg_Price   = ("Price (£/MWh)", "mean"),
        Min_Price   = ("Price (£/MWh)", "min"),
        Max_Price   = ("Price (£/MWh)", "max"),
        Avg_Demand  = ("Demand (GW)",   "mean"),
        Avg_Wind    = ("Wind (GW)",     "mean"),
        Avg_Solar   = ("Solar (GW)",    "mean"),
        Avg_Hydro   = ("Hydro (GW)",    "mean"),
        Avg_Nuclear = ("Nuclear (GW)",  "mean"),
        Avg_Biomass = ("Biomass (GW)",  "mean"),
        Gas_P_Therm = ("Gas (p/therm)", "mean"),
    ).reset_index()
    daily["Price_Range"] = daily["Max_Price"] - daily["Min_Price"]
    _peak = df.loc[df.groupby("Date")["Price (£/MWh)"].idxmax(),
                   ["Date","Settlement Period"]].set_index("Date")
    daily["Peak_Period"] = daily["Date"].map(_peak["Settlement Period"])

    ws.column_dimensions["A"].width = 14
    for col, w in zip("BCDEFGHIJ", [12,10,10,12,11,11,12,12,11]):
        ws.column_dimensions[col].width = w

    # Title
    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value = "GB Power Market Model — Daily Summary"
    c.font  = Font(name="Arial", bold=True, size=12, color=WHITE)
    c.fill  = fill(DARK_BLUE)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 24

    headers = ["Date", "Avg Price\n(£/MWh)", "Min Price\n(£/MWh)",
               "Max Price\n(£/MWh)", "Price Range\n(£/MWh)",
               "Avg Demand\n(GW)", "Avg Wind\n(GW)", "Avg Solar\n(GW)",
               "Peak SP", "NBP\n(p/therm)"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=i, value=h)
        c.font      = Font(name="Arial", bold=True, size=10, color=WHITE)
        c.fill      = fill(MID_BLUE)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                 wrap_text=True)
        c.border    = thin_border()
    ws.row_dimensions[2].height = 30

    for r_idx, row in enumerate(daily.itertuples(), 3):
        row_fill = WHITE if r_idx % 2 == 1 else PALE_BLUE
        vals = [str(row.Date), row.Avg_Price, row.Min_Price, row.Max_Price,
                row.Price_Range, row.Avg_Demand, row.Avg_Wind,
                row.Avg_Solar, row.Peak_Period, row.Gas_P_Therm]
        fmts = [None, "£#,##0.00","£#,##0.00","£#,##0.00",
                "£#,##0.00","0.0","0.0","0.0","0","0.00"]
        for c_idx, (v, f) in enumerate(zip(vals, fmts), 1):
            c = ws.cell(row=r_idx, column=c_idx, value=v)
            c.font   = cell_font(size=10)
            c.fill   = fill(row_fill)
            c.border = thin_border()
            c.alignment = Alignment(horizontal="center", vertical="center")
            if f: c.number_format = f

        # Colour avg price cell by level
        pc = ws.cell(row=r_idx, column=2)
        price = row.Avg_Price
        if   price < 60:  col = "92D050"
        elif price < 75:  col = "FFC000"
        elif price < 90:  col = "FF7C00"
        else:             col = "FF0000"
        pc.fill = fill(col)
        pc.font = Font(name="Arial", size=10, bold=True, color=WHITE)

    return daily


def write_charts_sheet(wb, df, daily):
    ws = wb.create_sheet("Charts")
    ws.sheet_view.showGridLines = False

    # We need a helper data sheet for charts
    data_ws = wb.create_sheet("_ChartData")
    data_ws.sheet_view.showGridLines = False

    # Write daily data for chart
    data_ws["A1"] = "Date"
    data_ws["B1"] = "Avg Price"
    data_ws["C1"] = "Min Price"
    data_ws["D1"] = "Max Price"
    data_ws["E1"] = "Avg Wind"
    data_ws["F1"] = "Avg Demand"

    for i, row in enumerate(daily.itertuples(), 2):
        data_ws.cell(i, 1, str(row.Date))
        data_ws.cell(i, 2, round(row.Avg_Price,  2))
        data_ws.cell(i, 3, round(row.Min_Price,  2))
        data_ws.cell(i, 4, round(row.Max_Price,  2))
        data_ws.cell(i, 5, round(row.Avg_Wind,   2))
        data_ws.cell(i, 6, round(row.Avg_Demand, 2))

    n = len(daily) + 1

    # --- Chart 1: Daily price range ---
    chart1 = LineChart()
    chart1.title    = "Daily Power Price (£/MWh)"
    chart1.style    = 10
    chart1.y_axis.title = "£/MWh"
    chart1.x_axis.title = "Date"
    chart1.width    = 18
    chart1.height   = 12

    cats = Reference(data_ws, min_col=1, min_row=2, max_row=n)

    for col,title in [(2,"Avg Price"),(3,"Min Price"),(4,"Max Price")]:
        data = Reference(data_ws, min_col=col, min_row=1, max_row=n)
        chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    for i,colour in enumerate(["2E75B6","70AD47","FF0000"]):
        chart1.series[i].graphicalProperties.line.solidFill = colour
        chart1.series[i].graphicalProperties.line.width     = 18000
    ws.add_chart(chart1, "A2")

    chart2 = LineChart()
    chart2.title="Daily Average Wind & Demand (GW)"; chart2.style=10
    chart2.y_axis.title="GW"; chart2.x_axis.title="Date"
    chart2.width=18; chart2.height=12
    for col,title in [(5,"Avg Wind"),(6,"Avg Demand")]:
        data = Reference(data_ws, min_col=col, min_row=1, max_row=n)
        chart2.add_data(data, titles_from_data=True)
    chart2.set_categories(cats)
    for i,colour in enumerate(["92D050","2E75B6"]):
        chart2.series[i].graphicalProperties.line.solidFill = colour
        chart2.series[i].graphicalProperties.line.width     = 18000
    ws.add_chart(chart2, "A24")

    # --- Chart 3: Hydro / Nuclear / Biomass availability ---
    data_ws["G1"] = "Avg Hydro"
    data_ws["H1"] = "Avg Nuclear"
    data_ws["I1"] = "Avg Biomass"
    for _ci, _row in enumerate(daily.itertuples(), 2):
        data_ws.cell(_ci, 7, round(_row.Avg_Hydro,   2))
        data_ws.cell(_ci, 8, round(_row.Avg_Nuclear, 2))
        data_ws.cell(_ci, 9, round(_row.Avg_Biomass, 2))
    chart3 = LineChart()
    chart3.title = "Daily Avg Dispatchable Generation (GW)"
    chart3.style = 10
    chart3.y_axis.title = "GW"
    chart3.x_axis.title = "Date"
    chart3.width = 18
    chart3.height = 12
    for col in [7, 8, 9]:
        d = Reference(data_ws, min_col=col, min_row=1, max_row=n)
        chart3.add_data(d, titles_from_data=True)
    chart3.set_categories(cats)
    for ci, colour in enumerate(["00B0F0", "7030A0", "C65911"]):
        chart3.series[ci].graphicalProperties.line.solidFill = colour
        chart3.series[ci].graphicalProperties.line.width     = 18000
    ws.add_chart(chart3, "A46")

    # Title
    ws.merge_cells("A1:O1")
    c = ws["A1"]
    c.value = "GB Power Market Model — Charts"
    c.font  = Font(name="Arial", bold=True, size=12, color=WHITE)
    c.fill  = fill(DARK_BLUE)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 24


def write_excel(df, ic_shapes, ic_level_source, gas_curve, ets_price, run_time):
    fname = f"gb_power_model_{run_time.strftime('%Y%m%d_%H%M')}.xlsx"

    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    print("Writing Assumptions sheet...")
    write_assumptions_sheet(wb, ic_shapes, ic_level_source, gas_curve, ets_price)

    print("Writing Daily Summary sheet...")
    daily = write_daily_summary_sheet(wb, df)

    print("Writing Half-Hourly Results sheet...")
    write_results_sheet(wb, df)

    print("Writing Charts sheet...")
    write_charts_sheet(wb, df, daily)

    # Set tab colours
    wb["Assumptions"].sheet_properties.tabColor          = DARK_BLUE
    wb["Daily Summary"].sheet_properties.tabColor        = MID_BLUE
    wb["Half-Hourly Results"].sheet_properties.tabColor  = "4472C4"
    wb["Charts"].sheet_properties.tabColor               = "70AD47"
    wb["_ChartData"].sheet_properties.tabColor           = "BFBFBF"

    # Set active sheet to Daily Summary
    wb.active = wb["Daily Summary"]

    wb.save(fname)
    print(f"\nSaved: {fname}")
    return fname


# ===========================================================================
# MAIN
# ===========================================================================

def model_dates() -> list:
    """Every date the model will cover, given today's rolling window."""
    start = date.today() + timedelta(days=MODEL_START_DAY)
    return [start + timedelta(days=i)
            for i in range(MODEL_END_DAY - MODEL_START_DAY + 1)]


def write_forward_template(path: str = None) -> str:
    """
    Write a forward_levels.json pre-dated for today's window, values nulled.

    The window rolls forward every day, so hand-typing dates goes stale
    immediately. Generate the template, paste your assessments in, run. Nulls
    are ignored by the loader, so a partly filled template is valid.
    """
    path = path or os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                FORWARD_LEVELS_FILE)
    if os.path.exists(path):
        raise SystemExit(f"{path} already exists — delete or edit it rather "
                         f"than overwriting assessments")
    dates = model_dates()
    blob = {
        "_comment": [
            f"Generated {date.today().isoformat()} for the window "
            f"{dates[0].isoformat()} to {dates[-1].isoformat()}.",
            "The window rolls daily, so regenerate rather than reusing this "
            "file on a later day; the model warns if the curve stops "
            "overlapping the window.",
            "Nulls are ignored and fall back to the assumptions in "
            "run_model.py. See forward_levels.example.json for field notes.",
        ],
        "currency": "EUR",
        "levels": {label: None for _, label, _, _ in INTERCONNECTORS},
        "nbp_p_therm": {d.isoformat(): None for d in dates},
        "uk_ets_gbp_tonne": None,
    }
    with open(path, "w") as fh:
        json.dump(blob, fh, indent=2)
        fh.write("\n")
    return path


if __name__ == "__main__":
    if "--init-forward" in sys.argv:
        p = write_forward_template()
        ds = model_dates()
        print(f"Wrote {p}")
        print(f"  {len(ds)} gas days, {ds[0]} to {ds[-1]}")
        print("  Fill in and re-run. This file is gitignored — keep licensed "
              "assessments out of the repository.")
        raise SystemExit(0)

    run_time = datetime.now()
    print(f"GB Power Model — {run_time.strftime('%Y-%m-%d %H:%M')}")
    print(f"Forecast window: day {MODEL_START_DAY} to day {MODEL_END_DAY} ahead\n")

    # Pull data
    forecasts = load_forecasts()

    # Forward assessments (gitignored file) then interconnector price profiles
    print("\nLoading forward price inputs...")
    fx      = fetch_eur_gbp()
    forward = load_forward_inputs(fx)

    print("\nBuilding interconnector price profiles from ENTSO-E...")
    ic_shapes, ic_level_source = build_ic_price_shapes(fx, forward["ic_levels"])

    gas_curve = resolve_gas_curve(
        forward["nbp"], forecasts.index.get_level_values("date"))
    ets_price = forward["ets"] if forward["ets"] is not None \
        else ETS_PRICE_GBP_TONNE
    print(f"  Gas: {gas_curve.min():.2f}-{gas_curve.max():.2f} p/therm "
          f"across {len(gas_curve)} days   |   UK ETS: £{ets_price:.2f}/tonne")

    # Run model — the stack is rebuilt per period as import prices are shaped
    print("\nRunning dispatch model...")
    df = run_model(forecasts, ic_shapes, gas_curve, ets_price)

    # Print console summary
    print("\n=== Summary ===")
    print(f"Periods modelled : {len(df)}")
    print(f"Avg price        : £{df['Price (£/MWh)'].mean():.2f}/MWh")
    print(f"Min price        : £{df['Price (£/MWh)'].min():.2f}/MWh")
    print(f"Max price        : £{df['Price (£/MWh)'].max():.2f}/MWh")
    print()
    for plant, count in df["Marginal Plant"].value_counts().items():
        print(f"  {plant:<25} {count:>4} periods ({count/len(df)*100:.1f}%)")

    # Write Excel
    print("\nWriting Excel output...")
    fname = write_excel(df, ic_shapes, ic_level_source,
                        gas_curve, ets_price, run_time)
    print(f"\nDone. Open: {fname}")
