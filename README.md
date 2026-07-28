# GB Power Dispatch Model

A merit-order model of the GB wholesale power market. It pulls live demand,
wind and plant availability forecasts, clears each half-hour against the
domestic stack and the interconnectors, and writes a formatted Excel workbook
with half-hourly marginal plant and clearing prices 3 to 14 days ahead.


```
python run_model.py
```

Output is `gb_power_model_YYYYMMDD_HHMM.xlsx` — assumptions, daily summary,
half-hourly results and charts.

## What it does

For every settlement period the model builds a supply stack and finds the
price at which supply meets demand.

**Gas** is split into eight tranches by thermal efficiency, from 60% down to
34%, with capacities mapped from NESO's 2026/27 Final Annual Load Factor data.
Each tranche gets a short-run marginal cost:

```
SRMC = fuel/efficiency + (ETS_factor/efficiency x UKA)
                       + (CPS_factor/efficiency x CPS) + VOM
```

UK ETS and Carbon Price Support are applied separately, with their own
emissions factors — UKGHGI 2025 for ETS, the BEIS condensed set for CPS —
because they are different instruments on different bases. CPS is frozen at
£18/tonne. At 80p/therm the carbon components add roughly £19/MWh to a 60%
efficient CCGT.

**Nuclear, biomass, hydro and interconnector availability** come from Elexon's
FOU2T14D forecast dataset, per fuel type per day, with the most recent
publication taken and installed capacity used as a fallback. France and
Ireland are summed across their several BMU fuel codes (IFA1, IFA2 and
ElecLink; Moyle, EWIC and Greenlink).

**Wind, solar and demand** come from NESO's 2-14 day ahead forecasts, with
transmission and embedded wind added together.

Output is the wholesale clearing price — the SRMC of the marginal plant.
BSUoS, TNUoS and DUoS are non-commodity charges recovered separately and are
deliberately excluded.

## Interconnectors

A link is not a generator. Power flows toward the dearer market, so at a
candidate GB price `p` a link supplies GB when its continental price is at or
below `p`, and draws from GB when it is above. Modelling links as supply only
silently forces net imports in every period and removes the export demand that
lifts the GB price at peak.

That makes the clearing price a fixed point: exports raise the GB price, and
the GB price decides whether a link exports. The model solves it directly.
Supply is non-decreasing in price and demand is non-increasing, so
supply − demand crosses zero exactly once; the model walks the candidate
prices — the distinct domestic marginal costs plus the continental prices —
and takes the first that clears.

The continental market is treated as a price taker: flow does not move its
price, so a link runs at full capacity whenever the two prices differ, and
sits partly loaded only when it is itself marginal. That is the coupled case,
where the GB price has converged onto that link's price.

Continental prices are shaped, not flat. The model pulls 30 days of ENTSO-E
day-ahead prices for each counterparty zone, computes the mean ratio to daily
mean by settlement period, and splits weekday from weekend because continental
weekend middays sit far lower. Hourly zones such as SEM are upsampled by
holding the price across the hour. EUR zones are converted at the ECB daily
reference rate.

## Forward price inputs

By default the model uses trailing realised means for the interconnector
levels and the hardcoded gas and carbon assumptions in `run_model.py`.

To run it against a forward view instead:

```
python run_model.py --init-forward   # writes forward_levels.json for today's window
```

Fill in what you have and re-run. Partial files are fine — anything null falls
back. The file takes three things:

| Field | Shape | Notes |
|---|---|---|
| `levels` | one **baseload** price per link | Not peak. The intraday shape is applied on top, so a peak assessment would count the daily shape twice. |
| `nbp_p_therm` | single number, or a daily curve | A curve is preferred. Gas sets the GB price in most periods, so this is the most sensitive input in the model — roughly £0.62/MWh per p/therm at the marginal 55% tranche. |
| `uk_ets_gbp_tonne` | one number | Carbon is an annual compliance instrument, so no shape. |

Neither gas nor carbon takes an intraday shape: a gas day is a single product,
so a plant's fuel cost is flat across all 48 settlement periods.

`forward_levels.json` is gitignored. Forward assessments are licensed data and
must not be committed to a public repository. See
`forward_levels.example.json` for the format.

The window rolls forward daily, so regenerate the template rather than reusing
an old one — the model warns if the curve stops overlapping the window.

## Setup

```
pip install requests pandas openpyxl
export ENTSOE_API_TOKEN=...   # free from transparency.entsoe.eu
python run_model.py
```

Without a token the model still runs, but falls back to flat interconnector
import prices, which materially flattens the GB price curve. NESO and Elexon
need no authentication.

## Tests

```
python test_dispatch.py
```

Ten cases over the clearing logic: import and export, simultaneous flows in
both directions on different links, price convergence onto a marginal link,
negative continental prices, surplus renewables clearing at zero, and unserved
energy. The last case probes 3,000 random stacks to confirm supply − demand
is monotonic in price, which is the property the solver relies on.

## Known limitations

- **No backtest.** The model produces a price and does not yet score it
  against day-ahead outturn. This is the next thing to build.
- **No storage.** Batteries and pumped-storage arbitrage are not modelled;
  pumped storage enters only as available capacity.
- **No unit commitment.** No start costs, minimum stable generation or
  ramp constraints, so the model will dispatch a CCGT for a single
  half-hour where a real plant would not start.
- **No transmission constraints.** GB is treated as a single zone, so
  constraint costs and the boundary flows behind them are invisible.
- **Nuclear and hydro VOM are assumptions.** No public DESNZ source exists;
  both are flagged in the code. They rarely set the price.

## Data sources

| Source | Used for |
|---|---|
| NESO Data Portal | 2-14 day demand, transmission wind, embedded wind and solar |
| Elexon Insights (FOU2T14D) | Nuclear, biomass, hydro and interconnector availability |
| ENTSO-E Transparency | Day-ahead prices for continental shaping |
| ECB | EUR/GBP daily reference rate |
| NESO Final ALF 2026/27 | Gas fleet capacities by efficiency band |
| DESNZ 2025 generation costs | Gas VOM |
| UKGHGI 2025 / BEIS 2025 | ETS and CPS emissions factors |
