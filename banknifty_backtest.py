"""
banknifty_backtest.py
=====================
Bank Nifty 09:20 Short Strangle Backtest — Hiring Assignment Submission
Author: Candidate
Data  : BANKNIFTY_SPOT.csv  (1-min OHLC, index/spot)
        Options_data_2023.csv (1-min OHLC, options)

STRATEGY SUMMARY
----------------
- Universe  : Week 1 of every expiry cycle (see ASSUMPTION #1 below)
- Entry     : 09:20 bar close — SELL the CE and PE whose close is closest to Rs 50
- Exit      : 15:20 close OR 50 % stop-loss (entry × 1.5) on High, whichever first
- Sizing    : 1 lot = 15 units, fixed, no compounding

ASSUMPTIONS (stated explicitly per the brief)
---------------------------------------------
ASSUMPTION #1 — "Week 1" definition
  Bank Nifty has WEEKLY expiries every Wednesday.  A new "expiry cycle" begins
  the day after each Wednesday expiry.  "Week 1" is defined as the calendar
  week (Mon–Sun block) that contains the FIRST Wednesday of that new expiry
  cycle.  Concretely: after expiry Wednesday W, the very next Thursday is the
  start of a new cycle; all trading days in the Mon–Fri block whose Wednesday
  falls after W (i.e., next Wednesday W+7) form "Week 1" of that cycle.
  This means we trade every trading day in that Mon–Fri block, including the
  new expiry Wednesday itself.
  This interpretation can be changed by redefining `get_week1_trading_days()`.

ASSUMPTION #2 — Tie-break when two strikes are equidistant from Rs 50
  Pick the strike with the LOWER premium (cheaper / further OTM).  This is
  the more conservative choice: lower premium -> smaller absolute loss if the
  stop triggers, and lower margin requirement.

ASSUMPTION #3 — Stop-loss exit price
  When a 1-min bar's High >= stop_level (entry x 1.5), we exit at the
  stop_level itself, NOT at the bar's close.  Rationale: we set the stop at
  that exact price, so we assume it was filled exactly there (conservative
  for a short: our loss is capped at exactly 50 % of premium received).

ASSUMPTION #4 — Starting capital
  Rs 10,00,000 (ten lakh).  This is a single configurable constant at the top
  of the script (STARTING_CAPITAL).  Position size does NOT scale with capital
  — it is always 1 lot (15 units).  The running "available capital" column is
  the starting capital plus cumulative realised P&L up to that trade row.

ASSUMPTION #5 — "09:20 bar" interpretation
  The options data timestamps are HH:MM:59 (each bar closes at :59 seconds).
  The "09:20 close" corresponds to Time == '09:20:59'.  This is one bar after
  '09:19:59', consistent with the bar labelling convention in the dataset.

ASSUMPTION #6 — Stop-loss monitoring window
  Stop-loss is checked from the bar AFTER entry (09:21:59 onward) up to and
  including 15:20:59.  The entry bar itself is excluded to avoid lookahead:
  we receive the 09:20:59 close price as our entry fill; we cannot have also
  observed whether High on that same bar exceeded stop_level before filling.

ASSUMPTION #7 — Spot price logging
  The spot close price is recorded at entry time (09:20:59).  The brief says
  "at entry, at minimum"; we log it only at entry to keep one row per leg.

LOT SIZE AND POSITION SIZING
-----------------------------
LOT_SIZE       = 15
STARTING_CAPITAL = 10_00_000  (Rs 10 lakh)
"""

import time
import warnings
import logging
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")          # headless — no GUI window needed
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

# =============================================================================
# GLOBAL CONSTANTS  — change these in one place; they propagate everywhere
# =============================================================================
SPOT_CSV       = Path("BANKNIFTY_SPOT.csv")
OPTIONS_CSV    = Path("Options_data_2023.csv")
OUTPUT_EXCEL   = Path("BankNifty_ShortStrangle_Backtest.xlsx")
CHART_DIR      = Path("charts")             # temp folder for embedded PNGs

LOT_SIZE         = 15                       # Bank Nifty lot size
STARTING_CAPITAL = 10_00_000               # Rs 10,00,000 — ASSUMPTION #4

TARGET_PREMIUM   = 50.0                    # Strike selection target (Rs)
ENTRY_TIME_STR   = "09:20:59"             # 09:20 bar close — ASSUMPTION #5
EXIT_TIME_STR    = "15:20:59"             # 15:20 bar close
SL_MULTIPLIER    = 1.5                    # stop level = entry x 1.5

# =============================================================================
# LOGGING
# =============================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("bnf_backtest")

warnings.filterwarnings("ignore", category=FutureWarning)


# =============================================================================
# MODULE 1: DATA LOADING & CLEANING
# =============================================================================

def load_spot(path: Path) -> pd.DataFrame:
    """
    Load and clean the Bank Nifty spot (index) 1-minute OHLC data.

    Returns a DataFrame with:
      - index  : datetime (Date + Time combined, no timezone)
      - columns: open, high, low, close
    Only rows between 09:15 and 15:30 are kept (market hours).
    Duplicate timestamps are dropped (first kept, duplicates logged).
    """
    log.info("Loading spot data from %s ...", path)
    df = pd.read_csv(path, usecols=["ts", "o", "h", "l", "c"])
    df.rename(columns={"ts": "datetime", "o": "open", "h": "high",
                        "l": "low",      "c": "close"}, inplace=True)

    # Parse timestamps
    df["datetime"] = pd.to_datetime(df["datetime"])
    df["date"] = df["datetime"].dt.date
    df["time"] = df["datetime"].dt.time

    # Filter to market hours
    t_start = pd.to_datetime("09:15:00").time()
    t_end   = pd.to_datetime("15:30:59").time()
    before  = len(df)
    df = df[(df["time"] >= t_start) & (df["time"] <= t_end)].copy()
    log.info("  Spot: %d rows after market-hours filter (dropped %d)",
             len(df), before - len(df))

    # Drop duplicates
    dups = df.duplicated(subset="datetime", keep="first")
    if dups.sum():
        log.warning("  Spot: dropping %d duplicate timestamps", dups.sum())
    df = df[~dups].copy()

    df.set_index("datetime", inplace=True)
    df.sort_index(inplace=True)
    log.info("  Spot: %d clean rows, %d unique trading days",
             len(df), df["date"].nunique())
    return df


def load_options(path: Path) -> pd.DataFrame:
    """
    Load and clean the Bank Nifty options 1-minute OHLC data.

    Column mapping from the raw CSV:
      Date, Ticker, Time, Open, High, Low, Close, Call/Put

    Processing:
      - Combine Date + Time into a single datetime column.
      - Parse Strike and OptionType from Ticker using regex.
      - Drop rows with missing OHLC values (warn + count).
      - Drop duplicate (datetime, Ticker) pairs (warn + count).
      - Keep only market hours (09:15:59 to 15:30:59).
      - Sort by datetime, Ticker.

    Returns a clean DataFrame with columns:
      datetime, date, time_str, ticker, strike, option_type,
      open, high, low, close
    """
    log.info("Loading options data from %s (~10M rows, may take 30s) ...", path)
    t0 = time.perf_counter()

    # Read with specific dtypes to save memory
    dtype_map = {
        "Open":  np.float32,
        "High":  np.float32,
        "Low":   np.float32,
        "Close": np.float32,
    }
    df = pd.read_csv(
        path,
        usecols=["Date", "Ticker", "Time", "Open", "High", "Low", "Close", "Call/Put"],
        dtype=dtype_map,
    )
    log.info("  Raw read: %d rows in %.1fs", len(df), time.perf_counter() - t0)

    # -- Rename columns -------------------------------------------------------
    df.rename(columns={
        "Date":     "date_str",
        "Ticker":   "ticker",
        "Time":     "time_str",
        "Open":     "open",
        "High":     "high",
        "Low":      "low",
        "Close":    "close",
        "Call/Put": "option_type",
    }, inplace=True)

    # -- Combine date + time -> datetime --------------------------------------
    # Both columns are strings; combine then parse once (fast path)
    df["datetime"] = pd.to_datetime(
        df["date_str"] + " " + df["time_str"], format="%Y-%m-%d %H:%M:%S"
    )
    df["date"] = df["datetime"].dt.normalize()   # date-only for groupby joins

    # -- Parse strike from ticker using fast string slicing (no regex) --------
    #    Ticker format is fixed: 'BANKNIFTY' (9 chars) + <digits> + 'CE'/'PE' (2 chars)
    #    e.g. 'BANKNIFTY37000PE' -> strike = ticker[9:-2] -> '37000' -> 37000
    #    This is ~10x faster than str.extract(regex) on 10M rows.
    #    We validate by checking the last 2 chars match the Call/Put column.
    TICKER_PREFIX_LEN = len("BANKNIFTY")   # 9
    TICKER_SUFFIX_LEN = 2                   # len('CE') or len('PE')
    ticker_suffix = df["ticker"].str[-TICKER_SUFFIX_LEN:]  # fast slice
    strike_str    = df["ticker"].str[TICKER_PREFIX_LEN:-TICKER_SUFFIX_LEN]

    # Validate: all tickers should start with 'BANKNIFTY'
    bad_prefix = ~df["ticker"].str.startswith("BANKNIFTY")
    if bad_prefix.sum():
        log.warning("  Options: %d rows with non-BANKNIFTY tickers -- dropped",
                    bad_prefix.sum())
        df = df[~bad_prefix].copy()
        ticker_suffix = ticker_suffix[~bad_prefix]
        strike_str    = strike_str[~bad_prefix]

    # Validate: strike_str should be all-digit
    bad_strikes = ~strike_str.str.isdigit()
    if bad_strikes.sum():
        log.warning("  Options: %d rows with non-numeric strike in ticker -- dropped",
                    bad_strikes.sum())
        df        = df[~bad_strikes].copy()
        strike_str = strike_str[~bad_strikes]

    df["strike"] = strike_str.astype(np.int32)
    # option_type already present in 'Call/Put' column; suffix parsed above
    # is for validation only -- we trust the Call/Put column as ground truth.

    # -- Drop rows with missing OHLC ------------------------------------------
    ohlc_cols = ["open", "high", "low", "close"]
    na_mask = df[ohlc_cols].isna().any(axis=1)
    if na_mask.sum():
        log.warning("  Options: dropping %d rows with NaN OHLC", na_mask.sum())
        df = df[~na_mask].copy()

    # -- Market-hours filter (09:15:59 to 15:30:59) ---------------------------
    before = len(df)
    df = df[(df["time_str"] >= "09:15:59") & (df["time_str"] <= "15:30:59")].copy()
    log.info("  Options: %d rows after market-hours filter (dropped %d)",
             len(df), before - len(df))

    # -- Drop duplicate (datetime, ticker) ------------------------------------
    dup_mask = df.duplicated(subset=["datetime", "ticker"], keep="first")
    if dup_mask.sum():
        log.warning("  Options: dropping %d duplicate (datetime, ticker) rows",
                    dup_mask.sum())
        df = df[~dup_mask].copy()

    # -- Sort -----------------------------------------------------------------
    df.sort_values(["datetime", "ticker"], inplace=True)
    df.reset_index(drop=True, inplace=True)

    log.info("  Options: %d clean rows, %d unique dates, %d unique tickers",
             len(df), df["date"].nunique(), df["ticker"].nunique())
    log.info("  Options load total: %.1fs", time.perf_counter() - t0)
    return df


# =============================================================================
# MODULE 2: EXPIRY & WEEK-1 DAY SELECTION
# =============================================================================

def get_week1_trading_days(trading_dates: pd.DatetimeIndex) -> list:
    """
    Return the list of trading dates that fall in "Week 1" of each calendar month.

    ASSUMPTION #1 - Week-1 definition (auditable -- change this function only):
    ----------------------------------------------------------
    PREVIOUS (WRONG) INTERPRETATION: treating each Wednesday-to-Wednesday
    span as one "expiry cycle" and calling the next span "Week 1."  With
    weekly Bank Nifty expiries EVERY Wednesday, that logic reduces to a no-op
    -- every single trading day satisfies the definition, because every week
    is trivially "Week 1" of the immediately preceding weekly cycle.  That
    is a circular definition that filters nothing.

    CORRECTED INTERPRETATION:
    "Week 1" means the first week of each calendar MONTH -- specifically, the
    subset of trading days from the 1st of the month up to and including the
    FIRST Wednesday of that month.  All trading days from the 1st through
    that first Wednesday are Week-1 days; everything from Thursday onward
    (weeks 2, 3, 4) is excluded.

    Why this interpretation:
      - "Trade only week 1" only makes sense as a restriction if it excludes
        weeks 2/3/4.  With monthly grouping, it does.
      - Bank Nifty monthly context: each calendar month has 4-5 Wednesday
        expiries.  The "first week" is a natural unit meaning the first
        mini-cycle of the month.
      - This reduces the tradeable universe to roughly 1/4 of all trading days
        (only Mon-Wed of the first week each month), which is a meaningful,
        auditable filter.

    Definition: for each calendar month, Week-1 days = every trading day d
    such that d <= first Wednesday of that month.
    (The first Wednesday is included -- it is an expiry day and still traded.)
    ----------------------------------------------------------

    Parameters
    ----------
    trading_dates : pd.DatetimeIndex
        Sorted unique trading dates available in the options dataset.

    Returns
    -------
    list of pd.Timestamp
        All trading dates that are in Week 1 of their respective calendar month.
    """
    td = pd.DatetimeIndex(sorted(set(trading_dates)))
    df = pd.DataFrame({"date": td})
    df["year_month"] = df["date"].dt.to_period("M")

    week1_days = []
    for ym, grp in df.groupby("year_month"):
        wednesdays_in_month = grp.loc[grp["date"].dt.dayofweek == 2, "date"]
        if wednesdays_in_month.empty:
            # Edge case: a month with no Wednesday trading day -- skip it.
            log.warning("  Month %s: no Wednesday trading day found -- skipped", ym)
            continue
        first_wed = wednesdays_in_month.min()   # earliest Wednesday in this month
        # Week 1 = every trading day from month-start up to (and incl.) first_wed
        month_week1 = grp.loc[grp["date"] <= first_wed, "date"].tolist()
        week1_days.extend(month_week1)
        log.info("  Month %s: first Wed = %s, Week-1 days = %d",
                 ym, first_wed.date(), len(month_week1))

    week1_days_sorted = sorted(set(week1_days))
    log.info("Total Week-1 trading days selected: %d (out of %d total trading days; "
             "excluded %d days from weeks 2-4 of each month)",
             len(week1_days_sorted), len(td), len(td) - len(week1_days_sorted))
    return week1_days_sorted


# =============================================================================
# MODULE 3: STRIKE SELECTION -- closest to Rs 50 at 09:20 bar
# =============================================================================

def select_strikes(opts_day: pd.DataFrame, trade_date: pd.Timestamp) -> dict:
    """
    For a single trading day, select the CE and PE strike whose 09:20:59
    close price is closest to Rs 50.

    Parameters
    ----------
    opts_day  : DataFrame -- all options rows for trade_date (already filtered).
    trade_date: pd.Timestamp

    Returns
    -------
    dict with keys 'CE' and 'PE', each a dict:
        {ticker, strike, option_type, entry_price}
    or None for a leg if no valid 09:20 bar exists for that option type.

    ASSUMPTION #2 -- Tie-break:
      If two strikes have identical |close - 50|, pick the one with the LOWER
      close price (cheaper, further OTM).  This is the more conservative choice
      for a short-premium strategy (lower potential loss, lower margin).
    """
    # Filter to the 09:20:59 bar only
    at_entry = opts_day[opts_day["time_str"] == ENTRY_TIME_STR].copy()

    result = {}
    for opt_type in ("CE", "PE"):
        leg = at_entry[at_entry["option_type"] == opt_type].copy()

        if leg.empty:
            log.warning("  %s: No %s options at %s -- leg skipped",
                        trade_date.date(), opt_type, ENTRY_TIME_STR)
            result[opt_type] = None
            continue

        # Distance from target premium
        leg["dist_from_target"] = (leg["close"] - TARGET_PREMIUM).abs()

        # Primary sort: dist_from_target (ascending)
        # Tie-break sort: close price (ascending) -- ASSUMPTION #2
        leg.sort_values(["dist_from_target", "close"], inplace=True)
        best = leg.iloc[0]

        result[opt_type] = {
            "ticker":      best["ticker"],
            "strike":      int(best["strike"]),
            "option_type": opt_type,
            "entry_price": float(best["close"]),
        }

    return result


# =============================================================================
# MODULE 4: SIGNAL GENERATION -- entry, stop-loss, time exit
# =============================================================================

def compute_exit(leg_info: dict, opts_day: pd.DataFrame, trade_date: pd.Timestamp) -> dict:
    """
    Given an entered leg (short option), find the exit price and time using:
      1. 50% stop-loss: first bar after entry where High >= entry x 1.5.
         Exit price = stop_level (ASSUMPTION #3).
      2. Time exit: 15:20:59 close if no stop is hit.

    STOP-LOSS MECHANICS (no lookahead -- ASSUMPTION #6):
    ---------------------------------------------------
    - The entry bar is 09:20:59.  We receive the close of that bar as our fill.
    - Stop monitoring begins at 09:21:59 (the NEXT bar).
    - We scan bars in chronological order; the first bar where High >= stop_level
      is the exit bar.  Exit price = stop_level (the level we set the stop at).
    - The exit bar is included up to 15:20:59 (the time-exit bar).
    - If no stop hit, exit at the 15:20:59 close.
    - If 15:20:59 bar is missing, we log a warning and use the last available
      bar at or before 15:20:59.

    Parameters
    ----------
    leg_info : dict with ticker, strike, option_type, entry_price
    opts_day : all options rows for that day
    trade_date: pd.Timestamp

    Returns
    -------
    dict with keys: exit_price, exit_time, exit_reason ('SL' or 'TIME')
    """
    ticker      = leg_info["ticker"]
    entry_price = leg_info["entry_price"]
    stop_level  = entry_price * SL_MULTIPLIER

    # Rows for this specific ticker on this day, after the entry bar
    ticker_rows = opts_day[
        (opts_day["ticker"] == ticker) &
        (opts_day["time_str"] > ENTRY_TIME_STR) &   # strictly AFTER entry bar
        (opts_day["time_str"] <= EXIT_TIME_STR)      # up to and including exit bar
    ].copy()

    # -- Stop-loss scan (vectorized comparison, then argmax for first hit) ----
    # Using High column as required: for a short position, rising price hurts;
    # High is the most adverse price within the 1-min bar for a short seller.
    if not ticker_rows.empty:
        sl_hit = ticker_rows["high"] >= stop_level   # boolean Series
        if sl_hit.any():
            first_sl_idx = sl_hit.idxmax()           # index of first True
            sl_row = ticker_rows.loc[first_sl_idx]
            return {
                "exit_price":  stop_level,           # ASSUMPTION #3
                "exit_time":   sl_row["time_str"],
                "exit_reason": "SL",
            }

    # -- Time exit at 15:20:59 ------------------------------------------------
    time_exit_rows = ticker_rows[ticker_rows["time_str"] == EXIT_TIME_STR]
    if not time_exit_rows.empty:
        exit_price = float(time_exit_rows.iloc[0]["close"])
        return {
            "exit_price":  exit_price,
            "exit_time":   EXIT_TIME_STR,
            "exit_reason": "TIME",
        }

    # -- Fallback: last available bar before 15:20:59 -------------------------
    # Handles missing 15:20 bar gracefully -- ASSUMPTION: use last close
    fallback_rows = opts_day[
        (opts_day["ticker"] == ticker) &
        (opts_day["time_str"] > ENTRY_TIME_STR) &
        (opts_day["time_str"] <= EXIT_TIME_STR)
    ]
    if not fallback_rows.empty:
        last_row = fallback_rows.iloc[-1]
        log.warning("  %s %s: 15:20 bar missing; using last available bar %s",
                    trade_date.date(), ticker, last_row["time_str"])
        return {
            "exit_price":  float(last_row["close"]),
            "exit_time":   last_row["time_str"],
            "exit_reason": "TIME_FALLBACK",
        }

    # -- No data at all after entry -- log and return NaN ---------------------
    log.warning("  %s %s: No post-entry data found -- P&L set to 0",
                trade_date.date(), ticker)
    return {
        "exit_price":  entry_price,   # effectively 0 P&L
        "exit_time":   EXIT_TIME_STR,
        "exit_reason": "NO_DATA",
    }


# =============================================================================
# MODULE 5: POSITION SIZING
# =============================================================================

def get_quantity() -> int:
    """
    Return fixed quantity per leg.
    1 lot x LOT_SIZE = 15 units.
    Position size does NOT scale with capital (fixed sizing).
    """
    return LOT_SIZE   # 15 units per leg, fixed, no compounding


# =============================================================================
# MODULE 6: P&L AND TRADE SHEET GENERATION
# =============================================================================

def build_trade_sheet(
    week1_days: list,
    opts: pd.DataFrame,
    spot: pd.DataFrame,
) -> pd.DataFrame:
    """
    Main backtest loop: iterates over ~50 Week-1 trading days (not over
    minute bars -- the per-bar work is vectorized inside compute_exit()).

    For each day:
      - Select strikes (Module 3).
      - Compute exit for each leg (Module 4).
      - Build one row per leg in the trade sheet.

    Returns
    -------
    pd.DataFrame with all required trade sheet columns (one row per leg per day).
    """
    qty   = get_quantity()
    rows  = []
    cum_pnl = 0.0
    available_capital = float(STARTING_CAPITAL)

    # -- Pre-extract spot 09:20 prices for all days (vectorized) -------------
    # We look them up by date in a dict for O(1) access per day.
    spot_reset = spot.reset_index()
    spot_reset["time_str"] = spot_reset["datetime"].dt.strftime("%H:%M:%S")
    spot_entry = spot_reset[spot_reset["time_str"] == ENTRY_TIME_STR].copy()
    spot_entry["date"] = spot_entry["datetime"].dt.normalize()
    spot_price_at_entry = spot_entry.set_index("date")["close"].to_dict()

    # -- Pre-group options by date for fast lookup ----------------------------
    opts_grouped = dict(tuple(opts.groupby("date")))

    log.info("Building trade sheet over %d Week-1 days ...", len(week1_days))

    for trade_date in week1_days:
        td = pd.Timestamp(trade_date)

        # -- Filter options to this day (O(1) dict lookup) --------------------
        opts_day = opts_grouped.get(td, None)

        if opts_day is None or opts_day.empty:
            log.warning("  %s: No options data found -- day skipped", td.date())
            continue

        # -- Strike selection -------------------------------------------------
        strikes = select_strikes(opts_day, td)

        # -- Spot price at entry ----------------------------------------------
        spot_entry_price = spot_price_at_entry.get(td, np.nan)

        # -- Build rows for CE and PE legs ------------------------------------
        for opt_type in ("CE", "PE"):
            leg = strikes.get(opt_type)
            if leg is None:
                continue   # already logged in select_strikes

            # Compute exit
            exit_info = compute_exit(leg, opts_day, td)

            entry_price = leg["entry_price"]
            exit_price  = exit_info["exit_price"]
            entry_value = entry_price * qty
            exit_value  = exit_price  * qty

            # For a SHORT position: P&L = sell_price - buy_back_price (per unit)
            # Gross P&L = entry_value - exit_value
            gross_pnl = entry_value - exit_value

            cum_pnl          += gross_pnl
            available_capital = STARTING_CAPITAL + cum_pnl

            rows.append({
                "entry_date":        td.date(),
                "exit_date":         td.date(),      # same day (intraday)
                "entry_time":        ENTRY_TIME_STR,
                "exit_time":         exit_info["exit_time"],
                "exit_reason":       exit_info["exit_reason"],
                "ticker":            leg["ticker"],
                "strike":            leg["strike"],
                "option_type":       opt_type,
                "entry_price":       round(entry_price, 4),
                "exit_price":        round(exit_price,  4),
                "quantity":          qty,
                "entry_value":       round(entry_value, 4),
                "exit_value":        round(exit_value,  4),
                "gross_pnl":         round(gross_pnl,   4),
                "cumulative_pnl":    round(cum_pnl,      4),
                "available_capital": round(available_capital, 4),
                "spot_at_entry":     (round(spot_entry_price, 2)
                                      if not np.isnan(spot_entry_price) else np.nan),
            })

    df = pd.DataFrame(rows)
    log.info("Trade sheet built: %d rows (%d days x 2 legs, minus skipped legs)",
             len(df), len(week1_days))
    return df


# =============================================================================
# MODULE 7: STATISTICS
# =============================================================================

def compute_statistics(trade_sheet: pd.DataFrame) -> dict:
    """
    Compute all required statistics from the trade sheet.

    Statistics computed:
    --------------------
    - CAGR: (end_NAV / start_NAV)^(1/years) - 1, where start_NAV = STARTING_CAPITAL
    - Max Drawdown: from running-peak equity curve (trade-wise)
    - Win/Loss counts and percentages: CE, PE, combined
    - Average % P&L: CE vs PE, expiry-day vs non-expiry-day
    - Monthly % P&L table
    - Equity curve (NAV = 100 base, updated trade-wise)

    Returns
    -------
    dict with keys: stats_summary, monthly_pnl, equity_curve, drawdown_curve,
                    win_loss_table, avg_pct_pnl_table
    """
    if trade_sheet.empty:
        log.warning("Trade sheet is empty -- no statistics to compute")
        return {}

    df = trade_sheet.copy()
    df["entry_date"] = pd.to_datetime(df["entry_date"])

    # -- Equity curve (trade-wise, base NAV = 100) ----------------------------
    # NAV is updated after every individual leg's P&L is realized.
    # equity = STARTING_CAPITAL + cumulative_pnl
    # NAV_index = equity / STARTING_CAPITAL x 100
    df = df.reset_index(drop=True)
    df["equity"]    = STARTING_CAPITAL + df["cumulative_pnl"]
    df["nav_index"] = df["equity"] / STARTING_CAPITAL * 100.0

    # -- Running peak and drawdown --------------------------------------------
    # Start the peak at NAV = 100 (before any trade)
    nav_series      = pd.concat([pd.Series([100.0]), df["nav_index"]])
    running_peak    = nav_series.cummax()
    drawdown_series = (nav_series - running_peak) / running_peak  # always <= 0
    max_drawdown    = drawdown_series.min()                        # most negative

    # -- CAGR -----------------------------------------------------------------
    start_date = df["entry_date"].min()
    end_date   = df["entry_date"].max()
    years      = max((end_date - start_date).days / 365.25, 1 / 365.25)
    start_nav  = STARTING_CAPITAL
    end_nav    = df["equity"].iloc[-1]
    cagr       = (end_nav / start_nav) ** (1 / years) - 1

    # -- Win / Loss -----------------------------------------------------------
    df["is_win"] = df["gross_pnl"] > 0

    def win_loss_stats(sub: pd.DataFrame, label: str) -> dict:
        n       = len(sub)
        wins    = sub["is_win"].sum()
        losses  = n - wins
        win_pct = wins / n * 100 if n else 0
        return {
            "label":    label,
            "trades":   n,
            "wins":     int(wins),
            "losses":   int(losses),
            "win_pct":  round(win_pct, 2),
            "loss_pct": round(100 - win_pct, 2),
        }

    wl_ce   = win_loss_stats(df[df["option_type"] == "CE"], "CE")
    wl_pe   = win_loss_stats(df[df["option_type"] == "PE"], "PE")
    wl_all  = win_loss_stats(df, "Combined")

    # -- Average % P&L --------------------------------------------------------
    # % P&L per trade = gross_pnl / entry_value
    df["pct_pnl"] = df["gross_pnl"] / df["entry_value"] * 100

    # Identify expiry days (Wednesdays in trading dates)
    df["is_expiry"] = df["entry_date"].dt.dayofweek == 2

    def avg_pct(sub, label):
        return {
            "label":   label,
            "count":   len(sub),
            "avg_pct": round(sub["pct_pnl"].mean(), 4) if len(sub) else np.nan,
        }

    avg_pnl_rows = [
        avg_pct(df[(df["option_type"] == "CE") & ~df["is_expiry"]],  "CE -- Non-Expiry"),
        avg_pct(df[(df["option_type"] == "CE") &  df["is_expiry"]],  "CE -- Expiry Day"),
        avg_pct(df[(df["option_type"] == "CE")],                      "CE -- Combined"),
        avg_pct(df[(df["option_type"] == "PE") & ~df["is_expiry"]],  "PE -- Non-Expiry"),
        avg_pct(df[(df["option_type"] == "PE") &  df["is_expiry"]],  "PE -- Expiry Day"),
        avg_pct(df[(df["option_type"] == "PE")],                      "PE -- Combined"),
        avg_pct(df[~df["is_expiry"]],                                 "All -- Non-Expiry"),
        avg_pct(df[ df["is_expiry"]],                                 "All -- Expiry Day"),
        avg_pct(df,                                                    "All -- Combined"),
    ]

    # -- Monthly P&L ----------------------------------------------------------
    # Group cumulative P&L by month-end NAV; compute month-over-month %
    df["ym"]           = df["entry_date"].dt.to_period("M")
    monthly_last_nav   = df.groupby("ym")["nav_index"].last()
    monthly_first_nav  = pd.Series(
        [100.0] + monthly_last_nav.iloc[:-1].tolist(),
        index=monthly_last_nav.index,
    )
    monthly_pct_pnl    = (monthly_last_nav - monthly_first_nav) / monthly_first_nav * 100
    monthly_df         = pd.DataFrame({
        "Month":          monthly_last_nav.index.astype(str),
        "End NAV":        monthly_last_nav.values.round(4),
        "Start NAV":      monthly_first_nav.values.round(4),
        "Monthly % P&L":  monthly_pct_pnl.values.round(4),
    })

    # -- Total P&L summary ----------------------------------------------------
    stats_summary = {
        "CAGR (%)":             round(cagr * 100, 4),
        "Max Drawdown (%)":     round(max_drawdown * 100, 4),
        "Total Gross P&L (Rs)": round(df["gross_pnl"].sum(), 2),
        "Starting Capital (Rs)":STARTING_CAPITAL,
        "Ending Capital (Rs)":  round(end_nav, 2),
        "Total Trades":         len(df),
        "Start Date":           str(start_date.date()),
        "End Date":             str(end_date.date()),
        "Years":                round(years, 4),
    }

    log.info("Statistics computed: CAGR=%.2f%%, MaxDD=%.2f%%, TotalPnL=%.0f",
             stats_summary["CAGR (%)"],
             stats_summary["Max Drawdown (%)"],
             stats_summary["Total Gross P&L (Rs)"])

    return {
        "stats_summary":   stats_summary,
        "monthly_pnl":     monthly_df,
        "equity_curve":    df[["entry_date", "ticker", "nav_index", "equity"]].copy(),
        "drawdown_series": drawdown_series.values * 100,   # in %
        "win_loss":        [wl_ce, wl_pe, wl_all],
        "avg_pct_pnl":     avg_pnl_rows,
        "trade_df":        df,
    }


# =============================================================================
# MODULE 7b: CHART GENERATION (helpers for Excel embedding)
# =============================================================================

def save_equity_chart(equity_curve: pd.DataFrame, drawdown_series: np.ndarray,
                      chart_dir: Path) -> tuple:
    """
    Generate and save:
      1. Equity curve chart (NAV index, base=100)
      2. Drawdown chart (with max drawdown point annotated)

    Returns paths to the two saved PNG files.
    """
    chart_dir.mkdir(exist_ok=True)

    # -- Style setup ----------------------------------------------------------
    plt.rcParams.update({
        "figure.facecolor": "#0d1117",
        "axes.facecolor":   "#161b22",
        "axes.edgecolor":   "#30363d",
        "axes.labelcolor":  "#c9d1d9",
        "xtick.color":      "#8b949e",
        "ytick.color":      "#8b949e",
        "text.color":       "#c9d1d9",
        "grid.color":       "#21262d",
        "grid.linestyle":   "--",
        "grid.alpha":       0.7,
        "font.family":      "DejaVu Sans",
    })

    x   = np.arange(len(equity_curve))
    nav = equity_curve["nav_index"].values

    # -- Chart 1: Equity Curve ------------------------------------------------
    fig1, ax1 = plt.subplots(figsize=(14, 6))
    ax1.plot(x, nav, color="#58a6ff", linewidth=1.5, label="NAV Index")
    ax1.axhline(100, color="#8b949e", linewidth=0.8, linestyle="--", label="Base NAV=100")
    ax1.fill_between(x, 100, nav, where=(nav >= 100), alpha=0.15, color="#3fb950")
    ax1.fill_between(x, 100, nav, where=(nav <  100), alpha=0.15, color="#f85149")
    ax1.set_title("Equity Curve -- BankNifty Short Strangle (09:20)",
                  fontsize=14, fontweight="bold", pad=12)
    ax1.set_xlabel("Trade #", fontsize=11)
    ax1.set_ylabel("NAV Index (Base = 100)", fontsize=11)
    ax1.legend(fontsize=10)
    ax1.grid(True)
    fig1.tight_layout()
    equity_path = chart_dir / "equity_curve.png"
    fig1.savefig(equity_path, dpi=150, bbox_inches="tight")
    plt.close(fig1)
    log.info("  Equity curve chart saved -> %s", equity_path)

    # -- Chart 2: Drawdown Curve ----------------------------------------------
    dd    = drawdown_series                            # in %, length = len(nav)+1
    x_dd  = np.arange(len(dd))
    max_dd_idx = int(np.argmin(dd))
    max_dd_val = dd[max_dd_idx]

    fig2, ax2 = plt.subplots(figsize=(14, 4))
    ax2.fill_between(x_dd, 0, dd, alpha=0.5, color="#f85149", label="Drawdown %")
    ax2.plot(x_dd, dd, color="#f85149", linewidth=1.0)
    ax2.scatter(max_dd_idx, max_dd_val, color="#ff7b72", s=80, zorder=5,
                label=f"Max DD = {max_dd_val:.2f}%")
    ax2.annotate(
        f"Max DD\n{max_dd_val:.2f}%",
        xy=(max_dd_idx, max_dd_val),
        xytext=(max_dd_idx + max(1, len(dd) // 20), max_dd_val * 0.8),
        color="#ff7b72", fontsize=9,
        arrowprops=dict(arrowstyle="->", color="#ff7b72"),
    )
    ax2.set_title("Drawdown Curve -- BankNifty Short Strangle",
                  fontsize=14, fontweight="bold", pad=12)
    ax2.set_xlabel("Trade # (including start point)", fontsize=11)
    ax2.set_ylabel("Drawdown (%)", fontsize=11)
    ax2.axhline(0, color="#8b949e", linewidth=0.8)
    ax2.legend(fontsize=10)
    ax2.grid(True)
    fig2.tight_layout()
    dd_path = chart_dir / "drawdown_curve.png"
    fig2.savefig(dd_path, dpi=150, bbox_inches="tight")
    plt.close(fig2)
    log.info("  Drawdown curve chart saved -> %s", dd_path)

    return str(equity_path), str(dd_path)


# =============================================================================
# MODULE 8: EXCEL OUTPUT
# =============================================================================

# -- Excel styling helpers ---------------------------------------------------

def _header_font():
    return Font(bold=True, color="FFFFFF", name="Calibri", size=11)

def _header_fill(hex_color="1F4E79"):
    return PatternFill("solid", fgColor=hex_color)

def _center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def _thin_border():
    side = Side(style="thin", color="CCCCCC")
    return Border(left=side, right=side, top=side, bottom=side)

def _write_header_row(ws, row_num, headers, fill_hex="1F4E79"):
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=row_num, column=col, value=h)
        cell.font      = _header_font()
        cell.fill      = _header_fill(fill_hex)
        cell.alignment = _center()
        cell.border    = _thin_border()

def _write_data_row(ws, row_num, values, number_formats=None):
    for col, val in enumerate(values, start=1):
        cell = ws.cell(row=row_num, column=col, value=val)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = _thin_border()
        if number_formats and col <= len(number_formats) and number_formats[col - 1]:
            cell.number_format = number_formats[col - 1]

def _auto_col_width(ws, min_width=10, max_width=35):
    for col in ws.columns:
        max_len = max(
            (len(str(cell.value)) if cell.value is not None else 0)
            for cell in col
        )
        ws.column_dimensions[get_column_letter(col[0].column)].width = \
            min(max_width, max(min_width, max_len + 4))


def write_guide_sheet(ws):
    """Write the Guide sheet with plain-English explanations of the backtest."""
    ws.title = "Guide"

    # Title
    ws.merge_cells("A1:E1")
    title_cell = ws["A1"]
    title_cell.value     = "BankNifty Short Strangle Backtest -- Guide & Interpretation"
    title_cell.font      = Font(bold=True, color="FFFFFF", size=14, name="Calibri")
    title_cell.fill      = PatternFill("solid", fgColor="1F4E79")
    title_cell.alignment = _center()
    ws.row_dimensions[1].height = 30

    guide_content = [
        ("", ""),
        ("OVERVIEW", ""),
        ("Strategy", "Bank Nifty 09:20 Short Strangle -- SELL one CE and one PE at 09:20 each day in Week 1 of every expiry cycle."),
        ("Data", "BANKNIFTY_SPOT.csv (1-min OHLC, index/spot) and Options_data_2023.csv (1-min OHLC, options, Jan 2023 to Jan 2024)."),
        ("Lot Size", "15 (1 lot). Fixed across all trades -- no compounding, no scaling with capital."),
        ("Starting Capital", f"Rs {STARTING_CAPITAL:,} (ten lakh). Single configurable constant. Used for CAGR and NAV calculation only; position size does not scale."),
        ("", ""),
        ("ASSUMPTION #1 -- Week 1 Definition", ""),
        ("", "CORRECTED interpretation: 'Week 1' means the first week of each calendar MONTH."),
        ("", "Specifically: for each month, find the first Wednesday of that month. "
              "Every trading day from the 1st of the month up to and including that first Wednesday is a 'Week-1' day. "
              "Thursdays onward (weeks 2, 3, 4) are excluded."),
        ("", "Why monthly, not weekly: Bank Nifty had weekly expiries (every Wednesday) in 2023. "
              "A weekly-cycle definition is a no-op -- since every week has a Wednesday, every week trivially "
              "satisfies 'first week after the last expiry.' The filter must operate on a monthly cycle to exclude anything."),
        ("", "This cuts the universe to ~1/4-1/5 of all trading days (typically Mon-Wed of the first week each month)."),
        ("", "The definition lives entirely in get_week1_trading_days() and can be changed in one place."),
        ("", ""),
        ("ASSUMPTION #2 -- Tie-Break Rule", ""),
        ("", "If two strikes are exactly equidistant from Rs 50, the one with the LOWER premium (cheaper, further OTM) is chosen."),
        ("", "Rationale: lower premium = smaller maximum loss on the short, and lower margin requirement. Conservative for risk management."),
        ("", ""),
        ("ASSUMPTION #3 -- Stop-Loss Exit Price", ""),
        ("", "When a 1-min bar's High >= entry_price x 1.5, the exit price is set to the stop_level (entry x 1.5), NOT the bar's close."),
        ("", "Rationale: the stop is a specific order placed at that price level; we assume it was filled exactly there. This gives a precise, reproducible P&L."),
        ("", ""),
        ("ASSUMPTION #4 -- Starting Capital", ""),
        ("", f"Rs {STARTING_CAPITAL:,}. Set as a single constant STARTING_CAPITAL at the top of the script. CAGR is computed relative to this base. Available capital = STARTING_CAPITAL + cumulative realized P&L."),
        ("", ""),
        ("ASSUMPTION #5 -- 09:20 bar Timestamp", ""),
        ("", "Options data timestamps are HH:MM:59 (bar closes at :59 seconds). The 09:20 bar = Time == '09:20:59'. Entry fill = close price of that bar."),
        ("", ""),
        ("ASSUMPTION #6 -- Stop-Loss Monitoring Window", ""),
        ("", "Monitoring begins at 09:21:59 (the bar AFTER entry). The entry bar (09:20:59) is excluded from stop-loss scanning to avoid lookahead bias. Scanning stops at 15:20:59 (the time-exit bar)."),
        ("", ""),
        ("ASSUMPTION #7 -- Spot Price Logging", ""),
        ("", "Spot close is logged only at entry time (09:20:59). The brief says 'at entry, at minimum'; exit-time spot is not logged to keep one row per leg."),
        ("", ""),
        ("HOW TO READ SHEET 2 -- TRADESHEET", ""),
        ("", "One row per leg per day (two rows per trading day: one CE, one PE)."),
        ("", "Gross P&L = entry_value - exit_value (positive = profit for the short seller)."),
        ("", "Cumulative P&L accumulates trade-by-trade (not daily)."),
        ("", "Available Capital = Starting Capital + Cumulative P&L at that row."),
        ("exit_reason codes", "SL = stop-loss triggered; TIME = time exit at 15:20; TIME_FALLBACK = 15:20 bar missing; NO_DATA = no post-entry data."),
        ("", ""),
        ("HOW TO READ SHEET 3 -- STATISTICS", ""),
        ("", "CAGR: annualized return from start_NAV to end_NAV over the dataset period."),
        ("", "Max Drawdown: computed from the running peak of the trade-wise NAV curve (not simple min-of-series)."),
        ("", "Equity Curve: NAV index starting at 100, updated after every individual leg trade."),
        ("", "Monthly % P&L: (end-of-month NAV - start-of-month NAV) / start-of-month NAV x 100."),
        ("", "Win % / Loss %: a trade is a 'win' if Gross P&L > 0."),
    ]

    for r, (key, val) in enumerate(guide_content, start=2):
        kc = ws.cell(row=r, column=1, value=key)
        vc = ws.cell(row=r, column=2, value=val)
        if key and key.isupper():
            kc.font = Font(bold=True, color="1F4E79", size=11, name="Calibri")
        elif key.startswith("ASSUMPTION"):
            kc.font = Font(bold=True, color="375623", size=10, name="Calibri")
        vc.alignment = Alignment(wrap_text=True, vertical="top")
        kc.alignment = Alignment(wrap_text=True, vertical="top")

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 100
    for r in range(2, len(guide_content) + 3):
        ws.row_dimensions[r].height = 18


def write_trade_sheet(ws, trade_df: pd.DataFrame):
    """Write the full trade sheet to the Excel worksheet."""
    ws.title = "Tradesheet"

    headers = [
        "Entry Date", "Exit Date", "Entry Time", "Exit Time", "Exit Reason",
        "Option Ticker", "Strike Price", "Option Type",
        "Entry Price (Rs)", "Exit Price (Rs)", "Quantity",
        "Entry Value (Rs)", "Exit Value (Rs)", "Gross P&L (Rs)",
        "Cumulative P&L (Rs)", "Available Capital (Rs)",
        "Spot at Entry (Rs)",
    ]
    _write_header_row(ws, 1, headers, fill_hex="1F4E79")

    col_map = [
        "entry_date", "exit_date", "entry_time", "exit_time", "exit_reason",
        "ticker", "strike", "option_type",
        "entry_price", "exit_price", "quantity",
        "entry_value", "exit_value", "gross_pnl",
        "cumulative_pnl", "available_capital", "spot_at_entry",
    ]

    # Number formats for numeric columns (index 0-based, 1-indexed in Excel)
    num_fmts = [
        None, None, None, None, None,           # dates/strings
        None, "#,##0", None,                     # ticker, strike, opt_type
        "#,##0.00", "#,##0.00", "#,##0",         # entry_price, exit_price, qty
        "#,##0.00", "#,##0.00",                  # entry_val, exit_val
        '#,##0.00;[Red]-#,##0.00',               # gross_pnl
        '#,##0.00;[Red]-#,##0.00',               # cum_pnl
        "#,##0.00",                               # avail_cap
        "#,##0.00",                               # spot
    ]

    for row_idx, (_, row) in enumerate(trade_df.iterrows(), start=2):
        values = [row[c] if c in row.index else "" for c in col_map]
        _write_data_row(ws, row_idx, values, num_fmts)
        # Colour CE rows light blue, PE rows light orange
        opt_fill = "DDEEFF" if row["option_type"] == "CE" else "FFF0E0"
        for col in range(1, len(headers) + 1):
            ws.cell(row=row_idx, column=col).fill = PatternFill("solid", fgColor=opt_fill)

    _auto_col_width(ws)
    ws.freeze_panes = "A2"


def write_statistics_sheet(ws, stats: dict, equity_path: str, dd_path: str):
    """Write the statistics sheet with tables and embedded charts."""
    ws.title = "Statistics"

    row = 1
    # -- Summary Stats --------------------------------------------------------
    ws.merge_cells(f"A{row}:B{row}")
    hdr = ws.cell(row=row, column=1, value="Summary Statistics")
    hdr.font      = Font(bold=True, color="FFFFFF", size=13, name="Calibri")
    hdr.fill      = PatternFill("solid", fgColor="1F4E79")
    hdr.alignment = _center()
    row += 1

    for k, v in stats["stats_summary"].items():
        kc = ws.cell(row=row, column=1, value=k)
        vc = ws.cell(row=row, column=2, value=v)
        kc.font    = Font(bold=True, name="Calibri", size=10)
        kc.fill    = PatternFill("solid", fgColor="EBF3FF")
        vc.alignment = Alignment(horizontal="right")
        kc.border  = _thin_border()
        vc.border  = _thin_border()
        row += 1

    row += 1

    # -- Win/Loss Table -------------------------------------------------------
    wl_headers = ["Segment", "Total Trades", "Wins", "Losses", "Win %", "Loss %"]
    _write_header_row(ws, row, wl_headers, fill_hex="375623")
    row += 1
    for wl in stats["win_loss"]:
        _write_data_row(ws, row, [
            wl["label"], wl["trades"], wl["wins"], wl["losses"],
            wl["win_pct"], wl["loss_pct"]
        ])
        row += 1

    row += 1

    # -- Average % P&L Table -------------------------------------------------
    avg_headers = ["Segment", "Trade Count", "Avg % P&L"]
    _write_header_row(ws, row, avg_headers, fill_hex="6B3FA0")
    row += 1
    for ap in stats["avg_pct_pnl"]:
        _write_data_row(ws, row, [ap["label"], ap["count"], ap["avg_pct"]])
        ws.cell(row=row, column=3).number_format = "0.0000"
        row += 1

    row += 1

    # -- Monthly P&L Table ---------------------------------------------------
    monthly_df = stats["monthly_pnl"]
    monthly_headers = list(monthly_df.columns)
    _write_header_row(ws, row, monthly_headers, fill_hex="9E3B00")
    row += 1
    for _, mrow in monthly_df.iterrows():
        vals = [mrow["Month"], round(mrow["End NAV"], 4),
                round(mrow["Start NAV"], 4), round(mrow["Monthly % P&L"], 4)]
        _write_data_row(ws, row, vals)
        pnl_cell = ws.cell(row=row, column=4)
        pnl_cell.number_format = "0.00"
        if vals[3] < 0:
            pnl_cell.font = Font(color="C00000", bold=True)
        else:
            pnl_cell.font = Font(color="375623", bold=True)
        row += 1

    row += 2

    # -- Embed charts ---------------------------------------------------------
    ws.cell(row=row, column=1, value="Equity Curve").font = \
        Font(bold=True, size=12, color="1F4E79", name="Calibri")
    row += 1
    try:
        img1 = XLImage(equity_path)
        img1.width  = 900
        img1.height = 380
        ws.add_image(img1, f"A{row}")
        row += 28   # approximate row advance for image height
    except Exception as e:
        log.warning("Could not embed equity chart: %s", e)

    ws.cell(row=row, column=1, value="Drawdown Curve").font = \
        Font(bold=True, size=12, color="C00000", name="Calibri")
    row += 1
    try:
        img2 = XLImage(dd_path)
        img2.width  = 900
        img2.height = 260
        ws.add_image(img2, f"A{row}")
    except Exception as e:
        log.warning("Could not embed drawdown chart: %s", e)

    # -- Equity Curve table (underlying data) -- placed in column G onward ---
    ec = stats["equity_curve"]
    ec_headers = ["Trade #", "Date", "Ticker", "NAV Index", "Equity (Rs)"]
    ec_start_col = 7   # column G
    for ci, h in enumerate(ec_headers, start=ec_start_col):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font      = _header_font()
        cell.fill      = _header_fill("1F4E79")
        cell.alignment = _center()
        cell.border    = _thin_border()
    for ri, (_, er) in enumerate(ec.iterrows(), start=2):
        row_vals = [
            ri - 1,
            str(er["entry_date"].date()),
            er["ticker"],
            round(er["nav_index"], 4),
            round(er["equity"], 2),
        ]
        for ci, val in enumerate(row_vals, start=ec_start_col):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.alignment = Alignment(horizontal="center")
            cell.border    = _thin_border()

    _auto_col_width(ws)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 22
    ws.freeze_panes = "A2"


def write_excel(trade_df: pd.DataFrame, stats: dict,
                equity_path: str, dd_path: str,
                output_path: Path):
    """
    Write the 3-sheet Excel workbook in the required order:
      Sheet 1: Guide
      Sheet 2: Tradesheet
      Sheet 3: Statistics
    """
    log.info("Writing Excel workbook -> %s ...", output_path)
    wb = Workbook()

    # Workbook starts with one default sheet -- rename to Guide
    ws_guide = wb.active
    write_guide_sheet(ws_guide)

    ws_trade = wb.create_sheet()
    write_trade_sheet(ws_trade, trade_df)

    ws_stats = wb.create_sheet()
    write_statistics_sheet(ws_stats, stats, equity_path, dd_path)

    wb.save(output_path)
    log.info("Excel saved: %s", output_path)


# =============================================================================
# TIMING WRAPPER
# =============================================================================

class Timer:
    """Simple context-manager timer that stores elapsed seconds."""
    def __init__(self, label):
        self.label   = label
        self.elapsed = 0.0

    def __enter__(self):
        self._start = time.perf_counter()
        return self

    def __exit__(self, *_):
        self.elapsed = time.perf_counter() - self._start
        log.info("  [STAGE] %s: %.2fs", self.label, self.elapsed)


# =============================================================================
# MAIN
# =============================================================================

def main():
    total_start = time.perf_counter()
    timings = {}

    print("\n" + "=" * 70)
    print("  BANK NIFTY SHORT STRANGLE BACKTEST -- 09:20 Entry")
    print("=" * 70 + "\n")

    # -- Stage 1: Data Load --------------------------------------------------
    with Timer("Data Load") as t_load:
        spot = load_spot(SPOT_CSV)
        opts = load_options(OPTIONS_CSV)
    timings["Data Load"] = t_load.elapsed

    # -- Stage 2: Week-1 Day Selection ---------------------------------------
    with Timer("Week-1 Selection") as t_week:
        trading_dates = opts["date"].drop_duplicates().sort_values()
        trading_dti   = pd.DatetimeIndex(trading_dates)
        week1_days    = get_week1_trading_days(trading_dti)
    timings["Week-1 Selection"] = t_week.elapsed

    # -- Stage 3: Backtest (trade sheet) -------------------------------------
    with Timer("Backtest") as t_bt:
        trade_df = build_trade_sheet(week1_days, opts, spot)
    timings["Backtest"] = t_bt.elapsed

    if trade_df.empty:
        log.error("No trades generated -- check data and Week-1 logic. Exiting.")
        return

    # -- Stage 4: Statistics -------------------------------------------------
    with Timer("Statistics") as t_stats:
        stats = compute_statistics(trade_df)
    timings["Statistics"] = t_stats.elapsed

    # -- Stage 5: Charts -----------------------------------------------------
    with Timer("Charts") as t_charts:
        equity_path, dd_path = save_equity_chart(
            stats["equity_curve"], stats["drawdown_series"], CHART_DIR
        )
    timings["Charts"] = t_charts.elapsed

    # -- Stage 6: Excel Output -----------------------------------------------
    with Timer("Excel Write") as t_xl:
        write_excel(trade_df, stats, equity_path, dd_path, OUTPUT_EXCEL)
    timings["Excel Write"] = t_xl.elapsed

    # -- Final Report --------------------------------------------------------
    total_elapsed = time.perf_counter() - total_start

    print("\n" + "=" * 70)
    print("  TIMING REPORT")
    print("=" * 70)
    for stage, secs in timings.items():
        print(f"  {stage:<25}: {secs:6.2f}s")
    print(f"  {'TOTAL':<25}: {total_elapsed:6.2f}s")
    print("=" * 70)

    print("\n  KEY RESULTS")
    print("=" * 70)
    for k, v in stats["stats_summary"].items():
        print(f"  {k:<35}: {v}")
    print("=" * 70)

    print("\n  WIN / LOSS")
    for wl in stats["win_loss"]:
        print(
            f"  {wl['label']:<12}: {wl['trades']} trades | "
            f"{wl['wins']} wins ({wl['win_pct']:.1f}%) | "
            f"{wl['losses']} losses ({wl['loss_pct']:.1f}%)"
        )

    print("\n  MONTHLY P&L (NAV-indexed)")
    print(stats["monthly_pnl"].to_string(index=False))
    print()

    # -- Self-check assertions (real computed checks, not cosmetic Trues) ------
    # Each check is an actual assertion evaluated against trade_df and stats.
    # A FAIL here means something in the code is wrong, not just undocumented.
    print("\n  SELF-CHECK (computed assertions against trade data)")
    print("=" * 70)

    week1_set = set(pd.Timestamp(d) for d in week1_days)
    td_full   = stats["trade_df"]

    def check(label: str, result: bool, detail: str = ""):
        status = "PASS" if result else "FAIL"
        suffix = f"  ({detail})" if detail else ""
        print(f"  [{status}] {label}{suffix}")
        return result

    all_passed = True

    # 1. Every traded date is a Week-1 date
    traded_dates = set(pd.Timestamp(d) for d in trade_df["entry_date"].unique())
    non_week1 = traded_dates - week1_set
    all_passed &= check(
        "Only Week-1 days traded",
        len(non_week1) == 0,
        f"{len(non_week1)} non-Week1 dates found" if non_week1 else "all dates verified",
    )

    # 2. SL exits: exit_time > entry_time (no lookahead — stop can't fire at or before entry)
    sl_rows = trade_df[trade_df["exit_reason"] == "SL"]
    if not sl_rows.empty:
        lookahead_sl = sl_rows[sl_rows["exit_time"] <= sl_rows["entry_time"]]
        all_passed &= check(
            "SL exits are strictly after entry time",
            len(lookahead_sl) == 0,
            f"{len(lookahead_sl)} lookahead SL rows" if not lookahead_sl.empty else f"{len(sl_rows)} SL exits verified",
        )
    else:
        check("SL exits are strictly after entry time", True, "no SL exits in dataset")

    # 3. Time exits: exit_time == EXIT_TIME_STR (or TIME_FALLBACK)
    time_rows = trade_df[trade_df["exit_reason"].isin(["TIME", "TIME_FALLBACK"])]
    wrong_time = time_rows[time_rows["exit_time"] > EXIT_TIME_STR]
    all_passed &= check(
        "Time exits at or before 15:20:59",
        len(wrong_time) == 0,
        f"{len(wrong_time)} exits after 15:20" if not wrong_time.empty else f"{len(time_rows)} time exits verified",
    )

    # 4. Fixed lot size — no scaling
    all_passed &= check(
        "Position size fixed at LOT_SIZE throughout",
        bool((trade_df["quantity"] == LOT_SIZE).all()),
        f"all {len(trade_df)} rows have qty={LOT_SIZE}",
    )

    # 5. One row per (date, option_type) — no duplicate legs
    dup_legs = trade_df.duplicated(subset=["entry_date", "option_type"], keep=False)
    all_passed &= check(
        "No duplicate (date, option_type) rows",
        not dup_legs.any(),
        f"{dup_legs.sum()} duplicate leg rows found" if dup_legs.any() else "unique",
    )

    # 6. Gross P&L sign sanity: short P&L = entry_value - exit_value
    recomputed_pnl = (trade_df["entry_value"] - trade_df["exit_value"]).round(4)
    pnl_mismatch = (recomputed_pnl - trade_df["gross_pnl"]).abs() > 0.01
    all_passed &= check(
        "Gross P&L = entry_value - exit_value (short formula)",
        not pnl_mismatch.any(),
        f"{pnl_mismatch.sum()} mismatches" if pnl_mismatch.any() else "all verified",
    )

    # 7. Cumulative P&L is monotonically consistent (each row = prior + gross_pnl)
    expected_cum = trade_df["gross_pnl"].cumsum().round(4)
    cum_mismatch = (expected_cum - trade_df["cumulative_pnl"]).abs() > 0.01
    all_passed &= check(
        "Cumulative P&L is consistent (trade-wise running sum)",
        not cum_mismatch.any(),
        f"{cum_mismatch.sum()} mismatches" if cum_mismatch.any() else "all verified",
    )

    # 8. Available capital = STARTING_CAPITAL + cumulative_pnl at every row
    expected_cap = (STARTING_CAPITAL + trade_df["cumulative_pnl"]).round(4)
    cap_mismatch = (expected_cap - trade_df["available_capital"]).abs() > 0.01
    all_passed &= check(
        "Available capital = start_cap + cumulative_pnl",
        not cap_mismatch.any(),
        f"{cap_mismatch.sum()} mismatches" if cap_mismatch.any() else "all verified",
    )

    # 9. Max drawdown computed from running peak (not just min of NAV series)
    #    Verify: drawdown at each point = (nav - running_peak) / running_peak
    nav_s = pd.concat([pd.Series([100.0]), td_full["nav_index"]])
    peak_s = nav_s.cummax()
    dd_s   = (nav_s - peak_s) / peak_s * 100
    reported_maxdd = stats["stats_summary"]["Max Drawdown (%)"]
    computed_maxdd = round(dd_s.min(), 4)
    all_passed &= check(
        "Max drawdown from running peak",
        abs(reported_maxdd - computed_maxdd) < 0.01,
        f"reported={reported_maxdd:.4f}% computed={computed_maxdd:.4f}%",
    )

    # 10. SL stop price = entry * 1.5 (not bar close)
    if not sl_rows.empty:
        expected_sl_price = (sl_rows["entry_price"] * SL_MULTIPLIER).round(4)
        sl_price_ok = (expected_sl_price - sl_rows["exit_price"]).abs() < 0.01
        all_passed &= check(
            "SL exit_price = entry_price * 1.5 (not bar close)",
            sl_price_ok.all(),
            f"{(~sl_price_ok).sum()} rows with wrong SL price" if not sl_price_ok.all() else f"{len(sl_rows)} SL exits verified",
        )
    else:
        check("SL exit_price = entry_price * 1.5", True, "no SL exits")

    print("=" * 70)
    overall = "ALL CHECKS PASSED" if all_passed else "SOME CHECKS FAILED -- review output above"
    print(f"  Overall: {overall}")
    print("=" * 70 + "\n")

    print(f"  Output saved: {OUTPUT_EXCEL.resolve()}\n")


if __name__ == "__main__":
    main()
