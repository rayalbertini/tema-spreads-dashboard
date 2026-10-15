#!/usr/bin/env python3
"""
Publish Tema ETF intraday bid/ask spreads to GitHub.

Run this on the machine that has a Bloomberg Terminal (logged in). Every five
minutes it:

  1. pulls 1-minute BID and ASK bars for each ETF from the Bloomberg API (blpapi),
  2. averages them into 5-minute buckets and takes ask minus bid
     (the same calculation as the BDH formulas in the Excel workbook),
  3. writes data/spreads.json,
  4. commits and pushes it to the `data` branch of your repo.

The dashboard (index.html on GitHub Pages) reads that file.

Usage
-----
  python publish.py --repo-dir ../tema-spreads-dashboard-data --loop     # run all day
  python publish.py --repo-dir ../tema-spreads-dashboard-data --once     # single run
  python publish.py --once --no-push --out data/spreads.json             # local test, no git
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import subprocess
import sys
import tempfile
import time
from datetime import date, datetime, time as dtime, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

ET = ZoneInfo("America/New_York")
UTC = timezone.utc
BUCKET_MIN = 5
SESSION_OPEN = dtime(9, 30)
SESSION_CLOSE = dtime(16, 0)
PUBLISH_UNTIL = dtime(16, 10)   # keep publishing a little after the close so the 16:00 bucket lands
BAR_SETTLE_SECONDS = 20         # wait after each 5-minute boundary before asking Bloomberg

log = logging.getLogger("publish")


# --------------------------------------------------------------------------- #
# Session / calendar helpers
# --------------------------------------------------------------------------- #
def session_window(now_et: datetime) -> tuple[date, datetime, datetime]:
    """Return (trading_date, open_utc, close_utc) for the session to publish.

    Before the open, or on a weekend, this falls back to the most recent weekday.
    Exchange holidays are not modelled: on a holiday Bloomberg returns no bars,
    and the publisher leaves the existing file untouched.
    """
    d = now_et.date()
    if now_et.time() < SESSION_OPEN:
        d -= timedelta(days=1)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    open_et = datetime.combine(d, SESSION_OPEN, tzinfo=ET)
    close_et = datetime.combine(d, SESSION_CLOSE, tzinfo=ET)
    return d, open_et.astimezone(UTC), close_et.astimezone(UTC)


def in_publish_window(now_et: datetime) -> bool:
    return now_et.weekday() < 5 and SESSION_OPEN <= now_et.time() <= PUBLISH_UNTIL


# --------------------------------------------------------------------------- #
# Bloomberg access (blpapi is imported lazily so the helpers below stay usable
# on machines without the Bloomberg SDK, e.g. export_from_workbook.py)
# --------------------------------------------------------------------------- #
def open_bloomberg(host: str, port: int):
    try:
        import blpapi
    except ImportError:
        sys.exit(
            "blpapi is not installed. Install it from Bloomberg's index:\n"
            "  pip install --index-url=https://blpapi.bloomberg.com/repository/releases/python/simple/ blpapi"
        )
    opts = blpapi.SessionOptions()
    opts.setServerHost(host)
    opts.setServerPort(port)
    session = blpapi.Session(opts)
    if not session.start():
        raise RuntimeError(
            f"Could not start a Bloomberg session on {host}:{port}. "
            "Is the Terminal running and logged in on this machine?"
        )
    if not session.openService("//blp/refdata"):
        session.stop()
        raise RuntimeError("Could not open //blp/refdata.")
    return session


def _responses(session):
    """Yield every message until the final RESPONSE event arrives."""
    import blpapi

    while True:
        event = session.nextEvent(15000)
        if event.eventType() == blpapi.Event.TIMEOUT:
            raise TimeoutError("Timed out waiting for Bloomberg.")
        for msg in event:
            if msg.hasElement("responseError"):
                raise RuntimeError(f"Bloomberg error: {msg.getElement('responseError')}")
            yield msg
        if event.eventType() == blpapi.Event.RESPONSE:
            return


def fetch_names(session, securities: list[str]) -> dict[str, str]:
    """Fund names via ReferenceDataRequest (same as =BDP(..., "NAME") in the workbook)."""
    svc = session.getService("//blp/refdata")
    req = svc.createRequest("ReferenceDataRequest")
    for sec in securities:
        req.append("securities", sec)
    req.append("fields", "NAME")
    session.sendRequest(req)
    names: dict[str, str] = {}
    for msg in _responses(session):
        if not msg.hasElement("securityData"):
            continue
        for sd in msg.getElement("securityData").values():
            sec = sd.getElementAsString("security")
            fields = sd.getElement("fieldData")
            names[sec] = fields.getElementAsString("NAME") if fields.hasElement("NAME") else ""
    return names


def fetch_bars(session, security: str, event_type: str,
               start_utc: datetime, end_utc: datetime) -> dict[datetime, float]:
    """1-minute bars of BID or ASK; returns {bar_start_utc: open_price}."""
    svc = session.getService("//blp/refdata")
    req = svc.createRequest("IntradayBarRequest")
    req.set("security", security)
    req.set("eventType", event_type)          # "BID" or "ASK"
    req.set("interval", 1)                    # minutes
    # Bloomberg treats naive datetimes as GMT.
    req.set("startDateTime", start_utc.replace(tzinfo=None))
    req.set("endDateTime", end_utc.replace(tzinfo=None))
    session.sendRequest(req)

    bars: dict[datetime, float] = {}
    for msg in _responses(session):
        if not msg.hasElement("barData"):
            continue
        for bar in msg.getElement("barData").getElement("barTickData").values():
            t = bar.getElementAsDatetime("time")
            t = t.replace(tzinfo=UTC) if t.tzinfo is None else t.astimezone(UTC)
            bars[t] = bar.getElementAsFloat("open")
    return bars


# --------------------------------------------------------------------------- #
# Spread calculation
# --------------------------------------------------------------------------- #
def five_minute_spread(bid: dict, ask: dict, open_utc: datetime,
                       close_utc: datetime, now_utc: datetime):
    """Ask minus bid per 5-minute bucket, labelled by bucket END time (UTC).

    Bars from 9:30-9:34 land in the bucket labelled 9:35, and so on. A bucket
    with no data repeats the previous bucket's spread, like the workbook does.
    Buckets that have not started yet are left out.
    """
    import pandas as pd

    def bucket(bars: dict):
        if not bars:
            return pd.Series(dtype="float64")
        s = pd.Series(list(bars.values()), index=pd.DatetimeIndex(list(bars.keys()))).sort_index()
        return s.resample(f"{BUCKET_MIN}min", origin=pd.Timestamp(open_utc),
                          closed="left", label="right").mean()

    spread = (bucket(ask) - bucket(bid)).dropna()
    if spread.empty:
        return spread

    step = pd.Timedelta(minutes=BUCKET_MIN)
    grid = pd.date_range(pd.Timestamp(open_utc) + step, pd.Timestamp(close_utc), freq=f"{BUCKET_MIN}min")
    grid = grid[(grid - step) <= pd.Timestamp(now_utc)]
    return spread.reindex(grid).ffill().dropna()


def build_payload(session, config: dict, names_cache: dict[str, str],
                  now_utc: datetime) -> dict | None:
    trading_date, open_utc, close_utc = session_window(now_utc.astimezone(ET))
    end_utc = min(now_utc, close_utc)
    suffix = config.get("suffix", "US Equity")
    tickers = config["tickers"]
    securities = {t["ticker"]: f"{t['ticker']} {suffix}" for t in tickers}

    missing = [s for s in securities.values() if s not in names_cache]
    if missing:
        names_cache.update(fetch_names(session, missing))

    etfs: dict[str, dict] = {}
    total_points = 0
    for t in tickers:
        tk = t["ticker"]
        sec = securities[tk]
        try:
            ask = fetch_bars(session, sec, "ASK", open_utc, end_utc)
            bid = fetch_bars(session, sec, "BID", open_utc, end_utc)
            spread = five_minute_spread(bid, ask, open_utc, close_utc, now_utc)
        except Exception as exc:  # keep going; one bad ticker should not blank the dashboard
            log.warning("%s: %s", tk, exc)
            spread = None
        series = []
        if spread is not None:
            series = [
                [ts.to_pydatetime().astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"), round(float(v), 4)]
                for ts, v in spread.items()
            ]
        total_points += len(series)
        etfs[tk] = {
            "name": names_cache.get(sec, ""),
            "lmm": t.get("lmm", ""),
            "series": series,
        }

    if total_points == 0:
        log.warning("No bars returned for any ETF (holiday or market closed?). Not overwriting the file.")
        return None

    return {
        "updated": now_utc.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "sessionDate": trading_date.isoformat(),
        "bucketMinutes": BUCKET_MIN,
        "order": [t["ticker"] for t in tickers],
        "etfs": etfs,
    }


# --------------------------------------------------------------------------- #
# Output + git
# --------------------------------------------------------------------------- #
def write_json_atomic(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(payload, f, separators=(",", ":"))
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)


def _git(repo: Path, *args: str, check: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(["git", "-C", str(repo), *args], check=check,
                          capture_output=True, text=True)


def git_publish(repo: Path, relpath: str, branch: str, message: str, keep_history: bool) -> None:
    """Commit relpath and push it to origin/<branch>.

    By default the data branch is kept to a single commit (amend + force push),
    so 78 updates a day do not bloat the repo. Force-pushing is refused unless
    the working copy is on a dedicated data branch.
    """
    current = _git(repo, "symbolic-ref", "--short", "HEAD").stdout.strip()
    if current != branch:
        raise RuntimeError(f"{repo} is on branch '{current}', expected '{branch}'. "
                           "Use a separate clone checked out on the data branch.")

    _git(repo, "add", relpath)
    if not _git(repo, "status", "--porcelain", "--", relpath).stdout.strip():
        log.info("Nothing changed; skipping commit.")
        return

    has_commits = _git(repo, "rev-parse", "--verify", "HEAD", check=False).returncode == 0
    force = False
    if has_commits and not keep_history:
        if branch in ("main", "master"):
            raise RuntimeError("Refusing to amend and force-push main/master.")
        _git(repo, "commit", "--amend", "-m", message)
        force = True
    else:
        _git(repo, "commit", "-m", message)

    push = ["push", "origin", f"HEAD:refs/heads/{branch}"]
    if force:
        push.insert(1, "--force-with-lease")
    result = _git(repo, *push, check=False)
    if result.returncode != 0:
        raise RuntimeError(f"git push failed: {result.stderr.strip()}")
    log.info("Pushed to origin/%s.", branch)


# --------------------------------------------------------------------------- #
# Main loop
# --------------------------------------------------------------------------- #
def run_once(session, config, names_cache, out_path: Path, args) -> bool:
    now_utc = datetime.now(UTC).replace(microsecond=0)
    payload = build_payload(session, config, names_cache, now_utc)
    if payload is None:
        return False
    write_json_atomic(out_path, payload)
    n = sum(len(e["series"]) for e in payload["etfs"].values())
    log.info("Wrote %s (%d ETFs, %d points).", out_path, len(payload["etfs"]), n)
    if not args.no_push:
        rel = os.path.relpath(out_path, args.repo_dir)
        git_publish(Path(args.repo_dir), rel, args.branch,
                    f"spreads {now_utc:%Y-%m-%d %H:%M}Z", args.keep_history)
    return True


def seconds_until_next_boundary(now: datetime) -> float:
    minute = (now.minute // BUCKET_MIN + 1) * BUCKET_MIN
    nxt = now.replace(minute=0, second=0, microsecond=0) + timedelta(minutes=minute)
    return max(1.0, (nxt - now).total_seconds() + BAR_SETTLE_SECONDS)


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    mode = p.add_mutually_exclusive_group()
    mode.add_argument("--once", action="store_true", help="publish one update and exit (default)")
    mode.add_argument("--loop", action="store_true",
                      help="keep running; publish every 5 minutes during market hours")
    p.add_argument("--config", default=str(Path(__file__).with_name("config.json")))
    p.add_argument("--repo-dir", default=".", help="clone of the repo checked out on the data branch")
    p.add_argument("--out", default=None, help="output file (default: <repo-dir>/data/spreads.json)")
    p.add_argument("--branch", default="data")
    p.add_argument("--no-push", action="store_true", help="write the file but skip git")
    p.add_argument("--keep-history", action="store_true",
                   help="normal commits instead of amend + force push")
    p.add_argument("--host", default="localhost")
    p.add_argument("--port", type=int, default=8194)
    p.add_argument("--log-file", default=None)
    args = p.parse_args()

    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.StreamHandler()] + ([logging.FileHandler(args.log_file)] if args.log_file else []),
    )
    config = json.loads(Path(args.config).read_text(encoding="utf-8"))
    out_path = Path(args.out) if args.out else Path(args.repo_dir) / "data" / "spreads.json"
    names_cache: dict[str, str] = {}
    session = None

    def with_session():
        nonlocal session
        if session is None:
            session = open_bloomberg(args.host, args.port)
        return session

    def attempt() -> None:
        nonlocal session
        try:
            run_once(with_session(), config, names_cache, out_path, args)
        except Exception:
            log.exception("Update failed; will retry on the next cycle.")
            if session is not None:
                try:
                    session.stop()
                except Exception:
                    pass
                session = None

    if not args.loop:
        attempt()
        return

    log.info("Publisher running. Press Ctrl+C to stop.")
    attempt()  # seed immediately so the dashboard is never empty
    try:
        while True:
            time.sleep(seconds_until_next_boundary(datetime.now(UTC)))
            if in_publish_window(datetime.now(ET)):
                attempt()
            else:
                log.debug("Outside market hours; idle.")
    except KeyboardInterrupt:
        log.info("Stopped.")


if __name__ == "__main__":
    main()
