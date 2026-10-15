# Tema ETFs intraday spreads dashboard

A GitHub Pages dashboard for the intraday bid/ask spread workbook. It refreshes every 5 minutes.

The numbers come from Bloomberg (`BDH`/`BDP`), which only runs on a machine with a logged-in
Terminal. GitHub can host and display the dashboard, but it cannot calculate it. So a small script
on the Bloomberg machine publishes the data and GitHub serves it:

```
Bloomberg machine ── publish.py (every 5 min) ──► `data` branch: data/spreads.json
                                                          │
GitHub Pages (main branch): index.html ◄── fetches it every 5 min
```

## Files

| File | Purpose |
|---|---|
| `index.html` | The dashboard. Single file, no build step, no libraries. |
| `data/spreads.json` | Starter data (the Oct 14 session from your workbook, marked as sample). |
| `publish.py` | Runs on the Bloomberg machine. Pulls bars, computes spreads, pushes JSON. |
| `config.json` | ETF tickers and LMMs. Edit this to add or remove ETFs. |
| `export_from_workbook.py` | Alternative publisher that reads values saved by Excel (option B). |
| `run_publisher.bat` | Starts the publisher loop on Windows. |
| `requirements.txt` | Python dependencies. |

## Setup

### 1. Put the dashboard on GitHub Pages

1. Create a repo and push everything in this folder to `main`.
2. **Settings, Pages**: deploy from branch `main`, folder `/ (root)`.
3. Open `https://<you>.github.io/<repo>/`. It works immediately with the sample data.

### 2. Create the data branch (on the Bloomberg machine)

The publisher pushes to its own `data` branch in a separate clone. Pushing there does not trigger a
Pages rebuild (GitHub soft-limits Pages to about 10 builds an hour, and you will push 12).

```
git clone https://github.com/<you>/<repo>.git tema-spreads-dashboard-data
cd tema-spreads-dashboard-data
git checkout --orphan data
git rm -rf .
```

Make sure `git push` works there without a password prompt (Git Credential Manager or a deploy key
with write access).

### 3. Install and test the publisher

```
pip install -r requirements.txt
pip install --index-url=https://blpapi.bloomberg.com/repository/releases/python/simple/ blpapi

python publish.py --once --no-push --out test.json                          # check the numbers locally
python publish.py --once --repo-dir ..\tema-spreads-dashboard-data          # first real push
```

### 4. Run it all day

Edit the path in `run_publisher.bat` and run it, or schedule it in Task Scheduler: weekdays at
9:25 AM ET, action `run_publisher.bat`. The loop publishes at each 5-minute mark during market
hours (plus a final run after the close) and idles otherwise. If the machine's clock is not Eastern,
the schedule time is what needs adjusting; the script itself works in ET internally.

### 5. Point the dashboard at live data

In `index.html`, set `githubRepo: "<you>/<repo>"` in `CONFIG` and commit. The page then reads
`data/spreads.json` from the `data` branch. If many viewers share one office IP address, use
`dataUrl` instead (see the comment in `CONFIG`).

## Option B: keep Excel as the calculation engine

If you would rather keep the workbook's own formulas, leave it open on the Bloomberg machine and add
this macro (run it once, or call it from `Workbook_Open`). It saves a copy every 5 minutes and
exports it:

```vba
Sub SaveAndExport()
    ThisWorkbook.SaveCopyAs "C:\spreads\snapshot.xlsx"
    Shell "cmd /c cd /d C:\spreads\tema-spreads-dashboard && python export_from_workbook.py " & _
          "C:\spreads\snapshot.xlsx --repo-dir C:\spreads\tema-spreads-dashboard-data", vbHide
    Application.OnTime Now + TimeValue("00:05:00"), "SaveAndExport"
End Sub
```

This assumes the machine's clock is US Eastern, because Excel stores local times.

## How the dashboard calculates things

Same definitions as the workbook. Current is the latest 5-minute bucket, Open is the first bucket,
and Session avg, Min and Max are over all buckets so far. "vs open" is Current minus Open.
Average, tightest, widest and median are across ETFs' current values.

Each bucket is labelled by its end time (9:35 covers 9:30 to 9:35), and an empty bucket repeats the
previous spread, as your sheets do.

## Troubleshooting

| What you see | Likely cause |
|---|---|
| "Stale" chip and warning banner | The publisher has not pushed for over 10 minutes: machine asleep, Terminal logged out, or script stopped. Check `publisher.log`. |
| "Could not load data: file not found" | `githubRepo` or branch is wrong, or the first push has not happened yet. |
| "GitHub rate limit" error | Too many viewers on one IP. Switch to `dataUrl`. |
| "Could not start a Bloomberg session" | Terminal is not running or not logged in on that machine. |
| Chart empty for an ETF | No bars yet (early in the session) or the ticker is wrong in `config.json`. |
| Nothing updates on a holiday | Expected. Bloomberg returns no bars, so the file is left alone. |

## Before you publish

- **Data licence.** Bloomberg data is usually restricted from redistribution. Check with your
  Bloomberg administrator before putting it on a public site.
- **Visibility.** GitHub Pages sites are public unless you use GitHub Enterprise Cloud with access
  control. A private repo does not make the page private, and the page should never contain an access token.
