# AIPO holdings treemap

Asks Defiance's public download endpoint for the current dated AIPO XLSX workbook, writes `aipo.txt`, and fetches the latest available **regular-hours one-minute bar** for every equity symbol with yfinance (falling back to daily bars when needed). It writes an interactive, searchable `aipo.html` treemap. Rectangle area uses Defiance's percentage weight, green/red compares Yahoo's latest bar with the prior trading session's final bar, and missing quotes appear gray with `N/A`. The source date and quote time are displayed in the chart.

## Use locally

Requires Python 3.12 (3.11 should also work).

```bash
python -m pip install -r requirements.txt
python generate_aipo.py
```

To process a saved Defiance export without downloading it again:

```bash
python generate_aipo.py --source /path/to/aipo-09-28-2026.xlsx
```

Open `aipo.html` in a browser. D3.js loads from jsDelivr, so the browser needs internet access. The text file is tab separated: `symbol`, `company`, and `weight_pct` (percentage points). Each published ticker is kept. Dot or slash tickers become Yahoo's dash form. No weight scaling is applied. The money-market sweep and `Cash & Other` rows are excluded.

Quotes come from yfinance one-minute regular-hours bars and, for tickers without intraday data, daily closing bars. The page shows the timestamp of the latest bar (Pacific time) or the date of a daily fallback. Requests run in batches of 10, pause two seconds between batches, and back off for 30 then 60 seconds when a batch returns nothing. Use `--quote-batch-size`, `--quote-delay`, and `--quote-retry-delay` to tune that pacing. If Yahoo does not return a quote for a few stocks, their tiles remain visible as `N/A`. If a full batch keeps failing or more than 10% of quotes fail, generation stops and leaves the previous outputs in place. Yahoo may still rate limit shared runners; rerun the workflow later if that happens.

## Publish on GitHub

1. Make a new GitHub repository with `main` as its default branch. Copy **all** project files and folders to its root and push them, including `.github/workflows/update-aipo.yml` and `treemap_template.html`.
2. Under **Settings → Pages → Build and deployment**, set **Source** to **GitHub Actions**.
3. Under **Settings → Actions → General**, allow workflows and grant **Read and write permissions** to `GITHUB_TOKEN` if your repository policy requires it. The workflow itself requests `contents: write`, `pages: write`, and `id-token: write`.
4. Open **Actions → Refresh AIPO holdings map → Run workflow** for the first run. Successful runs commit `aipo.txt` and `aipo.html` and deploy the chart directly to Pages. The site URL appears in **Settings → Pages** and in the workflow's deployment job. For a normal project repository it is `https://YOUR_USERNAME.github.io/YOUR_REPO/aipo.html` (the root URL also opens the chart).

The workflow runs daily at **6:35 AM America/Los_Angeles**, with daylight saving handled by GitHub. Runs can be delayed by GitHub's scheduler. A job will fail before committing or publishing if the download, parsing, or most quote lookups fail. Defiance's source data can lag the market on weekends and holidays.

The included starter `aipo.html` and `aipo.txt` were built from Defiance's official September 28, 2026 workbook. Prices are marked `N/A`. The first successful workflow run replaces both with the latest holdings and current Yahoo prices.

Sources: [Defiance AIPO full-holdings page](https://www.defianceetfs.com/aipo-full-holdings/), [GitHub schedule syntax](https://docs.github.com/en/actions/reference/workflows-and-actions/workflow-syntax#onschedule), [Pages publishing setup](https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site).
