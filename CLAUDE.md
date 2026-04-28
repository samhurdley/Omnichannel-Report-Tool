# Omnichannel Reporting Tool

Generates branded, interactive HTML slideshow presentations from Trade Desk omnichannel CSV exports.

## Architecture

```
CSV export + client_config.xlsx → process_report.py → HTML slideshow report
```

The Streamlit app (`app.py`) wraps this for a drag-and-drop UI with Google Sheets history tracking.

## Key Files

| File | Purpose |
|------|---------|
| `process_report.py` | Core engine: CSV parsing, metric calculation, HTML generation (~2,100 lines) |
| `app.py` | Streamlit web UI; handles file uploads, Google Sheets backend, ZIP download |
| `client_config.xlsx` | Maps client names to their conversion/traffic/revenue column names |

## Dev Commands

```bash
# Run the Streamlit UI
streamlit run app.py

# Generate preview reports from real test data (see Preview section below)
python process_report.py --test

# Process a single CSV directly
python process_report.py path/to/report.csv
```

## Output Preview

Running `python process_report.py --test` generates two HTML files at `~/Desktop/Reporting/`:

- `high_conv_test_report.html` — ≥20 conversions, CPA visible everywhere, upsell block fires
- `low_conv_test_report.html` — 5 conversions, Cost Per Visit replaces CPA throughout

**Requires:** Real Meatbox CSV and `client_config.xlsx` must exist at `~/Desktop/Reporting/`. If those files aren't present the command will silently skip (guarded by the post-edit hook).

A Claude Code hook automatically runs this test and opens the preview in your browser after every edit to `process_report.py` or `app.py` — as long as the CSV exists at the path above.

## Report Logic

- **High-conv mode** (≥20 conversions): shows CPA, Conversions sparkline, CPA benchmarks
- **Low-conv mode** (<20 conversions): shows Cost Per Visit, Site Traffic sparkline
- **Upsell block**: appears when current-month CPA improves >15% vs 2-month rolling average from history
- **MoM trends**: green/red arrows based on whether metric movement is favourable (direction varies by metric)

## Client Config (`client_config.xlsx`)

- **Config sheet**: maps client name → conversion column(s), site traffic column, revenue column
- **History sheet**: one row per client per month; used for MoM sparklines and upsell detection

## GitHub

Remote: `https://github.com/samhurdley/Omnichannel-Report-Tool`

```bash
git add -p          # stage changes selectively
git commit -m "..."
git push
```
