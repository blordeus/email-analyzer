# email-analyzer

A CLI tool that analyzes email campaign CSV exports from any platform — Mailchimp, Klaviyo, ConvertKit, and others. Auto-detects column names and exports a multi-sheet Excel report with embedded charts.

---

## Features

- **Auto-detects columns** — works with exports from Mailchimp, Klaviyo, ConvertKit, and more
- **Interactive column mapping** — `--map` flag for manual assignment when auto-detection falls short
- **Key metrics** — open rate, CTR, click-to-open rate, unsubscribe rate, bounce rate
- **Trend analysis** — monthly aggregates with best/worst campaign breakdowns
- **Excel report** — 5 sheets with color-coded performance and 3 embedded charts
- **Industry benchmarks** — open rate and CTR columns color-coded against standard thresholds

---

## Setup

**1. Clone the repo**
```bash
git clone https://github.com/blordeus/email-analyzer.git
cd email-analyzer
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

---

## Usage

```bash
# Analyze a CSV export
python email_analyzer.py --file campaigns.csv

# Custom output filename
python email_analyzer.py --file campaigns.csv --output q1_report.xlsx

# Manually map columns if auto-detection misses any
python email_analyzer.py --file campaigns.csv --map

# Try the included sample
python email_analyzer.py --file sample_campaigns.csv
```

---

## CSV Format

Your CSV needs at minimum these columns (exact names vary by platform — the tool handles common variants automatically):

| Field | Mailchimp | Klaviyo | ConvertKit |
|-------|-----------|---------|------------|
| Emails sent | `emails_sent` | `recipients` | `total_sent` |
| Emails opened | `emails_opened` | `unique_opens` | `opens` |
| Clicks | `unique_clicks` | `unique_link_clicks` | `clicks` |
| Unsubscribes | `unsubscribes` | `unsubscribed` | `opt_outs` |
| Bounces | `bounces` | `total_bounces` | `hard_bounces` |

Optional but recommended: `campaign_name`, `send_date`

---

## Output

### Terminal
```
📧 Email Campaign Analyzer
----------------------------------------
  📂 Loaded: campaigns.csv (36 rows)
  ✅ Auto-detected: campaign_name, send_date, emails_sent, emails_opened, ...

  Total Campaigns              36
  Avg Open Rate (%)            32.98
  Avg CTR (%)                  2.33
  Best Open Rate (%)           44.8
  ...
```

### Excel Report (5 sheets)

| Sheet | Contents |
|-------|----------|
| All Campaigns | Every campaign with color-coded open rate and CTR |
| Monthly Trends | Aggregated monthly stats + 3 charts |
| Top 10 Campaigns | Best performers by open rate |
| Bottom 10 Campaigns | Worst performers by open rate |
| Summary | Overall stats and date range |

### Color Coding (Open Rate)
- 🟢 Green — ≥ 30% (above average)
- 🟡 Yellow — 20–30% (industry average)
- 🔴 Red — < 20% (below average)

---

## Project Structure

```
email-analyzer/
├── email_analyzer.py     ← main script
├── sample_campaigns.csv  ← example data to test with
├── requirements.txt
├── .gitignore
└── README.md
```

---

## Tech Stack

- [pandas](https://pandas.pydata.org/) — data loading, calculation, aggregation
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel export, styling, charts
- Python standard library: `argparse`, `pathlib`

---

## License

MIT
