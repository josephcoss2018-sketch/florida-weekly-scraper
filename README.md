# Florida Weekly Scraper

Automated weekly scraper for the Florida **2026 General Election Candidate Report**.

**Source:** https://dos.elections.myflorida.com/candidates/

## What it does

- Fetches all candidates from the 2026 General Election listing
- Fetches each candidate detail page (20 parallel workers)
- Decodes Cloudflare-obfuscated email addresses
- Exports a formatted Excel spreadsheet to `florida_reports/florida_candidates_YYYYMMDD.xlsx`
- Automatically stops running after September 1, 2026

## Schedule

Runs every **Monday at 13:00 UTC** via GitHub Actions.
Can also be triggered manually via the Actions tab.

## Output

- **Sheet 1**: All candidates with 14 columns, frozen header, auto-filter, sorted by Office/District/Name
- **Sheet 2**: Summary with party, status, and office breakdowns

## Columns

| Column | Description |
|---|---|
| Account | Florida DOE account number |
| Name | Candidate full name |
| Office | Office being sought |
| District | District number (if applicable) |
| Party | Political party affiliation |
| Incumbent | "Yes" if incumbent |
| Address | Mailing address |
| Phone | Campaign phone number |
| Email | Campaign email (CF-decoded) |
| Website | Campaign website |
| Status | Active / Qualified / Withdrawn |
| Date Filed | Date candidacy was filed |
| Date Qualified | Date candidate qualified |
| Method | Qualifying method |

## Local usage

```bash
pip install -r requirements.txt
python florida_weekly_scraper.py
```
