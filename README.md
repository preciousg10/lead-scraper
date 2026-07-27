# Lead Scraper

A Python script that pulls business leads from the **Google Places API** across a
geographic grid, de-duplicates them, and exports the results, useful for building
local outreach/prospecting lists.

## What it does

- Sweeps a latitude/longitude range (and/or a list of cities) and queries the Google
  Places **Text Search** API for businesses matching your search terms.
- Collects name, rating, review count, website, and phone number for each place.
- Tracks already-seen places in `seen.json`, so re-runs only add new leads.
- Writes results to `leads_data.json` and can export to Excel (via `openpyxl`).
- Can run on a schedule (via the `schedule` library) for continuous collection.

## Stack

Python 3 · `requests` · `beautifulsoup4` · `openpyxl` · `schedule`

## Setup

```bash
pip install requests beautifulsoup4 openpyxl schedule
```

Set your Google Places API key (see `.env.example`):

```bash
export GOOGLE_PLACES_API_KEY=your-key-here
# PowerShell:  $env:GOOGLE_PLACES_API_KEY = "your-key-here"
```

> Requires a Google Cloud project with the **Places API** enabled.

## Run

```bash
python search.py
```

Results are written to `leads_data.json`.
