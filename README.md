# Lead Scraper

![Python](https://img.shields.io/badge/Python-3-3776AB?logo=python&logoColor=white)
![Google Places API](https://img.shields.io/badge/Google%20Places%20API-4285F4?logo=googlemaps&logoColor=white)

> Pull business leads from the Google Places API across a geographic grid, deduplicate them, and export ready-to-use prospect lists.

Useful for building local outreach lists: point it at an area and a set of search terms, and it collects matching businesses with their contact details.

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Environment Variables](#environment-variables)
- [Usage](#usage)
- [Output](#output)
- [License](#license)

## Features

- Sweeps a latitude and longitude range (and/or a list of cities), querying the Places Text Search API for your terms.
- Captures name, rating, review count, website, and phone number for each business.
- Tracks already-seen places in `seen.json`, so re-runs only add new leads.
- Exports to `leads_data.json` and to Excel via `openpyxl`.
- Optional scheduled runs via the `schedule` library for continuous collection.

## Prerequisites

- Python 3.9 or newer
- A Google Cloud project with the **Places API** enabled

## Installation

```bash
git clone https://github.com/preciousg10/lead-scraper.git
cd lead-scraper
pip install requests beautifulsoup4 openpyxl schedule
```

## Environment Variables

See `.env.example`. Set your Google Places API key before running:

```bash
export GOOGLE_PLACES_API_KEY=your-key-here
# PowerShell:
$env:GOOGLE_PLACES_API_KEY = "your-key-here"
```

## Usage

```bash
python search.py
```

Tune the search area and terms via the constants near the top of `search.py` (latitude and longitude bounds, search terms, and city list).

## Output

| File | Contents |
|------|----------|
| `leads_data.json` | All collected leads, deduplicated |
| `seen.json` | IDs already seen, used for incremental runs |

## License

© 2026 Precious G. All rights reserved. This repository is public for viewing and portfolio purposes only; please do not copy, reuse, or redistribute the code without permission.
