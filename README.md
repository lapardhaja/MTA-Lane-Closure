# Maine Turnpike Authority – Lane Closure Dashboard

A web app for analyzing traffic volume data and scheduling lane closures on the Maine Turnpike. Upload CSV or Excel files, filter by location and direction, and view volume heatmaps, closure schedules, and projections.

## Features

- **Data upload** – CSV, .xls, or .xlsx with columns: Location, Direction, Date, Time, Volume, Lanes
- **Filters** – Location, direction, day of week (All, Mon–Thu, or specific day), date range, holidays
- **Lane Closure Table** – Volume heatmap by date and time with color-coded closure thresholds (2-lane and 3-lane segments)
- **Closure Schedule** – Closure begin/end times by week (single or double lane closure)
- **Projections** – Target-year projections with configurable growth rates per location
- **Export** – HTML, PDF, or Excel

## Quick Start

1. Open `index.html` in a browser (or serve the folder locally).
2. Upload a CSV or Excel file via **Upload File**.
3. Select **Location** and **Direction**.
4. Use filters as needed; the table updates automatically.

## Data Format

Expected columns (names are flexible, e.g. `Location` / `Loc` / `Segment`):

| Column   | Example   |
|----------|-----------|
| Location | ML 52-53  |
| Direction| NB, SB    |
| Date     | 1/1/2025  |
| Time     | 12:00 AM  |
| Volume   | 500       |
| Lanes    | 2 or 3    |

## Documentation

- **AGENTS.md** – Implementation details, projection logic, holiday rules, and invariants for agents and developers. Read before modifying projection or holiday behavior.

## Project Structure

```
Dashboard/
├── index.html      # Main app
├── date-detail.html# Volume-by-hour detail view (linked from table cells)
├── app.js          # Logic, parsing, rendering
├── styles.css      # Styles
├── mock-data.csv   # Sample data
├── assets/         # Logo and other assets
├── AGENTS.md       # Agent/dev documentation (projection, holidays, invariants)
└── README.md
```

## Local Development

No build step. Use any static server, e.g.:

```bash
# Python
python -m http.server 8080

# Node (npx)
npx serve .
```

Then open `http://localhost:8080`.

## Hosting (e.g. GitHub Pages)

1. Push the repo to GitHub.
2. **Settings → Pages** → Source: **Deploy from a branch**.
3. Branch: `main`, folder: `/` (root).
4. Site will be at `https://<username>.github.io/<repo-name>/`.

## License

Internal use – Maine Turnpike Authority / HNTB.
