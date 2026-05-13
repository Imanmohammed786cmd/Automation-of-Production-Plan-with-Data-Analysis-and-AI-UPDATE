# Production Consumption Planner

Static webpage for uploading production quantity files and generating machine/material consumption from BOM data.

## GitHub Pages

GitHub Pages can host the frontend files:

- `index.html`
- `styles.css`
- `app.js`
- `data.js`
- `production_qty_template.csv`
- `production_qty_template.xlsx`

The Python server file `server.py` is for local Excel export only. GitHub Pages does not run Python, so on GitHub Pages the app falls back to browser CSV exports when Excel server export is unavailable.

## Local Server

Run locally for full Excel export:

```bash
python3 server.py
```

Then open:

```text
http://127.0.0.1:4175/index.html
```

## Deploy Steps

1. Create a new GitHub repository.
2. Push this folder to the repository.
3. In GitHub, open Settings > Pages.
4. Set Source to `Deploy from a branch`.
5. Choose branch `main` and folder `/root`.
6. Save and wait for the GitHub Pages URL.

