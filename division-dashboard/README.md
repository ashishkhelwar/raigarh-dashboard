# Raigarh Division Dashboard

A React + Vite dashboard for **Raigarh Forest Division** — reports, statistics, and
APO planning data drawn from the Working Plan 2020-21 → 2029-30. It renders
range-wise distribution, forest composition / site quality, the 10-year watershed
treatment schedule, and Forest Rights Act status, plus a document generator that
produces downloadable Word (`.docx`) reports in-browser.

> This is a separate app from the root `index.html` plantation-monitoring dashboard.
> Nothing here touches that page.

## Tech stack

- [Vite](https://vitejs.dev/) — dev server & build
- [React 18](https://react.dev/)
- [Recharts](https://recharts.org/) — charts
- [lucide-react](https://lucide.dev/) — icons
- [docx](https://docx.js.org/) — client-side Word report generation

## Getting started

```bash
cd division-dashboard
npm install
npm run dev      # start the dev server (http://localhost:5173)
```

## Build

```bash
npm run build    # outputs static site to division-dashboard/dist/
npm run preview  # preview the production build locally
```

The build uses a relative `base` (`./`), so the contents of `dist/` can be served
from a domain root or a subpath (e.g. a GitHub Pages project site) without further
configuration.

## Project layout

```
division-dashboard/
├── index.html                 # Vite HTML entry
├── vite.config.js
├── package.json
└── src/
    ├── main.jsx               # React entry point
    ├── index.css              # minimal reset
    └── DivisionDashboard.jsx  # the dashboard + docx report generator
```

## Data

All figures live as constants at the top of `src/DivisionDashboard.jsx`
(`DIVISION`, `RANGE_DATA`, `WC_COUNTS`, `SITE_QUALITY`, `WATERSHED_SCHEDULE`, `FRA`).
Update those objects to refresh the dashboard; totals are derived automatically.
