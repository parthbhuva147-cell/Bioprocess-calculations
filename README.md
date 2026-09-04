# BioProcess Calc

A fully static, free-to-host landing website with working browser-based bioprocess calculators.

## Included

- Specific growth rate (mu)
- Doubling time
- Biomass yield Yx/s
- Product yield Yp/s
- Volumetric productivity
- Carbon in from glucose
- Carbon out from CO2
- CSV upload and start/end batch analysis
- Simple process trend plot
- Summary CSV export
- No database and no backend required

## CSV format

Use this header:

time,biomass,substrate,product

Recommended units:
- time = h
- biomass = g/L
- substrate = g/L
- product = g/L

## Run locally

Open `index.html` directly in your browser.

For best results, you can also use a simple local server:

python -m http.server 8000

Then visit:

http://localhost:8000

## Publish free with GitHub Pages

1. Create a GitHub account.
2. Create a new public repository, e.g. `bioprocess-calc`.
3. Upload `index.html`, `styles.css`, and `app.js`.
4. Open repository Settings → Pages.
5. Under "Build and deployment", select "Deploy from a branch".
6. Choose the `main` branch and `/root`.
7. Save.
8. GitHub will give you a free public website URL.

## Important scientific limitation

This is a prototype using simplified interval calculations. It does not correct for:
- changing reactor volume
- sampling losses
- fed-batch substrate additions
- evaporation
- gas temperature/pressure normalization
- inlet CO2 subtraction
- biomass carbon composition
- multi-phase process boundaries
- measurement uncertainty

Those are good candidates for the next SaaS version.
