# BioProcess Calc V5 — Multi-file Process Analytics

This version is designed as a browser-based research/process-development workspace rather than a fixed-format calculator.

## Major capabilities

### Multi-file loading
Load multiple CSV or Excel files from the same fermentation at once:
- bioreactor process export
- OD/DCW sampling data
- residual glucose data
- HPLC product data
- feed/additions data

Each file gets:
- a role
- a selected time column
- its own time offset
- an optional forced acquisition interval

### Process-time handling
The process timeline is explicit and user-controlled.

For headers such as:

`TimeStamp(UTC+02:00 interval=5)`

the tool detects the acquisition interval. If Excel has corrupted ambiguous day/month timestamps and the timestamp span disagrees strongly with the logger interval, interval-derived elapsed time is used.

A forced interval can also be entered per file.

The user then defines the single analysis process window used for:
- KPIs
- gas integration
- carbon balance
- figures

### Data mapping
Signals can be selected from any uploaded file:
- DO
- pH
- temperature
- agitation
- air/O2/N2/CO2 inlet gas flows
- off-gas CO2 and O2
- OUR
- CER
- RQ
- feed
- OD600
- DCW
- residual glucose
- Product 1
- Product 2
- sample volume

### Process KPIs
Where data are available:
- process duration
- minimum DO
- mean pH
- mean temperature
- maximum agitation
- total feed
- max OD600
- max DCW
- final residual glucose
- final product titre
- approximate volumetric productivity
- estimated mu_max using local log-linear DCW fits with R² >= 0.95
- apparent concentration-based Yx/s and Yp/s

### CO2 calculation
Auto mode prefers CER if a usable non-zero CER signal exists.

CER integration:
`mol CO2 = integral[CER(t) * reactor_volume(t) dt]`

If CER is unavailable, an off-gas method can be used:
- off-gas CO2 concentration
- inlet gas flow
- optional off-gas O2

When off-gas O2 is available, outlet flow is estimated using an N2 balance.
Otherwise outlet flow is approximated from inlet flow and the result is clearly identified as approximate.

### Carbon ledger
Positive terms:
- initial glucose carbon
- feed glucose carbon
- gas-phase CO2 entering
- manual carbon additions

Negative terms:
- carbon removed in samples
- net CO2 carbon evolved
- manual removals

Final measured inventory:
- residual glucose carbon
- biomass carbon
- Product 1 carbon
- Product 2 carbon

The difference is reported as unaccounted carbon.

Default carbon fractions:
- glucose: 0.40001998 g C/g glucose
- biomass: 0.488 g C/g DCW (user editable)
- resveratrol-like Product 1 default: 0.73672 g C/g
- piceatannol-like Product 2 default: 0.68846 g C/g

## Important scientific limitations

This tool cannot make an incomplete measurement set into a closed mass balance.

Carbon recovery may omit:
- organic acids
- lipids
- extracellular metabolites
- unmeasured products
- evaporation/condensation effects
- unrecorded additions/removals
- biomass-composition changes
- analytical uncertainty

Gas calculations depend on the physical meaning and calibration basis of the uploaded gas signals.

The tool is intended for research and process-development analysis and should be validated before any regulated use.

## Figure export

Figures use Plotly and can be exported as:
- SVG vector graphics
- high-resolution PNG

Journal requirements differ. Final figure dimensions, fonts and resolution should be checked against the target journal's author guidelines.

## Publishing on GitHub Pages

Replace the existing repository files with:
- `index.html`
- `styles.css`
- `app.js`
- `README.md`

No GitHub Pages reconfiguration is required if Pages is already enabled.


### CO2 ledger bookkeeping

The gas routines calculate net CO2 evolution. For a mass-balance ledger, the software reports gas-phase CO2 entering as a positive term and reconstructs gross exhaust CO2 as:

`gross CO2 out = net CO2 evolution + CO2 entering`

This avoids double-counting the inlet CO2 term.
