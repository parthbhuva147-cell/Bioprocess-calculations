# BioProcess Calc V6

Integrated, client-side bioprocess analytics for GitHub Pages.

## What changed from V5

### Obvious multi-file workflow
V6 no longer relies on one generic `multiple` file picker. It has dedicated upload areas for:
- Bioreactor data
- OD / DCW data
- HPLC / product data
- Substrate / sampling data
- Additional files

Each area can accept multiple files, and **Load & merge selected files** loads them into the same experiment workspace.

### Cleaner interface
The confusing visible "online process signals" section was removed from the main workflow.

V6 now:
1. auto-detects measurements,
2. shows only detected variables as simple badges,
3. hides manual column mapping under an **Advanced** section.

### Process time
V6 uses one explicit process-time window for all calculations and plots.

Important behavior:
- Headers such as `TimeStamp(UTC+02:00 interval=5)` are recognized.
- If spreadsheet dates conflict with the acquisition interval, interval-derived time is used.
- Numeric process-time columns such as `time_h` are treated as actual process hours and are **not reset to zero**.
- Per-file offsets and forced logging intervals remain available.
- Optional process-phase boundaries can be shown on plots.

For the tested F29 export:
- 118,739 rows
- 5 second interval
- expected recorded duration ≈ 164.9139 h

### Restored quick calculators
V6 restores:
- specific growth rate
- doubling time
- biomass yield Yx/s
- product yield Yp/s
- substrate consumption efficiency
- volumetric product productivity
- biomass productivity
- glucose carbon
- CO2 carbon

### Integrated real-world KPIs
Where required data are available:
- process duration
- estimated μmax
- doubling time
- maximum OD600
- maximum DCW
- substrate utilization
- mass-based Yx/s
- mass-based Yp/s
- final product titre
- product productivity
- biomass productivity
- final residual glucose
- total feed
- minimum DO
- mean pH
- mean temperature
- maximum agitation

Mass-based substrate and yield calculations correct for mapped feed and sample withdrawals.

### Carbon balance
Carbon ledger:
- `+` initial glucose
- `+` feed/manual glucose
- `+` gas-phase CO2 entering
- `+` user-defined carbon additions
- `−` carbon removed by samples
- `−` CO2 carbon leaving in exhaust
- `−` user-defined carbon removals
- final measured residual glucose, biomass and products remain in the reactor inventory

Sample carbon is calculated at each sampling time from mapped concentrations and sample volume.

### CO2
Automatic CO2 mode:
1. use non-zero CER if available,
2. otherwise use off-gas CO2 plus gas flow,
3. report "not available" when the file has no usable gas signal.

The tested F29 file has zero values in both BlueVary vol.% channels and the BlueVary OUR/CER/RQ channels, so it cannot by itself provide a meaningful biological CO2 evolution value.

### Figures
Scientific figures are produced separately by variable to avoid mixing incompatible units.

Exports:
- vector SVG
- high-resolution PNG
- adjustable width / height / font / line width
- optional phase boundaries

Always confirm the final size/resolution requirements of the target journal.

## GitHub Pages

Replace:
- `index.html`
- `styles.css`
- `app.js`
- `README.md`

in the existing repository. GitHub Pages configuration does not need to be repeated.
