# BioProcess Calc V3

Browser-based bioprocess calculators and instrument-data analyzer.

## V3 improvements

- Recognizes instrument-style timestamp headers such as:
  `TimeStamp(UTC+02:00 interval=5)`
- Converts date/time values to elapsed process hours
- Supports CSV and Excel (`.xlsx`) uploads
- Detects common bioreactor signals:
  - DO
  - pH
  - temperature
  - stirrer/agitation
  - air, O2, N2 and CO2 flow
  - feed / SUBS_A
  - BlueVary vol.% channels
  - OUR
  - CER
  - RQ
- Preserves duplicate headers by displaying Excel-style column letters
- Lets the user select any detected numeric process signal for plotting
- Handles large files by downsampling the graph while retaining full data for summary statistics
- Displays only the first 100 rows in the browser preview
- Exports process summary statistics as CSV

## Important

Raw bioreactor files normally do not contain offline biomass, residual substrate, or HPLC product measurements.
Yx/s and Yp/s therefore require separate sampling data. A future version can merge sampling and reactor files by timestamp.


## V4 timestamp correction

V4 checks the acquisition interval embedded in instrument headers such as:

`TimeStamp(UTC+02:00 interval=5)`

If Excel has auto-converted ambiguous D/M/Y dates and the resulting calendar span is implausible compared with the row count and logging interval, BioProcess Calc uses:

`elapsed process time = (row index × acquisition interval)`

instead of reporting artificial month-long gaps.

For the tested 118,739-row export with a 5-second interval, the interval-based recorded duration is approximately 164.91 h.
