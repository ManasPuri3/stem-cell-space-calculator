# Stem Cell Manufacturing in Microgravity: Feasibility Calculator (MVP)

## Run it
Double-click `start_calculator.bat`. It opens http://localhost:8000. Close the black window to stop the server.
Nothing to install: the app is plain HTML/JavaScript served by the Python already on this computer.

## Pitch page (3-minute demo)
http://localhost:8000/pitch.html: one screen, four clicks (or arrow keys): Today, then microgravity growth, then robot factory, then cheaper rockets and return. Bars show cost per patient dose (150 million heart cells), Earth vs space. The last step shows the break-even conditions and three sliders (extra growth, rocket price, return cost).

## What's in it
- **Overview**: the competition question, a plain-English answer, "what would it take for space to win?" (max launch price, max return cost, max orbital manufacturing cost, minimum quality and growth ratios, and the 100%-purity ceiling), a viability map (launch price vs return cost), and every key input rated against published evidence (In range / Guess / Optimistic / Conservative / No evidence).
- **Scenarios**: space viability case (heart muscle cells on an autonomous free-flyer with projected launch/return prices; every bet listed), heart cells with today's ISS logistics, blood stem cells (HSC), PRD hypothetical example, empty.

## Files
| File | What it is |
|---|---|
| `Stem_Cell_Microeconomics.xlsx` | Earth vs space unit economics with live formulas, evidence status, subjective assumption probabilities, sources |
| `build_spreadsheet.py` | Single source of truth: builds the xlsx **and** `app/scenario_data.js` (the calculator's literature-informed scenario) |
| `app/engine.js` | Pure numerical engine (PRD sections 6, 7, 10) |
| `app/index.html`, `app/app.js`, `app/style.css` | UI (Overview, Assumptions, Evidence, Results, Methods) |
| `app/worker.js` | Runs the Monte Carlo simulation off the UI thread |
| `app/tests.html` | Automated engine tests: open http://localhost:8000/tests.html |
| `tests/test_spreadsheet.py` | Python tests; evaluates the workbook's Excel formulas with pycel |

## Change assumptions in the spreadsheet scenario
Edit the `add(...)` rows in `build_spreadsheet.py`, then run:
```
py build_spreadsheet.py
```
Editing the xlsx alone does not update the app.

## Tests
- Browser: http://localhost:8000/tests.html (22 tests: all PRD section 11 numbers, acceptance tests 1-12, viability thresholds, reasonableness labels, and JS-vs-Python cross-checks for both scenarios).
- Python: `py -m unittest discover -s tests -v` (8 tests; evaluates the workbook's real Excel formulas for both scenarios; needs `py -m pip install pycel openpyxl`).

## Deviations from the PRD (MVP, time-boxed)
- Plain JavaScript instead of React/TypeScript, because Node.js is not installed. The engine has no dependencies, so it can be ported later.
- Deferred: discrete distributions and joint profiles, duration and supply-mass rules, evidence editing, scenario persistence, JSON/CSV export, print report, heatmap, cost-ratio histogram, numerical break-even search, simple/advanced mode toggle, dose view.
- Observed performance: 50,000 draws take about 0.6 s in a Web Worker on this laptop, and Cancel terminates the worker.
