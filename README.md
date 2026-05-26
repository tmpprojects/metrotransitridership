# Metro Transit Ridership — D3 Multi-Line Chart

A learning project for the [tcplot Claude Code workshop](https://tcplot.com/claude-code/).
We build a D3 multi-line chart of Metro Transit (Twin Cities) ridership by mode,
using the FTA's National Transit Database Monthly Module data.

## Run it

Three ways, all work:

**A. GitHub Pages (no clone).** The chart is published as a static site — open it in any browser:

<https://tmpprojects.github.io/metrotransitridership/>

**B. Just open it.** Double-click `index.html` (or open with your browser). No server, no setup. The data is pre-extracted into `data/metro-transit.js` and D3 is bundled locally in `vendor/`, so `file://` works fine.

**C. Local HTTP server.** Slightly nicer for live-editing because you can refresh and devtools cooperate better:

```bash
python3 -m http.server 8000
```

Then open <http://localhost:8000/>.

## What's here

- `index.html` — the chart. D3 v7 (UMD) is loaded from `vendor/d3.v7.min.js`; ridership/service data is loaded from `data/metro-transit.js`.
- `vendor/d3.v7.min.js` — vendored D3 build so the page works offline / on `file://`.
- `data/` — raw and processed data:
  - `*.csv` — full CSV exports of the FTA Monthly Module workbook, mirrored from [tcplot.com/claude-code/](https://tcplot.com/claude-code/).
  - `metro-transit.js` — generated. Just the Metro Transit rows we plot (~76 KB instead of 8 MB of CSVs). Sets `window.METRO_DATA = {...}`.
- `build-data.py` — extracts the chart data from CSVs into `data/metro-transit.js`. Re-run after a CSV refresh:
  ```bash
  python3 build-data.py
  ```
  Or refresh straight from FTA (downloads the latest Monthly Module workbook, re-exports every tab to `data/*.csv`, then rebuilds). This path needs `openpyxl` — see [Setup](#setup) below:
  ```bash
  python3 build-data.py --fetch     # scrape FTA for the newest .xlsx
  python3 build-data.py --fetch --url https://www.transit.dot.gov/.../...xlsx  # or pin a release
  ```
- `requirements.txt` — Python deps for `--fetch` only (just `openpyxl`).
- `README.md` — you are here.

## Setup

The default rebuild (`python3 build-data.py`) is pure stdlib — no setup. You only need a Python environment for the `--fetch` path, which uses `openpyxl` to read FTA's `.xlsx` workbook.

**Option A — venv (portable, recommended):**

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 build-data.py --fetch
```

**Option B — Debian/Ubuntu system package (no venv):**

```bash
sudo apt install python3-openpyxl
python3 build-data.py --fetch
```

`.venv/` is gitignore-friendly and the recommended choice if you'll be cloning this on other machines.

## Data shape

`UPT-Table.csv` (and the sibling `VRM-Table.csv` we also use) are wide-format:
each row is `(Agency × Mode × Type of Service)`, and columns `1/2002 … 3/2026`
hold monthly values (comma-formatted strings, blank when not reported).

`build-data.py` filters those CSVs down to Metro Transit's rows (NTD ID 50027)
for the four modes below and writes a compact `data/metro-transit.js` of the form:

```js
window.METRO_DATA = {
  ntdId: "50027",
  upt: {
    MB: [{ date: "2002-01-01", value: 5500000 }, …],
    LR: [...], RB: [...], CR: [...],
  },
  vrmBus: [{ date: "2019-01-01", value: 2040000 }, …],
};
```

Metro Transit is **NTD ID 50027**, with four modes appearing in the file:

| Mode | Label                        |
| ---- | ---------------------------- |
| MB   | Bus                          |
| LR   | Light Rail                   |
| RB   | Bus Rapid Transit            |
| CR   | Commuter Rail (Northstar)    |

## Workshop progression

Following the [workshop prompts](https://tcplot.com/claude-code/):

1. **Sketch** — plan the chart (multi-line, one line per mode, monthly UPT).
2. **Get something on screen** — this file. ✓
3. Fix the first thing that's wrong.
4. Add pandemic context.
5. Make it interactive.
6. A judgment call Claude can't make for you.
7. Ship it.

## Source

Federal Transit Administration. *Monthly Module Adjusted Data Release.*
National Transit Database, U.S. Department of Transportation.
<https://www.transit.dot.gov/ntd/data-product/monthly-module-adjusted-data-release>

## License

This project is licensed under the [MIT License](LICENSE).
