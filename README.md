# Dealership service-campaign reports

This project creates English-language reports from the latest campaign and
dealership-zone Excel workbooks in the project folder. Campaign data is read
from the workbook sheet containing `Campaign Code`, `Campaign Name`, `Status`,
and `Responsible Store`. Zone assignments are read from the sheet containing
`Zone` and `Dealer Site`. Dashboard sheets are skipped, but another sheet in
the same workbook (such as `DATABASE`) can provide campaign data.

## Weekly Excel update

1. Put the updated campaign export or dashboard workbook and the updated
   zone-management workbook in this project folder, next to `web_app.py`.
2. Replace the previous source files, or leave old copies in place. When more
   than one workbook matches a data role, the app uses the most recently
   modified valid workbook for that role. Do not leave a temporary Excel lock
   file (`~$...xlsx`) as the only copy.
3. Start (or restart) the report server:

   ```powershell
   .\.venv\Scripts\python.exe web_app.py
   ```

4. Open `index.html` with VS Code Live Server. The app loads the latest valid
   workbooks on first request and reuses them for subsequent filter requests.
   It checks workbook file timestamps and sizes, then reloads automatically
   after an Excel file is replaced or updated.

The campaign workbook needs these columns: `Campaign Code`, `Campaign Name`,
`Status`, and `Responsible Store`. The zone workbook needs `Zone` and
`Dealer Site`. The currently configured personal reporting zone is Zone 4.
Zone comparisons are labelled Zone 1 through Zone 4, with no regional name
added to the comparison.

## Run the web form

1. Open this project folder in VS Code.
2. Install dependencies if needed:

   ```powershell
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```

3. Start the report generator and leave the terminal running:

   ```powershell
   .\.venv\Scripts\python.exe web_app.py
   ```

4. Open `http://127.0.0.1:8765` in your browser. Choose a zone, then generate a zone
   overview, dealership-group report, or individual-dealership report. Choose
   the France overview to combine Zones 1–4.
   **View report** displays an HTML preview; **Download PDF** downloads a
   meeting-ready PDF. Downloaded PDFs are also saved in `rapports`. Open
   `rapport.html` (**Report history**) to see saved reports, their generation
   dates, and links to open or download them.
   Use **My zone** in the navigation to view every dealership assigned to the
   configured personal zone, including sites without campaign records.
   Its France map highlights departments containing assigned dealerships and
   plots each known dealership at its commune center. The highlighted
   departments are a visual guide, not an official zone boundary.

The web page styles are in `style.css`. The local report service listens on
`127.0.0.1:8765`; keep `web_app.py` running while using the page. VS Code Live
Server can also serve the static files; start `web_app.py` first so its report
API and history links are available.
Web reports are limited to dealership records mapped to French Zones 1–4.
Choose a zone to generate a zone overview, group report, or individual
dealership report; the France overview combines all four zones.

## Report contents

- Campaign progress is calculated from the source `Status` field; `Finished`
  is counted as completed.
- France overview lists Zones 1–4 and the combined French total. Campaign rows
  not matched to a French zone assignment are excluded. Zone reports list every
  assigned dealership site, including sites with no campaign rows, separately
  from the ranking of dealerships that do have campaign records.
- Zone, dealership-group, and individual-dealership reports include a bar
  chart. France overview charts Finished records by zone.
- The France overview ranks zones by their total `Finished` cases and shows each
  zone's leading dealership. They also include a top-10 dealership ranking by
  `Finished` count and a top-10 group ranking.
- Dealership leaderboards show `Total campaigns`, `Finished`,
  `To be finished`, and `Completion rate`. The leading three places are
  highlighted in gold, silver, and bronze.
- Zone overviews include a top-10 dealership ranking by the number of
  `Finished` cases.
- Dealership-group reports include the full dealership ranking for that group.
- Individual dealership reports use the labels `Total campaigns`, `Finished`,
  `To be finished`, and `Completion rate`, and include a status bar chart in
  both the HTML preview and the PDF. The dealership selector does not require
  a group filter.
- PDFs are dated automatically. There is no date filter.

The `Finished` ranking is a count, not a quality or repair-speed measure.
Campaign volume differs between dealerships, so the reports show totals and
completion rates for context. Confirm operational causes in the source system
before drawing conclusions from open cases.

## Command-line report generator

The separate command-line workflow is also available:

```powershell
.\.venv\Scripts\python.exe app.py
```

It lists the available Excel workbooks and prompts for one to use.
