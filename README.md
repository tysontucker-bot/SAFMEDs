# SAFMEDs

A fully client-side SAFMEDs flashcard drill app.

## Features

- Load bundled `.xlsx` / `.xls` decks with `Term` and `Definition` columns
- Practice in shuffled 60-second timings with keyboard shortcuts
- Mark cards correct/incorrect and track per-minute rates
- Automatic results history (most recent 20 timings per deck)
- Progress chart with correct/min and incorrect/min trend lines
- Local-only storage via `localStorage` (no backend)
- Light/dark mode and responsive layout

## Run

Open `index.html` in a browser. No build step required.

## Spreadsheets

- Store workbook files in `/home/runner/work/SAFMEDs/SAFMEDs/spreadsheets`
- The app currently loads `spreadsheets/SAFMEDs 1.xlsx`
- To add more bundled spreadsheets later, place them in `spreadsheets/` and add their relative paths to `DEFAULT_DECK_FILES` in `/home/runner/work/SAFMEDs/SAFMEDs/app.js`
