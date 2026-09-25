# SAFMEDs

A fully client-side SAFMEDs flashcard drill app.

## Features

- Load bundled `.xlsx` / `.xls` decks with `Term` and `Definition` columns
- Practice in shuffled 60-second timings with keyboard shortcuts
- Preview cards before starting a timed run
- Mark cards correct/incorrect and track per-minute rates
- Automatic results history (most recent 20 timings per deck)
- Progress chart with correct/min and incorrect/min trend lines
- Local-only storage via `localStorage` (no backend)
- Highlight text on cards with `==highlighted text==`
- Edit card highlights in preview and save those edits locally on the current device
- Light/dark mode and responsive layout

## Run

Open `index.html` in a browser. No build step required.

## Spreadsheets

- Store workbook files in `spreadsheets/`
- The app currently loads `spreadsheets/SAFMEDs 1.xlsx`
- To add more bundled spreadsheets later, place them in `spreadsheets/` and add their relative paths to `DEFAULT_DECK_FILES` in `app.js`
- In Term/Definition cells, wrap words in `==double equals==` to highlight them during preview and practice
- You can also open Preview, choose Edit Card, and apply/remove `==double equals==` highlighting there; those edits are saved in browser storage on the current device and do not update the original spreadsheet file
