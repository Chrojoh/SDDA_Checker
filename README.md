# SDDA Title Tracker Mobile

A mobile-focused standalone browser application for reviewing the two CSV or XLSX downloads available in a member's SDDA account:

- Dog Titles Table
- Dog Scores Table

Open `index.html` directly or publish this folder with GitHub Pages. All processing happens in the browser; there is no server and no account login.

This edition uses a touch-friendly single-column workflow on phones, stacked score cards instead of a horizontally scrolling table, and a sticky print/save action. Wider screens and printed reports retain a conventional multi-column layout.

For XLSX account downloads, the tracker checks every worksheet and uses the one containing the expected Dog Titles or Dog Scores column headers. Introductory rows above the table are allowed.

The tracker follows the high-to-low Gold-score procedure in the supplied SDDA guide. It identifies titles recorded in the title download, reconstructs Advanced and Excellent titling-score sets, counts calculated Gold scores, and explains remaining championship requirements.

The **Print / save PDF** button opens the browser's print preview using the formatted report layout. Choose a PDF destination when the browser provides one, or print the report on paper. The page title—and therefore the suggested PDF filename in Chrome or Edge—is generated as `Dog Name - SDDA Title Tracker - YYYY-MM-DD`.

## Filling blank score dates

The tracker links to SDDA's current Trial Workbook and accepts the downloaded `.xlsx` file. It reads the `Processed and Future Trials` sheet, matches `Trial Number`, and fills blank score dates when a matching trial is present. SDDA only describes this as a list of recently processed and future trials, so older trial numbers may not be available in the current workbook.

## Important limitations

The supplied score download does not contain a judge column. Judge-diversity requirements cannot be confirmed. Blank Event Date values also prevent reliable cutoff calculations for MACH2/MACH3/MACH4 and corresponding Elite extensions. The tracker reports those limitations instead of guessing.
