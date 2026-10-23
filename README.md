# PDF to Excel — Box App Import Edition

## Run locally
1. Install Node.js 20+.
2. Open a terminal in this folder.
3. Run `npm install`.
4. Run `npm run dev` and open the local URL shown by Vite.

## Features
- Manually select multiple PDF quotations.
- Extract item quantities, item codes, dimensions, ply, and explicit DF/SF type.
- Multi-page PDFs and OCR for scanned pages.
- Review and edit extracted rows before export.
- Export the exact 27-column Box App Bulk Calculation Template.
- Export one combined workbook or separate workbooks per PDF.
- Fill blank template fields with the defaults supplied by the user.
- Process files locally in the browser.

## Defaults
Customer Name: Noble Art And Craft House
Dimension Unit: inch; Flute Type: B/C
Joint: 2; Deckle Margin: 2; Overlap Length: 1
Wastage: 6; Setup Cost: 0; Tooling Cost: 0
Printing Required: yes
Printing/Punching/Lamination/Pasting/Transport/Other per box: 0
Factory Overhead: 5
Paper Combo Name: 18/140+18/100+18/140+18/100+18/140+18/100+18/140

## Important
Defaults apply only to blank fields. Extracted values are preserved. Parser is heuristic and all rows should be verified against source PDFs before Box App import. This starter has not been validated against every RFQ layout.
