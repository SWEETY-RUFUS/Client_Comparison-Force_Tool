# Client Comparison Tool — Reverse Order + CL_ID + Built-in Force Mapping

## Files
- `index.html` — webpage
- `style.css` — layout and styling
- `app.js` — comparison and formula generation

## How to use
1. Upload the main client workbook.
2. Click **Start Comparison**. The `Force` sheet is reserved for mapping and is excluded from client comparisons.
3. Download the result workbook. It includes `Matched_Sheet`, `Match_Type`, `Matched_CL_ID`, and `PTY_LOCAL_CODE` in each valid client sheet.
4. Open the result workbook in Excel and go to the `Force` sheet.
5. Paste the current client's two columns into `Force!A:B`, including headers `Client_Unique_ID` and `PTY_LOCAL_CODE`.
6. Type the exact target client worksheet name into `Force!D2` (for example, `286859_Pernod`). Only that client sheet's PTY_LOCAL_CODE formulas will return codes.
7. For the next client, clear the previous mapping values in columns A:B, change `Force!D2`, and paste the next client's mapping.

## Formula behavior
Each client sheet contains a formula in `PTY_LOCAL_CODE`. It checks that `Force!D2` exactly equals that sheet's name, then uses an exact match between `Matched_CL_ID` and `Force!A:A` (`Client_Unique_ID`) to return the value from `Force!B:B` (`PTY_LOCAL_CODE`). Other client sheets remain blank. Excel recalculates formulas when the workbook is opened.

## Required headers
Each client sheet needs `CL_STORE_NAME` and `CL_PA_LINE1`. If available, `CL_ID` is read from the matched row and placed in `Matched_CL_ID`.

## Notes
- Comparison and formula generation run locally in the browser. The app does not upload workbooks to GitHub.
- Do not commit client workbooks to the public repository.
- Keep `index.html`, `style.css`, `app.js`, and `README.md` together in the repository root.
- Formula results are calculated by Excel, not by the browser-based XLSX writer.
- The Force mapping is temporary by design: clear A:B and update D2 before mapping another client.
