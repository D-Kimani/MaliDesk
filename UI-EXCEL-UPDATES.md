# MaliDesk UI and Excel export updates

This package is based on the uploaded `MaliDesk(3).rar` project. The Vite application entry point is `frontend/src/main.jsx`, which loads `frontend/src/MaliDesk.jsx`.

## UI changes
- Removed the Download CSV action from Import & Export.
- Kept Download Excel and styled it green (`#1F8A4C`).
- Removed the top-bar backup/download icon and notification bell; the date remains.
- Routed the older `App.jsx` and root `src/MaliDesk.jsx` compatibility entry points to the canonical frontend implementation so stale UI versions are not accidentally used.

## Excel export
- The supplied `rental_income_corrected.xlsx` is installed as `frontend/public/rental-income-template.xlsx` and is loaded as the export base.
- Exported unit ledgers now write to the template's actual columns and keep template formula cells (including the running-balance formula in column I) instead of replacing them with static balances.
- Existing formula cells are not cleared when the sample ledger values are removed.

## Verification
- Static source checks pass for the removed controls, green Excel button, active Vite entry point, and template reference.
- JSX syntax parsing passes for the canonical frontend and compatibility entry points.
- The bundled template matches the supplied spreadsheet byte-for-byte; workbook inspection confirms 24 sheets, formula cells, data validations, and the `UnitRegister` table are present in the template.
- A full Vite production build could not be run in this environment because the required npm packages were not installed and the package registry was unreachable. Please run `npm ci` then `npm run build` from `frontend` on a machine with registry access before deployment.

Note: the supplied workbook contains formulas, data validations, and tables, but no embedded Excel drawing/button objects were found. The project retains the template workbook as the export base; final runtime validation of Excel's interactive elements should be done in desktop Excel after exporting a populated workbook.
