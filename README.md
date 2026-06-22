# Mosaic Wellness WMS

Google Apps Script based Warehouse Management System for Mosaic Wellness operations.

## Main Files

- `code-production.gs` - production backend Apps Script code. Copy into Apps Script as `Code.gs`.
- `index.html` - desktop WMS dashboard.
- `mobile.html` - mobile floor app.

## Deployment

This app is designed to run from Google Apps Script because it uses `google.script.run` and Google Sheet services.

GitHub is used for source control. The live app should still be deployed from Apps Script unless the frontend/backend are later separated into a standalone web app.

## Notes

- Spreadsheet data files are ignored by Git to avoid publishing operational data.
- Backup files are ignored by Git.
- After updating Apps Script, deploy a new web app version and clear browser local data once if old cache causes issues.
