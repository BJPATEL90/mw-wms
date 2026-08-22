# MW WMS Modular Production Version

This folder is the separated-code version of the final production files.

The original final files remain untouched in `production-ready/`.

## Apps Script Backend

Upload every `.gs` file from `apps-script/` into the same Apps Script project.

Recommended Apps Script file list:

- `00-router.gs` - `doGet`, `doPost`, action router, shared spreadsheet constant
- `01-setup-auth.gs` - setup menu, sheet headers, login, users
- `02-legacy-warehouse.gs` - older warehouse flow and stock reports
- `03-master-data.gs` - SKU/Bin master save helpers and brand inference
- `04-wms-save-load.gs` - WMS data save/load functions
- `05-operations-current-inventory.gs` - adjustments, cycle count helpers, current inventory
- `06-refresh-cache.gs` - sheet clearing, cache helpers, full data refresh
- `07-bin-allocation-dashboard.gs` - brand/bin allocation and dashboard helpers
- `08-uniware-sync.gs` - Uniware inventory email sync
- `09-gatepass-sync.gs` - Gatepass email sync

After uploading the `.gs` files, run `createAllSheets()` once, then deploy a new Apps Script web app version.

Note: the live app frontend should be served from GitHub Pages using the `frontend/` folder below. If you still want to open the Apps Script URL directly as a full app page, keep an Apps Script `index.html` file in that project too. GitHub login/API calls only need the `.gs` backend.

## GitHub Frontend

Upload the contents of `frontend/` to GitHub Pages.

Required structure:

```text
index.html
assets/
  css/
    styles.css
  js/
    00-xlsx-loader.js
    01-core-state.js
    02-auth-dashboard.js
    03-inbound-putaway.js
    04-inventory-reports.js
    05-operations.js
    06-runtime-sync.js
    07-planning.js
```

Keep the JS file names as-is because `index.html` loads them in this order.

## Permission Setup

In the `USERS` sheet, use the `Allowed Sections` column.

Examples:

```text
ALL
INBOUND
OPERATIONS,PLANNING
INBOUND,OPERATIONS,PLANNING
```

Dashboard and Reports are available for every logged-in user. Inventory remains admin-only.

## Validation

This modular split was generated from the final production files and syntax checked:

- Combined Apps Script syntax OK
- Combined frontend JavaScript syntax OK
- Each frontend JavaScript file syntax OK
