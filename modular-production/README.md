# MW WMS Modular Production Version

Generated from the current live branch.

## GitHub Frontend

GitHub Pages serves the root files:

- index.html
- assets/css/styles.css
- assets/js/*.js

The same separated frontend is also saved under frontend/.

## Apps Script Backend

The backend source is split under apps-script/. Add all .gs files into the same Apps Script project if you want to manage Apps Script in separated files.

For this FSN compliance card, no backend logic change is required.

## FSN Compliance Card

The FSN Replenishment section now shows Compliant SKU Lines and the compliance percentage. The calculation is line-based, matching the existing Misplaced SKU Lines KPI.
