# Inventory Excel Report – Odoo 18 Custom Module

## Overview
Generates a dynamic, professionally formatted **Excel (.xlsx)** inventory report
from a configurable wizard directly inside Odoo 18.

---

## Features
| Feature | Detail |
|---|---|
| **Wizard Filters** | Date From, Date To, Location (mandatory), Company (read-only), Product Category (M2M), Product (M2M) |
| **Opening Stock** | Historical qty at the selected location **before** Date From |
| **Stock In** | One column per unique **source** location that sent goods to the report location within the date range |
| **Stock Out** | One column per unique **destination** location that received goods from the report location within the date range |
| **Closing Stock** | Opening + Total In − Total Out |
| **Prices** | Cost Price (`standard_price`) and Sales Price (`lst_price`) from the product |
| **Grand Total Row** | Summed at the bottom |
| **Freeze Panes** | Frozen at data start row and Stock-In column start |
| **Print Ready** | Landscape, A4, fit-to-width |

---

## Module Structure
```
inventory_excel_report/
├── __init__.py
├── __manifest__.py
├── security/
│   └── ir.model.access.csv
├── views/
│   └── inventory_report_wizard_view.xml
└── wizard/
    ├── __init__.py
    └── inventory_report_wizard.py
```

---

## Installation
1. Copy the `inventory_excel_report` folder into your Odoo `addons` directory.
2. Restart the Odoo service.
3. Activate **Developer Mode** (Settings → Developer Tools).
4. Go to **Apps**, click **Update Apps List**.
5. Search for **Inventory Excel Report** and click **Install**.

---

## Usage
**Inventory → Reporting → Inventory Excel Report**

1. Set **Date From** and **Date To**.
2. Select the **Location** (mandatory – must be an Internal location).
3. Optionally filter by **Product Category** and/or **Product Name**.
4. Click **Generate Excel Report** → the `.xlsx` file downloads immediately.

---

## Logic Details

### Opening Stock
Calculated as the net quantity at the chosen location **before** `Date From`:
```
Opening = SUM(qty into location before date_from)
        − SUM(qty out of location before date_from)
```
Uses a bulk `read_group` query for performance across many products.

### Stock In Columns (Dynamic)
- Queries `stock.move.line` where `location_dest_id = wizard_location` and `date` within range.
- Groups by unique source location (`location_id`).
- Each unique source location becomes its own column header.

### Stock Out Columns (Dynamic)
- Queries `stock.move.line` where `location_id = wizard_location` and `date` within range.
- Groups by unique destination location (`location_dest_id`).
- Each unique destination location becomes its own column header.

### Closing Stock
```
Closing = Opening + Total Stock In − Total Stock Out
```

---

## Odoo 18 Compatibility Notes
- Uses `quantity` field on `stock.move.line` (renamed from `qty_done` in Odoo 17).
- Uses `state = 'done'` filter on `stock.move.line` (related to `move_id.state`).
- Compatible with Odoo 18's `xlsxwriter` library (bundled with Odoo).
- Menu registered under `stock.menu_stock_reporting`.

---

## Permissions
| Group | Access |
|---|---|
| Stock User | Read, Write, Create, Unlink on wizard |
| Stock Manager | Read, Write, Create, Unlink on wizard |
