# Wine Supplier DB

Clean supplier database builder for wine price lists.

The project goal is to build a verified supplier database from Excel price lists
for restaurant wine-list analysis.

## Scope

- Source price lists: Excel only (`.xlsx`, `.xls`).
- CSV and other source formats are intentionally ignored for now.
- Generated artifacts such as SQLite databases, JSON exports, reports, and raw
  price lists are excluded from git by default.

## Planned Structure

```text
supplier_db/
  TECH_SPEC.md
  suppliers.yaml
  build_supplier_db.py
  suppliers.db
  suppliers.json
  rejected_rows.csv
  build_report.csv
  aliases.json
```

See [supplier_db/TECH_SPEC.md](supplier_db/TECH_SPEC.md) for the full technical
specification.
