# multi-shop-stock-check-toolkit
Excel + VBA toolkit for multi-branch retail stock checks and supplier ordering.

Multi-Shop Stock Check Toolkit (Excel + VBA)

This project is a production-used Excel + VBA toolkit built for a small retail chain to standardise stock checks and streamline supplier ordering across multiple shops.

Features

Master product catalogue
Single source of truth for all products (ID, category, brand, order unit size, global status).

Per-shop stock sheets
Identical table structures for each branch (ShopName_Stock), including:

Current stock (units)

Par levels

Local & global status

Last counted date

SortKey for custom ordering

Data validation & dependent dropdowns
Uses named ranges and INDIRECT+SUBSTITUTE+TRIM to provide robust dependent lists (e.g. category → brand → product) while tolerating spaces, hyphens and special characters.

Checklist export
For each *_Stock sheet, a button runs ExportOneShop, generating a clean Stock Check workbook:

Grouped by product category

Product ID + Display Name

Blank “Counted Units” column for staff

Only includes Active + Global Active items

Checklist import
Import_Returned_Checklists reads completed checklists from Imports/Checklists, identifies the target shop, matches rows by Product ID, and updates:

Current Stock (Units)

Last Counted

Supplier order list generator
Generate_Supplier_Order_From_Checklists aggregates requested quantities across all shops, uses the “Order Unit (count)” to convert units to boxes, and outputs a WhatsApp-ready text file:
