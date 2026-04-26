# Invoice & Inventory Management System

This repository contains a massive upgrade from a basic Automated Invoice Generator to a full **Hardware-Locked Invoice & Inventory Management System**. It features stock tracking, live database analytics, an automated in-sheet form, and robust security.

## New Architecture Overview

The system runs completely inside Excel without external database dependencies, utilizing 6 core sheets:

1. **Dashboard:** Live analytics featuring 6 KPI cards (Total Revenue, Stock Value, etc.) and a native Excel `FILTER` array pulling items requiring restocking.
2. **Invoice:** A premium, printable invoice generator. Features include live VLOOKUPs to pull prices from the master inventory and real-time "Stock Remaining" indicators to prevent overselling.
3. **Inventory:** The master stock database. Uses Conditional Formatting to highlight stock statuses (Red = Out of Stock, Orange = Low Stock, Green = In Stock).
4. **StockIn:** An automated ledger of stock additions. Features a clean, programmatic **In-Sheet Form** to add new products or top up existing ones easily.
5. **Records:** The database of all processed invoices.
6. **Settings:** Master configuration for company details, tax rates, prefixes, and dropdown lists for product categories.
7. **LicenseData (Hidden):** Stores the encrypted hardware fingerprint and expiry date.

## Developer Deployment Checklist

This system includes a robust hardware locking mechanism. To deploy to a client, you must follow these steps:

1. Open `Module2` (License_Module) and change the salt string in `HashFingerprint()` to your own secret value.
2. Open the workbook on the **CLIENT's machine**.
3. Run the `GenerateLicenseForClient()` macro and copy the output key.
4. Paste the license key into `LicenseData!B1`.
5. Fill the client's company name in `LicenseData!B2`.
6. Set the expiry date in `LicenseData!B3` (format: `DD-MMM-YYYY`).
7. Fill `Settings!B2:B5` with the client's company details.
8. Lock the VBA project: VBA Editor > Tools > VBAProject Properties > Protection.
9. Set a VBA project password (keep this password — do NOT give it to the client).
10. Save and close the workbook.
11. Reopen to confirm license validation passes.
12. Deliver the workbook to the client.

## Setup Instructions

If you are setting this up from scratch using the source code files:

1. Open a blank Excel workbook and save it as an `.xlsm` file.
2. Open the VBA Editor (`Alt + F11`).
3. Import the 3 `.bas` files and overwrite the code in `ThisWorkbook` with the `.cls` file.
4. Run the `SetupInvoiceSystem` macro. **WARNING: This will wipe any existing sheets.**
5. The system will build all 6 sheets, apply named ranges, create the tables, and insert the formulas.
6. Navigate to the `Dashboard` and insert the 4 required charts manually using the pivot table features.
