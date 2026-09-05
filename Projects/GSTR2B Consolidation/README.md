# GSTR-2B Consolidation & Pivot Automation 📊⚙️

## The Problem
Manually merging monthly GSTR-2B Excel files for GST reconciliation is highly tedious and error-prone. The process requires extracting data from multiple complex sheets (B2B, B2BA, B2B-CDNR, B2B-CDNRA), formatting it consistently, and manually adjusting numerical signs for credit/debit notes before analysis.

## The Solution
This Excel VBA macro fully automates the extraction, transformation, and loading (ETL) of GSTR-2B data from a user-selected folder of monthly files into a single master workbook. 

### Key Technical Features:
* **Automated Data Consolidation:** Loops through multiple files via `Scripting.FileSystemObject`, mapping complex column structures across standard, amendment, and credit note sheets.
* **Intelligent Data Transformation:** Automatically detects and converts values from Credit Notes and reduced amendments into negative numbers to ensure accurate summations.
* **Dynamic Highlighting & Formatting:** Highlights amended rows (B2BA/B2B-CDNRA) in yellow, tracks the exact month of data origin, and applies continuous borders and bolding.
* **Instant Pivot Reporting:** Programmatically generates a Pivot Table using a `PivotCache`, summarizing Taxable Values and all Integrated/Central/State taxes by month with automated Number Formatting.
* **Dynamic File Saving:** Calculates the financial year range based on the processed files and triggers a dynamic `GetSaveAsFilename` prompt.

## User Interface

The macro is tied to a developer command button on the main dashboard for single-click execution.

**Macro Execution Button:**
![Macro Button](image_009248.png)

**ActiveX Button Properties:**
![Button Properties](image_009283.png)
