# pandabear

**Excel schedule → Microsoft Teams Shifts workbook**

A desktop utility created to reduce manual schedule re-entry. It reads a schedule from a selected worksheet in an existing Excel workbook and creates a workbook formatted for import into Microsoft Teams Shifts.

> **Development provenance:** The initial application code was generated with substantial AI assistance. The workflow requirements and intended behavior were directed by the project creator. This repository documents the resulting implementation without claiming it was written entirely by hand.

## What it does

- Opens an `.xlsx` schedule using a desktop interface.
- Lets the operator choose an existing worksheet and output location.
- Reads scheduled work blocks from the worksheet's cell formatting and time columns.
- Maps employee names to work email addresses stored in a local JSON file.
- Generates a Shifts import workbook containing **Shifts**, **Time Off**, **Day Notes**, and **Members** worksheets.
- Warns when a name has no matching email address so the operator can correct the mapping before import.

This is a **template-specific converter**, not a universal Excel-to-Shifts converter. The conversion logic expects a particular layout, time grid, and color-coded schedule.

## Stack

Python · CustomTkinter · openpyxl · pandas · dateparser

## Project structure

```text
pandabear/
├── README.md
├── requirements.txt
└── pandabear/
    ├── TSC.py          # desktop interface, file dialogs, local email mapping
    └── xlsx_engine.py  # schedule parsing and Shifts export
```

## Running from source

1. Install Python 3.11+ and ensure the system supports Tkinter.
2. Create and activate a virtual environment.
3. Install dependencies with `python -m pip install -r requirements.txt`.
4. Run `python pandabear/TSC.py`.
5. Use **Employee Email** to maintain the local name-to-email mappings. Choose **Upload Schedule** to select a workbook and worksheet.
6. **Review the generated workbook before importing it into Teams Shifts**, especially when there are missing email matches.

The source expects a correctly structured input workbook. No general-purpose sample input or automated test suite is currently supplied. Dependencies and the commands above have not been verified in an automated clean-environment test.

## Important implementation details

- The schedule parser reads colored, 15-minute work slots from the input template, with a configured day boundary of 6:00 PM.
- Employee mappings are saved to `employee_emails.json` next to the script or packaged executable. This file can contain personal work email addresses and **should not be committed**.
- Missing email mappings currently produce placeholder values. Correct these before using an export operationally.
- The current export uses a fixed **Call Center** group label; adapt that assumption before reusing the project for another department.
- Optional icons and a packaged executable are not included in this repository.

## Current limitations and planned improvements

- [ ] Provide **synthetic** input and output example workbooks with no employee data.
- [ ] Extract parsing and validation into independently testable modules.
- [ ] Add tests for shifts crossing breaks, missing emails, and different workbook layouts.
- [ ] Replace fixed schedule/grid assumptions with configurable values.
- [ ] Add structured error reporting and a repeatable packaging process.

## Data handling

Do not upload real scheduling spreadsheets, employee directories, or credential files. Test with synthetic datasets. This is an independent portfolio project, not an official Microsoft product or integration.
