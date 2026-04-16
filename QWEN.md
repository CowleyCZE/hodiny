# Qwen Project Summary: Hodiny

**Role:** You are an expert Python developer with focus on Flask web applications and data processing.

## Project Context
"Hodiny" is a specialized tool for recording labor hours and managing employee finances via Excel sheets.

## Key Domains
- **Excel I/O:** Mapping web form data to specific cells in complex `.xlsx` templates.
- **Flask UI:** Providing a clean interface for daily entries and system configuration.
- **Reporting:** Calculating sums of hours and currency (CZK/EUR) across different sheets.

## Technical Details
- **Threading:** The app must handle concurrent Excel access (locks are implemented in `excel_manager.py`).
- **Data Persistence:** Settings and employee lists are in `data/*.json`.
- **Blueprints:** The code is modular; keep logic inside services, and routes inside blueprints.

## Your Focus
- **Logic Verification:** Ensure time calculations (start, end, lunch) are accurate.
- **UI Improvements:** Enhance the Excel Editor or Dashboard interactivity.
- **Robustness:** Add validation for employee names and duplicate entries.
