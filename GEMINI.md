# Gemini Project Summary: Hodiny

**Role:** You are a Python/Flask expert and Excel automation specialist.

## Project Overview
"Hodiny" is a comprehensive attendance tracking system that manages employee working hours, payments (advances), and reports directly within Excel files (`.xlsx`). It features a web-based UI for direct Excel editing and an AI-ready "voice command" processor.

## Core Features
- **Excel Automation:** Direct reading/writing to `Hodiny_Cap.xlsx` based on a strict template.
- **Employee Management:** CRUD operations stored in `data/employee_config.json`.
- **Excel Editor:** In-browser interactive cell editing (thread-safe).
- **Voice/Text Commands:** NLP processing for recording time (e.g., "work today 7 to 16").
- **Reporting:** Monthly summaries, Excel preview, and email exports (SMTP).

## Tech Stack
- **Backend:** Python, Flask (Blueprints architecture).
- **Excel Engine:** `openpyxl` (via `excel_manager.py`).
- **Frontend:** Jinja2 templates, Vanilla JS (AJAX for the editor).

## Strategic Mandates
1. **Excel Integrity:** Always use the file lock in `ExcelManager` when writing. Never break the template structure.
2. **Employee Selection:** Respect the "selected" state in `employee_config.json` for batch entries.
3. **Voice Logic:** Extend `utils/voice_processor.py` for more complex NLP if needed.

## Key Files
- `app.py`: Application bootstrap.
- `blueprints/`: Domain-specific route handlers.
- `excel_manager.py`: The heart of Excel I/O.
- `employee_management.py`: Logic for employee persistence.
- `utils/voice_processor.py`: Command parsing.
