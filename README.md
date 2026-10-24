# Activity Management Toolkit

A Python-based toolkit for managing event-program data and automating recurring event-operations documents.

The project brings together structured JSON data, Excel workbooks, Word/PDF document generation, speaker and influencer profiles, and a small Flask preview app. It is designed to reduce repetitive preparation work across event planning and administration.

> **Project status:** Personal workflow-automation project. The repository contains multiple scripts developed for specific tasks; it is not yet packaged as a single installable application. Some scripts depend on local templates, office documents, browser installations, or environment-specific settings.

## What it does

- **Structured event data:** Load and process program, speaker, influencer, and activity data stored in JSON, Excel, and related files.
- **Document generation:** Create or populate event materials such as agendas, speaker letters, publisher files, sign-in sheets, and Word/PDF outputs.
- **Data preparation:** Merge and map data from multiple sources, fill Excel templates, and prepare records for downstream document generation.
- **Speaker and influencer profiles:** Combine names in an event program with available profile details such as organization, title, education, and experience.
- **Browser-based preview:** Use a Flask app to render a selected event program with Jinja templates before using or adapting the output.
- **Reusable configuration:** Resolve data, template, and output directories through `config/paths.json`.

## Repository layout

| Path | Purpose |
|---|---|
| `scripts/actions/` | Task-specific scripts for document generation, Excel processing, profiles, sign-in sheets, and email workflows |
| `scripts/core/` | Shared path setup, data-loading utilities, schema helpers, and merge functions |
| `config/` | Path settings and JSON schemas |
| `data/shared/` | Shared program and influencer JSON data |
| `data/activities/` | Activity-specific source files and workbooks |
| `data/attachments/` | Supporting documents used by workflows |
| `templates/` | Document and activity templates |
| `output/` | Generated files, merged data, and backups |
| `tests/` | Automated tests for selected data and profile functions |
| `static/` | Static assets used by the preview app, if present |



| Python File name | Function |
|---|---|
|add_program_data_to_excel|Add program_data at the every row of an Excel file.|
|app|HTML format of program_data. host=127.0.0.1:5000 (default)and seperate data into subpage: Cover, notes, chairs, speechers|
|siginin_table_from_excel|Read excel file and export them into Word file with specific format (table size, font size, row rules,etc)|


## Requirements

- Python 3.10 or newer is recommended.
- Install the dependencies required by the workflow you want to run. The repository does not currently provide a single authoritative `requirements.txt`.
- The Flask preview app requires Flask and Jinja2.
- Excel and Word workflows may require packages such as `openpyxl` and `python-docx`; PDF rendering or browser-based workflows may have additional system dependencies.

For example, to run the preview app:

```bash
python -m pip install Flask Jinja2
```

Install any additional packages required by the specific script you use.

## Run the event preview app

From the repository root:

```bash
python scripts/actions/app.py --serve
```

Then open [http://127.0.0.1:5000/](http://127.0.0.1:5000/) in your browser.

To use a different port:

```bash
python scripts/actions/app.py --serve --port 5050
```

To render a specific program by ID, use the `event_id` query parameter:

```text
http://127.0.0.1:5000/?event_id=2
```

The app reads `data/shared/program_data.json` and `data/shared/influencer_data.json`, and renders `templates/template.html`. The available routes include:

- `/` — event preview
- `/event/<event_id>` — preview a specific event
- `/cover`, `/notes`, `/chairs`, `/speakers` — section-oriented views
- `/_ctx` — development endpoint that exposes the generated template context as JSON

The `/_ctx` endpoint is intended for local debugging. Do not expose it on a public server if the data contains personal or confidential information.

## Run the tests

Install pytest if needed:

```bash
python -m pip install pytest
python -m pytest
```

The current tests cover selected behavior for influencer-profile construction, matching people to program data, program lookup, and Excel-column collection. They do not yet cover every script in the repository.

## Configuration

Edit `config/paths.json` to configure the data, template, output, activity-data, and Chrome paths used by shared utilities. Relative paths are resolved from the project root; the Chrome path can also be discovered from environment variables or common installation locations.

Before running a script, confirm that:
1. Its expected input files exist in the configured data directories.
2. The relevant template and output directories are present.
3. Required Python packages and external tools are installed.
4. The selected program or activity ID matches the data you intend to process.

Individual scripts may have their own command-line options and assumptions. Check a script's `--help` output or its argument parser before running it.

## Data handling and privacy

This project can process names, professional profiles, contact lists, registration workbooks, and other event-related information. Before publishing or sharing the repository:

- Replace real personal data with synthetic examples or remove it from version control.
- Check `data/`, `output/`, `output/backups/`, documents, spreadsheets, and screenshots for personal or confidential information.
- Keep credentials, SMTP passwords, tokens, and other secrets out of source code and Git history.
- Do not expose the Flask development server or the `/_ctx` endpoint to untrusted users.

## Current limitations and next steps

- Dependencies are not yet consolidated into a reproducible environment file.
- Scripts are task-oriented and may require different inputs, templates, or local software.
- Automated test coverage is limited to selected functions.
- Generated outputs and historical backups are mixed into the repository and could be separated from source code.

Potential improvements include adding a maintained `requirements.txt`, documenting each script's inputs and outputs, using synthetic sample data, expanding tests, and separating source files from generated artifacts.

---

## 中文簡介

這是一個以 Python 建立的活動行政流程自動化工具集，主要用來處理活動議程、講者／專業人士資料、Excel 資料填入與合併，以及 Word、PDF、簽到表和相關活動文件的產生。

專案包含共用資料處理模組、JSON schema、文件範本、個別工作腳本，以及用 Flask 預覽活動資料的簡易網頁。各腳本目前仍依任務分開執行，實際需求可能因範本、輸入檔案及本機環境而不同。

執行前請確認依賴套件、資料路徑與範本均已準備完成；若要公開作為作品集，建議先移除或匿名化真實姓名、聯絡資料、報名名單及其他內部文件。
