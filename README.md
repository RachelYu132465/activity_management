# Activity Management Toolkit
**For user who is required to produce strict format of document, many are in Microsoft word format.**

The project brings together structured JSON data, Excel workbooks, Word/PDF document generation, speaker and influencer profiles, and a small Flask preview app. It is designed to reduce repetitive preparation work across event planning and administration.

> NOTE: All files other than python extension is an example file for user to check firstly if the result is what you need. If you are administrative assitants searching for office-automation tool, this library is intended to keep raw data into json format, and then export them into Microsoft, pdf, png, etc. So check `program_data.json` first to see if the data structure etc suit you.
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

## Key Function file
| Python File name | Function |
|---|---|
|add_program_data_to_excel|Add program_data to every row of an Excel file.|
|app|HTML format of program_data host=127.0.0.1:5000 (default) and separate data into subpages: Cover, notes, chairs, speakers|
|siginin_table_from_excel|Read Excel file and export them into Word file with specific format (table size, font size, row rules, etc).|
|mergeExcelByKey| Merge multiple Excels into one Excel using one column value.|
|test_printer|Test your selected printer first, and then print.|
|print_file|Batch print files in specified folder.|
|send_email_with_attachments|Input template Word file and program_data to form email. SMTP_test.py is at the root folder to test your connection first, and then set your SMTP host (Gmail or Outlook, etc), sender email address and password.|

Below are of the same function, export program_data.json into other format:|generate_agenda, parse_agenda_docx, render_to_docx_2.py, render_to_pdf.py, render_with_chrome.py, schedule_table.py, siginin_table_from_excel.py, siginin_table_from_json.py |format written in the code||make_speaker_letters| You have template Word file. ||create_PNG_deskcard|png file|


## Requirements

- Python 3.10 or newer is recommended.
- Install the dependencies required by the workflow you want to run. The repository does not currently provide a single authoritative `requirements.txt`.
- The Flask preview app requires Flask and Jinja2.
- Excel and Word workflows may require packages such as `openpyxl` and `python-docx`; PDF rendering or browser-based workflows may have additional system dependencies.

For example, to run the preview app:

```bash
python -m pip install Flask Jinja2
```



## 中文簡介

這是一個以 Python 建立的活動行政流程自動化工具集，主要用來處理活動議程、講者／專業人士資料、Excel 資料填入與合併，以及 Word、PDF、簽到表和相關活動文件的產生。

專案包含共用資料處理模組、JSON schema、文件範本、個別工作腳本，以及用 Flask 預覽活動資料的簡易網頁。各腳本目前仍依任務分開執行，實際需求可能因範本、輸入檔案及本機環境而不同。

執行前請確認依賴套件、資料路徑與範本均已準備完成；若要公開作為作品集，建議先移除或匿名化真實姓名、聯絡資料、報名名單及其他內部文件。
