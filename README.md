# Print Shop Automation

A two-script automation pipeline for running a print business from a phone. Sorts incoming print jobs by file type, calculates page counts and pricing, and generates a ready-to-send customer invoice, plus keeps a full job history archive automatically.

## How it works

**1. Sorter Bot** (`file_sorter_bot.py`)
Scans an inbox folder of freshly-received files and automatically sorts them into the correct category:
- PowerPoint files (`.ppt` / `.pptx`) → PPTX folder
- PDF files → detected as **portrait** (standard documents) or **landscape** (slide decks) based on page dimensions, then routed accordingly

**2. Printing Price Calculator** (`printing_price_calculator.py`)
Processes the sorted files for a named client and:
- Counts pages/slides per document
- Pads standard documents to the next even page count (for double-sided printing), and slide decks to the next multiple of 8 (4-per-page layout)
- Calculates total A4 sheets needed and the resulting price, using configurable cost and selling price per sheet
- Optionally merges multiple PDFs into single output files
- Generates a formatted invoice breakdown (ready to copy into WhatsApp) alongside an internal cost/profit record
- Archives all source files and the invoice into a dated job history folder, keeping the workspace clean for the next job

## Why it exists

Manually counting pages, calculating prices, and tracking jobs for a print business is repetitive and error-prone at volume. This pipeline turns a folder of raw customer files into a calculated invoice and an organized job archive in one run.

## Built with

- Python
- `pypdf` (PDF page counting, orientation detection, merging)
- `python-pptx` (PowerPoint slide counting)
- Built and run entirely on Android using Pydroid 3

## Setup

1. Install dependencies in Pydroid's pip menu: `pypdf`, `python-pptx`
2. Set `FOLDER_ROOT` in both scripts to your working folder
3. Adjust `PRICE_PER_SHEET_SELLING` and `PRICE_PER_SHEET_COST` in `printing_price_calculator.py` to your own rates
4. Drop incoming files into the `_Inbox_Dump` folder, run `file_sorter_bot.py` to sort them
5. Run `printing_price_calculator.py`, enter the client's name when prompted, and get a calculated invoice plus organized output

## Note

Built as a personal tool for an actual running print business (StephenPRINTS), not a demo project. File paths in the config reflect a personal folder structure and should be changed to match your own setup.
