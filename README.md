# Crosstab Studio

**Turn SPSS banner files into clean, client-ready tables in seconds.**

![Python](https://img.shields.io/badge/python-3.10+-blue) ![Streamlit](https://img.shields.io/badge/streamlit-1.x-red) ![License](https://img.shields.io/badge/license-MIT-green)

Analysts spend days manually extracting crosstabs from SPSS-generated banner files — reformatting headers, relocating net rows, stripping artifacts, rebuilding the whole thing when the wave-2 file arrives slightly differently. Crosstab Studio eliminates that work. Upload a banner file, select your questions and columns, and export a formatted deliverable in under five minutes.

## Features

- **Auto format detection** — matches against a library of known banner profiles (KP, KP Omni, IData, Corporate Reputation, Global Brand Identity); unknown formats trigger a deep scan that learns and saves the structure
- **Multi-wave comparison** — upload W1, W2, and W3 simultaneously; output includes colour-coded comparison tables for wave-over-wave shifts
- **Clean table structure** — net rows moved to bottom with a visual separator, base N displayed in column headers, table of contents auto-generated
- **Format wizard** — scans the first 15 sheets, shows a colour-coded row preview, saves detected format as a named reusable profile
- **Dual export** — formatted Excel (`.xlsx`) or Word (`.docx`) tables; output filename set automatically as `BannerName_reformatted.xlsx`
- **Extensible profile library** — run the deep scan once on a new vendor format, name it, and every future file from that source is handled automatically

## How it works

1. **Upload** your SPSS-generated banner Excel file (single wave or multi-wave)
2. **Detect** — Crosstab Studio matches the file against its profile library, or runs a deep scan if the format is unrecognised
3. **Select** questions and banner columns from a clean UI; preview row structure before exporting
4. **Export** to formatted Excel or Word — output is named and structured automatically

## Supported formats

| Profile | Vendor | Key characteristics |
|---|---|---|
| KP | Knowledge Panel | Fixed header rows, specific net row labelling, KP column ordering |
| KP Omni | Knowledge Panel (omnibus) | Shared-banner layout, condensed column headers |
| IData | IData | Multi-level column headers, percentage and n rows interleaved |
| Corporate Reputation | Custom | Attribute grid structure, stakeholder segment banners |
| Global Brand Identity | Custom | Multi-market layout, market-as-column structure |

New formats can be saved as named profiles via the Format Wizard.

## Tech

- Python 3.10+
- Streamlit
- pandas
- openpyxl
- python-docx

## Run locally

```bash
git clone https://github.com/camahoy/crosstab-studio.git
cd crosstab-studio
pip install -r requirements.txt
streamlit run app.py
```

Opens at `http://localhost:8501`.
