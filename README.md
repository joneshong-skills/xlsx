[English](README.md) | [繁體中文](README.zh.md)

# XLSX Skill

A Claude Code skill for creating, reading, editing, and analyzing Excel spreadsheets (.xlsx, .xlsm, .csv, .tsv) with professional formatting, formulas, and financial modeling support.

## Features

- **Read & Analyze**: Load spreadsheets with pandas for data analysis, statistics, and visualization
- **Create**: Build new Excel files with openpyxl including formulas, formatting, and charts
- **Edit**: Modify existing workbooks while preserving templates, styles, and conventions
- **Financial Modeling**: Industry-standard color coding, number formatting, and formula construction
- **Formula Recalculation**: Automated recalc via LibreOffice with error detection and reporting
- **Data Cleaning**: Restructure messy tabular data (malformed rows, misplaced headers) into proper spreadsheets
- **Format Conversion**: Convert between .xlsx, .xlsm, .csv, and .tsv formats

## Installation

```bash
# Copy to your skills directory
cp -r xlsx/ ~/.claude/skills/xlsx/
```

### Python Dependencies

```bash
pip install openpyxl pandas
```

### System Dependencies (macOS)

```bash
brew install --cask libreoffice
```

## Usage

Simply ask Claude to work with a spreadsheet:

- "Create a financial model for Q4 projections"
- "Add a summary sheet with formulas to this Excel file"
- "Clean up this messy CSV and export as xlsx"
- "Read and analyze the data in sales.xlsx"
- "Format this spreadsheet with proper number formatting"

## File Structure

```
xlsx/
  SKILL.md          -- Main skill instructions
  LICENSE.txt       -- License terms
  references/       -- (reserved for reference materials)
  scripts/
    recalc.py       -- Formula recalculation via LibreOffice
    office/         -- Office document utilities (pack, unpack, validate)
      helpers/      -- XML processing helpers
      schemas/      -- OOXML XSD schemas for validation
      validators/   -- Format-specific validators
  assets/           -- (reserved for future assets)
```

## License

Proprietary. See LICENSE.txt for complete terms.

## Source

Adapted from [anthropics/skills](https://github.com/anthropics/skills/tree/main/skills/xlsx).
