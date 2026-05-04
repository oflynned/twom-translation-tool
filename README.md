# TWoM Translation Tool

A small desktop tool for translating `This War of Mine` language files. It imports binary `.lang` / `.lngp` files, creates an editable CSV, helps you fill in translations, and packs the result back into a `.lang` file for use with Storyteller.

## Features

- Import TWoM `.lang` or `.lngp` files
- Export and reopen a working CSV with `key`, `source`, and `translation` columns
- Browse strings by group and search/filter untranslated entries
- Edit translations in the app or externally in Excel/LibreOffice
- Pack translated CSV content back into a binary `.lang` file
- Optional AI translation through the Anthropic Messages API

## Requirements

- Windows
- Python 3.9 or newer
- No third-party Python packages are required; the app uses the Python standard library

## Quick Start

1. Install Python 3.9+ from [python.org](https://www.python.org/).
2. During installation, enable **Add Python to PATH**.
3. Run `Launch.bat`.
4. Click **Import .lang** and select a source language file.
5. Save the generated working CSV.
6. Translate strings in the app, or edit the CSV in a spreadsheet editor.
7. Click **Pack to .lang** to create the translated language file.
8. Open Storyteller, select your mod, pack it, then test in game.

You can also run the app directly:

```powershell
python twom_translator.py
```

## CSV Format

Working CSV files use these columns:

| Column | Description |
| --- | --- |
| `key` | Internal string identifier from the game file |
| `source` | Original source text |
| `translation` | Your translated text |

Leave `key` unchanged. Empty `translation` cells will fall back to the source text when packing unless they are filled later.

## AI Translation

The app can translate one string at a time or batch-translate empty strings using Anthropic.

1. Open the settings window in the app.
2. Enter your Anthropic API key.
3. Set the target language.
4. Use the single or batch AI translation actions.

The API key and target language are stored locally in:

```text
~/.twom_v5.json
```

AI translation uses API credits. Review generated translations before packing them into a release.

## Project Files

```text
twom_translator.py  Main Tkinter desktop application
lang_parser.py      Binary .lang parser, builder, and CSV helpers
Launch.bat          Windows launcher with a Python availability check
```

## Development Notes

The `.lang` parser expects the TWoM language binary format:

```text
[data_size: uint32 LE] [record_count: uint32 LE]
per record:
  [key_len: uint16 LE] [key: ASCII bytes]
  [val_len: uint16 LE] [value: UTF-16 LE]
```

Generated CSV files are written with UTF-8 BOM encoding so spreadsheet tools on Windows open them cleanly.
