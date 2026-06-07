# TWoM Translation Tool

A small desktop tool for translating `This War of Mine` language files. It imports binary `.lang` / `.lngp` files, creates an editable CSV, helps you fill in translations, and packs the result back into a `.lang` file for use with Storyteller.

## Features

- Import TWoM `.lang` or `.lngp` files
- Export and reopen a working CSV with `key`, `source`, and `translation` columns
- Automatically reopen the last working CSV on app launch when it is still available
- Browse strings by group and work through queue filters for `All`, `Untranslated`, `Draft`, `Reviewed`, and `Warnings`
- Edit translations in the app or externally in Excel/LibreOffice
- Review inline warnings for placeholders, line breaks, whitespace, and inconsistent duplicates
- Reuse suggestions from identical source strings or similar keys in the same group
- Copy the source text into the translation field, then tweak only the words that changed
- Insert source or project-wide placeholders like `{0}`, `{ms|}`, `%s`, `\n`, or `^CharacterName^` from the editor menu
- Pack translated CSV content back into a binary `.lang` file
- Optional AI translation with Anthropic, OpenAI, Google Gemini, or OpenRouter using your own API key

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

The app also writes an optional sidecar metadata file next to the CSV:

```text
<your-file>.csv.twom-meta.json
```

This stores review states plus UI state like the last selected group, filter, search, and active key. The CSV format itself stays unchanged.

## Keyboard Shortcuts

- `Ctrl+I` import `.lang`
- `Ctrl+O` open CSV
- `Ctrl+S` save CSV
- `Ctrl+P` pack to `.lang`
- `Ctrl+F` or `F6` focus search
- `Ctrl+E` focus the translation editor
- `Ctrl+Enter` save current translation
- `Ctrl+Shift+Enter` save and mark reviewed
- `Ctrl+Up` / `Ctrl+Down` move to previous or next visible row
- `Ctrl+G` jump to the next untranslated string
- `Ctrl+Shift+W` jump to the next warning
- `Ctrl+1` to `Ctrl+5` switch queue filters: All, Untranslated, Draft, Reviewed, Warnings

## AI Translation

The app can translate one string at a time, review a draft translation with AI feedback, or batch-translate empty strings using multiple providers.

Supported providers:

- Anthropic
- OpenAI
- Google Gemini
- OpenRouter

1. Open the settings window in the app.
2. Pick the active AI provider.
3. Enter that provider's API key.
4. Optionally change the model id for that provider.
5. Set the target language in the main window.
6. Use `AI suggest` for a draft translation, `AI feedback` for review/ideation on your current translation, or the batch AI action for empty strings.

The selected provider, per-provider API keys, per-provider model ids, and target language are stored locally in:

```text
~/.twom.json
```

AI translation uses your own provider credits. Review generated translations before packing them into a release.

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

## Git Setup

This folder currently contains the project files, but it may still need to be initialized as a Git repository:

```powershell
git init
git add .
git commit -m "Initial commit"
```

## License

No license has been added yet. Add one before publishing if you want others to use, modify, or redistribute the project.
