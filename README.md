# santranslit

Transliterate Sanskrit (Devanagari) text inside Microsoft Word `.docx` documents into Tamil script while preserving original formatting and non-Sanskrit content.

## Overview

`santranslit` scans a Word document, detects Devanagari Sanskrit text, and transliterates only those segments into Tamil script. All other text, formatting, styles, tables, and document structure remain unchanged.

The tool treats the source document as read-only and writes results to a separate output file.

## Capabilities

* Transliterate Devanagari Sanskrit → Tamil script
* Preserve formatting (bold, italics, styles, tables, etc.)
* Process paragraphs and nested tables
* Transliterate only Sanskrit spans within mixed-language runs
* Optional restriction by run style
* Optional Tamil font application to fully Sanskrit runs
* Dry-run analysis mode
* CLI-based workflow suitable for automation

## Requirements

* Python ≥ 3.13
* Microsoft Word `.docx` input files

## Dependencies

* `python-docx`
* `indic-transliteration`

## Usage

```bash
python santranslit/sansliterate.py input.docx -o output.docx
```

### Dry Run (No File Modification)

```bash
python santranslit/sansliterate.py input.docx --dry-run
```

### Enable Verbose Logging

```bash
python santranslit/sansliterate.py input.docx -v
```

### Restrict Transliteration to Specific Style

```bash
python santranslit/sansliterate.py input.docx --style Sanskrit
```

---

### Specify Tamil Font

```bash
python santranslit/sansliterate.py input.docx --tamil-font Vijaya
```

---

### Combined Example

```bash
python santranslit/sansliterate.py input.docx -o output.docx --style Sanskrit --tamil-font Vijaya -v
```

---

## How It Works

1. The input `.docx` is copied to an output path.
2. The document is traversed paragraph-by-paragraph.
3. Runs are scanned for Devanagari Unicode characters.
4. Only detected Sanskrit spans are transliterated.
5. Formatting is preserved.
6. The output document is saved and validated.

## Transliteration Rules

### Sanskrit Detection

Sanskrit text is identified using the Devanagari Unicode block:

```
U+0900 – U+097F
```

### Mixed Runs

Only Devanagari spans are transliterated. All other characters remain unchanged.

Example:

```
Original: This is मन्त्र text
Result:   This is மந்திர text
```

---

## Font Handling

Tamil font is applied only when:

* The run contains Sanskrit characters, and
* The run contains no non-Sanskrit alphabetic characters

This prevents font corruption in mixed-language runs.

## CLI Reference

### Positional Arguments

| Argument | Description                |
| -------- | -------------------------- |
| `INPUT`  | Path to input `.docx` file |

### Options

| Option            | Description                                               |
| ----------------- | --------------------------------------------------------- |
| `-o`, `--output`  | Output `.docx` file path                                  |
| `--style`         | Restrict transliteration to runs with matching style name |
| `--tamil-font`    | Font applied to fully Sanskrit runs (default: `Vijaya`)   |
| `--dry-run`       | Report modifications without writing output               |
| `-v`, `--verbose` | Enable debug logging                                      |

## Output Behavior

Default output naming:

```
<input_stem>_ta.docx
```

Example:

```
manuscript.docx → manuscript_ta.docx
```

## Limitations

* Only `.docx` format is supported
* Sanskrit must be encoded in Devanagari
* Non-Devanagari Sanskrit (e.g. IAST) is not processed
* Font application depends on Word font availability
