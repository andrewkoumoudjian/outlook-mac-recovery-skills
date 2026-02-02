# Skill: olm-convert

## Description

Use the open-source `olm-convert` toolkit to convert an Outlook for Mac `.olm` archive into a tree of standard `.eml` or `.html` message files for downstream parsing or ingestion.

## Environment

- OS: macOS or Linux
- Requirements:
  - `python3` in PATH
  - `olm-convert` installed (via `pip` or cloned repo)
    - Example: `pip install olm-convert` or local clone of the GitHub project

## Capabilities

- Converts a `.olm` file to `.eml` or `.html`.
- Preserves folder hierarchy in the output directory.
- Optionally skips attachments for lighter exports.
- Can also export contacts to vCard format and calendar to ICS format.

## Tools

### Tool: convert_olm_to_eml

- **Type**: shell
- **Entry point**: `python3 olmConvert.py`
- **Arguments**:
  - `olmPath` (string) – path to source `.olm`.
  - `outputDir` (string) – directory where converted mail will be written.
  - `--format {eml,html}` (optional, default `eml`).
  - `--noAttachments` (optional, flag).
  - `--verbose` (optional, flag).
- **Output**:
  - File tree like:
    - `<outputDir>/user@example.com/Inbox/Subject - 2025-02-02 21.04.56.eml`

### Example command

```bash
python3 olmConvert.py \
  --format eml \
  --verbose \
  "/path/to/archive.olm" \
  "/path/to/output"
```

## Usage

- Run this skill after `export-cache-to-olm` has created an `.olm`.
- Feed resulting `.eml` files into other agents for parsing, deduping, or indexing.

## References

- `olm-convert` GitHub project (free, MIT-licensed OLM converter) with CLI usage and module API: [https://github.com/PeterWarrington/olm-convert]
- Example usage and options taken from the project's README, including `olmConvert.py` CLI syntax and `convertOLM(olmPath, outputDir, ...)` module entry point.

### Citation IDs
- [web:47] PeterWarrington/olm-convert - GitHub
