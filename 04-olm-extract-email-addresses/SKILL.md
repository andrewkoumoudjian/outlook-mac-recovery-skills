# Skill: olm-extract-email-addresses

## Description

Given an Outlook for Mac `.olm` archive, extract unique email addresses (senders, recipients) using the `outlook-olm-email-exporter` Python tool for quick contact recovery or graph building.

## Environment

- OS: macOS or Linux
- Requirements:
  - `python3` in PATH
  - Clone or install `outlook-olm-email-exporter` from GitHub

## Capabilities

- Parses `.olm` file contents.
- Extracts email addresses from:
  - `From`
  - `To`
  - `Cc`
  - `Bcc`
- Outputs a deduplicated list to stdout or a CSV file (depending on implementation).

## Tools

### Tool: extract_addresses_from_olm

- **Type**: python
- **Entry point**: `python3 extract_olm_emails.py`
- **Arguments**:
  - `--olm` (string) – path to `.olm` file.
  - `--output` (string, optional) – path to write a CSV or TXT of addresses.
- **Output**:
  - Either a newline-separated list on stdout, or a CSV file with columns like `email,address_source`.

### Example implementation wrapper (extract_olm_emails.py)

```python
#!/usr/bin/env python3
"""
Thin wrapper around outlook-olm-email-exporter functionality.
"""

import argparse, sys

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--olm", required=True, help="Path to OLM file")
    parser.add_argument("--output", help="Optional output file for unique addresses")
    args = parser.parse_args()

    # NOTE: In a real setup, import the library code from the cloned repo instead:
    # from olm_exporter import extract_addresses
    #
    # addresses = extract_addresses(args.olm)
    # For this template, just show the interface.
    addresses = set()  # placeholder

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            for addr in sorted(addresses):
                f.write(addr + "\n")
    else:
        for addr in sorted(addresses):
            print(addr)

if __name__ == "__main__":
    main()
```

(Replace the placeholder logic with the actual API calls from the cloned project.)

## Usage

- Call this skill after you have an `.olm` file (either from Outlook export or from another machine).
- Use the resulting email list for:
  - Rebuilding contact lists.
  - Cross-referencing with other datasets.
  - Seeding further OSINT or graph analysis.

## References

- `outlook-olm-email-exporter` repository: "extract and export e-mail addresses from outlook .olm files using python" [https://github.com/eercanayar/outlook-olm-email-exporter]
- Additional OLM-to-EML converter scripts (e.g., `convert_olm_to_eml.py` gists) can complement this skill for full message export.

### Citation IDs
- [web:82] eercanayar/outlook-olm-email-exporter - GitHub
- [web:79] Convert Outlook for Mac (OLM) to EML files as PST - GitHub Gist
- [web:88] Convert Outlook for Mac (OLM) to EML files - GitHub Gist
