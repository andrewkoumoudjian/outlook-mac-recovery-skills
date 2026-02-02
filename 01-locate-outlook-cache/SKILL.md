# Skill: locate-outlook-cache

## Description

Locate and list Outlook for Mac profile/cache directories on macOS, including Classic Outlook profile folders such as `Outlook 15 Profiles/Main Profile`, so that downstream tools can copy or process the cached data.

## Environment

- OS: macOS
- Requirements:
  - `python3` in PATH
  - Permission to read `~/Library`

## Capabilities

- Detects installed Outlook profile roots for:
  - Outlook 2016/2019/2021/365 (Classic Outlook)
  - Outlook 2011 (legacy, best-effort)
- Returns absolute paths to candidate profile/cache directories.

## Tools

### Tool: list_outlook_profiles

- **Type**: python
- **Entry point**: `python3 locate_outlook_profiles.py`
- **Arguments**:
  - `--include-legacy` (optional, flag) – also search Outlook 2011 paths.
- **Output**:
  - JSON array of discovered profile objects:
    ```json
    [
      {
        "variant": "classic",
        "path": "/Users/USER/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile"
      },
      {
        "variant": "legacy",
        "path": "/Users/USER/Documents/Microsoft User Data/Office 2011 Identities/Main Identity"
      }
    ]
    ```

### Example implementation (locate_outlook_profiles.py)

```python
#!/usr/bin/env python3
import argparse, json, os, pathlib

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--include-legacy", action="store_true")
    args = parser.parse_args()

    home = pathlib.Path.home()
    candidates = []

    classic = home / "Library" / "Group Containers" / "UBF8T346G9.Office" / "Outlook" / "Outlook 15 Profiles"
    if classic.is_dir():
        for p in classic.iterdir():
            if p.is_dir():
                candidates.append({"variant": "classic", "path": str(p)})

    if args.include_legacy:
        legacy = home / "Documents" / "Microsoft User Data" / "Office 2011 Identities"
        if legacy.is_dir():
            for p in legacy.iterdir():
                if p.is_dir():
                    candidates.append({"variant": "legacy", "path": str(p)})

    print(json.dumps(candidates))

if __name__ == "__main__":
    main()
```

## Usage

- Call `list_outlook_profiles` first in any recovery pipeline.
- Feed returned `path` values into later skills that copy/export or inspect the data.

## References

- Microsoft docs describing Outlook for Mac profile cache locations: [https://learn.microsoft.com/en-us/troubleshoot/outlook/installation/outlook-for-mac-is-a-locally-cached-email-client]
- Profile path: `~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles` for Outlook 2016+
- Legacy Outlook 2011 identity path under `Office 2011 Identities`

### Citation IDs
- [web:18] Outlook for Mac is a locally cached email client - Microsoft Learn
- [web:87] Where is Outlook Main Profile information stored? - Microsoft Q&A
