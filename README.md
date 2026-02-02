# Outlook Mac Recovery Skills

This repository contains OpenCode-style SKILL files for recovering emails from Outlook for Mac cache files.

## Overview

This collection of skills provides a pipeline for recovering emails from Outlook for Mac when you have access to the local cache but not the account credentials.

## Pipeline

```
01-locate-outlook-cache → 02-export-cache-to-olm → 03-olm-convert → 04-olm-extract-email-addresses
```

### Skills

1. **01-locate-outlook-cache**: Find Outlook profile/cache directories on macOS
   - Locates Classic Outlook (2016/2019/2021/365) and legacy (2011) profiles
   - Returns JSON array of candidate profile paths

2. **02-export-cache-to-olm**: Drive Classic Outlook to export `.olm` archive
   - Uses AppleScript to trigger Outlook's export wizard
   - Note: Requires human-in-the-loop for export completion

3. **03-olm-convert**: Convert `.olm` to `.eml` or `.html`
   - Uses `olm-convert` Python tool
   - Preserves folder hierarchy
   - Optional attachment handling

4. **04-olm-extract-email-addresses**: Extract email addresses from `.olm`
   - Uses `outlook-olm-email-exporter`
   - Extracts From/To/Cc/Bcc addresses
   - Deduplicated output

## Important Limitations

- **New Outlook (preview)**: This pipeline does NOT work with the "New Outlook" for Mac, which is web-based and stores data differently. You must use Classic Outlook.
- **Account access**: If Outlook profile cache is encrypted and you cannot open Outlook, third-party tools cannot bypass the encryption.
- **Export step**: The `.olm` export from Classic Outlook requires user interaction with the Export wizard - it cannot be fully automated via scripting.

## References

### Microsoft Documentation
- Outlook for Mac is a locally cached email client: https://learn.microsoft.com/en-us/troubleshoot/outlook/installation/outlook-for-mac-is-a-locally-cached-email-client
- Export items to archive file (.olm): https://support.microsoft.com/en-us/office/export-items-to-an-archive-file-in-outlook-for-mac-281a62bf-cc42-46b1-9ad5-6bda80ca3108

### Open Source Tools Referenced
- `olm-convert` (MIT): https://github.com/PeterWarrington/olm-convert
- `outlook-olm-email-exporter` (Python): https://github.com/eercanayar/outlook-olm-email-exporter

## License

These SKILL files are provided as templates for agent frameworks. The referenced third-party tools have their own licenses.
