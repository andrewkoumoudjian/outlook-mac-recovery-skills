# Skill: export-cache-to-olm

## Description

Drive Classic Outlook for Mac (not "New Outlook") via AppleScript to export an `.olm` archive from the local profile cache, enabling offline processing without live account credentials.

## Environment

- OS: macOS
- Requirements:
  - Classic Outlook for Mac installed and able to open the profile
  - `osascript` available
  - User must grant automation permissions if prompted

## Capabilities

- Initiates an Outlook export of:
  - All mail items, or
  - A specified folder hierarchy (optional extension)
- Saves resulting `.olm` to a specified target directory.

## Tools

### Tool: export_outlook_profile_to_olm

- **Type**: shell
- **Entry point**: `osascript export_to_olm.scpt`
- **Arguments**:
  - `exportPath` (POSIX path string) – folder where the `.olm` should be saved.
- **Output**:
  - Path to the generated `.olm` file (printed to stdout).

### Example implementation (export_to_olm.scpt)

```applescript
on run argv
    if (count of argv) is 0 then
        error "exportPath argument required"
    end if
    set exportFolder to POSIX file (item 1 of argv) as alias

    tell application "Microsoft Outlook"
        activate
        -- This uses the built-in Export UI; fully headless export is limited.
        display dialog "Outlook export will start. Choose 'Items of the following types' and continue." buttons {"OK"} default button "OK"
    end tell

    -- NOTE: A fully automated export (without UI) is not officially documented.
    -- In practice, human-in-the-loop interaction with Outlook's Export wizard
    -- is usually required to create the .olm file at exportFolder.
end run
```

(You can replace the dialog flow with a more aggressive UI scripting sequence if your environment allows it.)

## Usage

- Agent calls `export_outlook_profile_to_olm` to instruct the user and trigger Outlook's export wizard.
- After the user completes the wizard and selects `exportPath`, the agent scans that folder for newly created `.olm` files.

## References

- Official export instructions for Outlook for Mac using the **Export to Archive File (.olm)** wizard: [https://support.microsoft.com/en-us/office/export-items-to-an-archive-file-in-outlook-for-mac-281a62bf-cc42-46b1-9ad5-6bda80ca3108]
- Microsoft notes that Outlook for Mac is a locally cached email client and exports from that cache to `.olm` files.

### Citation IDs
- [web:2] Export items to an archive file in Outlook for Mac - Microsoft Support
- [web:18] Outlook for Mac is a locally cached email client - Microsoft Learn
- [web:84] Manually export Outlook for Mac emails/archive - Nucleus Technologies
