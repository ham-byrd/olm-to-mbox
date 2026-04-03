# olm-to-mbox

Convert Outlook for Mac `.olm` archives to `.mbox` files for importing into Thunderbird, Gmail, or other email clients.

## What it does

- Extracts the `.olm` ZIP archive
- Parses email XML files (headers, body, HTML, attachments)
- Writes standard `.mbox` files organized by folder (Inbox, Sent Items, etc.)

## Requirements

- Python 3.7+
- No external dependencies (uses only Python standard library)

## Usage

```bash
python3 olm_to_mbox.py input.olm [output_directory]
```

If no output directory is specified, it creates one named `<input>_mbox` next to the input file.

### Example

```bash
python3 olm_to_mbox.py ~/Desktop/MyEmail.olm ~/Desktop/converted
```

Output:
```
Found 1523 emails in 8 folders.

  Inbox -> Inbox.mbox (892 messages)
  Sent Items -> Sent_Items.mbox (431 messages)
  Drafts -> Drafts.mbox (12 messages)
  ...

Done! Converted 1523 emails, 0 failed.
```

## Importing into Gmail via Thunderbird

1. Install [Thunderbird](https://www.thunderbird.net/)
2. Add your Gmail account to Thunderbird (IMAP)
3. Install the **ImportExportTools NG** add-on in Thunderbird
4. Right-click a local folder > **ImportExportTools NG** > **Import mbox file**
5. Select the `.mbox` files from the output directory
6. Drag/copy the imported emails into your Gmail IMAP folders
7. Wait for sync to complete -- emails will appear in Gmail

## License

MIT
