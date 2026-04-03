#!/usr/bin/env python3
"""
Convert Outlook for Mac .olm archive to .mbox files.

Usage:
    python3 olm_to_mbox.py input.olm [output_directory]

OLM files are ZIP archives containing emails as XML files.
This script extracts them and writes standard .mbox files
that can be imported into Thunderbird via ImportExportTools NG.
"""

import sys
import os
import zipfile
import mailbox
import xml.etree.ElementTree as ET
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import formatdate, formataddr, parsedate_to_datetime
from email import encoders
from datetime import datetime, timezone
import base64
import re
import html


def parse_olm_address(addr_elem):
    """Parse an OLM address element into (name, email) tuple."""
    if addr_elem is None:
        return None
    # Address elements can contain sub-elements or direct text
    address_list = addr_elem.findall('.//emailAddress')
    if address_list:
        results = []
        for a in address_list:
            name = a.findtext('OPFContactEmailAddressName', '').strip()
            email = a.findtext('OPFContactEmailAddressAddress', '').strip()
            if email:
                results.append((name, email))
        return results
    # Try direct text
    text = addr_elem.text
    if text and '@' in text:
        return [('', text.strip())]
    return []


def parse_olm_date(date_str):
    """Parse OLM date string to datetime."""
    if not date_str:
        return None
    date_str = date_str.strip()
    # OLM uses various date formats
    formats = [
        '%Y-%m-%dT%H:%M:%S',
        '%Y-%m-%dT%H:%M:%S%z',
        '%Y-%m-%d %H:%M:%S %z',
        '%Y-%m-%d %H:%M:%S',
        '%a, %d %b %Y %H:%M:%S %z',
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    # Try removing fractional seconds
    date_str_clean = re.sub(r'\.\d+', '', date_str)
    for fmt in formats:
        try:
            return datetime.strptime(date_str_clean, fmt)
        except ValueError:
            continue
    return None


def format_address_list(addresses):
    """Format list of (name, email) tuples into a header string."""
    if not addresses:
        return ''
    parts = []
    for name, email in addresses:
        if name:
            parts.append(formataddr((name, email)))
        else:
            parts.append(email)
    return ', '.join(parts)


def xml_to_email(xml_content, attachments_data=None):
    """Convert OLM XML email content to a Python email.message.Message."""
    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError:
        return None

    # Handle namespace if present
    # OLM emails can have different root tags
    # Try common element names with and without namespace

    def find_text(tag):
        """Find text for a tag, trying with and without common prefixes."""
        elem = root.find(f'.//{tag}')
        if elem is not None and elem.text:
            return elem.text.strip()
        return ''

    subject = find_text('OPFMessageCopySubject')
    html_body = find_text('OPFMessageCopyHTMLBody')
    text_body = find_text('OPFMessageCopyBody')
    sent_time = find_text('OPFMessageCopySentTime')
    recv_time = find_text('OPFMessageCopyReceivedTime')
    message_id = find_text('OPFMessageCopyMessageID')

    # Parse addresses
    from_addrs = parse_olm_address(root.find('.//OPFMessageCopyFromAddresses'))
    to_addrs = parse_olm_address(root.find('.//OPFMessageCopyToAddresses'))
    cc_addrs = parse_olm_address(root.find('.//OPFMessageCopyCCAddresses'))
    bcc_addrs = parse_olm_address(root.find('.//OPFMessageCopyBCCAddresses'))

    # Also try sender address as fallback
    if not from_addrs:
        sender = root.find('.//OPFMessageCopySenderAddress')
        if sender is not None:
            from_addrs = parse_olm_address(sender)

    # Parse date
    date = parse_olm_date(sent_time) or parse_olm_date(recv_time)

    # Find attachment references
    attachment_elems = root.findall('.//messageAttachment')

    has_attachments = bool(attachment_elems and attachments_data)

    # Build the email message
    if has_attachments or (html_body and text_body):
        msg = MIMEMultipart('mixed' if has_attachments else 'alternative')
        if html_body and text_body:
            alt = MIMEMultipart('alternative') if has_attachments else msg
            alt.attach(MIMEText(text_body, 'plain', 'utf-8'))
            alt.attach(MIMEText(html_body, 'html', 'utf-8'))
            if has_attachments:
                msg.attach(alt)
        elif html_body:
            msg.attach(MIMEText(html_body, 'html', 'utf-8'))
        elif text_body:
            msg.attach(MIMEText(text_body, 'plain', 'utf-8'))
    elif html_body:
        msg = MIMEText(html_body, 'html', 'utf-8')
    elif text_body:
        msg = MIMEText(text_body, 'plain', 'utf-8')
    else:
        msg = MIMEText('', 'plain', 'utf-8')

    # Set headers
    if subject:
        msg['Subject'] = subject
    if from_addrs:
        msg['From'] = format_address_list(from_addrs)
    if to_addrs:
        msg['To'] = format_address_list(to_addrs)
    if cc_addrs:
        msg['Cc'] = format_address_list(cc_addrs)
    if bcc_addrs:
        msg['Bcc'] = format_address_list(bcc_addrs)
    if message_id:
        msg['Message-ID'] = message_id if message_id.startswith('<') else f'<{message_id}>'
    if date:
        msg['Date'] = formatdate(date.timestamp(), localtime=True)
    else:
        msg['Date'] = formatdate(localtime=True)

    # Add attachments
    if has_attachments and attachments_data:
        for att_elem in attachment_elems:
            att_name = att_elem.findtext('OPFAttachmentName', '')
            att_url = att_elem.findtext('OPFAttachmentURL', '')
            att_mime = att_elem.findtext('OPFAttachmentContentType', 'application/octet-stream')

            if att_url and att_url in attachments_data:
                att_data = attachments_data[att_url]
                maintype, _, subtype = att_mime.partition('/')
                if not subtype:
                    maintype = 'application'
                    subtype = 'octet-stream'
                part = MIMEBase(maintype, subtype)
                part.set_payload(att_data)
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition', 'attachment',
                    filename=att_name or 'attachment'
                )
                msg.attach(part)

    return msg


def convert_olm_to_mbox(olm_path, output_dir):
    """Convert an .olm file to .mbox files organized by folder."""
    if not os.path.isfile(olm_path):
        print(f"Error: File not found: {olm_path}")
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)

    print(f"Opening {olm_path}...")
    try:
        zf = zipfile.ZipFile(olm_path, 'r')
    except zipfile.BadZipFile:
        print("Error: Not a valid .olm (ZIP) file.")
        sys.exit(1)

    # Categorize files by folder
    email_files = {}   # folder_path -> list of xml filenames
    all_files = zf.namelist()

    # Identify email XML files (they're typically in message folders)
    # OLM structure: account/Data/FolderName/message.xml or similar
    for name in all_files:
        if name.endswith('.xml') and '/Data/' in name:
            # Skip contacts, calendar, etc.
            lower = name.lower()
            if any(skip in lower for skip in ['/contacts/', '/calendar/', '/tasks/', '/notes/']):
                continue
            # Extract folder name from path
            parts = name.split('/Data/')
            if len(parts) >= 2:
                folder_path = os.path.dirname(parts[1])
                if not folder_path:
                    folder_path = 'Root'
                if folder_path not in email_files:
                    email_files[folder_path] = []
                email_files[folder_path].append(name)

    if not email_files:
        # Fallback: try all XML files
        for name in all_files:
            if name.endswith('.xml'):
                folder_path = os.path.dirname(name) or 'Root'
                # Clean up folder path for use as filename
                folder_path = folder_path.replace('/', '_').strip('_')
                if not folder_path:
                    folder_path = 'Root'
                if folder_path not in email_files:
                    email_files[folder_path] = []
                email_files[folder_path].append(name)

    if not email_files:
        print("No email XML files found in the archive.")
        print(f"Archive contains {len(all_files)} files.")
        if all_files:
            print("Sample paths:")
            for f in all_files[:10]:
                print(f"  {f}")
        sys.exit(1)

    total_emails = sum(len(v) for v in email_files.values())
    print(f"Found {total_emails} emails in {len(email_files)} folders.")

    # Pre-load attachment data referenced by emails
    attachment_data = {}
    for name in all_files:
        if not name.endswith('.xml'):
            # Could be an attachment file
            attachment_data[name] = None  # lazy load marker

    converted = 0
    failed = 0

    for folder_name, xml_files in sorted(email_files.items()):
        # Clean folder name for filesystem
        safe_name = re.sub(r'[^\w\s\-.]', '_', folder_name).strip('_')
        if not safe_name:
            safe_name = 'Unknown'
        mbox_path = os.path.join(output_dir, f"{safe_name}.mbox")

        print(f"\n  {folder_name} -> {safe_name}.mbox ({len(xml_files)} messages)")

        mbox = mailbox.mbox(mbox_path)
        mbox.lock()

        for xml_file in xml_files:
            try:
                xml_content = zf.read(xml_file).decode('utf-8', errors='replace')

                # Load any attachments referenced in this email's directory
                email_dir = os.path.dirname(xml_file)
                local_attachments = {}
                for name in all_files:
                    if name.startswith(email_dir) and not name.endswith('.xml'):
                        try:
                            local_attachments[name] = zf.read(name)
                        except Exception:
                            pass

                email_msg = xml_to_email(xml_content, local_attachments)
                if email_msg:
                    mbox.add(email_msg)
                    converted += 1
                else:
                    failed += 1
            except Exception as e:
                failed += 1
                print(f"    Warning: Failed to convert {os.path.basename(xml_file)}: {e}")

        mbox.unlock()
        mbox.close()

    zf.close()

    print(f"\nDone! Converted {converted} emails, {failed} failed.")
    print(f"Output directory: {output_dir}")
    print(f"\nTo import into Thunderbird:")
    print(f"  1. Install 'ImportExportTools NG' add-on")
    print(f"  2. Right-click a local folder > ImportExportTools NG > Import mbox file")
    print(f"  3. Select the .mbox files from: {output_dir}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 olm_to_mbox.py <input.olm> [output_directory]")
        print("\nConverts Outlook for Mac .olm archive to .mbox files.")
        sys.exit(1)

    olm_path = os.path.expanduser(sys.argv[1])
    if len(sys.argv) >= 3:
        output_dir = os.path.expanduser(sys.argv[2])
    else:
        base = os.path.splitext(os.path.basename(olm_path))[0]
        output_dir = os.path.join(os.path.dirname(olm_path) or '.', f"{base}_mbox")

    convert_olm_to_mbox(olm_path, output_dir)


if __name__ == '__main__':
    main()
