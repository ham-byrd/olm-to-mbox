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
from email.utils import formatdate, formataddr
from email import encoders
from datetime import datetime
from collections import defaultdict
import re
import time


def progress_bar(current, total, folder_name=''):
    """Print a short progress bar that overwrites in place."""
    if current % 50 != 0 and current != total:
        return  # only update every 50 emails to reduce flicker
    pct = current / total if total else 0
    filled = int(20 * pct)
    bar = '#' * filled + '-' * (20 - filled)
    elapsed = time.time() - progress_bar.start_time
    rate = current / elapsed if elapsed > 0 else 0
    eta = int((total - current) / rate) if rate > 0 else 0
    eta_m, eta_s = divmod(eta, 60)
    sys.stdout.write(f'\r  [{bar}] {pct:.1%} ~{eta_m}m{eta_s:02d}s')
    sys.stdout.flush()

progress_bar.start_time = time.time()


def parse_olm_address(addr_elem):
    """Parse an OLM address element into (name, email) tuple."""
    if addr_elem is None:
        return None
    address_list = addr_elem.findall('.//emailAddress')
    if address_list:
        results = []
        for a in address_list:
            # OLM stores addresses as XML attributes, not child text
            name = a.get('OPFContactEmailAddressName', '').strip()
            email = a.get('OPFContactEmailAddressAddress', '').strip()
            if not email:
                # Fallback: try as child elements
                name = a.findtext('OPFContactEmailAddressName', '').strip()
                email = a.findtext('OPFContactEmailAddressAddress', '').strip()
            if email:
                results.append((name, email))
        return results
    text = addr_elem.text
    if text and '@' in text:
        return [('', text.strip())]
    return []


def parse_olm_date(date_str):
    """Parse OLM date string to datetime."""
    if not date_str:
        return None
    date_str = date_str.strip()
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

    def find_text(tag):
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

    from_addrs = parse_olm_address(root.find('.//OPFMessageCopyFromAddresses'))
    to_addrs = parse_olm_address(root.find('.//OPFMessageCopyToAddresses'))
    cc_addrs = parse_olm_address(root.find('.//OPFMessageCopyCCAddresses'))
    bcc_addrs = parse_olm_address(root.find('.//OPFMessageCopyBCCAddresses'))

    if not from_addrs:
        sender = root.find('.//OPFMessageCopySenderAddress')
        if sender is not None:
            from_addrs = parse_olm_address(sender)

    date = parse_olm_date(sent_time) or parse_olm_date(recv_time)

    attachment_elems = root.findall('.//messageAttachment')
    has_attachments = bool(attachment_elems and attachments_data)

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

    # Clean up any stale lock files from previous interrupted runs
    import glob as globmod
    for lockfile in globmod.glob(os.path.join(output_dir, '*.lock*')):
        os.remove(lockfile)

    print(f"Opening {olm_path}...")
    try:
        zf = zipfile.ZipFile(olm_path, 'r')
    except zipfile.BadZipFile:
        print("Error: Not a valid .olm (ZIP) file.")
        sys.exit(1)

    # Categorize files by folder
    email_files = {}   # folder_path -> list of xml filenames
    all_files = zf.namelist()

    # Build a directory index for fast attachment lookup (O(1) instead of O(n))
    print("Indexing archive...")
    dir_index = defaultdict(list)  # directory -> list of non-xml files in it
    for name in all_files:
        dirname = os.path.dirname(name)
        if not name.endswith('.xml'):
            dir_index[dirname].append(name)

    for name in all_files:
        if name.endswith('.xml') and '/Data/' in name:
            lower = name.lower()
            if any(skip in lower for skip in ['/contacts/', '/calendar/', '/tasks/', '/notes/']):
                continue
            parts = name.split('/Data/')
            if len(parts) >= 2:
                folder_path = os.path.dirname(parts[1])
                if not folder_path:
                    folder_path = 'Root'
                if folder_path not in email_files:
                    email_files[folder_path] = []
                email_files[folder_path].append(name)

    if not email_files:
        for name in all_files:
            if name.endswith('.xml'):
                folder_path = os.path.dirname(name) or 'Root'
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
    print(f"Found {total_emails:,} emails in {len(email_files)} folders.\n")

    converted = 0
    failed = 0
    start_time = time.time()
    progress_bar.start_time = start_time

    for folder_name, xml_files in sorted(email_files.items()):
        safe_name = re.sub(r'[^\w\s\-.]', '_', folder_name).strip('_')
        if not safe_name:
            safe_name = 'Unknown'
        mbox_path = os.path.join(output_dir, f"{safe_name}.mbox")

        # Clear progress line, print folder header, then progress resumes below
        sys.stdout.write(f'\r{" " * 90}\r')
        print(f"  {folder_name} ({len(xml_files):,} msgs) -> {safe_name}.mbox")

        mbox = mailbox.mbox(mbox_path)

        for xml_file in xml_files:
            try:
                xml_content = zf.read(xml_file).decode('utf-8', errors='replace')

                # Fast attachment lookup using pre-built index
                email_dir = os.path.dirname(xml_file)
                local_attachments = {}
                for att_file in dir_index.get(email_dir, []):
                    try:
                        local_attachments[att_file] = zf.read(att_file)
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

            progress_bar(converted + failed, total_emails, safe_name)

        mbox.close()

    elapsed = time.time() - start_time
    mins, secs = divmod(int(elapsed), 60)

    # Clear progress line
    sys.stdout.write(f'\r{" " * 90}\r')
    sys.stdout.flush()

    zf.close()

    print(f"\nDone! Converted {converted:,} emails, {failed:,} failed in {mins}m {secs}s.")
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
