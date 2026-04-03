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
        return
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
    """Parse an OLM address element into list of (name, email) tuples."""
    if addr_elem is None:
        return None
    address_list = addr_elem.findall('.//emailAddress')
    if address_list:
        results = []
        for a in address_list:
            # OLM stores addresses as XML attributes
            name = a.get('OPFContactEmailAddressName', '').strip()
            email = a.get('OPFContactEmailAddressAddress', '').strip()
            if not email:
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


def simplify_folder_name(raw_path):
    """Extract a clean folder name from OLM internal paths.

    OLM paths look like:
      Local/com.microsoft.__Messages/Outlook.../hbyrd@.../Sent Items/message.xml
      Accounts/hbyrd@noteslive.vip/com.microsoft.__Messages/Archive/message.xml

    We want just: 'Sent Items', 'Archive', 'Inbox', etc.
    """
    parts = raw_path.replace('\\', '/').split('/')

    # Walk backwards to find the meaningful folder name
    # Skip: message filenames, com.microsoft.* internals, account names, 'Local', 'Accounts', 'Data'
    skip = {'local', 'accounts', 'data', 'com.microsoft.__messages', 'com.microsoft.__attachments'}
    for p in reversed(parts):
        lower = p.lower()
        if lower in skip:
            continue
        if p.startswith('message_') and p.endswith('.xml'):
            continue
        if p.endswith('.xml'):
            continue
        if '@' in p:  # email address like hbyrd@noteslive.vip
            continue
        if p.startswith('Outlook for Mac Archive'):
            continue
        if p == '':
            continue
        return p

    return 'Other'


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
    # Fallback: use Preview if both body fields are empty
    if not html_body and not text_body:
        text_body = find_text('OPFMessageCopyPreview')
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

    # Fallback: use DisplayTo if no structured To addresses
    if not to_addrs:
        display_to = find_text('OPFMessageCopyDisplayTo')
        if display_to:
            to_addrs = [('', display_to)]

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

    # Add attachments — OLM stores metadata as XML attributes
    if has_attachments and attachments_data:
        for att_elem in attachment_elems:
            att_name = att_elem.get('OPFAttachmentName', '') or att_elem.findtext('OPFAttachmentName', '')
            att_url = att_elem.get('OPFAttachmentURL', '') or att_elem.findtext('OPFAttachmentURL', '')
            att_mime = att_elem.get('OPFAttachmentContentType', '') or att_elem.findtext('OPFAttachmentContentType', 'application/octet-stream')
            if not att_mime:
                att_mime = 'application/octet-stream'

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

    all_files = zf.namelist()

    # Build a set of all files for fast attachment URL lookup
    print("Indexing archive...")
    all_files_set = set(all_files)

    # Find all email XML files and group by simplified folder name
    email_files = defaultdict(list)
    seen_message_ids = defaultdict(set)  # for deduplication

    for name in all_files:
        if not name.endswith('.xml'):
            continue
        # Skip non-email XMLs
        lower = name.lower()
        if any(skip in lower for skip in ['/contacts/', '/calendar/', '/tasks/', '/notes/',
                                           'categories.xml', '/address book/']):
            continue
        if 'message' not in lower and 'messages' not in os.path.dirname(lower):
            continue

        folder = simplify_folder_name(name)
        email_files[folder].append(name)

    if not email_files:
        print("No email XML files found in the archive.")
        sys.exit(1)

    # Deduplicate: OLM often stores two copies of each email
    # (one under Local/, one under Accounts/). Prefer Accounts/ copies
    # as they have more complete body content.
    print("Deduplicating emails...")
    deduped = {}
    total_before = 0
    for folder, xml_list in email_files.items():
        total_before += len(xml_list)
        # Sort so Accounts/ paths come first (preferred), Local/ second
        xml_list.sort(key=lambda x: (0 if x.startswith('Accounts/') else 1, x))
        unique = []
        seen = set()
        for xml_file in xml_list:
            try:
                content = zf.read(xml_file).decode('utf-8', errors='replace')
                m = re.search(r'OPFMessageCopyMessageID[^>]*>([^<]+)<', content)
                if m:
                    mid = m.group(1).strip()
                    if mid in seen:
                        continue
                    seen.add(mid)
                unique.append(xml_file)
            except Exception:
                unique.append(xml_file)
        deduped[folder] = unique

    email_files = deduped
    total_emails = sum(len(v) for v in email_files.values())
    print(f"Found {total_emails:,} unique emails in {len(email_files)} folders (was {total_before:,} before dedup).\n")

    converted = 0
    failed = 0
    start_time = time.time()
    progress_bar.start_time = start_time

    for folder_name, xml_files in sorted(email_files.items()):
        safe_name = re.sub(r'[^\w\s\-.]', '_', folder_name).strip('_')
        if not safe_name:
            safe_name = 'Unknown'
        mbox_path = os.path.join(output_dir, f"{safe_name}.mbox")

        sys.stdout.write(f'\r{" " * 90}\r')
        print(f"  {folder_name} ({len(xml_files):,} msgs) -> {safe_name}.mbox")

        mbox = mailbox.mbox(mbox_path)

        for xml_file in xml_files:
            try:
                xml_content = zf.read(xml_file).decode('utf-8', errors='replace')

                # Load attachments referenced by URL in this email
                referenced = {}
                root = ET.fromstring(xml_content)
                for a in root.findall('.//messageAttachment'):
                    url = a.get('OPFAttachmentURL', '') or a.findtext('OPFAttachmentURL', '')
                    if url and url in all_files_set:
                        try:
                            referenced[url] = zf.read(url)
                        except Exception:
                            pass

                email_msg = xml_to_email(xml_content, referenced if referenced else None)
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
