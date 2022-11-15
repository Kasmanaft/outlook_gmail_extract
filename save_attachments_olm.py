import sys
import os
import zipfile
from lxml import etree
import pandas as pd
import re


def load_attachment(zip, name):
    fh = zip.open(name)
    return fh


def get_attachments(zip, email):
    attachments = []
    tag_attachments = email.find('.//OPFMessageCopyAttachmentList')
    if tag_attachments is not None:
        for attachment in tag_attachments.findall('.//messageAttachment'):
            name = attachment.get('OPFAttachmentName')
            file = {'file_name': name}
            url = attachment.get('OPFAttachmentURL')
            if url is not None:
                fh = load_attachment(zip, url)
                file['file_handle'] = fh
                attachments.append(file)
    return attachments


def parse_message(zip, name):
    attachments = []
    doc = None
    fh = zip.open(name)
    try:
        doc = etree.parse(fh)
    except etree.XMLSyntaxError:
        p = etree.XMLParser(huge_tree=True)
        try:
            doc = etree.parse(fh, p)
        except etree.XMLSyntaxError:
            # probably corrupt
            pass

    if doc is None:
        return

    for email in doc.findall('//email'):
        attachments += get_attachments(zip, email)

    return attachments


def save_attachments(attachments):
    for attachment in attachments:
        if 'file_handle' in attachment:
            with open(re.sub(r"[/\\?%*:|\"<>\x7F\x00-\x1F]", "-", attachment['file_name']), 'wb+') as f:
                f.write(attachment['file_handle'].read())


def create_and_open_folder(folder_name):
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    os.chdir(folder_name)


def main():
    attachments = []
    fn = sys.argv[1]
    folder_name = sys.argv[2]

    zf = zipfile.ZipFile(fn, 'r')
    for info in zf.namelist():
        if 'com.microsoft.__Attachments' in info:
            continue
        if 'message_' not in info:
            continue

        attachments += parse_message(zf, info)

    attachments_unique = pd.DataFrame(attachments).drop_duplicates(subset=['file_name']).to_dict('records')

    create_and_open_folder(folder_name)
    save_attachments(attachments_unique)


if __name__ == '__main__':
    main()
