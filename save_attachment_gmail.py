import os
import sys
import zipfile
import re
import email
from email import policy
from email.parser import BytesParser
import pandas as pd


def save_attachments(attachments):
    for attachment in attachments:
        if 'content' in attachment:
            with open(re.sub(r"[/\\?%*:|\"<>\x7F\x00-\x1F]", "-", attachment['file_name']), 'wb+') as f:
                f.write(attachment['content'])


def create_and_open_folder(folder_name):
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    os.chdir(folder_name)


def extract_attachments_from_email(msg):
    attachments = []
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        file_name = part.get_filename()
        if bool(file_name):
            file = {'file_name': file_name}
            file['content'] = part.get_payload(decode=True)
            attachments.append(file)
    return attachments


def main():
    attachments = []
    fn = sys.argv[1]
    folder_name = sys.argv[2]

    zf = zipfile.ZipFile(fn, 'r')
    for info in zf.namelist():
        if info.endswith('.eml'):
            with zf.open(info) as fp:
                msg = BytesParser(policy=policy.default).parse(fp)
                attachments += extract_attachments_from_email(msg)

    if len(attachments) == 0:
        print("No attachments found.")
    else:
        attachments_unique = pd.DataFrame(attachments).drop_duplicates(subset=['file_name']).to_dict('records')

        create_and_open_folder(folder_name)
        save_attachments(attachments_unique)


if __name__ == '__main__':
    main()
