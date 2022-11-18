import sys
import zipfile
import csv
import email
from email import policy
from email.parser import BytesParser
import pandas as pd


def array_of_dicts_to_csv(array, filename):
    with open(filename, 'w+') as f:
        w = csv.DictWriter(f, array[0].keys())
        w.writeheader()
        w.writerows(array)


def get_empty_record():
    return dict.fromkeys(['email', 'emails', 'firstname', 'lastname', 'fullname', 'company', 'department',
                          'company_phone', 'cell_phone', 'fax', 'city', 'country', 'state', 'street', 'zip', 'website'])


def contacts_from_string(string):
    contacts = []
    if string is not None:
        for address in string.split(','):
            record = get_empty_record()
            record['fullname'], record['email'] = email.utils.parseaddr(address)
            contacts.append(record)
    return contacts


def extract_emails_from_email(msg):
    emails = []
    if msg['from'] is not None:
        emails += contacts_from_string(msg['from'])
    if msg['to'] is not None:
        emails += contacts_from_string(msg['to'])
    if msg['cc'] is not None:
        emails += contacts_from_string(msg['cc'])
    if msg['bcc'] is not None:
        emails += contacts_from_string(msg['bcc'])
    return emails


def main():
    emails = []
    fn = sys.argv[1]
    output_name = sys.argv[2]

    zf = zipfile.ZipFile(fn, 'r')
    for info in zf.namelist():
        if info.endswith('.eml'):
            with zf.open(info) as fp:
                msg = BytesParser(policy=policy.default).parse(fp)
                emails += extract_emails_from_email(msg)

    if len(emails) == 0:
        print("No emails found.")
    else:
        unique_emails = pd.DataFrame(emails).drop_duplicates(subset=['email']).to_dict('records')
        array_of_dicts_to_csv(unique_emails, output_name)


if __name__ == '__main__':
    main()
