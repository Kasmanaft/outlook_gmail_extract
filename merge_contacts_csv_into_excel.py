import sys
import csv
import os
import pandas as pd
import re

pat = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
used_emails = {}


def files_in_dir(dir):
    for root, dirs, files in os.walk(dir):
        for file in files:
            if file.endswith(".csv"):
                yield os.path.join(root, file)


def csv_to_array_of_dicts(csv_file):
    with open(csv_file, 'r') as f:
        reader = csv.DictReader(f)
        return [row for row in reader]


def array_of_dicts_to_excel(contacts, output_name):
    df = pd.DataFrame(contacts)
    df.to_excel(output_name, index=False)


def set_uniq_email(contact, key_email, uniq):
    if contact['emails'] is not None:
        for email in contact['emails'].split(','):
            second_email = email.strip().lower()
            if second_email != '' and re.match(pat, second_email):
                print(second_email)
                if second_email not in used_emails:
                    used_emails[second_email] = key_email
                else:
                    uniq = used_emails[second_email]

    return uniq


def contact_merge(original, contact):
    emails = original['emails'].split(',')
    # emails.append(contact['email'])
    emails += contact['emails'].split(',')
    emails = list(filter(None, emails))
    original['emails'] = ','.join(set(emails))

    if original['firstname'].strip() == '':
        original['firstname'] = contact['firstname']
    if original['lastname'].strip() == '':
        original['lastname'] = contact['lastname']
    if original['fullname'].strip() == '':
        original['fullname'] = contact['fullname']
    if original['company'].strip() == '':
        original['company'] = contact['company']
    if original['department'].strip() == '':
        original['department'] = contact['department']
    if original['company_phone'].strip() == '':
        original['company_phone'] = contact['company_phone']
    if original['cell_phone'].strip() == '':
        original['cell_phone'] = contact['cell_phone']
    if original['fax'].strip() == '':
        original['fax'] = contact['fax']
    if original['city'].strip() == '':
        original['city'] = contact['city']
    if original['country'].strip() == '':
        original['country'] = contact['country']
    if original['state'].strip() == '':
        original['state'] = contact['state']
    if original['street'].strip() == '':
        original['street'] = contact['street']
    if original['zip'].strip() == '':
        original['zip'] = contact['zip']
    if original['website'].strip() == '':
        original['website'] = contact['website']

    return original


def main():
    raw_contacts = []
    contacts = {}

    dir = sys.argv[1]
    output_name = sys.argv[2]

    for csv_file in files_in_dir(dir):
        raw_contacts += csv_to_array_of_dicts(csv_file)

    for contact in raw_contacts:
        uniq = False
        key_email = contact['email'].strip().lower()
        contact['email'] = key_email
        if key_email != '' and re.match(pat, key_email):
            if key_email not in used_emails:
                used_emails[key_email] = key_email
                uniq = set_uniq_email(contact, key_email, uniq)
            else:
                set_uniq_email(contact, key_email, uniq)
                uniq = used_emails[key_email]

            if uniq is False:
                contacts[key_email] = contact
            else:
                contacts[uniq] = contact_merge(contacts[uniq], contact)

    if len(contacts) == 0:
        print("No contatcs found.")
    else:
        array_of_dicts_to_excel(contacts.values(), output_name)


if __name__ == '__main__':
    main()
