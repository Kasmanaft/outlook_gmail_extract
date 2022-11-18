import sys
import csv
import os
import pandas as pd


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


def main():
    raw_contacts = []
    contacts = []
    used_emails = []
    dir = sys.argv[1]
    output_name = sys.argv[2]

    for csv_file in files_in_dir(dir):
        raw_contacts += csv_to_array_of_dicts(csv_file)

    for contact in raw_contacts:
        if contact['email'] not in used_emails:
            used_emails.append(contact['email'])
            if contact['emails'] is not None:
                used_emails += contact['emails'].split(',')
            contacts.append(contact)

    if len(contacts) == 0:
        print("No contatcs found.")
    else:
        array_of_dicts_to_excel(contacts, output_name)


if __name__ == '__main__':
    main()
