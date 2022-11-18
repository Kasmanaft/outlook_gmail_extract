import sys
import zipfile
import csv
from lxml import etree
import pandas as pd


def get_addresses(email):
    tag_from = email.find('.//OPFMessageCopyFromAddresses')
    tag_sender = email.find('.//OPFMessageCopySenderAddress')
    tag_to = email.find('.//OPFMessageCopyToAddresses')
    tag_cc = email.find('.//OPFMessageCopyCCAddresses')
    tag_bcc = email.find('.//OPFMessageCopyBCCAddresses')

    emails = get_contacts(tag_from)
    emails += get_contacts(tag_sender)
    emails += get_contacts(tag_to)
    emails += get_contacts(tag_cc)
    emails += get_contacts(tag_bcc)

    return emails


def get_empty_record():
    return dict.fromkeys(['email', 'emails', 'firstname', 'lastname', 'fullname', 'company', 'department',
                          'company_phone', 'cell_phone', 'fax', 'city', 'country', 'state', 'street', 'zip', 'website'])


def get_contacts(addresses):
    emails = []
    if addresses is not None:
        for address in addresses.findall('.//emailAddress'):
            record = get_empty_record()
            record['email'] = address.get('OPFContactEmailAddressAddress')
            record['fullname'] = address.get('OPFContactEmailAddressName')
            if record['fullname'] is not None and record['fullname'] != record['email']:
                emails.append(record)

    return emails


def parse_message(zip, name):
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
        return get_addresses(email)


def array_of_dicts_to_csv(array, filename):
    with open(filename, 'w+') as f:
        w = csv.DictWriter(f, array[0].keys())
        w.writeheader()
        w.writerows(array)


def main():
    emails = []
    fn = sys.argv[1]
    output_name = sys.argv[2]

    zf = zipfile.ZipFile(fn, 'r')
    for info in zf.namelist():
        if 'com.microsoft.__Attachments' in info:
            continue
        if 'message_' not in info:
            continue

        emails += parse_message(zf, info)

    unique_emails = pd.DataFrame(emails).drop_duplicates(subset=['email']).to_dict('records')
    array_of_dicts_to_csv(unique_emails, output_name)


if __name__ == '__main__':
    main()
