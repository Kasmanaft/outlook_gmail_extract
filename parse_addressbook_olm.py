import sys
import zipfile
import csv
from lxml import etree
import pandas as pd


def get_emails(doc):
    emails = []
    for address in doc.findall('.//contactEmailAddress'):
        emails.append(address.get('OPFContactEmailAddressAddress').lower())

    return list(set(emails))


def get_safe(doc, attr):
    el = doc.find(".//"+attr)
    if el is not None:
        return el.text.strip()


def get_contact(doc):
    emails = get_emails(doc)
    if len(emails) > 0:
        contact = {
            'email': emails.pop(),
            'emails': ','.join(emails),
            'firstname': get_safe(doc, 'OPFContactCopyFirstName'),
            'lastname': get_safe(doc, 'OPFContactCopyLastName'),
            'fullname': get_safe(doc, 'OPFContactCopyDisplayName'),
            'company': get_safe(doc, 'OPFContactCopyBusinessCompany'),
            'department': get_safe(doc, 'OPFContactCopyBusinessDepartment'),
            'company_phone': get_safe(doc, 'OPFContactCopyBusinessPhone'),
            'cell_phone': get_safe(doc, 'OPFContactCopyCellPhone'),
            'fax': get_safe(doc, 'OPFContactCopyHomeFax'),
            'city': get_safe(doc, 'OPFContactCopyHomeCity'),
            'country': get_safe(doc, 'OPFContactCopyHomeCountry'),
            'state': get_safe(doc, 'OPFContactCopyHomeState'),
            'street': get_safe(doc, 'OPFContactCopyHomeStreetAddress'),
            'zip': get_safe(doc, 'OPFContactCopyHomeZip'),
            'website': get_safe(doc, 'OPFContactCopyHomeWebPage')
        }
        return contact
    else:
        return


def parse_addressbook(zip, name):
    doc = None
    contacts = []
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

    for el in doc.findall('//contact'):
        contact = get_contact(el)
        if contact is not None:
            contacts.append(contact)

    return contacts


def array_of_dicts_to_csv(array, filename):
    with open(filename, 'w+') as f:
        w = csv.DictWriter(f, array[0].keys())
        w.writeheader()
        w.writerows(array)


def main():
    contacts = []
    fn = sys.argv[1]
    output_name = sys.argv[2]

    zf = zipfile.ZipFile(fn, 'r')
    for info in zf.namelist():
        if info.endswith('.xml') and 'Contacts' in info:
            contacts += parse_addressbook(zf, info)

    unique_contacts = pd.DataFrame(contacts).drop_duplicates(subset=['email']).to_dict('records')
    array_of_dicts_to_csv(unique_contacts, output_name)


if __name__ == '__main__':
    main()
