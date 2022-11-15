import sys
import zipfile
import csv
from lxml import etree


def get_emails(doc):
    emails = []
    for address in doc.findall('.//contactEmailAddress'):
        emails.append(address.get('OPFContactEmailAddressAddress').lower())

    return list(set(emails))


def get_contact(doc):
    emails = get_emails(doc)
    if len(emails) > 0:
        contact = {
            'email': emails.pop(),
            'emails': ','.join(emails),
            'firstname': doc.get('OPFContactCopyFirstName'),
            'fullname': doc.get('OPFContactCopyDisplayName'),
            'company': doc.get('OPFContactCopyBusinessCompany'),
            'department': doc.get('OPFContactCopyBusinessDepartment'),
            'company_phone': doc.get('OPFContactCopyBusinessPhone'),
            'cell_phone': doc.get('OPFContactCopyCellPhone'),
            'fax': doc.get('OPFContactCopyHomeFax'),
            'city': doc.get('OPFContactCopyHomeCity'),
            'country': doc.get('OPFContactCopyHomeCountry'),
            'state': doc.get('OPFContactCopyHomeState'),
            'street': doc.get('OPFContactCopyHomeStreetAddress'),
            'zip': doc.get('OPFContactCopyHomeZip'),
            'website': doc.get('OPFContactCopyHomeWebPage')
        }
        return contact
    else:
        return


def array_of_dicts_to_csv(array, filename):
    with open(filename, 'w+') as f:
        w = csv.DictWriter(f, array[0].keys())
        w.writeheader()
        w.writerows(array)


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


def main():
    fn = sys.argv[1]

    zf = zipfile.ZipFile(fn, 'r')
    for info in zf.namelist():
        if info.endswith('.xml') and 'Contacts' in info:
            parse_addressbook(zf, info)


if __name__ == '__main__':
    main()
