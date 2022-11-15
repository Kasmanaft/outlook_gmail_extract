import sys
import zipfile
from lxml import etree
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def get_addresses(email):
    tag_from = email.find('.//OPFMessageCopyFromAddresses')
    tag_sender = email.find('.//OPFMessageCopySenderAddress')
    tag_to = email.find('.//OPFMessageCopyToAddresses')
    tag_cc = email.find('.//OPFMessageCopyCCAddresses')
    tag_bcc = email.find('.//OPFMessageCopyBCCAddresses')

    emails = get_contacts(tag_from)
    emails.update(get_contacts(tag_sender))
    emails.update(get_contacts(tag_to))
    emails.update(get_contacts(tag_cc))
    emails.update(get_contacts(tag_bcc))

    return emails


def get_contacts(addresses):
    emails = {}
    if addresses is not None:
        for address in addresses.findall('.//emailAddress'):
            email = address.get('OPFContactEmailAddressAddress')
            name = address.get('OPFContactEmailAddressName')
            if name is not None and name != email:
                emails[email] = name

    return emails


def get_contact(contact):
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

    print(etree.tostring(doc, pretty_print=False))


def parse_addressbook(zip, name):
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

    for contact in doc.findall('//contact'):
        return get_contact(contact)


def main():
    emails = {}
    fn = sys.argv[1]

    zf = zipfile.ZipFile(fn, 'r')
    for info in zf.namelist():
        if info.endswith('.xml') and 'Contacts' in info:
            parse_addressbook(zf, info)


if __name__ == '__main__':
    main()
