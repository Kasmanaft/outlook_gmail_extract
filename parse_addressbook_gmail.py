import sys
import zipfile
import csv
import vobject
import pandas as pd


def get_phone_numbers(vCard):
    cell = home = work = None
    for tel in vCard.tel_list:
        if vCard.version.value == '2.1':
            if 'CELL' in tel.singletonparams:
                cell = str(tel.value).strip()
            elif 'WORK' in tel.singletonparams:
                work = str(tel.value).strip()
            elif 'HOME' in tel.singletonparams:
                home = str(tel.value).strip()
            else:
                logging.warning("Warning: Unrecognized phone number category in `{}'".format(vCard))
                tel.prettyPrint()
        elif vCard.version.value == '3.0':
            if 'CELL' in tel.params['TYPE']:
                cell = str(tel.value).strip()
            elif 'WORK' in tel.params['TYPE']:
                work = str(tel.value).strip()
            elif 'HOME' in tel.params['TYPE']:
                home = str(tel.value).strip()
            else:
                logging.warning("Unrecognized phone number category in `{}'".format(vCard))
                tel.prettyPrint()
        else:
            raise NotImplementedError("Version not implemented: {}".format(vCard.version.value))
    return cell, home, work


def get_info_list(vCard, vcard_filepath):
    vcard = collections.OrderedDict()
    for column in column_order:
        vcard[column] = None
    name = cell = work = home = email = note = None
    vCard.validate()
    for key, val in list(vCard.contents.items()):
        if key == 'fn':
            vcard['fullname'] = vCard.fn.value
        elif key == 'n':
            name = str(vCard.n.valueRepr()).replace('  ', ' ').strip()
            vcard['name'] = name
        elif key == 'tel':
            cell, home, work = get_phone_numbers(vCard)
            vcard['cell'] = cell
            vcard['Home phone'] = home
            vcard['Work phone'] = work
        elif key == 'email':
            email = str(vCard.email.value).strip()
            vcard['email'] = email
        else:
            # An unused key, like `adr`, `title`, `url`, etc.
            pass
    if name is None:
        logging.warning("no name for vCard in file `{}'".format(vcard_filepath))
    if all(telephone_number is None for telephone_number in [cell, work, home]):
        logging.warning("no telephone numbers for file `{}' with name `{}'".format(vcard_filepath, name))

    return vcard


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


def array_of_dicts_to_csv(array, filename):
    with open(filename, 'w+') as f:
        w = csv.DictWriter(f, array[0].keys())
        w.writeheader()
        w.writerows(array)


def get_vcards(zip, vcard_filepath):
    with zip.open(vcard_filepath) as fp:
        all_text = fp.read()
    for vCard in vobject.readComponents(all_text):
        yield vCard


def main():
    contacts = []
    fn = sys.argv[1]
    output_name = sys.argv[2]

    zf = zipfile.ZipFile(fn, 'r')
    for info in zf.namelist():
        if info.endswith('.vcf'):
            for vcard in get_vcards(zf, info):
                vcard_info = get_info_list(vcard, vcard_path)

    unique_contacts = pd.DataFrame(contacts).drop_duplicates(subset=['email']).to_dict('records')
    array_of_dicts_to_csv(unique_contacts, output_name)


if __name__ == '__main__':
    main()
