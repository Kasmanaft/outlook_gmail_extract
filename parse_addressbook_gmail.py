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
        elif vCard.version.value == '3.0':
            if 'CELL' in tel.params['TYPE']:
                cell = str(tel.value).strip()
            elif 'WORK' in tel.params['TYPE']:
                work = str(tel.value).strip()
            elif 'HOME' in tel.params['TYPE']:
                home = str(tel.value).strip()
        else:
            raise NotImplementedError("Version not implemented: {}".format(vCard.version.value))
    return cell, home, work


def get_empty_record():
    return dict.fromkeys(['email', 'emails', 'firstname', 'lastname', 'fullname', 'company', 'department',
                          'company_phone', 'cell_phone', 'fax', 'city', 'country', 'state', 'street', 'zip', 'website'])


def get_info_list(vCard, vcard_filepath):
    vcard = get_empty_record()
    name = cell = work = home = email = note = None
    vCard.validate()
    for key, val in list(vCard.contents.items()):
        if key == 'fn':
            vcard['fullname'] = vCard.fn.value
        elif key == 'n':
            name = str(vCard.n.valueRepr()).replace('  ', ' ').strip()
            vcard['firstname'] = name
        elif key == 'tel':
            cell, home, work = get_phone_numbers(vCard)
            vcard['cell_phone'] = cell | home
            vcard['company_phone'] = work
        elif key == 'email':
            email = str(vCard.email.value).strip()
            vcard['email'] = email
        else:
            # An unused key, like `adr`, `title`, `url`, etc.
            pass

    if vcard['fullname'] is None and name is not None:
        vcard['fullname'] = name
        vcard['firstname'] = name.split(' ')[0]
        vcard['lastname'] = name.split(' ')[-1]

    return vcard


def array_of_dicts_to_csv(array, filename):
    with open(filename, 'w+') as f:
        w = csv.DictWriter(f, array[0].keys())
        w.writeheader()
        w.writerows(array)


def get_vcards(zip, vcard_filepath):
    with zip.open(vcard_filepath) as fp:
        all_text = bytes_to_str(fp.read())
    for vCard in vobject.readComponents(all_text):
        yield vCard


def bytes_to_str(b):
    if isinstance(b, bytes):
        return b.decode('utf-8')
    return b


def main():
    contacts = []
    fn = sys.argv[1]
    output_name = sys.argv[2]

    zf = zipfile.ZipFile(fn, 'r')
    for info in zf.namelist():
        if info.endswith('.vcf'):
            for vcard in get_vcards(zf, info):
                contacts.append(get_info_list(vcard, info))

    if len(contacts) == 0:
        print("No contacts found.")
    else:
        unique_contacts = pd.DataFrame(contacts).drop_duplicates(subset=['email']).to_dict('records')
        array_of_dicts_to_csv(unique_contacts, output_name)


if __name__ == '__main__':
    main()
