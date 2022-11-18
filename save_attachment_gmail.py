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


def get_info_list(vCard, vcard_filepath):
    vcard = {}
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
            vcard['cell_phone'] = cell | home
            vcard['company_phone'] = work
        elif key == 'email':
            email = str(vCard.email.value).strip()
            vcard['email'] = email
        else:
            # An unused key, like `adr`, `title`, `url`, etc.
            pass

    return vcard


def array_of_dicts_to_csv(array, filename):
    with open(filename, 'w+') as f:
        w = csv.DictWriter(f, array[0].keys())
        w.writeheader()
        w.writerows(array)


def save_attachments(attachments):
    for attachment in attachments:
        if 'file_handle' in attachment:
            with open(re.sub(r"[/\\?%*:|\"<>\x7F\x00-\x1F]", "-", attachment['file_name']), 'wb+') as f:
                f.write(attachment['file_handle'].read())


def create_and_open_folder(folder_name):
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    os.chdir(folder_name)


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
