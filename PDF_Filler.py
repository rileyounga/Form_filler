import subprocess
import json
import pandas as pd
import pdfrw
from pdfrw import PdfReader
from pdfrw.objects import pdfstring

# change these to match your file locations
CLIENT_FILE = r"C:\Users\riley\OneDrive\Documents\Projects\Automation\Carson\Form Population for AY.csv"
FORM_FILE = r"C:\Users\riley\OneDrive\Documents\Projects\Automation\Carson\Austin Young Shared Folder\Concorde Forms\QCAF.pdf"
ADOBE_ACROBAT_PATH = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"

#FORM_FILE = input("")

def get_data(client_file, form_file):
    """
    Retrieves client data and a PDF template based on the provided form.

    Args:
        form: The PDF form to retrieve data for.

    Returns:
        A tuple containing the PDF template, client data, and representative data.
    """
    clients = pd.read_csv(client_file)

    representative = pd.DataFrame({
        'Number': ['4619'],
        'Name': ['Austin Young'],
        'Client': ['Carson Stoltz'],
    })

    name = representative['Client'].values[0]

    try:
        client = clients[clients['Full Name'] == name]
    except IndexError:
        print(f'No client found with name {name}')
        exit()

    try:
        template = PdfReader(form_file)
    except pdfrw.PdfReaderError:
        print(f'{form_file} is not a valid PDF')
        exit()

    form = form_file.split('\\')[-1].split('.')[0]
    return template, client, form

def form_filler(template, client, form, debug=False):
    """
    Fills a PDF template with data from a client.

    Args:
        template (pdfrw.PdfReader): The PDF template to fill.
        client (pd.DataFrame): The client data to fill the template with.
        debug (bool, optional): Whether to print debug information. Defaults to True.

    Returns:
        None
    """               
    i = 0
    try:
        field_map = json.load(open('form_fields.json'))[form]
    except KeyError:
        print(f'Form {form} not found in form_fields.json')
        exit()

    for page in template.pages:
        annotations = page['/Annots']
        if annotations is None:
            continue

        for annotation in annotations:
            i += 1
            if not annotation['/T']:
                kid = annotation
                annotation = annotation['/Parent']

            ft = annotation['/FT']
            ff = annotation['/Ff']
            key = annotation['/T'].to_unicode()

            if ft == '/Tx':
                if debug:
                    print(f'Text field: {key} {i}')
                
                def text_box(annotation, value):
                    pdfstr = pdfstring.PdfString.encode(value)
                    annotation.update(pdfrw.PdfDict(V=pdfstr))

                try:
                    rep_col = field_map[key]['key']
                    transform = field_map[key]['lambda']
                    value = client[rep_col].values[0]
                    if debug:
                        print(f'{value} {transform}')
                        
                    if transform != 'None':
                        value = eval(transform)(value)
                    
                except KeyError:
                    if debug:
                        print(f'KeyError: {key}')
                    continue
                
                text_box(annotation, value)

                
            elif ft == '/Btn':
                if ff and int(ff) & 1 << 15:
                    if debug:
                        print(f'Radio button field: {key} {i}')

                    #TODO: radio button
                else:
                    if debug:
                            print(f'Checkbox field: {key} {i}')
                    
                    def checkmark(annotation):
                        try:
                            for child in annotation['/Kids']:
                                keys = child['/AP']['/N'].keys()
                                try:
                                    keys.remove('/Off')
                                except:
                                    pass
                                export = keys[0]

                                val_str = pdfrw.objects.pdfname.BasePdfName(export)
                                if child == kid:
                                    child.update(pdfrw.PdfDict(AS=val_str))
                                    break

                            annotation.update(pdfrw.PdfDict(V=val_str))

                        except:
                            pass
                    try:
                        check_list = field_map['Check_list']
                    except KeyError:
                        if debug:
                            print(f'KeyError: {key}')
                        continue

                    if i in check_list:
                        checkmark(annotation)
                    
            elif ft == '/Ch':
                #TODO: combo box

                if ff and int(ff) & 1 << 17:
                    if debug:
                        print('combo')
                else:
                    if debug:
                        print('list')

    template.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
    pdfrw.PdfWriter().write('filled.pdf', template)


if __name__ == '__main__':
    template, client, form = get_data(CLIENT_FILE, FORM_FILE)
    form_filler(template, client, form, debug=False)
    try:
        subprocess.Popen([ADOBE_ACROBAT_PATH, '/A', 'open', 'filled.pdf'], shell=True)
    except FileNotFoundError:
        print(f'{ADOBE_ACROBAT_PATH} not found')
        exit()
