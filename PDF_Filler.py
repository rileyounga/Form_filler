import subprocess
import json
import time
import re
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import numpy as np
import pdfrw
from pdfrw import PdfReader
from pdfrw.objects import pdfstring

ADOBE_ACROBAT_PATH = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"

def get_files():
    """
    Function to create a GUI for selecting files using file dialogs.

    Args:
        None

    Returns:
        dict: A dictionary containing file locations for form and client files.
    """

    # Create the main window
    root = tk.Tk()
    root.title("File Upload Form")

    # Add padding (margins) to the main window
    root.configure(padx=20, pady=20)

    # Define variables to store file paths
    form_file_location = tk.StringVar()
    client1_file_location = tk.StringVar()
    client2_file_location = tk.StringVar()
    client3_file_location = tk.StringVar()
    client4_file_location = tk.StringVar()

    # Status variables for indicating file attachment
    form_status = tk.StringVar()
    client1_status = tk.StringVar()
    client2_status = tk.StringVar()
    client3_status = tk.StringVar()
    client4_status = tk.StringVar()

    # Functions to open file dialogs and update status labels
    def select_form_file():
        file_path = filedialog.askopenfilename(initialdir="")
        if file_path:
            form_file_location.set(file_path)
            form_status.set("File Attached")
        else:
            form_status.set("")

    def select_client1_file():
        file_path = filedialog.askopenfilename()
        if file_path:
            client1_file_location.set(file_path)
            client1_status.set("File Attached")
        else:
            client1_status.set("")

    def select_client2_file():
        file_path = filedialog.askopenfilename()
        if file_path:
            client2_file_location.set(file_path)
            client2_status.set("File Attached")
        else:
            client2_status.set("")

    def select_client3_file():
        file_path = filedialog.askopenfilename()
        if file_path:
            client3_file_location.set(file_path)
            client3_status.set("File Attached")
        else:
            client3_status.set("")

    def select_client4_file():
        file_path = filedialog.askopenfilename()
        if file_path:
            client4_file_location.set(file_path)
            client4_status.set("File Attached")
        else:
            client4_status.set("")

    # Function to handle Submit button
    def submit():
        # Close the GUI
        root.quit()

    # Create a style for green text
    style = ttk.Style()
    style.configure("Green.TLabel", foreground="green")

    # Create the GUI elements using ttk for improved styling
    # FORM Label and File Upload button
    form_label = ttk.Label(root, text="FORM:")
    form_label.grid(row=0, column=0, padx=10, pady=5, sticky='e')
    form_button = ttk.Button(root, text="File Upload", command=select_form_file)
    form_button.grid(row=0, column=1, padx=10, pady=5)
    form_status_label = ttk.Label(root, textvariable=form_status, style="Green.TLabel")
    form_status_label.grid(row=0, column=2, padx=10, pady=5, sticky='w')

    # Spacer between FORM and CLIENTS
    spacer1 = ttk.Separator(root, orient='horizontal')
    spacer1.grid(row=1, column=0, columnspan=3, pady=10, sticky='ew')

    # CLIENT 1
    client1_label = ttk.Label(root, text="CLIENT 1:")
    client1_label.grid(row=2, column=0, padx=10, pady=5, sticky='e')
    client1_button = ttk.Button(root, text="File Upload", command=select_client1_file)
    client1_button.grid(row=2, column=1, padx=10, pady=5)
    client1_status_label = ttk.Label(root, textvariable=client1_status, style="Green.TLabel")
    client1_status_label.grid(row=2, column=2, padx=10, pady=5, sticky='w')

    # CLIENT 2
    client2_label = ttk.Label(root, text="CLIENT 2:")
    client2_label.grid(row=3, column=0, padx=10, pady=5, sticky='e')
    client2_button = ttk.Button(root, text="File Upload", command=select_client2_file)
    client2_button.grid(row=3, column=1, padx=10, pady=5)
    client2_status_label = ttk.Label(root, textvariable=client2_status, style="Green.TLabel")
    client2_status_label.grid(row=3, column=2, padx=10, pady=5, sticky='w')

    # CLIENT 3
    client3_label = ttk.Label(root, text="CLIENT 3:")
    client3_label.grid(row=4, column=0, padx=10, pady=5, sticky='e')
    client3_button = ttk.Button(root, text="File Upload", command=select_client3_file)
    client3_button.grid(row=4, column=1, padx=10, pady=5)
    client3_status_label = ttk.Label(root, textvariable=client3_status, style="Green.TLabel")
    client3_status_label.grid(row=4, column=2, padx=10, pady=5, sticky='w')

    # CLIENT 4
    client4_label = ttk.Label(root, text="CLIENT 4:")
    client4_label.grid(row=5, column=0, padx=10, pady=5, sticky='e')
    client4_button = ttk.Button(root, text="File Upload", command=select_client4_file)
    client4_button.grid(row=5, column=1, padx=10, pady=5)
    client4_status_label = ttk.Label(root, textvariable=client4_status, style="Green.TLabel")
    client4_status_label.grid(row=5, column=2, padx=10, pady=5, sticky='w')

    # Spacer between CLIENTS and Submit button
    spacer2 = ttk.Separator(root, orient='horizontal')
    spacer2.grid(row=6, column=0, columnspan=3, pady=10, sticky='ew')

    # Submit Button, right-aligned
    submit_button = ttk.Button(root, text="Submit", command=submit)
    submit_button.grid(row=7, column=2, pady=10, sticky='e')

    root.mainloop()

    # After the GUI is closed, get the values
    form_file = form_file_location.get()
    client1_file = client1_file_location.get()
    client2_file = client2_file_location.get()
    client3_file = client3_file_location.get()
    client4_file = client4_file_location.get()

    # Store file locations
    file_locations = {
        'form_file': form_file,
        'client1_file': client1_file,
        'client2_file': client2_file,
        'client3_file': client3_file,
        'client4_file': client4_file
    }

    # Output the selected files
    print("Selected Files:")
    for key, value in file_locations.items():
        if value:
            print(f"{key}: {value}")
        else:
            print(f"{key}: No file selected")

    # Return the file locations if needed
    return file_locations

def form_filler(file_locations, debug=False):
    """
    Fills a PDF form with data from a number of CSV files.

    Args:
        file_locations (dict): A dictionary containing file locations for form and client files.
        debug (bool, optional): If True, print debug messages. Defaults to False.

    Returns:
        None
    """
    def prepare_data(file_locations, debug):
        """
        Prepares data for filling a PDF form.

        Args:
            file_locations (dict): A dictionary containing file locations for form and client files.
            debug (bool, optional): If True, print debug messages. Defaults to False.

        Returns:
            tuple: A tuple containing the template PDF, general field map, field map for the current form, and 
            a list of check boxes to be filled.
        """
        client1 = None
        client2 = None
        client3 = None
        client4 = None
        template = PdfReader(file_locations['form_file'])
        form = file_locations['form_file'].split('/')[-1].split('.')[0]

        if file_locations['client1_file']:
            client1 = pd.read_csv(file_locations['client1_file']).replace(np.nan, None)
        else:
            if debug:
                print('No client file selected')
            return
        if file_locations['client2_file']:
            client2 = pd.read_csv(file_locations['client2_file']).replace(np.nan, None)
        if file_locations['client3_file']:
            client3 = pd.read_csv(file_locations['client3_file']).replace(np.nan, None)
        if file_locations['client4_file']:
            client4 = pd.read_csv(file_locations['client4_file']).wreplace(np.nan, None)

        try:
            general_map = json.load(open('form_fields.json'))
        except KeyError:
            if not debug:
                print('\x1b[2J\x1b[H')
            print('Critical error: form_fields.json not found.' + '\n' + 'Make sure the file is in the same directory as this script.')
            wait = input('Press any key to exit')
            return
        
        try:
            field_map = general_map[form]
        except KeyError:
            if not debug:
                print('\x1b[2J\x1b[H')
            print(f'Critical error: Form {form} not found in form_fields.json' + '\n' + 'Make sure the file is in the same directory \
                as this script, and that the form name is correct.')
            wait = input('Press any key to exit')
            return

        check_list = field_map['Check_list']
        check_list_eval = []
        # Utilize eval to convert strings to variables
        for x in check_list:
            try:
                x = eval(x)
            except:
                continue
            if x is not None:
                check_list_eval.append(x)

        return template, general_map, field_map, check_list_eval, client1, client2, client3, client4
    
    def checkmark(annotation, kid):
        """
        Checks a checkbox in a PDF by setting the /AS value of the annotation to the name of the /On appearance stream.

        Args:
            annotation (dict): The annotation dictionary for the checkbox annotation.

        Returns:
            None
        """
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

    def text_box(annotation, value):
        """
        Fills in a text box in a PDF by setting the /V value of the annotation to the provided value.

        Args:
            annotation (dict): The annotation dictionary for the text box annotation.
            value (str): The value to fill in the text box.
        
        Returns:
            None
        """
        if value is None or re.search(r'\b(nan|None)\b', str(value)) is not None:
            return
        pdfstr = pdfstring.PdfString.encode(value)
        annotation.update(pdfrw.PdfDict(V=pdfstr))

    template, general_map, field_map, check_list, client1, client2, client3, client4 = prepare_data(file_locations, debug)
    json_rep = ''
    i = 0
    # For each page
    for page in template.pages:
        annotations = page['/Annots']
        if annotations is None:
            continue

        # For each item on the page
        for annotation in annotations:
            i += 1
            # If an annotation has no text, defer to its parent
            if not annotation['/T']:
                kid = annotation
                annotation = annotation['/Parent']

            ft = annotation['/FT']
            ff = annotation['/Ff']
            key = annotation['/T'].to_unicode()

            if ft == '/Tx':
                if debug:
                    print(f'Text field: {key} {i}')

                race_condition = True
                # Search for fields defined specifically for the form
                try:
                    json_rep = field_map[key]
                    race_condition = False
                except Exception:
                    if debug:
                        print(f'{key} not found in form_fields.json')

                # If not found, search for general fields
                if race_condition:
                    try:
                        json_rep = general_map['General'][key]
                    except Exception:
                        if debug:
                            print(f'{key} not found in general fields')
                        continue

                # If found, use eval to convert strings to variables
                try:
                    value = eval(json_rep)
                except Exception:
                    if debug:
                        print(f'Error evaluating {json_rep}')
                    continue

                text_box(annotation, value)

            elif ft == '/Btn':
                if ff and int(ff) & 1 << 15:
                    if debug:
                        print(f'Radio button field: {key} {i}')
                        continue
                else:
                    if debug:
                        print(f'Checkbox field: {key} {i}')

                    if i in check_list:
                        checkmark(annotation, kid)

            elif ft == '/Ch':
                if ff and int(ff) & 1 << 17:
                    if debug:
                        print(f'Combo box field: {key} {i}')
                        continue
                else:
                    if debug:
                        print(f'Dropdown field: {key} {i}')
                        continue

    try:
        subprocess.run(["taskkill", "/f", "/im", "Acrobat.exe"], check=True)
        time.sleep(1)
    except subprocess.CalledProcessError:
        if debug:
            print(f"Error closing Acrobat")

    template.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
    pdfrw.PdfWriter().write('filled.pdf', template)

def main():
    file_locations = get_files()
    form_filler(file_locations, debug=False)

    try:
        subprocess.Popen([ADOBE_ACROBAT_PATH, '/A', 'open', 'filled.pdf'], shell=True)
    except FileNotFoundError:
        print(f'{ADOBE_ACROBAT_PATH} not found')
        exit()

if __name__ == '__main__':
    main()
