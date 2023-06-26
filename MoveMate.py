import os
import shutil

import pandas as pd
import PySimpleGUI as sg
from PrivateFunctions import folders
import sys
import logging


try:
    os.remove("C:/QCCenter/Logs/MoveMate.log")
    logging.basicConfig(filename='C:/QCCenter/Logs/MoveMate.log')
except FileNotFoundError:
    logging.basicConfig(filename='C:/QCCenter/Logs/MoveMate.log')


def exception_hook(exc_type, exc_value, exc_traceback):
    logging.error(
        "Uncaught exception",
        exc_info=(exc_type, exc_value, exc_traceback)
    )


sys.excepthook = exception_hook

def _onKeyRelease(event):
    ctrl = (event.state & 0x4) != 0
    if event.keycode == 88 and ctrl and event.keysym.lower() != "x":
        event.widget.event_generate("<<Cut>>")

    if event.keycode == 86 and ctrl and event.keysym.lower() != "v":
        event.widget.event_generate("<<Paste>>")

    if event.keycode == 67 and ctrl and event.keysym.lower() != "c":
        event.widget.event_generate("<<Copy>>")

    if event.keycode == 65 and ctrl and event.keysym.lower() != "a":
        event.widget.event_generate("<<SelectAll>>")


def LoadExcel(excel):
    # df = pd.read_excel(excel)
    # headers = df.columns[:20].tolist()
    # print(headers)
    # sheets = df.sheet_names()
    # print(sheets)
    read_excel = pd.ExcelFile(excel)
    # headers = df.columns[:20].tolist()
    # print(headers)
    sheets = read_excel.sheet_names
    print(sheets)

    if len(sheets) > 1:
        sheet_layout = [
            [sg.Listbox(sheets, expand_y=True, size=(50, 25), key="-SHEETS-")],
            [sg.Button("Choose", key="-CHOOSE-")]
        ]

        sheet_chooser = sg.Window("Choose Sheet", sheet_layout)

        # Event loop to handle GUI events
        while True:
            event, values = sheet_chooser.read()
            if event == sg.WIN_CLOSED:
                break
            if event == '-CHOOSE-':
                sheets = values["-SHEETS-"][0]
                break

        sheet_chooser.close()
        print(sheets)
        df = pd.read_excel(excel, sheets)
        headers = df.columns[:20].tolist()

    else:
        df = pd.read_excel(excel)
        headers = df.columns[:20].tolist()

    # chosen_headers = ['Country', 'DocumentType', 'State']  # Replace with your actual header names
    # num_headers = 3
    def StartProccess(chosen, column, dst_folder, isHierarcy):
        transfer = []
        fixed_chosen = [sublist[0] for sublist in chosen]
        flag = False
        rows_count = df.count()[0]
        for line_number, (_, row) in enumerate(df.iterrows(), start=1):
            print(f"{line_number} / {rows_count}")
            concat = ""
            # Iterate through the chosen headers and check the values
            for header in fixed_chosen:
                value = row[header]  # Get the value in the current header
                if len(concat) == 0:
                    if isHierarcy:
                        concat = f"{value}\\"
                    else:
                        concat = f"{value}"
                elif (not flag) and isHierarcy:
                    concat = f"{concat}{value}"
                    flag = True
                else:
                    concat = f"{concat}_{value}"
                # print(f"Header: {header}, Value: {value}")
            concat = concat.replace(" ","").replace("nan", "ANY")
            path = row[column]
            print(f"{path} -> {concat}")
            transfer.append([path, concat])
            flag = False

        for line in transfer:
            folders.verify_folder(f"{dst_folder}\\{line[1]}", create=True)
            try:
                shutil.move(line[0], f"{dst_folder}\\{line[1]}")
            except FileNotFoundError:
                continue
        sg.popup("DONE!")

    selected = []
    col1 = [
        # [sg.Input("", key="-SearchKeys-", enable_events=True)],
        [sg.Listbox(headers, expand_y=True, size=(50, 25), key="-KeysList-")],
        [sg.Button("Add")]
    ]
    col2 = [
        # [sg.Input("", key="-SearchSelected-", enable_events=True)],
        [sg.Listbox(selected, expand_y=True, size=(50, 25), key="-SelectedList-")],
        [sg.Button("Remove")]
    ]

    # layout.append([sg.Frame(sg.Column(col1, scrollable=True, expand_y=True))])
    layout_viewer = [
        [sg.Button('Start')],
        [sg.Column(col1, expand_y=True), sg.Column(col2, expand_y=True)],
        [sg.Text("Path Columns"), sg.InputText(key='-PATH_COLUMN-', size=(13, 1)),
         sg.Text("Destination Folder"), sg.InputText(key='-DST-', size=(50, 1))],
        [sg.Checkbox("First Folder Hierarchy", key="-HIERARCHY-", default=False)]
    ]

    # Create PySimpleGUI window
    window_viewer = sg.Window('MoveMate', layout_viewer, icon=r'N:\Images\Shahaf\Projects\Assests\MoveMate.ico')

    # Event loop to handle GUI events
    while True:
        event, values = window_viewer.read()
        if event == sg.WIN_CLOSED:
            break
        if event == 'Start':
            StartProccess(selected, values["-PATH_COLUMN-"], values["-DST-"], values["-HIERARCHY-"])

        if event == "Add":
            selected.append(values["-KeysList-"])
            window_viewer.Element("-SelectedList-").Update(selected)

        if event == "Remove":
            selected.remove(values["-KeysList-"])
            window_viewer.Element("-SelectedList-").Update(selected)

    window_viewer.close()

sg.theme('DarkAmber')

layout = [[sg.Text('Select Your Desire Excel File')],
          [sg.FileBrowse(key="file_name", file_types=(("ALL XLSX Files", "*.xlsx"),)),
           sg.InputText(key='myfile', )],
          [sg.Button('Load Excel'), sg.Button('Cancel')]]

window = sg.Window('MoveMate', layout, icon=r'N:\Images\Shahaf\Projects\Assests\MoveMate.ico', finalize=True)
window.TKroot.bind_all("<Key>", _onKeyRelease, "+")

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel':
        break
    if event == 'Load Excel':
        if "xlsx" not in values["myfile"]:
            sg.Popup("Only Use Xlsx files here!")
        else:
            LoadExcel(values["myfile"])
            window.close()

window.close()
