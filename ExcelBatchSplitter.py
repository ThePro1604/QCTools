import os
import pandas as pd
from PrivateFunctions import folders
import PySimpleGUI as sg
import sys
import logging


try:
    os.remove("C:/QCCenter/Logs/MasterSplitter.log")
    logging.basicConfig(filename='C:/QCCenter/Logs/MasterSplitter.log')
except FileNotFoundError:
    logging.basicConfig(filename='C:/QCCenter/Logs/MasterSplitter.log')


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


def ExcelBatchSplitter(output, excel, lines, sheet, column, isSheet):
    def create_list_from_excel(excel_name, column_name, sheet_name):
        # Read the Excel file into a DataFrame
        df = pd.read_excel(excel_name, sheet_name=sheet_name)
        # Check if the column name exists in the DataFrame
        if column_name not in df.columns:
            print(f"Column '{column_name}' does not exist in the sheet '{sheet_name}'.")
            return []

        # Retrieve the column values as a list
        column_values = df[column_name].tolist()

        return column_values


    # Prompt the user for column and sheet names
    column_name = column
    sheet_name = sheet
    excel_name = excel

    def create_txt_files_from_list(data_list, lines_per_file, output_dst, sht_name):
        num_files = (len(data_list) + lines_per_file - 1) // lines_per_file
        folders.verify_folder(f"{output_dst}\\{sht_name}", create=True)
        for i in range(num_files):
            start_index = i * lines_per_file
            end_index = (i + 1) * lines_per_file
            file_lines = data_list[start_index:end_index]

            with open(f"{output_dst}\\{sht_name}\\{sht_name}_{i+1}.bat", "w") as file:
                for line in file_lines:
                    file.write(str(line) + "\n")

    if isSheet:
        # Get all the sheet names
        xls = pd.ExcelFile(excel_name)
        sheets = xls.sheet_names
        print(sheets)
        for sht in sheets:
            result = create_list_from_excel(excel_name, column_name, sht)
            lines_per_file = int(lines)
            output_dst = output
            create_txt_files_from_list(result, lines_per_file, output_dst, sht)
    else:
        # Create the list from Excel column
        result = create_list_from_excel(excel_name, column_name, sheet_name)
        lines_per_file = int(lines)
        output_dst = output
        create_txt_files_from_list(result, lines_per_file, output_dst, sheet_name)
    sg.popup("DONE!")


left = [
    [sg.Text("Choose File")],
    [sg.FileBrowse(key="file_name", file_types=(("ALL Files", "*.*"), ("ALL CSV Files", "*.csv"), ("ALL XLSX Files", "*.xlsx"),)),
     sg.InputText(key='myfile', )],
]

right = [
    [sg.Text("Output Folder")],
    [sg.FolderBrowse(key="folder_name"), sg.InputText(key='myfolder')],

]

middle = [
    [sg.Text("Batch Line Column"), sg.InputText(key='-BATCH_COLUMN-', size=(10, 1)),
    sg.Text("Sheet Name"), sg.InputText(key='-SHEET-', size=(13, 1)),
    sg.Text("Lines Per Batch"), sg.InputText(key='-LINES-', size=(13, 1))],
    [sg.Checkbox("All Sheets", key="-ALL_SHEET-", default=False)]
]

buttons = [
    [sg.Button('Start'), sg.Button('Close')],
]

copyright = [
    [sg.Text("Ⓒ By Shahaf Stossel", text_color="#A20909")],
]

description = [
    [sg.Text("This tool is used to create separate batch file out of an lines in a specific column")]
    # [sg.Text("the DocumentID cell is the one that is being used as the Name of the File")]

]

layout = [
    [sg.Push(), sg.Column(description), sg.Push()],
    [sg.HSeparator()],
    [
        sg.Column(left, vertical_alignment="top", key='-COL1-'),
        sg.VSeparator(),
        sg.Column(right, vertical_alignment="top", key='-COL1-')
    ],
    [sg.VPush()],
    [sg.Push(), sg.Column(middle), sg.Push()],
    [sg.VPush()],
    [sg.Push(), sg.Column(buttons), sg.Push()],
    [sg.Push(), sg.Column(copyright), sg.Push()],

]

window = sg.Window("MasterSplitter", layout, icon=r'N:\Images\Shahaf\Projects\Assests\ExcelSplitter.ico', finalize=True)
window.TKroot.bind_all("<Key>", _onKeyRelease, "+")

while True:
    event, values = window.read()
    if event == "Start" and len(values['myfolder']) > 1:
        ExcelBatchSplitter(values['myfolder'], values['file_name'], values['-LINES-'], values['-SHEET-'], values["-BATCH_COLUMN-"], values["-ALL_SHEET-"])
    if event == sg.WIN_CLOSED or event == "Close":
        break

window.close()
