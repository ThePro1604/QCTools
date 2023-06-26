import csv
import os
import json
import PySimpleGUI as sg
import sys
import logging


try:
    os.remove("C:/QCCenter/Logs/Jsons2Excel.log")
    logging.basicConfig(filename='C:/QCCenter/Logs/Jsons2Excel.log')
except FileNotFoundError:
    logging.basicConfig(filename='C:/QCCenter/Logs/Jsons2Excel.log')


def ToolRun_Jsons2Excel():
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

    def Jsons2Excel(folder):
        print(folder)
        data_list = []
        def get_list_of_json_files():
            list_of_files = os.listdir(folder)
            return list_of_files

        def get_list_of_json_files_nested(nested):
            list_of_files = os.listdir(folder + "\\" + nested)
            return list_of_files

        def create_list_from_json(jsonfile):
            with open(jsonfile, 'rb') as f:
                data = json.loads(f.read().decode('utf-8'))

            print(jsonfile)
            data_list = []  # create an empty list

            # append the items to the list in the same order.
            data_list.append(jsonfile[jsonfile.rfind("\\")+1:jsonfile.rfind("_")])
            try:
                data_list.append(data['DocumentNumber'])
            except KeyError:
                data_list.append("")

            try:
                data_list.append(data['FirstName'])
            except KeyError:
                data_list.append("")

            try:
                data_list.append(data['MiddleName'])
            except KeyError:
                data_list.append("")

            try:
                data_list.append(data['LastName'])
            except KeyError:
                data_list.append("")

            try:
                data_list.append(data['DateOfBirth'])
            except KeyError:
                data_list.append("")

            try:
                data_list.append(data['NationalityIso3'])
            except KeyError:
                data_list.append("")

            try:
                data_list.append(data['PersonalNumber'])
            except KeyError:
                data_list.append("")

            try:
                data_list.append(data['CountryIso3'])
            except KeyError:
                data_list.append("")

            data_list.append(jsonfile[:jsonfile.rfind("\\")])
            return data_list

        def write_csv():
            nested_list_of_files = []
            list_of_files = get_list_of_json_files()
            first_row = []
            first_row.append('RequestID')
            first_row.append('Document Number')
            first_row.append('First Name')
            first_row.append('Middle Name')
            first_row.append('Last Name')
            first_row.append('Date Of Birth')
            first_row.append('Nationality')
            first_row.append('Personal Number')
            first_row.append('Country')
            first_row.append('Path')
            with open(folder + '\output.csv', "a", newline='') as c:
                writer = csv.writer(c)
                writer.writerow(first_row)

            for file in list_of_files:
                if "." not in file:
                    list_of_files_nested = get_list_of_json_files_nested(file)
                    for nested_file in list_of_files_nested:
                        if ".json" in nested_file:
                            row = create_list_from_json(
                                folder + "\\" + file + "\\" +nested_file)  # create the row to be added to csv for each file (json-file)
                            with open(folder + '\output.csv', "a", newline='') as c:
                                writer = csv.writer(c)
                                writer.writerow(row)
                            c.close()

                if ".json" in file:
                    row = create_list_from_json(folder+"\\"+file)  # create the row to be added to csv for each file (json-file)
                    with open(folder + '\output.csv', "a", newline='') as c:
                        writer = csv.writer(c)
                        writer.writerow(row)
                    c.close()
        write_csv()

        layout_popup = [
            [sg.Text("Done!")],
            [sg.Button('Close')]
        ]

        popup = sg.Window("Complete", layout_popup, icon=r'N:\Images\Shahaf\Projects\Assests\json2excel.ico')

        while True:
            event, values = popup.read()
            if event == "Close" or event == sg.WIN_CLOSED:
                break
        popup.close()


    top = [
        [sg.Text("Folder")],
        [sg.FolderBrowse(key="folder_name"), sg.InputText(key='myfolder')],
    ]

    buttons = [
        [sg.Button('Start'), sg.Button('Close')],
    ]

    copyright = [
        [sg.Text("Ⓒ By Shahaf Stossel", text_color="#A20909")],
    ]

    description = [
        [sg.Text("Convert all the Jsons in a folder to Excel")],
    ]

    layout = [
        [sg.Push(), sg.Column(description), sg.Push()],
        [sg.HSeparator()],
        [
            sg.Column(top, vertical_alignment="top", key='-COL1-'),
        ],
        [sg.Push(), sg.Column(buttons), sg.Push()],
        [sg.Push(), sg.Column(copyright), sg.Push()],
    ]

    window = sg.Window("Jsons2Excel", layout, icon=r'N:\Images\Shahaf\Projects\Assests\json2excel.ico', finalize=True)
    window.TKroot.bind_all("<Key>", _onKeyRelease, "+")

    while True:
        event, values = window.read()
        if event == "Start" and len(values['myfolder']) > 1:
            Jsons2Excel(values['myfolder'])
        if event == sg.WIN_CLOSED or event == "Close":
            break

    window.close()


ToolRun_Jsons2Excel()
