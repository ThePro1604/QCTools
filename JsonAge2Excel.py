import csv
import os
import PySimpleGUI as sg

def ToolRun_JsonAge2Excel():
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

    def JsonAge2Excel(folder, filetype):
        # folder = r"N:\Images\Shahaf\PycharmTest\Test"

        data_list = []

        def get_list_of_json_files():
            directory = []
            filename = []
            list_of_files = []
            for (root, dirs, file) in os.walk(folder):
                for f in file:
                    directory.append(root)
                    filename.append(f)
                    if "json" in f:
                        # print(f)
                        list_of_files.append(os.path.join(root, f))
            return list_of_files

        def create_list_from_json(XMLfile):
            data_list = []  # create an empty list

            if os.path.getsize(XMLfile) != 0:
                print(XMLfile)

                text_file = open(XMLfile, encoding="utf8")
                data = text_file.read()
                text_file.close()
                age_verification = data[data.rfind("AgeVerificationReport"):data.rfind("AgeEstimationCompletionStatus")]
                # data_list.append(XMLfile[XMLfile.rfind("\\") + 1:XMLfile.index(".")])

                if "DocumentId" in data:
                    temp = data[data.find("DocumentId"):data.find("DocumentScope")]
                    data_list.append(temp[temp.find(": ")+3:temp.find(",")-1])
                else:
                    data_list.append("")


                if "AgeEstimation" in age_verification:
                    temp = age_verification[age_verification.find("AgeEstimation"):age_verification.find("AgeEstimationCompletionStatus")]
                    data_list.append(temp[temp.find(": ")+1:temp.find(",")])
                else:
                    data_list.append("")

            #path
            data_list.append(XMLfile[:XMLfile.rfind("\\")])
            return data_list

        def write_csv():
            nested_list_of_files = []
            list_of_files = get_list_of_json_files()
            first_row = []
            first_row.append('DocumentID')
            first_row.append('Age')
            first_row.append('Path')



            with open(folder + '\output.csv', "a", newline='') as c:
                writer = csv.writer(c)
                writer.writerow(first_row)

            for file in list_of_files:
                row = create_list_from_json(file)  # create the row to be added to csv for each file (json-file)
                with open(folder + '\output.csv', "a", newline='') as c:
                    try:
                        writer = csv.writer(c)
                        writer.writerow(row)
                    except:
                        continue
                    # writer = csv.writer(c)
                    # writer.writerow(row)
                c.close()
        write_csv()

        layout_popup = [
            [sg.Text("Done!")],
            [sg.Button('Close')]
        ]
        popup = sg.Window("Complete", layout_popup, icon=r'N:\Images\Shahaf\Projects\Assests\jsonage2excel.ico')

        while True:
            event, values = popup.read()
            if event == "Close" or event == sg.WIN_CLOSED:
                break
        popup.close()

    top = [
        [sg.Text("Source Folder")],
        [sg.FolderBrowse(key="folder_name"), sg.InputText(key='myfolder')],
    ]

    middle = [
        [sg.Text("File Type (txt, json, xml)"), sg.InputText(key='filetype', size=(5, 1))],
    ]

    buttons = [
        [sg.Button('Start'), sg.Button('Close')],
    ]

    copyright = [
        [sg.Text("Ⓒ By Shahaf Stossel", text_color="#A20909")],
    ]

    description = [
        [sg.Text("Extracts the estimated age from json files to excel ")],
    ]

    layout = [
        [sg.Push(), sg.Column(description), sg.Push()],
        [sg.HSeparator()],
        [
            sg.Column(top, vertical_alignment="top", key='-COL1-'),
        ],
        [sg.Push(), sg.Column(middle), sg.Push()],
        [sg.Push(), sg.Column(buttons), sg.Push()],
        [sg.Push(), sg.Column(copyright), sg.Push()],
    ]

    window = sg.Window("JsonAge2Excel", layout, icon=r'N:\Images\Shahaf\Projects\Assests\jsonage2excel.ico', finalize=True)
    window.TKroot.bind_all("<Key>", _onKeyRelease, "+")

    while True:
        event, values = window.read()
        if event == "Start" and len(values['myfolder']) > 1:
            JsonAge2Excel(values['myfolder'], values['filetype'])
        if event == sg.WIN_CLOSED or event == "Close":
            break

    window.close()


ToolRun_JsonAge2Excel()
