import csv
import os
import PySimpleGUI as sg


def ToolRun_POAnJsonResults2Excel():
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

    def POAnJsonResults2Excel(folder, filetype):
        # folder = r"N:\Images\Shahaf\PycharmTest\Test"

        data_list = []

        def get_list_of_json_files():
            list_of_files = os.listdir(folder)
            return list_of_files

        def create_list_from_json(XMLfile):
            data_list = []  # create an empty list

            if os.path.getsize(XMLfile) != 0:
                text_file = open(XMLfile)
                data = text_file.read()
                text_file.close()
                input_info = data[data.rfind("<PersonalDataForDataVerification>"):data.rfind("</PersonalDataForDataVerification>")]
                poa_info = data[data.find("<ProofOfAddressRequest>"):data.find("</ProofOfAddressRequest>")]
                print(XMLfile)
                data_list.append(XMLfile[XMLfile.rfind("\\") + 1:XMLfile.index(".")])

                if "<FirstName>" in input_info:
                    temp = input_info[input_info.find("<FirstName>"):input_info.rfind("</FirstName>")]
                    data_list.append(temp[temp.find("<Value>") + len("<Value>"):temp.rfind("</Value>")])
                else:
                    data_list.append("")

                if "<LastName>" in input_info:
                    temp = input_info[input_info.find("<LastName>"):input_info.rfind("</LastName>")]
                    data_list.append(temp[temp.find("<Value>") + len("<Value>"):temp.rfind("</Value>")])
                else:
                    data_list.append("")

                if "<Address>" in input_info:
                    temp = input_info[input_info.find("<Address>"):input_info.rfind("</Address>")]
                    data_list.append(temp[temp.find("<Value>") + len("<Value>"):temp.rfind("</Value>")])
                else:
                    data_list.append("")

                # ----------------------------------------

                if "<FirstName>" in poa_info:
                    temp = poa_info[poa_info.find("<FirstName>"):poa_info.rfind("</FirstName>")]
                    data_list.append(temp[temp.find("<Value>") + len("<Value>"):temp.rfind("</Value>")])
                else:
                    data_list.append("")

                if "<LastName>" in poa_info:
                    temp = poa_info[poa_info.find("<LastName>"):poa_info.rfind("</LastName>")]
                    data_list.append(temp[temp.find("<Value>") + len("<Value>"):temp.rfind("</Value>")])
                else:
                    data_list.append("")

                if "<Address>" in poa_info:
                    temp = poa_info[poa_info.find("<Address>"):poa_info.rfind("</Address>")]
                    data_list.append(temp[temp.find("<Value>") + len("<Value>"):temp.rfind("</Value>")])
                else:
                    data_list.append("")


                if "<IsFullNameMatch>" in data:
                    if data[data.find("<IsFullNameMatch>")+len("<IsFullNameMatch>"):data.find("</IsFullNameMatch>")] == "true":
                        data_list.append("Match")
                    else:
                        data_list.append("Not Match")
                else:
                    data_list.append("NULL")

                if "<IsAddressMatch>" in data:
                    if data[data.find("<IsAddressMatch>")+len("<IsAddressMatch>"):data.find("</IsAddressMatch>")] == "true":
                        data_list.append("Match")
                    else:
                        data_list.append("Not Match")
                else:
                    data_list.append("NULL")

            #path
            # data_list.append(jsonfile[:jsonfile.rfind("\\")])
            return data_list

        def write_csv():
            nested_list_of_files = []
            list_of_files = get_list_of_json_files()
            first_row = []
            first_row.append('DocumentID')
            first_row.append('Json First Name')
            first_row.append('Json Last Name')
            first_row.append('Json Address')

            first_row.append('POA First Name')
            first_row.append('POA Last Name')
            first_row.append('POA Address')

            first_row.append('FullName Status')
            first_row.append('Address Status')



            with open(folder + '\output.csv', "a", newline='') as c:
                writer = csv.writer(c)
                writer.writerow(first_row)

            for file in list_of_files:
                if "."+filetype in file:
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
        popup = sg.Window("Complete", layout_popup, icon=r'N:\Images\Shahaf\Projects\Assests\POAnJsonResults2Excel.ico')

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
        [sg.Text("Covert ClientDemo XML/Jason POA to UserInput results file to Excel")],
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

    window = sg.Window("POAnJsonResults2Excel", layout, icon=r'N:\Images\Shahaf\Projects\Assests\POAnJsonResults2Excel.ico', finalize=True)
    window.TKroot.bind_all("<Key>", _onKeyRelease, "+")

    while True:
        event, values = window.read()
        if event == "Start" and len(values['myfolder']) > 1:
            POAnJsonResults2Excel(values['myfolder'], values['filetype'])
        if event == sg.WIN_CLOSED or event == "Close":
            break

    window.close()


ToolRun_POAnJsonResults2Excel()
