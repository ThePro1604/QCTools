import csv
import os
import PySimpleGUI as sg


def ToolRun_IDnJsonResults2Excel():
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

    def IDnJsonResults2Excel(folder, filetype):
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
                input_info = data[data.rfind("<IdentityDataInput>"):data.rfind("</IdentityDataInput>")]
                bos_info = data[data.find("<DocumentData"):data.find("</DocumentData>")]
                bos_info2 = data[data.find("<DocumentData2>"):data.find("</DocumentData2>")]
                print(XMLfile)
                fullname = False

                data_list.append(XMLfile[XMLfile.rfind("\\") + 1:XMLfile.index(".")])

                if "<FirstName>" in input_info:
                    data_list.append(input_info[input_info.rfind("<FirstName>")+len("<FirstName>"):input_info.find("</FirstName")])
                else:
                    data_list.append("")

                if "<LastName>" in input_info:
                    data_list.append(input_info[input_info.rfind("<LastName>")+len("<LastName>"):input_info.find("</LastName")])
                else:
                    data_list.append("")

                if "<DateOfBirth>" in input_info:
                    date = input_info[input_info.rfind("<DateOfBirth>")+len("<DateOfBirth>"):input_info.find("T00")]
                    date = date[len(date)-2:] + "." + date[5:7] + "." + date[:4]
                    data_list.append(date)
                else:
                    data_list.append("NULL")

                if "<DateOfIssue>" in input_info:
                    data_list.append(input_info[input_info.rfind("<DateOfIssue>")+len("<DateOfIssue>"):input_info.find("</DateOfIssue")])
                else:
                    data_list.append("NULL")

                if "<DateOfExpiry>" in input_info:
                    data_list.append(input_info[input_info.rfind("<DateOfExpiry>")+len("<DateOfExpiry>"):input_info.find("</DateOfExpiry")])
                else:
                    data_list.append("NULL")


                if "FirstName=" in bos_info:
                    temp = bos_info[bos_info.find("FirstName="):bos_info.find(">")]
                    temp = temp[len("FirstName="):temp.find('" ')+1]
                    data_list.append(temp[1:-1])
                elif "<FullName>" in bos_info2:
                    data_list.append("")
                    data_list.append("")
                    temp = bos_info2[bos_info2.find("<FullName>"):bos_info2.rfind("</FullName>")]
                    data_list.append(temp[temp.find("<Value>") + len("<Value>"):temp.rfind("</Value>")])
                    fullname = True
                else:
                    data_list.append("")

                if not fullname:
                    if "LastName=" in bos_info:

                        temp = bos_info[bos_info.find("LastName="):bos_info.find(">")]
                        if "Nationality" in temp:
                            temp = temp[len("LastName="):temp.find('" ')+1]
                        else:
                            temp = temp[len("LastName="):]
                        data_list.append(temp[1:-1])
                    else:
                        data_list.append("")
                if not fullname:
                    data_list.append("")

                if "<Value>" in bos_info[bos_info.find("<DateOfBirth>"):bos_info.find("</DateOfBirth>")]:
                    temp = bos_info[bos_info.rfind("<DateOfBirth>"):bos_info.find("</DateOfBirth>")]
                    temp = temp[temp.rfind("<Value>")+len("<Value>"):temp.find("</Value>")]
                    temp = temp.replace("/", ".")
                    data_list.append(temp)
                else:
                    data_list.append("NULL")

                if "<Value>" in bos_info[bos_info.find("<DateOfIssue>"):bos_info.find("</DateOfIssue>")]:
                    temp = bos_info[bos_info.rfind("<DateOfIssue>"):bos_info.find("</DateOfIssue>")]
                    temp = temp[temp.rfind("<Value>")+len("<Value>"):temp.find("</Value>")]
                    temp = temp.replace("/", ".")
                    data_list.append(temp)
                else:
                    data_list.append("NULL")

                if "<Value>" in bos_info[bos_info.find("<DateOfExpiry>"):bos_info.find("</DateOfExpiry>")]:
                    temp = bos_info[bos_info.rfind("<DateOfExpiry>"):bos_info.find("</DateOfExpiry>")]
                    temp = temp[temp.rfind("<Value>")+len("<Value>"):temp.find("</Value>")]
                    temp = temp.replace("/", ".")
                    data_list.append(temp)

                else:
                    data_list.append("NULL")



                if "<FullNameStatus>" in data:
                    if data[data.find("<FullNameStatus>")+len("<FullNameStatus>"):data.find("<FullNameStatus>")+len("<FullNameStatus>")+1] == "0":
                        data_list.append("Match")
                    else:
                        data_list.append("Not Match")
                else:
                    data_list.append("NULL")

                if "<DateOfBirthStatus>" in data:
                    if data[data.find("<DateOfBirthStatus>")+len("<DateOfBirthStatus>"):data.find("<DateOfBirthStatus>")+len("<DateOfBirthStatus>")+1] == "0":
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
            first_row.append('Input First Name')
            first_row.append('Input Last Name')
            first_row.append('Input Date of Birth')
            first_row.append('Input Date of Issue')
            first_row.append('Input Date of Expiry')

            first_row.append('Output First Name')
            first_row.append('Output Last Name')
            first_row.append('Output Full Name')
            first_row.append('Output Date of Birth')
            first_row.append('Output Date of Issue')
            first_row.append('Output Date of Expiry')

            first_row.append('Full Name Status')
            first_row.append('Date Of Birth Status')



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
        popup = sg.Window("Complete", layout_popup, icon=r'N:\Images\Shahaf\Projects\Assests\idnjson2excel.ico')

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
        [sg.Text("Converts ClientDemo XML/Jason Id's to UserInput results files to Excel")],
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

    window = sg.Window("IDnJsonResults2Excel", layout, icon=r'N:\Images\Shahaf\Projects\Assests\idnjson2excel.ico', finalize=True)
    window.TKroot.bind_all("<Key>", _onKeyRelease, "+")

    while True:
        event, values = window.read()
        if event == "Start" and len(values['myfolder']) > 1:
            IDnJsonResults2Excel(values['myfolder'], values['filetype'])
        if event == sg.WIN_CLOSED or event == "Close":
            break

    window.close()


ToolRun_IDnJsonResults2Excel()
