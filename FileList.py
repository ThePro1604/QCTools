import os
import pandas as pd
import PySimpleGUI as sg
import sys
import logging


try:
    os.remove("C:/QCCenter/Logs/FileList.log")
    logging.basicConfig(filename='C:/QCCenter/Logs/FileList.log')
except FileNotFoundError:
    logging.basicConfig(filename='C:/QCCenter/Logs/FileList.log')


def ToolRun_FileList():
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

    #path of the file you want to enemurate
    def FileList(path):
        excel_name = os.path.basename(os.path.normpath(path))
        directory =[]
        filename=[]

        for (root,dirs, file) in os.walk(path):
            for f in file:
                directory.append(root)
                filename.append(f)
                print(f)

        #column name of the sheet
        df=pd.DataFrame(list(zip(directory,filename)),columns=['Directory',"filename"])
        #change the file of exccl sheet
        df.to_csv(path + f"/{excel_name}.csv")

        layout_popup = [
            [sg.Text("Done!")],
            [sg.Button('Close')]
        ]
        popup = sg.Window("Complete", layout_popup, icon=r'N:\Images\Shahaf\Projects\Assests\filelist.ico')

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
        [sg.Text("Create an excel file from all the documents in a specific folder")],
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

    window = sg.Window("FileList", layout, icon=r'N:\Images\Shahaf\Projects\Assests\filelist.ico', finalize=True)
    window.TKroot.bind_all("<Key>", _onKeyRelease, "+")


    while True:
        event, values = window.read()
        if event == "Start" and len(values['myfolder']) > 1:
            FileList(values['myfolder'])
        if event == sg.WIN_CLOSED or event == "Close":
            break

    window.close()


ToolRun_FileList()
