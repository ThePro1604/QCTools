import os
import re
import shutil
import PySimpleGUI as sg
from collections import defaultdict
import pandas as pd
import sys
import logging


try:
    os.remove("C:/QCCenter/Logs/MoveFiles.log")
    logging.basicConfig(filename='C:/QCCenter/Logs/MoveFiles.log')
except FileNotFoundError:
    logging.basicConfig(filename='C:/QCCenter/Logs/MoveFiles.log')


def ToolRun_MoveFiles():
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

    def GetHeader(excel):
        df = pd.read_excel(file)
        headers = df.columns[:20].tolist()

    def MoveFiles(path, file):
        df = pd.read_excel(file)
        pattern = re.compile(r'^[A-Z\d]+$')
        data_dict = defaultdict(list)


        for root, subdirs, files in os.walk(path):
            for name in files:
                key_name = name[:name.rfind("_")]
                data_dict[key_name].append([os.path.join(root, name), name])


        # Iterate through each row in the dataframe
        for index, row in df.iterrows():
            file_name = row['DocumentId']
            file_name = file_name.replace("-", "")
            folder_name = row['classification']
            folder_path = os.path.join(path, str(folder_name))

            # Check if the folder exists, create it if not
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            # Check if the file exists
            # print(file_name)
            if file_name.upper() in data_dict:
                for file in data_dict[file_name.upper()]:
                    file_path = file[0]
                    if os.path.exists(file_path) and os.access(file_path, os.W_OK) and os.access(folder_path, os.W_OK):
                        shutil.move(str(file_path), os.path.join(folder_path, str(file[1])))
                    else:
                        print("You do not have permission to move the file: ", file_path)

    layout = [
        [sg.Text("Excel file"), sg.Input(), sg.FileBrowse()],
        [sg.Text("Folder to search"), sg.Input(), sg.FolderBrowse()],
        [sg.Submit(), sg.Cancel()]
    ]

    window = sg.Window("Move Files", layout, icon=r'N:\Images\Shahaf\Projects\Assests\MoveFiles.ico', finalize=True)
    # window.TKroot.bind_all("<Key>", _onKeyRelease, "+")

    # event, values = window.Read()
    while True:
        event, values = window.read()
        if event == "Submit":
            file = values[0]
            path = values[1]
            MoveFiles(path, file)
            sg.Popup("Files have been moved.")
        if event == sg.WIN_CLOSED or event == "Close":
            break

    window.close()

    # if event == "Submit":
    #     file = values[0]
    #     path = values[1]
    #     MoveFiles(path, file)
    #     sg.Popup("Files have been moved.")
    #
    # window.Close()


ToolRun_MoveFiles()
