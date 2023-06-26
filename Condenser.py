import os
import shutil
import PySimpleGUI as sg
import sys
import logging


try:
    os.remove("C:/QCCenter/Logs/Condenser.log")
    logging.basicConfig(filename='C:/QCCenter/Logs/Condenser.log')
except FileNotFoundError:
    logging.basicConfig(filename='C:/QCCenter/Logs/Condenser.log')


# root_path = r"N:\Images\Shahaf\QCSupport\PaySafe"
# new_root = os.path.join(root_path, "Organized")
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

def Condenser(root_path):
    new_root = os.path.join(root_path, "Organized")

    # identify the directories to organize
    directories_to_organize = []
    for folder_name, subfolders, _ in os.walk(root_path):
        if not subfolders:
            directories_to_organize.append(folder_name)

    os.makedirs(new_root, exist_ok=True)

    # organize the directories
    for folder_name in directories_to_organize:
        print(f"Organizing {folder_name}")
        new_name = folder_name[len(root_path)+1:].replace("\\", "_").replace(" ", "")
        os.makedirs(os.path.join(new_root, new_name), exist_ok=True)
        for file_name in os.listdir(folder_name):
            if 'Thumbs.db' in file_name:
                continue
            source_file = os.path.join(folder_name, file_name)
            destination_file = os.path.join(os.path.join(new_root, new_name), file_name)
            shutil.move(source_file, destination_file)

    sg.Popup('Ok clicked', keep_on_top=True)


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
    [sg.Text("Condense all nested folders to one")],
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

window = sg.Window("Condenser", layout, icon=r'N:\Images\Shahaf\Projects\Assests\Condenser.ico', finalize=True)
window.TKroot.bind_all("<Key>", _onKeyRelease, "+")


while True:
    event, values = window.read()
    if event == "Start" and len(values['myfolder']) > 1:
        Condenser(values['myfolder'])
    if event == sg.WIN_CLOSED or event == "Close":
        break

window.close()
