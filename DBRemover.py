import os
import subprocess
import PySimpleGUI as sg


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


def DBRemover(path, excluded):
    folder_list = excluded.split('\n')
    # dir = 'N:\Images\Shahaf'
    # exclude = ['Tools', 'Projects', "SplitExcel"]

    os.system("taskkill /im explorer.exe /F")
    count = 0
    for root, dirs, files in os.walk(path):
        dirs[:] = [d for d in dirs if d not in folder_list]
        for file in files:
            print(os.path.join(root, file))
            if file.endswith('Thumbs.db'):
                try:
                    os.remove(os.path.join(root, file))
                except PermissionError:
                    os.system(f'cmd /c "del /s /q {os.path.join(root, file)}"')
                count += 1
                print("---------------------------------------- " + os.path.join(root, file))
    print(str(count) + " Thumbs file were removed")
    subprocess.Popen("explorer.exe")
    sg.popup("Done")


top = [
    [sg.Text("Folder")],
    [sg.FolderBrowse(key="folder_name"), sg.InputText(key='myfolder')],
    [sg.Text("Excluded Folders")],
    [sg.Multiline(key="-EX_FOLDERS-", size=(10, 10), expand_x=True)],

]

buttons = [
    [sg.Button('Start'), sg.Button('Close')],
]

copyright = [
    [sg.Text("Ⓒ By Shahaf Stossel", text_color="#A20909")],
]

description = [
    [sg.Text("Remove all Thumbs.db files from a folder (folder and sub folders)")],
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
window = sg.Window("DB Remover", layout, icon=r'N:\Images\Shahaf\Projects\Assests\DBRemover.ico', finalize=True)
window.TKroot.bind_all("<Key>", _onKeyRelease, "+")

while True:
    event, values = window.read()
    if event == "Start" and len(values['myfolder']) > 1:
        DBRemover(values['myfolder'], values["-EX_FOLDERS-"])
    if event == sg.WIN_CLOSED or event == "Close":
        break

window.close()
