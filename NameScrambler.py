import os
import random
import string
import PySimpleGUI as sg


def ToolRun_scrambler():
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

    def scrambler(folder):
        for count, filename in enumerate(os.listdir(folder)):
            num = random.randrange(0, 1000)
            letters = string.ascii_letters
            name = ''.join(random.choice(letters) for i in range(10))
            name += repr(num)
            dst = f"{str(name)}_page0.jpg"
            src = f"{folder}/{filename}"  # foldername/filename, if .py file is outside folder
            dst = f"{folder}/{dst}"
            try:
                os.rename(src, dst)
            except:
                continue

        layout_popup = [
            [sg.Text("Done!")],
            [sg.Button('Close')]
        ]
        popup = sg.Window("Complete", layout_popup, icon=r'N:\Images\Shahaf\Projects\Assests\namescramble.ico')

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
        [sg.Text("Scramble all the Images names in a specific folder")],
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

    window = sg.Window("NameScrambler", layout, icon=r'N:\Images\Shahaf\Projects\Assests\namescramble.ico', finalize=True)
    window.TKroot.bind_all("<Key>", _onKeyRelease, "+")

    while True:
        event, values = window.read()
        if event == "Start" and len(values['myfolder']) > 1:
            scrambler(values['myfolder'])
        if event == sg.WIN_CLOSED or event == "Close":
            break

    window.close()


ToolRun_scrambler()
