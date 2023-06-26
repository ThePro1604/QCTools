import json
import PySimpleGUI as sg
import os
import pandas as pd
import sys
import logging


try:
    os.remove("C:/QCCenter/Logs/JsonExtractor.log")
    logging.basicConfig(filename='C:/QCCenter/Logs/JsonExtractor.log')
except FileNotFoundError:
    logging.basicConfig(filename='C:/QCCenter/Logs/JsonExtractor.log')


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


def extractionGui(json_list):
    def handle_combo_keyboard_events(event, values):
        if event == '-SearchKeys-' and values['-SearchKeys-'] != '':
            letter_pressed = values['-SearchKeys-'][-1].upper()
            if letter_pressed.isalpha() or letter_pressed == "BackSpace":
                updated_data = list(filter(lambda k: values['-SearchKeys-'].lower() in k.lower(), leaf_keys))
                window_viewer["-KeysList-"].update(values=updated_data)

        if event == '-SearchSelected-' and values['-SearchSelected-'] != '':
            letter_pressed = values['-SearchSelected-'][-1].upper()
            if letter_pressed.isalpha() or letter_pressed == "BackSpace":
                updated_data = list(filter(lambda k: values['-SearchSelected-'].lower() in k.lower(), selected))
                window_viewer["-SelectedList-"].update(values=updated_data)

    def infoExtraction(jsonList, keys):
        # Load JSON data from file
        cols = []
        header = []
        header.append("Path")
        for key in keys:
            header.append(key[0])

        for file in jsonList:
            with open(file, encoding='utf-8') as f:
                data1 = json.load(f)
                values = []
                values.append(file)
                for key in keys:
                    if key[0] in leaf_keys:
                        value = data1
                        for k in key[0].split('.'):
                            if '[' in k:
                                idx = int(k[k.index('[') + 1:k.index(']')])
                                k = k[:k.index('[')]
                                value = value[k][idx]
                            else:
                                try:
                                    value = value[k]
                                except KeyError:
                                    value = ""
                        values.append(value)
                    else:
                        values.append("")
            cols.append(values)
        df1 = pd.DataFrame(cols, columns=header)
        df1.to_excel(os.path.join(folder_path, 'data.xlsx'), index=False)
        sg.Popup("DONE")

    # Load JSON data from file
    with open(json_list[0], encoding='utf-8') as f:
        data = json.load(f)

    # Flatten JSON object into a dictionary with dotted keys
    # def flatten_json(obj, parent_key='', sep='.'):
    #     items = []
    #     for key, value in obj.items():
    #         new_key = f"{parent_key}{sep}{key}" if parent_key else key
    #         if isinstance(value, dict):
    #             items.extend(flatten_json(value, new_key, sep=sep).items())
    #         else:
    #             items.append((new_key, value))
    #     return dict(items)
    #
    # flat_data = flatten_json(data)

    def get_leaf_keys(json_obj, parent_key=''):
        """
        Recursively traverse the JSON object and return a list of all the leaf keys.
        """
        if isinstance(json_obj, dict):
            leaf_keys = []
            for key in json_obj:
                new_key = f"{parent_key}.{key}" if parent_key else key
                leaf_keys.extend(get_leaf_keys(json_obj[key], new_key))
            return leaf_keys
        elif isinstance(json_obj, list):
            leaf_keys = []
            for i, item in enumerate(json_obj):
                new_key = f"{parent_key}[{i}]" if parent_key else f"[{i}]"
                leaf_keys.extend(get_leaf_keys(item, new_key))
            return leaf_keys
        else:
            return [parent_key]

    leaf_keys = get_leaf_keys(data)
    selected = []

    col1 = [
        [sg.Input("", key="-SearchKeys-", enable_events=True)],
        [sg.Listbox(leaf_keys, expand_y=True, size=(100, None), key="-KeysList-")],
        [sg.Button("Add")]
    ]
    col2 = [
        [sg.Input("", key="-SearchSelected-", enable_events=True)],
        [sg.Listbox(selected, expand_y=True, size=(100, None), key="-SelectedList-")],
        [sg.Button("Remove")]
    ]

    # layout.append([sg.Frame(sg.Column(col1, scrollable=True, expand_y=True))])
    layout_viewer = [
        [sg.Button('Print Values')],
        [sg.Column(col1, expand_y=True), sg.Column(col2, expand_y=True)]
    ]

    # Create PySimpleGUI window
    window_viewer = sg.Window('JSON Viewer', layout_viewer, size=(1600, 600), icon=r'N:\Images\Shahaf\Projects\Assests\JsonExtractor.ico')

    # Event loop to handle GUI events
    while True:
        event, values = window_viewer.read()
        if event == sg.WIN_CLOSED:
            break
        if event == 'Print Values':
            infoExtraction(json_list, selected)

        if event == "Add":
            selected.append(values["-KeysList-"])
            window_viewer.Element("-SelectedList-").Update(selected)

        if event == "Remove":
            selected.remove(values["-KeysList-"])
            window_viewer.Element("-SelectedList-").Update(selected)

        try:
            handle_combo_keyboard_events(event, values)
        except AttributeError:
            continue

    # Close the window and exit the program
    window_viewer.close()


def find_json_files(path):
    json_files = []
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.endswith(".json"):
                json_files.append(os.path.join(root, file))
    return json_files


sg.theme('DarkAmber')

layout = [[sg.Text('Select a folder to search for JSON files')],
          [sg.Input(key='-FOLDER-'), sg.FolderBrowse()],
          [sg.Button('Load Jsons'), sg.Button('Cancel')]]

window = sg.Window('JSON Extractor', layout, icon=r'N:\Images\Shahaf\Projects\Assests\JsonExtractor.ico', finalize=True)
window.TKroot.bind_all("<Key>", _onKeyRelease, "+")

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel':
        break
    if event == 'Load Jsons':
        folder_path = values['-FOLDER-']
        json_files = find_json_files(folder_path)
        window.close()
        extractionGui(json_files)

window.close()
