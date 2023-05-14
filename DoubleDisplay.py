import PySimpleGUI as sg
import pyperclip as pyperclip
from PIL import Image
import io
import os.path
from xlsx2csv import Xlsx2csv
from io import StringIO
import pandas as pd
from warnings import simplefilter


simplefilter(action="ignore", category=pd.errors.PerformanceWarning)


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


def ToolRun_DoubleDisplay():
    def warning():
        layout_popup = [
            [sg.Text("Make Sure The Excel File Is Closed!")],
            [sg.Button('Close')]
        ]
        popup_warning = sg.Window("Warning!", layout_popup)

        while True:
            event_warning, values_warning = popup_warning.read()
            if event_warning == "Close" or event_warning == sg.WIN_CLOSED:
                break
        popup_warning.close()

    def read_excel(path: str, sheet_name: str) -> pd.DataFrame:
        buffer = StringIO()
        Xlsx2csv(path, outputencoding="utf-8", sheet_name=sheet_name).convert(buffer)
        buffer.seek(0)
        df = pd.read_csv(buffer)
        return df

    def open_window(loaded):
        status = loaded
        global img_list2, img_list1, column_name2
        sg.theme('LightBlue')
        original = [
            [sg.Text("Column Title")],
            [sg.InputText(key='column_name1')],
            [sg.Multiline(size=(60, 30), key="box1")]
        ]

        conflict = [
            [sg.Text("Column Title")],
            [sg.InputText(key='column_name2')],
            [sg.Multiline(size=(60, 30), key="box2")]
        ]
        original_frame = [[sg.Frame('Original Documents', layout=original, expand_x=True, expand_y=True, font="Arial 15")]]
        conflict_frame = [[sg.Frame('Conflict Documents', layout=conflict, expand_x=True, expand_y=True, font="Arial 15")]]
        buttons_frame = [[sg.Button('Load'), sg.Button('Exit')]]
        secondary_layout = [
            [sg.Column(original_frame, element_justification='c'),
             sg.Column(conflict_frame, element_justification='c')],
            [sg.Frame('Conflict Documents', layout=buttons_frame, expand_x=True, )]
        ]

        # layout2 = [
        #     [sg.Text('Path Lists')],
        #     [
        #         sg.Multiline(size=(60, 50), key="box1"),
        #         sg.Multiline(size=(60, 50), key="box2"),
        #     ],
        #     [sg.Frame([sg.Button('Load'), sg.Button('Exit')])],
        # ]
        window2 = sg.Window("Double Image Display", secondary_layout, modal=True)
        while True:
            lines1 = ''
            lines2 = ''
            event2, values2 = window2.read()
            if event2 == "Exit" or event2 == sg.WIN_CLOSED:
                break
            if event2 == "Load":
                img_list1 = []
                img_list2 = []
                for letter in values2["box1"]:
                    lines1 += letter
                    img_list1 = lines1.splitlines()

                for letter in values2["box2"]:
                    lines2 += letter
                    img_list2 = lines2.splitlines()

                # for line in values2["box2"]:
                #     img_list2.append(line)
                status = True
                break
        window2.close()
        return (img_list1, img_list2, status, values2["column_name1"], values2["column_name2"])

    loaded_files = False
    img_list1 = []
    img_list2 = []
    global_list = []
    current = 0
    header_row = []
    excel_conversion = False
    sg.theme('LightBlue')

    instinct_remarks = [
        'One Doc Already Expired',
        'Different pic on Doc',
        'Same Details ',
        'Different Ids of same person',
        'Front and Back'
    ]

    analysis_remarks = [
        'Poor Quality',
        'Bad Capture',
        'NonID ',
        'Unsupported',
        'Inconclusive',
        'Potential Bug',
        'Wrong Extraction',
        'Wrong Input',
        'No json',
    ]

    list1 = [
        [sg.Listbox(global_list, size=(20, 30), expand_x=True, enable_events=True, select_mode=sg.SELECT_MODE_SINGLE, key='document_list', change_submits=True)]
    ]

    image1 = [
        [sg.Text('Document ID', key="doc1", font="Arial 15", enable_events=True)],
        [sg.Image(background_color='gray', expand_x=True, expand_y=True, key='display_img1'), ]
    ]

    image2 = [
        [sg.Text('Document ID', key="doc2", font="Arial 15", enable_events=True)],
        [sg.Image(background_color='gray', expand_x=True, expand_y=True, key='display_img2')]
    ]

    controls = [
        [sg.Button('Load', auto_size_button=True, font="Arial 15"),
         sg.Button('Previous', auto_size_button=True, font="Arial 15"),
         sg.Button('Next', auto_size_button=True, font="Arial 15"),
         sg.Button('Exit', auto_size_button=True, font="Arial 15")]
    ]

    saved_info = [
        [sg.Text("Justified/Unjustified", key="just_unjust", font="Arial 15"),
         sg.Text("Notes", key="notes_text", font="Arial 15")],
        [sg.Text("More Details", key="moredetails_text", font="Arial 15")]
    ]

    excel_controls1 = [
        # [sg.FileBrowse('Browse', target='_FILEBROWSE_', file_types=(("ALL Files", "*.*"), ("ALL CSV Files", "*.csv"), ("ALL XLSX Files", "*.xlsx"),)),
        #  sg.Text('', key='_FILEBROWSE_', enable_events=True, font="Arial 10", expand_x=True),
        #  sg.Button('Load Excel', auto_size_button=True, font="Arial 10")],
        [sg.Button('Justified', auto_size_button=True, font="Arial 10")],
        [sg.Button('Unjustified', auto_size_button=True, font="Arial 10")],
    ]

    excel_controls2 = [
        [sg.Text('Notes', key="excel_file", font="Arial 15")],
        [sg.Button('Apply', key="N_Apply", auto_size_button=True, font="Arial 10"),
         sg.Combo(instinct_remarks, key="note_combo")]
    ]

    excel_controls3 = [
        [sg.Text('More Details', font="Arial 15")],
        [sg.Button('Apply', key="MD_Apply", auto_size_button=True, font="Arial 10"), sg.InputText(key='more_details', font="Arial 15")]
    ]

    load_excel = [
        [sg.FileBrowse(key='file_name', file_types=(("ALL Files", "*.*"), ("ALL CSV Files", "*.csv"), ("ALL XLSX Files", "*.xlsx"),)),
         sg.InputText(key='myfile'), ],
    ]

    extra_settings = [
        [sg.Button('Analysis', auto_size_button=True, font="Arial 10", key='analysis'),
         sg.Button('Instinct', auto_size_button=True, font="Arial 10", key='instinct'),
         sg.Button('Custom', auto_size_button=True, font="Arial 10", key='custom')]
    ]

    excel_controls_combine = [
        [sg.Column(load_excel, element_justification='center'),
         sg.Column([[sg.Button('Load Excel', auto_size_button=True, font="Arial 10"), sg.Button('Save Progress', auto_size_button=True, font="Arial 10")]], element_justification='center')],
        [sg.HSeparator()],

        # [sg.FileBrowse(key='file_name', file_types=(("ALL Files", "*.*"), ("ALL CSV Files", "*.csv"), ("ALL XLSX Files", "*.xlsx"),)),
        #  sg.InputText(key='myfile'),
        # sg.Button('Load Excel', auto_size_button=True, font="Arial 10")],
        [sg.Column(excel_controls1, vertical_alignment='bottom', key="excel_controls1", visible=False),
         sg.Column(excel_controls2, vertical_alignment='bottom', key="excel_controls2", visible=False),
         sg.Column(excel_controls3, vertical_alignment='bottom', key="excel_controls3", visible=False)]
    ]

    copyright = [
        [sg.Text("Ⓒ By Shahaf Stossel", text_color="#A20909")],
    ]

    column1 = [[sg.Frame('File List', layout=list1, font="Arial 15")]]

    column2 = [[sg.Frame('Main Image', layout=image1, expand_x=True, expand_y=True, size=(550, 550), font="Arial 15")]]

    column3 = [[sg.Frame('Conflict Image', layout=image2, expand_x=True, expand_y=True, size=(550, 550), font="Arial 15")]]

    bottom_controls = [
        [sg.Frame('Control Buttons', layout=controls, expand_x=True, expand_y=True, font="Arial 15")],
        [sg.Frame('Investigation Info', layout=saved_info, expand_x=True, expand_y=True, font="Arial 15", element_justification='c', key="saved-info", visible=False)]
    ]

    excel_controls_section = [[sg.Frame('Excel Control Panel', layout=excel_controls_combine, expand_x=True, expand_y=True, font="Arial 15")]]
    # Create layout with two columns using precreated frames
    layout = [
        [sg.Column(column1, element_justification='c'),
         sg.Column(column2, element_justification='c'),
         sg.Column(column3, element_justification='c')],
        [sg.Column(bottom_controls, element_justification='c'),
         sg.Column(excel_controls_section, element_justification='c')],
        [sg.Frame('Extra Settings', layout=extra_settings, expand_x=True, expand_y=True, font="Arial 15", key="extra_settings", visible=False)],
        [sg.HSeparator()],
        [sg.Push(), sg.Column(copyright), sg.Push()]
    ]

    window = sg.Window('Double Image Display', layout, icon=r"N:\Images\Shahaf\Projects\Assests\DoubleDisplay.ico", return_keyboard_events=True, use_default_focus=False, finalize=True, resizable=True)
    window['document_list'].bind("<Button-1>", "")
    window['document_list'].bind("<Up>", "_Up")
    window['document_list'].bind("<Down>", "_Down")
    window['display_img1'].bind("<Double-1>", "")
    window['display_img2'].bind("<Double-1>", "")
    window['doc1'].bind("<Double-1>", "")
    window['doc2'].bind("<Double-1>", "")
    # window['display_img1'].bind("<MouseWheel>", "_Wheel")
    # window['display_img2'].bind("<MouseWheel>", "_Wheel")
    window.TKroot.bind_all("<Key>", _onKeyRelease, "+")

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            if excel_conversion:
                df.to_csv(csv, index=False, encoding='utf-8')
            break
        if event == 'Load Excel':
            excel_conversion = True
            print(values['file_name'])
            file_path = values['file_name']
            window['myfile'].update(file_path[file_path.rfind("/") + 1:])
            file_name = file_path[file_path.rfind("/") + 1:file_path.rfind(".xlsx")]
            folder_path = file_path[:file_path.rfind("/")]
            csv = os.path.join(folder_path, "Instinct.csv")
            try:
                if '.xlsx' in file_path:
                    df = read_excel(file_path, file_name)
                else:
                    df = pd.read_csv(file_path)

                header_row = df.columns.values.tolist()
                if 'Manual Remarks' not in header_row:
                    df['Manual Remarks'] = pd.Series([], dtype=pd.StringDtype())
                    df = df.fillna(" ")

                if 'Notes' not in header_row:
                    df['Notes'] = pd.Series([], dtype=pd.StringDtype())
                    df = df.fillna(" ")

                if 'More Details' not in header_row:
                    df['More Details'] = pd.Series([], dtype=pd.StringDtype())

                    df = df.fillna(" ")

                df.to_csv(csv, index=False, encoding='utf-8')
            except IOError:
                warning()
                excel_conversion = False

            if excel_conversion:
                window["extra_settings"].update(visible=True)
                window["saved-info"].update(visible=True)
                window["excel_controls1"].update(visible=True)
                window["excel_controls2"].update(visible=True)
                window["excel_controls3"].update(visible=True)


        elif (event == 'Load') or (not loaded_files):
            zoom_level = 1
            zoom_increment = 0.1
            img_compare = []
            img_list1, img_list2, loaded_files, column_name1, column_name2 = open_window(loaded_files)
            if loaded_files:
                global_list = []
                for i in range(len(img_list1)):
                    global_list.append(f'Document_{i}')
                current = 0
                window["document_list"].update(values=global_list)
                window['document_list'].update(set_to_index=current)
                img_compare.append(img_list1[current])
                img_compare.append(img_list2[current])

                document1 = img_compare[0]
                document1 = document1[document1.rfind("\\") + 1:document1.rfind(".")]
                document1 = document1[:document1.find("_")]
                document2 = img_compare[1]
                document2 = document2[document2.rfind("\\") + 1:document2.rfind(".")]
                document2 = document2[:document2.find("_")]

                window["doc1"].update(document1)
                window["doc2"].update(document2)

                try:
                    image1 = Image.open(img_compare[0])
                    image1.thumbnail((550, 550))
                    bio1 = io.BytesIO()
                    image1.save(bio1, format="PNG")
                except FileNotFoundError:
                    image1 = Image.open(r'N:\Images\Shahaf\Tools\QCToolBox\Resc\NotFound.JPG')
                    image1.thumbnail((550, 550))
                    bio1 = io.BytesIO()
                    image1.save(bio1, format="PNG")

                try:
                    image2 = Image.open(img_compare[1])
                    image2.thumbnail((550, 550))
                    bio2 = io.BytesIO()
                    image2.save(bio2, format="PNG")
                except FileNotFoundError:
                    image2 = Image.open(r'N:\Images\Shahaf\Tools\QCToolBox\Resc\NotFound.JPG')
                    image2.thumbnail((550, 550))
                    bio2 = io.BytesIO()
                    image2.save(bio2, format="PNG")

                window["display_img1"].update(data=bio1.getvalue())
                window["display_img2"].update(data=bio2.getvalue())

        if loaded_files:
            if (event == 'document_list') or (event == 'Next') or (event == 'Previous') or (event == 'document_list_Down') or (event == 'document_list_Up') or (event == "display_img1") or (event == "display_img2") or (event == "doc1") or (event == "doc2"):
                img_compare = []
                if event == "doc1":
                    pyperclip.copy(document1)

                if event == "doc2":
                    pyperclip.copy(document2)

                if event == "display_img1":
                    im = Image.open(img_list1[current])
                    im.show()

                if event == "display_img2":
                    im = Image.open(img_list2[current])
                    im.show()

                if event == 'Next':
                    current += 1
                    window['document_list'].update(set_to_index=current)
                    window['document_list'].update(scroll_to_index=current)
                    zoom_level = 1
                    zoom_increment = 0.1

                if event == 'Previous':
                    current -= 1
                    window['document_list'].update(set_to_index=current)
                    window['document_list'].update(scroll_to_index=current)
                    zoom_level = 1
                    zoom_increment = 0.1

                if event == 'document_list':
                    current = window['document_list'].get_indexes()[0]
                    window['document_list'].update(set_to_index=current)

                # if event == 'display_img1_Wheel':
                #     print("wheel")
                #     zoom_image1 = sg.Image(filename=img_list1[current]).get_size()
                #     if values[event] == b"\x00":
                #         # scroll up to zoom in
                #         zoom_level += zoom_increment
                #     elif values[event] == b"\xff":
                #         # scroll down to zoom out
                #         zoom_level -= zoom_increment
                #     new_size = (int(zoom_image1[0] * zoom_level), int(zoom_image1[1] * zoom_level))
                #     window["display_img1"].update(filename=img_list1[current], size=new_size)
                #
                #
                # if event == 'display_img2_Wheel':
                #     print("wheel2")
                #     zoom_image2 = sg.Image(filename=img_list2[current]).get_size()
                #     if values[event] == b"\x00":
                #         # scroll up to zoom in
                #         zoom_level += zoom_increment
                #     elif values[event] == b"\xff":
                #         # scroll down to zoom out
                #         zoom_level -= zoom_increment
                #     new_size = (int(zoom_image2[0] * zoom_level), int(zoom_image2[1] * zoom_level))
                #     window["display_img2"].update(filename=img_list2[current], size=new_size)
                #
                if current < len(img_list1) - 1:
                    if event == 'document_list_Down':
                        zoom_level = 1
                        zoom_increment = 0.1

                        current = window['document_list'].get_indexes()[0]
                        current += 1
                        # update selected item in BOX2
                        window['document_list'].update(set_to_index=current)
                        window['document_list'].update(scroll_to_index=current)

                if current > 0:
                    if event == 'document_list_Up':
                        zoom_level = 1
                        zoom_increment = 0.1

                        current = window['document_list'].get_indexes()[0]
                        current -= 1
                        # update selected item in BOX2
                        window['document_list'].update(set_to_index=current)
                        window['document_list'].update(scroll_to_index=current)

                img_compare.append(img_list1[current])
                img_compare.append(img_list2[current])

                document1 = img_compare[0]
                document1 = document1[document1.rfind("\\") + 1:document1.rfind(".")]
                document1 = document1[:document1.find("_")]
                document2 = img_compare[1]
                document2 = document2[document2.rfind("\\") + 1:document2.rfind(".")]
                document2 = document2[:document2.find("_")]

                if excel_conversion:
                    justification = df.iloc[df[df['Selfie'] == document2].index[0], df.columns.get_loc('Manual Remarks')]
                    note = df.iloc[df[df['Selfie'] == document2].index[0], df.columns.get_loc('Notes')]
                    details = df.iloc[df[df['Selfie'] == document2].index[0], df.columns.get_loc('More Details')]

                    if len(justification) > 2:
                        window["just_unjust"].update(justification, background_color="#006600", text_color="White")
                    else:
                        window["just_unjust"].update("Justified/Unjustified", background_color="LightBlue", text_color="Black")
                    if len(note) > 2:
                        window["notes_text"].update(note, background_color="#006600", text_color="White")
                    else:
                        window["notes_text"].update("Notes", background_color="LightBlue", text_color="Black")
                    if len(details) > 2:
                        window["moredetails_text"].update(details, background_color="#006600", text_color="White")
                    else:
                        window["moredetails_text"].update("More Details", background_color="LightBlue", text_color="Black")

                window["doc1"].update(document1)
                window["doc2"].update(document2)

                try:
                    image1 = Image.open(img_compare[0])
                    image1.thumbnail((550, 550))
                    bio1 = io.BytesIO()
                    image1.save(bio1, format="PNG")
                except FileNotFoundError:
                    image1 = Image.open(r'N:\Images\Shahaf\Tools\QCToolBox\Resc\NotFound.JPG')
                    image1.thumbnail((550, 550))
                    bio1 = io.BytesIO()
                    image1.save(bio1, format="PNG")

                try:
                    image2 = Image.open(img_compare[1])
                    image2.thumbnail((550, 550))
                    bio2 = io.BytesIO()
                    image2.save(bio2, format="PNG")
                except FileNotFoundError:
                    image2 = Image.open(r'N:\Images\Shahaf\Tools\QCToolBox\Resc\NotFound.JPG')
                    image2.thumbnail((550, 550))
                    bio2 = io.BytesIO()
                    image2.save(bio2, format="PNG")

                window["display_img1"].update(data=bio1.getvalue())
                window["display_img2"].update(data=bio2.getvalue())
            if (event == "Justified") and excel_conversion:
                df.iloc[df[df['Conflicting DocId'] == document2].index[0], df.columns.get_loc('Manual Remarks')] = "Justified"
            if (event == "Unjustified") and excel_conversion:
                df.iloc[df[df['Conflicting DocId'] == document2].index[0], df.columns.get_loc('Manual Remarks')] = "Unjustified"
            if (event == "N_Apply") and excel_conversion:
                df.iloc[df[df['Conflicting DocId'] == document2].index[0], df.columns.get_loc('Notes')] = values['note_combo']
            if (event == "MD_Apply") and excel_conversion:
                df.iloc[df[df['Conflicting DocId'] == document2].index[0], df.columns.get_loc('More Details')] = values['more_details']

            if (event == "analysis") and excel_conversion:
                print(event)
                window['note_combo'].update(value=" ", values=analysis_remarks)

            if (event == "instinct") and excel_conversion:
                window['note_combo'].update(value=" ", values=instinct_remarks)

            if (event == 'Save Progress') and excel_conversion:
                df.to_csv(csv, index=False, encoding='utf-8')
    window.close()

ToolRun_DoubleDisplay()
