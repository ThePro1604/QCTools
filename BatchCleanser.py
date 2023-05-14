import glob
import os
import shutil

import PIL
import fitz
import PySimpleGUI as sg
import magic
from PIL import Image
import pathlib


def ToolRun_BatchCleanser():
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


    def BatchCleanser(folder, isJson, isPhoto, isAudio, isVideo, isCombined):
        def ClearJson(folder):
            list_of_files = get_list_of_json_files()
            for file in list_of_files:
                file_fullpath = os.path.join(folder, file)
                if "." not in file:
                    nested_folder = file
                    list_of_files_nested = get_list_of_json_files_nested(file)
                    for nested_file in list_of_files_nested:
                        nested_file_fullpath = os.path.join(folder, nested_folder, nested_file)
                        if ".json" in nested_file:
                            os.remove(nested_file_fullpath)
                else:
                    if ".json" in file:
                        os.remove(file_fullpath)

        def ClearPhoto(folder):
            list_of_files = get_list_of_json_files()
            for file in list_of_files:
                file_fullpath = os.path.join(folder, file)
                if "." not in file:
                    nested_folder = file
                    list_of_files_nested = get_list_of_json_files_nested(file)
                    for nested_file in list_of_files_nested:
                        nested_file_fullpath = os.path.join(folder, nested_folder, nested_file)
                        if "photo" in nested_file.lower():
                            os.remove(nested_file_fullpath)
                else:
                    if "photo" in file.lower():
                        os.remove(file_fullpath)

        def ClearAudio(folder):
            list_of_files = get_list_of_json_files()
            for file in list_of_files:
                file_fullpath = os.path.join(folder, file)
                if "." not in file:

                    nested_folder = file
                    list_of_files_nested = get_list_of_json_files_nested(file)
                    for nested_file in list_of_files_nested:
                        nested_file_fullpath = os.path.join(folder, nested_folder, nested_file)
                        if "." in nested_file:
                            mime = magic.Magic(mime=True)
                            nested_filename = mime.from_file(nested_file_fullpath)
                            if nested_filename.find('audio') != -1:
                                os.remove(nested_file_fullpath)

                else:
                    mime = magic.Magic(mime=True)
                    filename = mime.from_file(file_fullpath)
                    if filename.find('audio') != -1:
                        os.remove(file_fullpath)

        def ClearVideo(folder):
            list_of_files = get_list_of_json_files()
            for file in list_of_files:
                file_fullpath = os.path.join(folder, file)
                if "." not in file:
                    nested_folder = file
                    list_of_files_nested = get_list_of_json_files_nested(file)
                    for nested_file in list_of_files_nested:
                        nested_file_fullpath = os.path.join(folder, nested_folder, nested_file)
                        if "." in nested_file:
                            mime = magic.Magic(mime=True)
                            nested_filename = mime.from_file(nested_file_fullpath)
                            if nested_filename.find('video') != -1:
                                os.remove(nested_file_fullpath)
                else:
                    mime = magic.Magic(mime=True)
                    filename = mime.from_file(file_fullpath)
                    if filename.startswith('video/') != -1:
                        os.remove(file_fullpath)

        def CombineFolders(folder):
            file_list = []
            path = folder + '\\**\\*.*'
            for filepath in glob.glob(path, recursive=True):
                file_list.append(filepath)
                try:
                    shutil.move(filepath, folder)
                except shutil.Error:
                    continue
                except PermissionError:
                    os.system('cmd /c "attrib -r -h -s {path}"'.format(path=filepath))
                    shutil.move(filepath, folder)

            print("-------------------------")
            for f in os.listdir(folder):
                if os.path.isdir(os.path.join(folder, f)):
                    try:
                        os.remove(os.path.join(folder, f))
                    except PermissionError:
                        os.system('cmd /c "attrib -r -h -s "{path}""'.format(path=os.path.join(folder, f)))
                        shutil.rmtree(os.path.join(folder, f))


        def get_list_of_json_files():
            list_of_files = os.listdir(folder)
            return list_of_files

        def get_list_of_json_files_nested(nested):
            list_of_files = os.listdir(folder + "\\" + nested)
            return list_of_files


        list_of_files = get_list_of_json_files()
        for file in list_of_files:
            file_fullpath = os.path.join(folder, file)
            if ".jpeg" in file:
                continue
            if "." not in file:
                nested_folder = file
                list_of_files_nested = get_list_of_json_files_nested(file)
                for nested_file in list_of_files_nested:
                    nested_file_fullpath = os.path.join(folder, nested_folder, nested_file)
                    if ".jpeg" in nested_file:
                        continue
                    if ".pdf" in nested_file:
                        try:
                            doc = fitz.open(nested_file_fullpath)  # open document
                            i = 0
                            for page in doc:
                                if "page" in nested_file:
                                    name = nested_file[0:nested_file.lower().index("page")]
                                    if os.path.isfile(f"{folder}/{nested_folder}/{name}page0.jpeg"):
                                        pix = page.get_pixmap()  # render page to an image
                                        pix.save(f"{folder}/{nested_folder}/{name}page1.jpeg", 'JPEG')
                                    else:
                                        pix = page.get_pixmap()  # render page to an image
                                        pix.save(f"{folder}/{nested_folder}/{name}page{i}.jpeg", 'JPEG')
                                        i += 1
                                elif "supp" in nested_file:
                                    name = nested_file[0:nested_file.lower().index("supp")]
                                    pix = page.get_pixmap()  # render page to an image
                                    pix.save(f"{folder}/{nested_folder}/{name}supp{i}.jpeg", 'JPEG')
                                    i += 1
                                else:
                                    name = nested_file[0:nested_file.lower().index(".pdf")]
                                    pix = page.get_pixmap()  # render page to an image
                                    pix.save(f"{folder}/{nested_folder}/{name}_{i}.jpeg", 'JPEG')
                                    i += 1

                            doc.close()
                            os.remove(nested_file_fullpath)
                        except:
                            continue
                    elif ".png" in nested_file.lower():
                        try:
                            # try:
                            #     doc = Image.open(folder + "/" + nested_folder + "/" + nested_file)  # open document
                            #     doc.save(folder + "/" + nested_folder + "/" + nested_file[0:nested_file.lower().index(".")] + ".jpeg")
                            #     os.remove(folder + "/" + nested_folder + "/" + nested_file)
                            # except OSError:
                            doc = Image.open(nested_file_fullpath)  # open document
                            rgb_doc = doc.convert('RGB')
                            rgb_doc.save(nested_file_fullpath[0:nested_file.lower().index(".")] + ".jpeg")
                            os.remove(nested_file_fullpath)
                        except PIL.UnidentifiedImageError:
                            continue
                    elif ".jpg" in nested_file.lower():
                        try:
                            # doc = Image.open(folder + "/" + nested_folder + "/" + nested_file)  # open document
                            # doc.save(folder + "/" + nested_folder + "/" + nested_file[0:nested_file.lower().index(".")] + ".jpeg")
                            # os.remove(folder + "/" + nested_folder + "/" + nested_file)
                            os.rename(nested_file_fullpath, nested_file_fullpath[0:nested_file.lower().index(".")] + ".jpeg")

                        except:
                            continue
            elif ".pdf" in file:
                try:
                    doc = fitz.open(file_fullpath)  # open document
                    i = 0
                    for page in doc:
                        if "page" in file:
                            name = file[0:file.lower().index("page")]
                            if os.path.isfile(f"{folder}/{name}page0.jpeg"):
                                pix = page.get_pixmap()  # render page to an image
                                pix.save(f"{folder}/{name}page1.jpeg", 'JPEG')
                            else:
                                pix = page.get_pixmap()  # render page to an image
                                pix.save(f"{folder}/{name}page{i}.jpeg", 'JPEG')
                                i += 1
                        elif "supp" in file:
                            name = file[0:file.lower().index("supp")]
                            pix = page.get_pixmap()  # render page to an image
                            pix.save(f"{folder}/{name}supp{i}.jpeg", 'JPEG')
                            i += 1
                        else:
                            name = file[0:file.lower().index(".pdf")]
                            pix = page.get_pixmap()  # render page to an image
                            pix.save(f"{folder}/{name}_{i}.jpeg", 'JPEG')
                            i += 1

                    doc.close()
                    os.remove(file_fullpath)
                except:
                    continue
            elif ".png" in file.lower():
                print("png " + file)
                try:
                    doc = Image.open(file_fullpath)  # open document
                    rgb_doc = doc.convert('RGB')
                    rgb_doc.save(file_fullpath[0:file.lower().index(".")] + ".jpeg")
                    doc.close()
                    os.remove(file_fullpath)
                except PermissionError:
                    os.chmod(file_fullpath, 0o777)  # change file permissions
                    os.remove(file_fullpath)

                    # os.remove(folder + "\\" + file)
            elif ".jpg" in file.lower():
                print("jpg " + file)
                try:
                    os.rename(file_fullpath, file_fullpath[0:file.lower().index(".")] + ".jpeg")
                    # doc = Image.open(folder + "/" + file)  # open document
                    # doc.save(folder + "/" + file[0:file.lower().index(".")] + ".jpeg")
                    # os.remove(folder + "/" + file)
                except:
                    continue

        if isJson:
            ClearJson(folder)

        if isPhoto:
            ClearPhoto(folder)

        if isAudio:
            ClearAudio(folder)

        if isVideo:
            ClearVideo(folder)

        if isCombined:
            CombineFolders(folder)

        layout_popup = [
            [sg.Text("Done!")],
            [sg.Button('Close')]
        ]
        popup = sg.Window("Complete", layout_popup, icon=r'N:\Images\Shahaf\Projects\Assests\pdf2jpg.ico')

        while True:
            event, values = popup.read()
            if event == "Close" or event == sg.WIN_CLOSED:
                break
        popup.close()

    top = [
        [sg.Text("Folder")],
        [sg.FolderBrowse(key="folder_name"), sg.InputText(key='myfolder')],
        [sg.Checkbox("Combine All to One Folder", key="-COMBINE-", default=False)],
        [sg.Text("Choose Extensions To Delete and")],
        [sg.Checkbox("Json", key="-JSON-"), sg.Checkbox("Photo", key="-PHOTO-"), sg.Checkbox("Video", key="-VIDEO-"), sg.Checkbox("Audio", key="-AUDIO-")]
    ]

    buttons = [
        [sg.Button('Start'), sg.Button('Close')],
    ]

    copyright = [
        [sg.Text("Ⓒ By Shahaf Stossel", text_color="#A20909")],
    ]

    description = [
        [sg.Text("Prepare Folder For Working Process")],
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
    window = sg.Window("BatchCleanser", layout, icon=r'N:\Images\Shahaf\Projects\Assests\BatchCleanser.ico', finalize=True)
    window.TKroot.bind_all("<Key>", _onKeyRelease, "+")

    while True:
        event, values = window.read()
        if event == "Start" and len(values['myfolder']) > 1:
            # if values["-COMBINE-"]:
            #     # os.makedirs(values['myfolder'] + "/Combine")
            BatchCleanser(values['myfolder'], values["-JSON-"], values["-PHOTO-"], values["-AUDIO-"], values["-VIDEO-"], values["-COMBINE-"])
        if event == sg.WIN_CLOSED or event == "Close":
            break

    window.close()


ToolRun_BatchCleanser()
