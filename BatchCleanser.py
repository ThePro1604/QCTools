import glob
import os
import shutil
import fitz
import PySimpleGUI as sg
from PIL import Image
from PrivateFunctions import folders
import traceback


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

    def BatchCleanser(folder, isJson, isPhoto, isAudio, isVideo, isCombined, isChange):
        def ChangeJPG(jpg_change):
            for jpg_file in jpg_list:
                os.rename(jpg_file, jpg_file[0:jpg_file.lower().index(".")] + ".jpeg")

        def ClearJson(json_list):
            for file in json_list:
                os.remove(file)

        def ClearPhoto(selfie_list):
            for file in selfie_list:
                os.remove(file)

        def ClearAudio(audio_list):
            for file in audio_list:
                os.remove(file)

        def ClearVideo(video_list):
            for file in video_list:
                os.remove(file)

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


        def get_list_of_json_files(folder):
            file_list = folders.file_list(folder, recursive=True, full_path=True)
            pdf_list = []
            png_list = []
            json_list = []
            jpg_list = []
            selfie_list = []
            audio_list = []
            video_list = []

            for file in file_list:
                if "photo" in file.lower():
                    selfie_list.append(file)
                elif file.lower().endswith(".pdf"):
                    pdf_list.append(file)
                elif file.lower().endswith(".png"):
                    png_list.append(file)
                elif file.lower().endswith(".json") or file.lower().endswith(".xml") or file.lower().endswith(".txt"):
                    json_list.append(file)
                elif file.lower().endswith(".jpg"):
                    jpg_list.append(file)
                elif file.lower().endswith(".wav"):
                    audio_list.append(file)
                elif file.lower().endswith(".webm") or file.lower().endswith(".mp4") or file.lower().endswith(".mpeg") or file.lower().endswith(".mov"):
                    video_list.append(file)

            return file_list, selfie_list, pdf_list, png_list, json_list, jpg_list, audio_list, video_list

        with open("log.txt", "w") as log:
            try:
                list_of_files, selfie_list, pdf_list, png_list, json_list, jpg_list, audio_list, video_list = get_list_of_json_files(folder)
                if len(pdf_list) > 0:
                    for pdf_file in pdf_list:
                        try:
                            doc = fitz.open(pdf_file)  # open document
                            i = 0
                            for page in doc:
                                if "page" in pdf_file:
                                    name = pdf_file[0:pdf_file.lower().index("page")]
                                    if os.path.isfile(f"{name}page0.jpeg"):
                                        pix = page.get_pixmap()  # render page to an image
                                        pix.save(f"{name}page1.jpeg", 'JPEG')
                                    else:
                                        pix = page.get_pixmap()  # render page to an image
                                        pix.save(f"{name}page{i}.jpeg", 'JPEG')
                                        i += 1
                                elif "supp" in pdf_file:
                                    name = pdf_file[0:pdf_file.lower().index("supp")]
                                    pix = page.get_pixmap()  # render page to an image
                                    pix.save(f"{name}supp{i}.jpeg", 'JPEG')
                                    i += 1
                                else:
                                    name = pdf_file[0:pdf_file.lower().index(".pdf")]
                                    pix = page.get_pixmap()  # render page to an image
                                    pix.save(f"{name}_{i}.jpeg", 'JPEG')
                                    i += 1

                            doc.close()
                            os.remove(pdf_file)
                        except:
                            continue

                if len(png_list) > 0:
                    for png_file in png_list:
                        png2jpeg = png_file[0:png_file.lower().rfind(".png")] + ".jpeg"
                        image = Image.open(png_file)
                        image_rgb = image.convert("RGB")
                        image_rgb.save(png2jpeg, "JPEG", quality=90)
                        image.close()
                        os.remove(png_file)

                if isChange and len(jpg_list) > 0:
                    ChangeJPG(jpg_list)

                if isJson and len(json_list) > 0:
                    ClearJson(json_list)

                if isPhoto and len(selfie_list) > 0:
                    ClearPhoto(selfie_list)

                if isAudio and len(audio_list) > 0:
                    ClearAudio(audio_list)

                if isVideo and len(video_list) > 0:
                    ClearVideo(video_list)

                if isCombined:
                    CombineFolders(folder)
            except Exception:
                traceback.print_exc(file=log)

        sg.popup("DONE!")

    top = [
        [sg.Text("Folder")],
        [sg.FolderBrowse(key="folder_name"), sg.InputText(key='myfolder')],
        [sg.Checkbox("Combine All to One Folder", key="-COMBINE-", default=False)],
        [sg.Checkbox("JPG -> JPEG", key="-JPG-", default=True)],
        [sg.Text("Choose Extensions To Delete and")],
        [sg.Checkbox("Json", key="-JSON-", default=True), sg.Checkbox("Photo", key="-PHOTO-", default=True), sg.Checkbox("Video", key="-VIDEO-"), sg.Checkbox("Audio", key="-AUDIO-")]
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
            BatchCleanser(values['myfolder'], values["-JSON-"], values["-PHOTO-"], values["-AUDIO-"], values["-VIDEO-"], values["-COMBINE-"], values["-JPG-"])
        if event == sg.WIN_CLOSED or event == "Close":
            break

    window.close()


ToolRun_BatchCleanser()
