import os
import shutil


def verify_folder(path, create=False, delete=False, response=True):
    status = os.path.exists(path)
    if create and not status:
        os.makedirs(path)
    if delete and status:
        shutil.rmtree(path)
    return status


def file_list(path, file_array=None, recursive=True, full_path=True):
    if file_array is None:
        file_array = []
    file_names = os.listdir(path)
    for file_name in file_names:
        file_path = os.path.join(path, file_name)
        if recursive and os.path.isdir(file_path):
            file_list(file_path, file_array=file_array, recursive=True, full_path=True)
        elif full_path:
            file_array.append(file_path)
        else:
            file_array.append(file_name)
    return file_array


def remove_empty(folder_path):
    if os.path.exists(folder_path):
        for root, dirs, files in os.walk(folder_path, topdown=False):
            for name in dirs:
                dir_path = os.path.join(root, name)
                if not os.listdir(dir_path): # check if directory is empty
                    print(dir_path)
                    os.rmdir(dir_path) # delete empty directory




