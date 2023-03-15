import win32com
import os
import shutil

def remove_win32genpy() :

    path = win32com.__gen_path__

    os.chdir(path)

    listing = os.listdir()

    first_element = listing[0]
    # print(first_element)

    full_path = f'{path}\{first_element}'

    try:
        shutil.rmtree(full_path)
        print(f"Directory '{full_path}' has been successfully deleted along with its contents.")
    except OSError as e:
        print(f"Error: {e.filename} - {e.strerror}.")
