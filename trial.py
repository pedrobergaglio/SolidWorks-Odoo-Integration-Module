import os
import datetime
from tkinter import messagebox

def rename_folder(folder_path, ensamble_name):
    try:
        new_folder_path = " ".join([folder_path, ensamble_name])
        os.rename(folder_path, new_folder_path)
        folder_path = new_folder_path
        print(f'Renamed folder to {folder_path}')
    except Exception as e:
        messagebox.showerror("Error al renombrar archivo", f"No se pudo renombrar la carpeta{folder_path}. Error: {str(e)}")
        print(datetime.datetime.now(), f"Error al renombrar la carpeta{folder_path}. Error: {str(e)}")


# Test the rename_folder function
folder_path = "/path/to/folder"
ensamble_name = "new_folder_name"
rename_folder(folder_path, ensamble_name)