import os
import win32com.client
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
import time
import tkinter.messagebox as messagebox

# Specify your macro name and part file path
macro_name = r".\Envío de piezas a Odoo\main.swp"
# "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe" "C:\Users\Usuario\Downloads\08978 Puerta Tablero.SLDPRT"

sldprt_files = []
sldasm_files = []
swApp = None

def run_solidworks_macro(swApp, macro_name):
    try:
        # Connect to SolidWorks
        swApp.Visible = False

        # Open the file
        #swModel = swApp.OpenDoc(part_file_path, 1)  # 1 = swDocumentPart

        # Run your VBA macro
        macro_full_path = os.path.join(os.path.expanduser("~"), "AppData\\Roaming\\SolidWorks\\SolidWorks 2019\\macros", macro_name)
        macro_full_path = r"C:\Users\Usuario\Documents\Pedro\Solid Module\Envío de piezas a Odoo\main.swp"
        swApp.RunMacro(macro_full_path, "main1", "main1")

        # Close the SolidWorks document
        #swApp.CloseDoc(swModel.GetTitle())

    except Exception as e:
        print(f"Error: {str(e)}")
        messagebox.showerror("SolidWorks Error", str(e))
        return
        
def get_text_file_content(file_name):
            file_path = os.path.join(os.getcwd(), "Envío de piezas a Odoo", file_name + ".txt")
            with open(file_path, 'r') as file:
                content = file.read()
            return content

def clean_text_file_content(file_name):
    file_path = os.path.join(os.getcwd(), "Envío de piezas a Odoo", file_name + ".txt")
    with open(file_path, 'w') as file:
        file.write('')

def clean_data_files():

    clean_text_file_content("Masa")
    clean_text_file_content("Volumen")
    clean_text_file_content("Superficie")
    clean_text_file_content("Ancho")
    clean_text_file_content("Largo")
    clean_text_file_content("Grosor")
    clean_text_file_content("Error")

def ensamble_odoo(file, masa, volumen, superficie, url):
    print("Ensamble: ", file)
    print(masa, "Kg", volumen, "mm3", superficie, "mm2")

def pieza_odoo(file, masa, volumen, superficie, ancho, largo, grosor, url):
    print("Pieza: ", file)
    print(masa, "Kg", volumen, "mm3", superficie, "mm2")
    print("Ancho:", ancho, "mm. Largo:", largo, "mm Espesor:", grosor, "mm.")

def process_sldasm(sldasm_files, folder_path):

    #escribir la ruta en el archivo input
        path_file = r".\Envío de piezas a Odoo\Ruta.txt"
        sldasm_file_path = os.path.join(folder_path, sldasm_files[0])
        sldasm_file_path = sldasm_file_path.replace("\\", "/")  # Replace backslashes with forward slashes
        
        # Clean the file in path_file and write the sldasm file there
        with open(path_file, 'w') as file:
            file.write(sldasm_file_path)

        #sldasm_file_path = r"C:\Users\Usuario\Downloads\04955 GAB-PEX-11\04955 GAB-PEX-11-B V2.2 ENSAMBLAJE.SLDASM"

        #clean data files
        clean_data_files()
        
        #correr el macros
        """
        el macros va a:
        -buscar la ruta del archivo en Ruta.txt
        -abrir el archivo
        -obtener y guardar los datos de masa en los archivos
        -obtener el bounding box
        -cerrar el archivo
        """
        run_solidworks_macro(swApp, macro_name)

        #abrir el archivo de error

        error_text = get_text_file_content("Error")
        if error_text:
            if error_text:
                messagebox.showerror("SolidWorks Error", error_text)
                return

        #recopilar los datos guardados

        masa = get_text_file_content("Masa").strip()
        volumen = get_text_file_content("Volumen").strip()
        superficie = get_text_file_content("Superficie").strip()
        #ancho = get_text_file_content("Ancho")
        #largo = get_text_file_content("Largo")
        #grosor = get_text_file_content("Grosor")

        #generar url ruta
        sldasm_file_path_url = sldasm_file_path.replace(" ", "%20")
        sldasm_file_path_url = "file:///" + sldasm_file_path_url  # Update URL format
        #enviar los datos a odoo
        #no enviar grosor para los ensambles

        ensamble_odoo(sldasm_files[0], masa, volumen, superficie, sldasm_file_path_url)

        #renombrar la carpeta con el codigo del ensamble

def ordenar_valores (ancho, largo, grosor):

    #turn into float, fst strip, then replace comma with dot
    ancho = float(ancho.strip().replace(",", "."))
    largo = float(largo.strip().replace(",", "."))
    grosor = float(grosor.strip().replace(",", "."))
     
    if ancho > largo:
        aux = ancho
        ancho = largo
        largo = aux

    if largo < grosor:
        aux = largo
        largo = grosor
        grosor = aux

    if ancho < grosor:
        aux = ancho
        ancho = grosor
        grosor = aux

    return ancho, largo, grosor

def process_sldprt(sldprt_file, folder_path):

    #escribir la ruta en el archivo input
        path_file = r".\Envío de piezas a Odoo\Ruta.txt"
        sldprt_file_path = os.path.join(folder_path, sldprt_file)
        sldprt_file_path = sldprt_file_path.replace("\\", "/")

        with open(path_file, 'w') as file:
            file.write(sldprt_file_path)

        #clean files
        clean_data_files()

        #correr el macros
        """
        el macros va a:
        -buscar la ruta del archivo en Ruta.txt
        -abrir el archivo
        -obtener y guardar los datos de masa en los archivos
        -aplanar la pieza
        -obtener el bounding box
        -cerrar el archivo
        """
        run_solidworks_macro(swApp, macro_name)

        #abrir el archivo de error

        error_text = get_text_file_content("Error")
        if error_text:
            if error_text:
                messagebox.showerror("SolidWorks Error", error_text)
                return

        #recopilar los datos guardados

        masa = get_text_file_content("Masa").strip()
        volumen = get_text_file_content("Volumen").strip()
        superficie = get_text_file_content("Superficie").strip()
        ancho = get_text_file_content("Ancho").strip()
        largo = get_text_file_content("Largo").strip()
        grosor = get_text_file_content("Grosor").strip()

        #generar url ruta
        sldprt_file_path_url = sldprt_file_path.replace(" ", "%20")
        sldprt_file_path_url = "file:///" + sldprt_file_path_url  # Update URL format

        #codigo del ensamble parent

        #ordenar los valores
        ancho, largo, grosor = ordenar_valores(ancho, largo, grosor)

        #enviar los datos a odoo
        pieza_odoo(sldprt_file, masa, volumen, superficie, ancho, largo, grosor, sldprt_file_path_url)

def folder(folder_path):

    global swApp

     # Connect to an existing SolidWorks instance or create a new one if not available
    try:
        swApp = win32com.client.GetObject("SldWorks.Application")
    except:

        swApp = win32com.client.Dispatch("SldWorks.Application")

        """print("SolidWorks no está abierto. Abriendo SolidWorks...")
        swApp = win32com.client.Dispatch("SldWorks.Application")
        start_time = time.time()

        while True:
            try:
                swApp = win32com.client.GetObject("SldWorks.Application")
                break
            except:
                time.sleep(1)

                # Check if 1 minute has passed
                elapsed_time = time.time() - start_time
                if elapsed_time > 60:
                    messagebox.showerror("SolidWorks Error", "SolidWorks no pudo ser iniciado.")
                    return
                    break
                continue"""

    file_names = os.listdir(folder_path)

    global sldasm_files
    global sldprt_files
            
    sldprt_files = [file_name for file_name in file_names if file_name.endswith('.SLDPRT')]
    sldasm_files = [file_name for file_name in file_names if file_name.endswith('.SLDASM')]

    #check if there is a sldasm file
    if sldasm_files:
        process_sldasm(sldasm_files, folder_path)
           
    #procesar cada pieza sldprt
    for sldprt_file in sldprt_files:
        process_sldprt(sldprt_file, folder_path)

    #finish program
    messagebox.showinfo("SolidWorks", "Proceso finalizado.")
    print("Proceso finalizado.")
    return

    return

# folder(r"C:\Users\Usuario\Downloads\04955 GAB-PEX-11")

