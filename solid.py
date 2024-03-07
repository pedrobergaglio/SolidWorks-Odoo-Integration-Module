import requests
import json
import os
import win32com.client
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
import time
import tkinter.messagebox as messagebox
import pandas as pd
import os

# Specify your macro name and part file path
macro_name = r".\Envío de piezas a Odoo\main.swp"
# "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe" "C:\Users\Usuario\Downloads\08978 Puerta Tablero.SLDPRT"

sldprt_files = []
sldasm_files = []
swApp = None

ensamble = {}
piezas = []

#importar referencias
espesores = pd.read_excel(r".\resources\espesores.xlsx")
insumos_piezas = pd.read_excel(r".\resources\insumos-piezas.xlsx")
peso_especifico = pd.read_excel(r".\resources\peso-especifico.xlsx")

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

def ensamble_odoo(ensamble, folder_path):
    #print("Ensamble: ", file)
    #print(masa, "Kg", volumen, "mm3", superficie, "mm2")

    #send request to odoo
    url = "http://localhost:8069"
    db = "odoo"
    username = "admin"
    password = "admin"
    
    #do request

    # Convert data to JSON format
    json_data = json.dumps(ensamble)

    # Send POST request
    response = requests.post(url, auth=(username, password), data=json_data)

    # Check response
    if response.status_code != 200:
        print(f"Request failed. Status code: {response.status_code}")
        return

    # Get response data
    response_data = response.json()
    ensamble["id"] = response_data.pop("id")

    # Split the folder path
    folder_path_parts = folder_path.split(os.sep)
    
    # Edit the last part of the folder path
    folder_path_parts[-1] = ensamble["id"]+ " " + folder_path_parts[-1]
    
    # Join the parts back together
    new_folder_path = os.sep.join(folder_path_parts)

    os.rename(folder_path, new_folder_path)




def insert_pieza_odoo(pieza):
    print("Pieza: ", pieza.name)
    print(pieza.weight, "Kg", pieza.volume, "mm3", pieza.surface, "mm2")
    #print("Ancho:", ancho, "mm. Largo:", largo, "mm Espesor:", grosor, "mm.")

    #send request to odoo

    url = "http://localhost:8069"
    db = "odoo"
    username = "admin"
    password = "admin"
    
    #do request

    # Convert data to JSON format
    json_data = json.dumps(pieza)

    # Send POST request
    response = requests.post(url, auth=(username, password), data=json_data)

    # Check response
    if response.status_code != 200:
        print(f"Request failed. Status code: {response.status_code}")
        return

    # Get response data
    response_data = response.json()
    pieza["id"] = response_data.pop("id")


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
                messagebox.showerror("Error de SolidWorks en el Ensamblaje", error_text)
                return

        #recopilar los datos guardados
        volumen = get_text_file_content("Volumen").strip()
        superficie = get_text_file_content("Superficie").strip()

        #obtener tag_id a partir del nombre del archivo
        #calcular masa a partir del peso específico
        material = sldasm_files.split(" ")[0]
        try:
            material_tag = peso_especifico[peso_especifico["REFERENCIA"] == material]["TAG"]
            if material_tag == "" or material_tag == None:
                messagebox.showerror("Error", f"Error al encontrar el tag del material para archivo {sldasm_files}, por favor verifique que tiene asignado un valor correcto en TAG.")
                return
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al encontrar la referencia al material en el archivo {sldasm_files}, por favor verifique que la referencia es correcta: {str(e)}")
            return

        #calcular masa a partir del peso de las piezas
        global ensamble
        global piezas
        #iterar sobre las piezas y sumar el net_weight y el gross_weight para obtener ambos valores totales
        net_weight = 0
        gross_weight = 0
        ids = []
        for pieza in piezas:
            net_weight += pieza["weight"]
            gross_weight += pieza["gross_weight"]
            ids.append(pieza["id"])
        
        #save everything in a dictionary
            
        ensamble = {
            "name": sldasm_files[0].split(".")[0],
            "product_tag_ids": "Conjunto",
            "weight": net_weight,
            "gross_weight": gross_weight,
            "volume": volumen,
            "surface": superficie,
            "categ_id": material_tag,
            "sale_ok": "true",
            "purchase_ok": "false",
            "product_route": "Fabricar",
            "tracking": "N° de CNC"
        }


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
        volumen = get_text_file_content("Volumen").strip()
        superficie = get_text_file_content("Superficie").strip()
        ancho = get_text_file_content("Ancho").strip()
        largo = get_text_file_content("Largo").strip()
        espesor = get_text_file_content("Grosor").strip()

        #calcular masa a partir del peso específico
        material = sldprt_file.split(" ")[0]
        try:
            peso_especifico_value = float(peso_especifico[peso_especifico["REFERENCIA"] == material]["VALOR"]) / 1

            if peso_especifico_value == 0 or peso_especifico_value == None:
                messagebox.showerror("Error", f"Error al encontrar el peso específico del material en el archivo {sldprt_file}, por favor verifique que tiene asignado un valor correcto en VALOR.")
                return

            net_weight = float(volumen) * peso_especifico_value
            gross_weight = float(espesor) * float(ancho) * float(largo) * peso_especifico_value

            
        except Exception as e:
            messagebox.showerror("Error", f"Error al encontrar la referencia al material en el archivo {sldprt_file}, por favor verifique que la referencia es correcta: {str(e)}")
            return

        try:
            material_tag = peso_especifico[peso_especifico["REFERENCIA"] == material]["TAG"]

            if material_tag == "" or material_tag == None:
                messagebox.showerror("Error", f"Error al encontrar el tag del material para archivo {sldprt_file}, por favor verifique que tiene asignado un valor correcto en TAG.")
                return
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al encontrar la referencia al material en el archivo {sldprt_file}, por favor verifique que la referencia es correcta: {str(e)}")
            return

        #ordenar los valores
        ancho, largo, espesor = ordenar_valores(ancho, largo, espesor)

        #calcular espesor
        espesor_values = espesores['ESPESOR'].values
        # Find the closest value in espesor_values to espesor
        espesor_found = min(espesor_values, key=lambda x:abs(float(x)-float(espesor)))
        #get the value in col STRING in the same row as espesor_found
        espesor_string = espesores[espesores['ESPESOR'] == espesor_found]['STRING']

        #seleccion de insumo para la pieza
        insumo = insumos_piezas[(insumos_piezas["ESPESOR"] == espesor_string) & (insumos_piezas["MATERIAL"] == material_tag)]["INSUMO"]
        
        #save everything in a dictionary
        global piezas

        piezas.append({
            "name": sldprt_file.split(".")[0],
            "product_tag_ids": "Piezas",
            "weight": net_weight,
            "gross_weight": gross_weight,
            "volume": volumen,
            "surface": superficie,
            "broad": ancho,
            "long": largo,
            "categ_id": material_tag,
            "thickness": espesor_string,
            "sale_ok": "true",
            "purchase_ok": "false",
            "product_route": "Fabricar",
            "tracking": "N° de CNC"
        })

#main donde inicia el procesamiento de la carpeta
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
    
    global ensamble
    global piezas

    for pieza in piezas:
        insert_pieza_odoo(pieza)

    ensamble_odoo(ensamble, folder_path)



    #finish program
    messagebox.showinfo("SolidWorks", "Proceso finalizado.")
    print("Proceso finalizado.")
    return

    return

# folder(r"C:\Users\Usuario\Downloads\04955 GAB-PEX-11")

