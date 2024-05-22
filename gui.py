import time
import os
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image, ImageTk
import socket
import tkinter.messagebox as messagebox
import requests
import json
#import win32com.client
import pandas as pd
import random
import sys
import datetime

# Open the log file in append mode
log_file = open(r"C:\SolidWorks Data\Envío de piezas a Odoo\logfile.log", 'a')

#sys.stdout = log_file
#sys.stderr = log_file

# Specify your macro name agui.pynd part file path
macro_name = r"C:\SolidWorks Data\Envío de piezas a Odoo\main.swp"
# "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe" "C:\Users\Usuario\Downloads\08978 Puerta Tablero.SLDPRT"

sldprt_files = []
sldasm_files = []
swApp = None

ensamble = {}
piezas = []

error = False

new_folder_path = ""
folder_path = ""

#importar referencias
espesores = pd.read_excel("espesores.xlsx")
#print(espesores.head())
insumos_piezas = pd.read_excel("insumos-piezas.xlsx")
#print(insumos_piezas.head())
peso_especifico = pd.read_excel("peso-especifico.xlsx")
#print(peso_especifico.head())

create_url = "http://ec2-3-15-193-242.us-east-2.compute.amazonaws.com:8069/itec-api/create/product"
update_url = "http://ec2-3-15-193-242.us-east-2.compute.amazonaws.com:8069/itec-api/update/product"

def run_solidworks_macro(swApp, macro_name):
    global error
    try:
        # Connect to SolidWorks
        swApp.Visible = False

        # Open the file
        #swModel = swApp.OpenDoc(part_file_path, 1)  # 1 = swDocumentPart

        # Run your VBA macro
        #macro_full_path = os.path.join(os.path.expanduser("~"), "AppData\\Roaming\\SolidWorks\\SolidWorks 2019\\macros", macro_name)
        macro_full_path = r"C:\SolidWorks Data\Envío de piezas a Odoo\main.swp"
        print(datetime.datetime.now(), macro_full_path)
        swApp.RunMacro(macro_full_path, "main1", "main1")

        # Close the SolidWorks document
        #swApp.CloseDoc(swModel.GetTitle())

    except Exception as e:
         
        print(datetime.datetime.now(), f"Error: {str(e)}")
        messagebox.showerror("SolidWorks Error", str(e))
        error = True
        return
        
def get_text_file_content(file_name):
            global error
            file_path = r"C:\SolidWorks Data\Envío de piezas a Odoo\\" + file_name + ".txt"
            print(datetime.datetime.now(), "searching content in: ", file_path)
            with open(file_path, 'r') as file:
                content = file.read()
            return content

def clean_text_file_content(file_name):
    global error
    file_path = r"C:\SolidWorks Data\Envío de piezas a Odoo\\" + file_name + ".txt"
    with open(file_path, 'w') as file:
        file.write('')

def clean_data_files():
    global error

    clean_text_file_content("Masa")
    clean_text_file_content("Volumen")
    clean_text_file_content("Superficie")
    clean_text_file_content("Ancho")
    clean_text_file_content("Largo")
    clean_text_file_content("Grosor")
    clean_text_file_content("Error")

def enviar_pieza(pieza):
    global error

    #send request to odoo

    global create_url

    #pop quantity out of the dictionary
    data = pieza.copy()
    data.pop("quantity")
    #print(pieza)
    
    #do request
    params_pieza = {"params": data}
    #add accept header
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    # Convert data to JSON format
    json_data = json.dumps(params_pieza)
    print(datetime.datetime.now(), json_data)

    

    # Check response
    try:
        # Send POST request
        response = requests.get(create_url, data=json_data, headers=headers)    

        if 404 == response.status_code:
            messagebox.showerror("Error en el envío", f"Se debe activar la integración del módulo con Odoo.")
            error_message = response.json()['result']['message']
            print(datetime.datetime.now(), error_message)
            error = True
            return
        
        if response.json()["result"]['status'] == "Error":
            
            error_message = response.json()['result']['message']
            if "Expected singleton" in error_message:
                error_message = f"La pieza {data['name']} ya existe en la base de datos."
                messagebox.showerror("Error en el envío", error_message )
                print(datetime.datetime.now(), error_message)
            else:
                messagebox.showerror("Error en el envío", f"{error_message}")
                print(datetime.datetime.now(), error_message)
            return
            
            

        if response.json()["result"]['status'] != "Ok":
            error_message = response.json()['result']['message']
            if "Expected singleton" in error_message:
                error_message = f"La pieza {data['name']} ya existe en la base de datos."
                messagebox.showerror("Error en el envío", error_message )
                print(datetime.datetime.now(), error_message)
            else: 
                messagebox.showerror("Request Failed", f"Request failed. Error status: {error_message}")
                print(datetime.datetime.now(), f"Request failed. Error message: {error_message}")
                error = True
            return
    except Exception as e:
        if response.json()["result"]['status'] != "Ok":
            error_message = response.json()['result']['message']
            if "Expected singleton" in error_message:
                error_message = f"La pieza {data['name']} ya existe en la base de datos."
                messagebox.showerror("Error en el envío", error_message)
                print(datetime.datetime.now(), error_message)
            else: 
                messagebox.showerror("Request Failed", f"Request failed. Error status: {error_message}")
                print(datetime.datetime.now(), f"Request failed. Error message: {error_message}")
                error = True
            return
        if 404 == response.status_code:
            messagebox.showerror("Error en el envío", f"Se debe activar la integración del módulo con Odoo.")
            print(datetime.datetime.now(), response.status_code)
            error = True
            return
        messagebox.showerror("Error en el envío", f"El envío sufrió un error inesperado. Por favor intente más tarde")
        print(datetime.datetime.now(), f"Request failed. Error status: {str(e)}")
        print(json_data)
        error = True
        return

    # Get response data
    #response_data = response.json()
    pieza["id"] = response.json()["result"]["default_code"]
    #actualizo el nombre de la pieza agregando el codigo recibido al final del nombre
    pieza["old_name"] = pieza["name"]
    pieza["name"] = pieza["name"] + " " + pieza["id"]

    #call function to update the url
    update_url(pieza)

def enviar_ensamble(ensamble, folder_path):
    global error
    
    global create_url
    # Convert data to JSON format
    json_data = json.dumps({"params": ensamble})
    print(ensamble)
    
    print(datetime.datetime.now(), json_data)
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    try:
        # Send get request
        response = requests.get(create_url, data=json_data, headers=headers)

        if 404 == response.status_code:
            messagebox.showerror("Error en el envío", f"Se debe activar la integración del módulo con Odoo.")
            print(datetime.datetime.now(), response.status_code)
            error = True
            return

        if response.json()["result"]['status'] != "Ok":
            error_message = response.json()['result']['message']
            messagebox.showerror("Error en el envío", f"Mensaje del error: {error_message}")
            print(datetime.datetime.now(), error_message)
            error = True
            return

    except Exception as e:

        try:

            error_message = response.json()['result']['message']
            if "Expected singleton" in error_message:
                error_message = f"El ensamble {ensamble['name']} ya existe en la base de datos."
                messagebox.showerror("Error en el envío", error_message )
                print(datetime.datetime.now(), error_message)
                
            if response:
                messagebox.showerror("Error en el envío", f"Falló el envío: {error_message}")
            messagebox.showerror("Error en el envío", f"Falló el envío: {error_message}")
            print(datetime.datetime.now(), f"Request failed. Error status: {error_message}, {str(e)}")
            print(json_data)
            error = True
            return
        except:
            messagebox.showerror("Error en el envío", f"Falló el envío")
            print(datetime.datetime.now(), f"Request failed. Error status: {str(e)}")
            print(json_data)
            error = True

    # Get response data
    #response_data = response.json()
    ensamble["id"] = response.json()["result"]["default_code"]
    #actualizo el nombre de la pieza agregando el codigo recibido al final del nombre
    ensamble["old_name"] = ensamble["name"]
    ensamble["name"] = ensamble["name"] + " " + ensamble["id"]   

    #call function to update the url
    update_url(ensamble)  

    return
    # Get response data
    #response_data = response.json()
    ensamble["id"] = response.json()["result"]["default_code"]

    # Split the folder path
    folder_path_parts = folder_path.split(os.sep)
    
    # Edit the last part of the folder path
    folder_path_parts[-1] = ensamble["id"]+ " " + folder_path_parts[-1]
    
    # Join the parts back together
    global new_folder_path
    new_folder_path = os.sep.join(folder_path_parts)

    os.rename(folder_path, new_folder_path)

def ordenar_valores (ancho, largo, grosor):
    global error

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

def procesar_ensamble(sldasm_files, folder_path):
        global error

    #escribir la ruta en el archivo input
        path_file = r"C:\SolidWorks Data\Ruta.txt"
        sldasm_file_path = os.path.join(folder_path, sldasm_files[0])
        sldasm_file_path = sldasm_file_path.replace("\\", "/")  # Replace backslashes with forward slashes

        #make the file path url
        sldasm_file_path_url = ("file:///" + sldasm_file_path.replace(" ", "%20"))

        
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
                error = True
                return

        #recopilar los datos guardados
        volumen = get_text_file_content("Volumen").strip().replace(",", ".")
        superficie = get_text_file_content("Superficie").strip().replace(",", ".")

        #obtener tag_id a partir del nombre del archivo
        #calcular masa a partir del peso específico
        material = sldasm_files[0].split(" ")[0]
        try:
            material_tag = peso_especifico.loc[peso_especifico["REFERENCIA"] == material, "TAG"].item()
            if material_tag == "" or material_tag == None:
                 
                messagebox.showerror("Error", f"Error al encontrar el tag del material para archivo {sldasm_files}, por favor verifique que tiene asignado un valor correcto en TAG.")
                error = True
                return
            
        except Exception as e:
            
            messagebox.showerror("Error", f"Error al encontrar la referencia al material en el archivo {sldasm_files}, por favor verifique que la referencia es correcta: {str(e)}")
            error = True
            return
        #calcular masa a partir del peso de las piezas
        global ensamble
        global piezas
        #iterar sobre las piezas y sumar el net_weight y el gross_weight para obtener ambos valores totales
        net_weight = 0
        gross_weight = 0
        superficie = 0
        ids = []
        print(piezas)
        try:
            
            
            for pieza in piezas:
                print(pieza)
                print(pieza["quantity"])
                net_weight += float(pieza["weight"])
                gross_weight += float(pieza["gross_weight"])
                superficie += float(pieza["superficie"])
                print(pieza["quantity"])
                ids.append({
                    "product_qty": pieza['quantity'],
                    "default_code": pieza["id"]
                })
        except Exception as e:
            

            if "id" in str(e):

                print("Error:", str(e))
                messagebox.showerror("Error", "Las piezas del ensamble no están siendo correctamente cargadas por el sistema. Se recomienda cargar manualmente este ensamble con sus piezas.")
                
                error = True
                return
             
            print("Error:", str(e))
            messagebox.showerror("Error", "Error procesando el ensamble: " + str(e))
            
            error = True
            return
        
        #save everything in a dictionary
            
        ensamble = {
            'name': sldasm_files[0].split(".")[0],
            "product_tag_ids": "Conjunto",
            "weight": net_weight,
            "gross_weight": gross_weight,
            "volume": volumen,
            "surface": superficie,
            "categ_id": material_tag,
            "sale_ok": "true",
            "purchase_ok": "false",
            "product_route": "Fabricar",
            #"tracking": "N° de CNC",
            "product_route": sldasm_file_path_url,
            "bill_of_materials": ids
        }
        print(ensamble)

def procesar_pieza(sldprt_file, folder_path):
        global error

    #escribir la ruta en el archivo input
        path_file = r"C:\SolidWorks Data\Ruta.txt"
        sldprt_file_path = os.path.join(folder_path, sldprt_file)
        sldprt_file_path = sldprt_file_path.replace("\\", "/")

        #make the file path url
        sldprt_file_path_url = "file:///" + sldprt_file_path.replace(" ", "%20")

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
            global error
            if error_text:
                 
                messagebox.showerror("SolidWorks Error", error_text)
                
                error = True
                return

        #recopilar los datos guardados
        volumen = get_text_file_content("Volumen").strip().replace(",", ".")
        superficie = get_text_file_content("Superficie").strip().replace(",", ".")
        ancho = get_text_file_content("Ancho").strip().replace(",", ".")
        largo = get_text_file_content("Largo").strip().replace(",", ".")
        espesor = get_text_file_content("Grosor").strip().replace(",", ".")

        print(datetime.datetime.now(),"archivo:", sldprt_file,  "volumen: ", volumen, "superficie: ", superficie, "ancho: ", ancho, "largo: ", largo, "espesor: ", espesor) 

        #calcular masa a partir del peso específico
        #check if file name starts with a number
        if sldprt_file.split(" ")[0].isdigit():
            print(sldprt_file, "is digit")
            material = sldprt_file.split(" ")[1]
        else:
            material = sldprt_file.split(" ")[0]
        try:
            
            peso_especifico_value = float(peso_especifico.loc[peso_especifico["REFERENCIA"] == material, "VALOR"].item())

            if peso_especifico_value == 0 or peso_especifico_value == None:
                 
                messagebox.showerror("Error", f"Error al encontrar el peso específico del material en el archivo {sldprt_file}, por favor verifique que tiene asignado un valor correcto en VALOR.")
            
                error = True
                return
            
            peso_especifico_value = peso_especifico_value / 1000000

            net_weight = float(volumen) * peso_especifico_value
            gross_weight = float(espesor) * float(ancho) * float(largo) * peso_especifico_value

            
        except Exception as e:
            
             
            messagebox.showerror("Error", f"Error al encontrar la referencia al material en el archivo {sldprt_file}, por favor verifique que la referencia es correcta: {str(e)}")
            
            error = True
            return

        try:
            
            material_tag = peso_especifico.loc[peso_especifico["REFERENCIA"] == material, "TAG"].item()
            

            if material_tag == "" or material_tag == None:
                 
                messagebox.showerror("Error", f"Error al encontrar el tag del material para archivo {sldprt_file}, por favor verifique que tiene asignado un valor correcto en TAG.")
                
                error = True
                return
            
        except Exception as e:
            
             
            messagebox.showerror("Error", f"Error al encontrar la referencia al material en el archivo {sldprt_file}, por favor verifique que la referencia es correcta: {str(e)}")
            
            error = True
            return
        try:
            #ordenar los valores
            ancho, largo, espesor = ordenar_valores(ancho, largo, espesor)

            #calcular espesor
            espesor_values = espesores['ESPESOR'].values
            # Find the closest value in espesor_values to espesor
            
            espesor_found = min(espesor_values, key=lambda x:abs(float(x)-float(espesor)))
            
            #get the value in col STRING in the same row as espesor_found
            espesor_string = espesores.loc[espesores['ESPESOR'] == espesor_found, 'STRING'].item()
            print(espesor_found)

        except Exception as e:
            
             
            messagebox.showerror("Error", f"Error al encontrar el espesor del archivo {sldprt_file}, por favor verifique que está cargado en el archivo de espesores: {str(e)}")
            print(f"Error al encontrar el espesor del archivo {sldprt_file}, por favor verifique que está cargado en el archivo de espesores: {str(e)}")
            error = True
            return

        try:
                    
            #seleccion de insumo para la pieza
            print(insumos_piezas)
            print("inusmoo", insumos_piezas.loc[(insumos_piezas["ESPESOR"] == espesor_found) & (insumos_piezas["MATERIAL"] == material), "INSUMO"].values)
            insumo = insumos_piezas.loc[(insumos_piezas["ESPESOR"] == espesor_found) & (insumos_piezas["MATERIAL"] == material), "INSUMO"].item()
            print("insumo", insumo)
            quantity = 1

        except Exception as e:
            
            
            messagebox.showerror("Error", f"Error al encontrar el insumo de la pieza {sldprt_file}, por favor verifique que está cargado en el archivo de insumos correctamente: {str(e)}")

            print(f"Error al encontrar el insumo de la pieza {sldprt_file}, por favor verifique que está cargado en el archivo de insumos correctamente: {str(e)}")
            error = True
            return

        #check if file name starts with a number
        if sldprt_file.split(" ")[0].isdigit():
            print(sldprt_file, "has digit, updating quantity")
            
            quantity = sldprt_file.split(" ")[0]
            #sldprt_file = " ".join(sldprt_file.split(" ")[1:])

        #save everything in a dictionary
        global piezas
        volumen = float(volumen)/1000000
        default_code = random.randint(0, 1000000)

        pieza = {
            'name': sldprt_file.split(".")[0],
            "quantity": quantity,
            "default_code": default_code,
            "product_tag_ids": "Piezas",
            "weight": net_weight,
            "gross_weight": gross_weight,
            "volume": volumen,
            "superficie": superficie,
            "broad": ancho,
            "long": largo,
            "categ_id": material_tag,
            "thickness": espesor_string,
            "sale_ok": "true",
            "purchase_ok": "false",
            "product_route": "Fabricar",
            #"tracking": "N° de CNC",
            "product_route": sldprt_file_path_url,
            "bill_of_materials": [
                {
                    "default_code": insumo,
                    "product_qty": gross_weight
                    }
                ],
        }

        piezas.append(pieza)

        print(pieza)

def update_url(archivo):
    global error

    #send request to odoo
    try:
        global update_url

        if archivo["product_tag_ids"] == "Conjunto":
            extension = ".SLDASM"
        else:
            extension = ".SLDPRT"

        #rename the file with the new name
        old_name = archivo["old_name"] + extension
        new_name = archivo["name"] + extension
        try:
            os.rename(os.path.join(folder_path, old_name), os.path.join(folder_path, new_name))
        except Exception as e:
            messagebox.showerror("Error al renombrar archivo", f"No se pudo renombrar el archivo {old_name}. Error: {str(e)}")
            print(datetime.datetime.now(), f"Error al renombrar archivo {old_name}. Error: {str(e)}")
            error = True
            return
        
        #generar url
        global new_folder_path
        file_url = new_folder_path + "/" + archivo["name"] + extension

        archivo["product_route"] = file_url

        # Convert data to JSON format
        json_data = json.dumps(archivo)

        # Send POST request
        response = requests.get(update_url, data=json_data)

        # Check response
        if response.status_code != "Ok":
            print(datetime.datetime.now(), f"Request failed. Status code: {response.status_code}")
            return
        
    except Exception as e:
        messagebox.showerror("Error en el envío", f"Se modificó el nombre pero no se pudo actualizar la URL del archivo {archivo['name']} en Odoo. Por favor, actualice manualmente.")
        print(datetime.datetime.now(), f"Request failed. Error status: {str(e)}")
        print(json_data)
        error = True
"""
#main donde inicia el procesamiento de la carpeta
def folder(input_folder_path):
    global error

    global folder_path
    folder_path = input_folder_path
    global swApp
    file_names = os.listdir(folder_path)
    global sldasm_files
    global sldprt_files    
    sldprt_files = [file_name for file_name in file_names if file_name.endswith('.SLDPRT')]
    sldasm_files = [file_name for file_name in file_names if file_name.endswith('.SLDASM')]
    global ensamble
    global piezas

    # Connect to an existing SolidWorks instance or create a new one if not available
    try:
        swApp = win32com.client.GetObject("SldWorks.Application")
    except:
        global error
        swApp = win32com.client.Dispatch("SldWorks.Application")

        
           
    #procesar cada pieza sldprt
    for sldprt_file in sldprt_files:
        if not sldprt_file.startswith("~") and not sldprt_file.startswith("$"):
            procesar_pieza(sldprt_file, folder_path)
            if error == True:
                return
            print("Pieza analizada")

    for pieza in piezas:
        print(datetime.datetime.now(), pieza['name'])
        enviar_pieza(pieza)
        if error == True:
            return
        print("Pieza enviada")
    
    
    
    #check if there is a sldasm file
    if sldasm_files:

        procesar_ensamble(sldasm_files, folder_path)
        if error == True:
            return
        print("Ensamble analizado")
        enviar_ensamble(ensamble, folder_path)
        if error == True:
            return
        print("Ensamble enviado")

         
    print(datetime.datetime.now(), "Proceso finalizado.")
    messagebox.showinfo("SolidWorks", "Proceso finalizado.")
    return
"""


class SimpleGUI(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("METALUX - SOLIDWORKS A ODOO")
        self.geometry("600x300")

        # Set the background color
        self.configure(bg="white")

        # Add window icon
        icon_path = r".\resources\metalux-logo.ico"
        self.iconbitmap(icon_path)

        # Company logo label
        logo_path = r".\resources\metalux brand.png"
        self.load_logo(logo_path)

        # Drag and drop area
        self.drop_area = tk.Label(self, text="Soltar carpeta de proyecto", bg="lightgray", pady=75, padx=180)
        self.drop_area.place(relx=0.5, rely=0.6, anchor="center")

        # Enable drag and drop functionality
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.on_drop)

        if not self.check_wifi_connection():
            messagebox.showinfo("Sin conexión a Internet", "Por favor conéctese a Internet y reintente nuevamente.")

    def load_logo(self, logo_path):
        try:
            image = Image.open(logo_path)
            image = image.resize((320, 82)) 
            self.logo_image = ImageTk.PhotoImage(image)
            self.logo_label = tk.Label(self, image=self.logo_image, bg="white")
            self.logo_label.place(relx=1, rely=0, anchor="ne")  # Adjusted position to the right
        except Exception as e:
             
            messagebox.showerror("Error", f"Error loading logo: {e}")
            
            error = True
            return

    

    def on_drop(self, event):

        self.drop_area.config(text="Procesando carpeta...")
        
        
        folder_path = event.data#[1:-1]

        if folder_path[0] == '{' or folder_path[0] == '$':
            folder_path = folder_path[1:]

        if folder_path[-1] == '}' or folder_path[-1] == '$':
            folder_path = folder_path[:-1]

        print(datetime.datetime.now(), folder_path)
        
        if os.path.isdir(folder_path):
            
            
            file_names = os.listdir(folder_path)
            
            sldprt_files = [file_name for file_name in file_names if file_name.endswith('.SLDPRT') and not file_name.startswith('~$')]
            sldasm_files = [file_name for file_name in file_names if file_name.endswith('.SLDASM') and not file_name.startswith('~$')]
            notfiltered_sldasm = sldasm_files
            sldasm_files = []
            for sldprt_file in notfiltered_sldasm:
                if not sldprt_file.startswith("~") and not sldprt_file.startswith("$"):
                    sldasm_files.append(sldprt_file)
            
            #if there are more than one sldasm file, show a message error
            if len(sldasm_files) > 1:
                self.drop_area.config(text="Hay más de un ensamblaje en la carpeta, envíe sólo uno.")
            elif sldprt_files:
                
                time.sleep(1)
            
                # Call the function from another_module with the file path as a parameter
                folder(folder_path)
                time.sleep(1)
                #finish program
                self.destroy
                time.sleep(1)
                return

            else:
                self.drop_area.config(text=f"Ingrese una carpeta con piezas de SolidWorks con sólo un ensamblaje. {folder_path}")
        else:
            self.drop_area.config(text=f"Ingrese una carpeta con piezas de SolidWorks para continuar. {folder_path} no es una carpeta.")

    def check_wifi_connection(self):
        try:
            # Create a socket object
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)  # Set timeout to 1 second

            # Connect to a remote host (in this case, google.com) on port 80
            result = sock.connect_ex(('google.com', 80))

            if result == 0:
                return True  # WiFi connection is available
            else:
                return False  # No WiFi connection
        except Exception:
            return False  # No WiFi connection

"""
if __name__ == "__main__":
    app = SimpleGUI()
    app.mainloop()
    log_file.close()  

"""
#folder(r"C:\Users\Usuario\Pictures\09131 Bandeja - copia")

piezas = []
ensamble = {}

folder_path = "/Users/pedrobergaglio/Downloads"
"/Users/pedrobergaglio/Downloads/pueba99.SLDPRT"

enviar_pieza({
    'name': 'prueba99',
    'quantity': 1,
    'default_code': 1,
    'product_tag_ids': 'Piezas',
    'weight': 3.0,
    'gross_weight': 0.0,
    'volume': 500.0,
    'superficie': 200.0,
    'broad': 140.0,
    'long': 250.0,
    'categ_id': 'Chapa Galvanizada SAE 1010',
    'thickness': '0.9',
    'sale_ok': 'true',
    'purchase_ok': 'false',
    'product_route': 'Bandeja gVggf',
    'bill_of_materials': [{'default_code': 20013, 'product_qty': 1.0}]
})
