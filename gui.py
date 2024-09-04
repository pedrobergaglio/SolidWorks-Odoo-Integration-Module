import time
import os
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image, ImageTk
import socket
import tkinter.messagebox as messagebox
import requests
import json
import win32com.client
import pandas as pd
import random
import sys
import datetime
import re

# Open the log file in append mode
log_file = open(r"C:\SolidWorks Data\Envío de piezas a Odoo\logfile.log", 'a')

sys.stdout = log_file
sys.stderr = log_file

# Specify your macro name agui.pynd part file path
macro_name = r"C:\SolidWorks Data\Envío de piezas a Odoo\main.swp"
# "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe" "C:\Users\Usuario\Downloads\08978 Puerta Tablero.SLDPRT"

sldprt_files = []
sldasm_files = []
swApp = None

ensamble = {}
piezas = []
dont_replace = []

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
update_odoo_url = "http://ec2-3-15-193-242.us-east-2.compute.amazonaws.com:8069/itec-api/update/product"

def run_solidworks_macro(swApp, macro_name):
    global error
    try:
        # Connect to SolidWorks
        swApp.Visible = False

        # Run your VBA macro
        #macro_full_path = os.path.join(os.path.expanduser("~"), "AppData\\Roaming\\SolidWorks\\SolidWorks 2019\\macros", macro_name)
        macro_full_path = r"C:\SolidWorks Data\Envío de piezas a Odoo\main.swp"
        print(datetime.datetime.now(), macro_full_path)
        swApp.RunMacro(macro_full_path, "main1", "main1")

    except Exception as e:
         
        print(datetime.datetime.now(), f"Error: {str(e)}")
        messagebox.showerror("SolidWorks Error", str(e))
        error = True
        return

def find_product_code(error_message):
    # Define the regex pattern to find the code starting with 'W' followed by digits
    pattern = r"W\d+"
    
    # Search for the pattern in the error message
    match = re.search(pattern, error_message)
    
    # If a match is found, return the code
    if match:
        return match.group(0)
    else:
        return None

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
    global dont_replace

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
            

        if response.json()["result"]['status'] != "Ok":
            error_message = response.json()['result']['message']

            #response for duplicated without info about the code
            if "Expected singleton" in error_message:
                error_message = f"La pieza {data['name']} ya existe en la base de datos. No será agregada en la lista de materiales del ensamble"
                #messagebox.showerror("Error en el envío", error_message )
                dont_replace.append(data['name'])
                print(datetime.datetime.now(), error_message)

            #response for duplicated with info about the code
            elif "El código de producto" in error_message:
                #find the code in the error message
                default_code = find_product_code(error_message)

                if default_code:
                    #update the default code in the pieza dictionary
                    pieza["default_code"] = default_code

                    #update the error message
                    error_message = f"La pieza {data['name']} ya existe en la base de datos. Será agregada igualmente en la lista de materiales del ensamble"
                else:
                    error_message = f"La pieza {data['name']} ya existe en la base de datos. No será agregada en la lista de materiales del ensamble"   

                
                #messagebox.showerror("Error en el envío", error_message )
                dont_replace.append(data['name'])
                print(datetime.datetime.now(), error_message)
            else: 
                messagebox.showerror("Error en el envío", f"Error en el envío. Estado del error: {error_message}")
                print(datetime.datetime.now(), f"Request failed. Error message: {error_message}")
                error = True
            return
        
    except Exception as e:
        
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

    # save the default code in the pieza dictionary
    pieza["default_code"] = response.json()["result"]["default_code"]
    #actualizo el nombre de la pieza agregando el codigo recibido al final del nombre
    pieza["old_name"] = pieza["name"]
    pieza["name"] = pieza["name"] + " " + pieza["default_code"]
    print(pieza["default_code"])

    #call function to update the url
    #update_url(pieza)

def enviar_ensamble():

    global ensamble
    global folder_path
    
    #calcular masa a partir del peso de las piezas
    global piezas
    global error
    #iterar sobre las piezas
    ids = []
    print(piezas)

    #iterate over the piezas and sum the net_weight and gross_weight to get both total values
    #then add them to the material list
    try:
        for pieza in piezas:
            print(pieza)

            #if the pieza has no default_code, dont add it but count its weight
            if not pieza["default_code"] :
                continue
            #add the pieza to the material list
            ids.append({
                "product_qty": pieza['quantity'],
                "default_code": pieza["default_code"]
            })
            
    except Exception as e:

        if "id" in str(e):

            print("Error:", str(e))
            messagebox.showerror("Error", "Las piezas del ensamble no están siendo correctamente cargadas por el sistema. Se recomienda cargar manualmente este ensamble con sus piezas.")
            return
            
        print("Error:", str(e))
        messagebox.showerror("Error", "Error procesando el ensamble: " + str(e))
        
        error = True
        return
    
    #save the new info in the dictionary
    ensamble["bill_of_materials"] = ids
    
    print(ensamble)
    
    global create_url
    # Convert data to JSON format
    json_data = json.dumps({"params": ensamble})
    print(ensamble)
    
    print(datetime.datetime.now(), json_data)
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    print("NOMBRE DEL ENSAMBLE::::::::::::::::")
    #print(ensamble["name"])
    print("NOMBRE DEL ENSAMBLE::::::::::::::::")

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
            if "Expected singleton" in error_message:
                error_message = f"El ensamble {ensamble['name']} ya existe en la base de datos. Las nuevas piezas fueron subidas al sistema pero no se agregaron a ningún ensamble"
                messagebox.showinfo("Error en el envío", error_message )
                dont_replace.append(ensamble['name'])
                print(datetime.datetime.now(), error_message)
            elif "El código de producto" in error_message:
                error_message = f"El ensamble {ensamble['name']} ya existe en la base de datos. Las nuevas piezas fueron subidas al sistema pero no se agregaron a ningún ensamble"
                messagebox.showinfo("Error en el envío", error_message )
                dont_replace.append(ensamble['name'])
                print(datetime.datetime.now(), error_message)
            else: 
                messagebox.showerror("Falló el envío", f"Falló el envío. Las nuevas piezas fueron subidas al sistema pero no se agregaron a ningún ensamble. Estado de error: {error_message}")
                print(datetime.datetime.now(), f"Request failed. Error message: {error_message}")
                error = True
            return

    except Exception as e:

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
    ensamble["default_code"] = response.json()["result"]["default_code"]
    #actualizo el nombre de la pieza agregando el codigo recibido al final del nombre
    ensamble["old_name"] = ensamble["name"]
    ensamble["name"] = ensamble["name"] + " " + ensamble["default_code"]  

     

    #call function to update the url
    #update_url(ensamble)  
    return

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
        global ensamble
        global piezas
        global error

        #escribir la ruta en el archivo input
        path_file = r"C:\SolidWorks Data\Ruta.txt"
        sldasm_file_path = os.path.join(folder_path, sldasm_files[0])
        sldasm_file_path = sldasm_file_path.replace("\\", "/")  # Replace backslashes with forward slashes

        #make the file path url
        #sldasm_file_path_url = ("file:///" + sldasm_file_path.replace(" ", "%20"))
        sldasm_file_path_url = (sldasm_file_path)
        #acaca

        
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
        global piezas
        print(piezas)

        print("****************************************")
        print(sldasm_files[0].split(".")[0])
        
        #save its data in a dictionary before the send, data about its pieces will be added
        ensamble = {
            "name": sldasm_files[0].split(".")[0],
            "product_tag_ids": "Conjunto",
            "volume": volumen,
            "sheet_type": material_tag,
            "categ_id":"Producto Fabricado",
            "sale_ok": "true",
            "purchase_ok": "false",
            "surface": superficie,
            #"tracking": "N° de CNC",
            "product_route": sldasm_file_path_url
        }

        

        print("NOMBRE DEL ENSAMBLE::::::::::::::::")
        print(ensamble["name"])
        print("NOMBRE DEL ENSAMBLE::::::::::::::::")
        
        #calcular masa a partir del peso de las piezas
        
    
        #iterar sobre las piezas y sumar el net_weight y el gross_weight para obtener ambos valores totales
        net_weight = 0
        gross_weight = 0
        superficie = 0
        volumen = 0
        ids = []
        print(piezas)

        #iterate over the piezas and sum the net_weight and gross_weight to get both total values
        #then add them to the material list
        try:
            for pieza in piezas:
                print(pieza)
                #cumulate the values
                net_weight += float(pieza["weight"])
                gross_weight += float(pieza["gross_weight"])
                superficie += float(pieza["superficie"])
                volumen += float(pieza["volume"])
        except Exception as e:

            if "id" in str(e):

                print("Error:", str(e))
                messagebox.showerror("Error", "Las piezas del ensamble no están siendo correctamente cargadas por el sistema. Se recomienda cargar manualmente este ensamble con sus piezas.")
                return
                
            print("Error:", str(e))
            messagebox.showerror("Error", "Error procesando el ensamble: " + str(e))
            
            error = True
            return
        
        #save the new info in the dictionary
        ensamble["weight"] = net_weight
        ensamble["gross_weight"] = gross_weight
        ensamble["surface"] = superficie
        ensamble["volume"] = volumen
        
        print(ensamble)
        
        print("****************************************")
        print(ensamble["name"])
        print("****************************************")

def procesar_pieza(sldprt_file, folder_path):
        global error

    #escribir la ruta en el archivo input
        path_file = r"C:\SolidWorks Data\Ruta.txt"
        sldprt_file_path = os.path.join(folder_path, sldprt_file)
        sldprt_file_path = sldprt_file_path.replace("\\", "/")

        #make the file path url
        #sldprt_file_path_url = "file:///" + sldprt_file_path.replace(" ", "%20")
        sldprt_file_path_url = sldprt_file_path

        with open(path_file, 'w') as file:
            file.write(sldprt_file_path)

        #limpiar los archivos mediante los cuales el script se comunica con el macros de solidworks
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
            # Ordenar los valores
            ancho, largo, espesor = ordenar_valores(ancho, largo, espesor)

            # Calcular espesor

            # Filtrar los espesores permitidos para el material
            print("Material:", material)
            print("Espesor:", espesor)
            
            espesores_material = espesores[espesores['MATERIALES HABILITADOS'].apply(lambda x: material in x)]
            
            print("Espesores:")
            print(espesores.head())
            
            print("Espesores Material:")
            print(espesores_material.head())

            # Tomar los valores de los espesores
            espesor_values = espesores_material['ESPESOR'].values
            print("Espesor Values:", espesor_values)

            # Find the closest value in espesor_values to espesor
            espesor_found = min(espesor_values, key=lambda x: abs(float(x) - float(espesor)))
            print("Espesor Found:", espesor_found)

            if abs(espesor_found - espesor) >= 0.15:
                messagebox.showerror("Error", f"No se encontró el insumo de la pieza {sldprt_file} con el espesor correcto. Por favor verifique que la chapa con el espesor de la pieza está cargado en el archivo de insumos correctamente.")
                print(f"No se encontró el insumo de la pieza {sldprt_file} con el espesor correcto. Por favor verifique que la chapa con el espesor de la pieza está cargado en el archivo de insumos correctamente.")
                error = True
                return

            # Get the value in col STRING in the same row as espesor_found
            matched_rows = espesores_material.loc[espesores_material['ESPESOR'] == espesor_found, 'STRING']
            print("Matched Rows:", matched_rows)

            # Ensure only one row is matched
            if len(matched_rows) != 1:
                raise ValueError(f"Expected a single match for espesor_found {espesor_found}, but found {len(matched_rows)} matches")

            # Proceed with extracting the single value
            espesor_string = matched_rows.item()
            print("Espesor String:", espesor_string)

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

        pieza = {
            'name': sldprt_file.split(".")[0],
            "default_code": 0,
            "quantity": quantity,
            "product_tag_ids": "Piezas",
            "weight": net_weight,
            "gross_weight": gross_weight,
            "volume": volumen,
            "categ_id":"Producto Fabricado",
            "superficie": superficie,
            "broad": ancho,
            "long": largo,
            "sheet_type": material_tag,
            "thickness": espesor_string,
            "sale_ok": "true",
            "purchase_ok": "false",
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
    global folder_path

    #send request to odoo
    
    global update_odoo_url

    if archivo["product_tag_ids"] == "Conjunto":
        extension = ".SLDASM"
    else:
        extension = ".SLDPRT"

    #rename the file with the new name
    old_name = archivo["old_name"] + extension
    new_name = archivo["name"] + extension

    

    try:
        rename1 = os.path.join(folder_path, old_name)#.replace("/", os.sep)
        rename2 = os.path.join(folder_path, new_name)#.replace("/", os.sep)
        os.rename(rename1, rename2)
        print(os.path.join(folder_path, new_name))
    except Exception as e:
        messagebox.showerror("Error al renombrar archivo", f"No se pudo renombrar el archivo {old_name}. Error: {str(e)}")
        print(datetime.datetime.now(), f"Error al renombrar archivo {old_name}. Error: {str(e)}")
        error = True
        return
    
    #generar url
    #global folder_path
    #file_url = ("file:///"+ folder_path.replace(os.sep, "/") + "/" + archivo["name"] + extension).replace(" ", "%20")
    #os.path.join("file:///", folder_path, archivo["name"]+extension)
    #file_url = (folder_path.replace(os.sep, "/") + "/" + archivo["name"] + extension)#.replace(os.sep, "/").replace("\\", "/")
    file_url = (folder_path.replace(os.sep, "/") )#+ "/" + archivo["name"] + extension)#.replace(os.sep, "/").replace("\\", "/")
    archivo["product_route"] = file_url

    print(archivo["product_route"])

    #add accept header
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    
    # Convert data to JSON format
    json_data = json.dumps({"params":archivo})

    try:

        # Send POST request
        response = requests.get(update_odoo_url, data=json_data, headers=headers)

        # Check response
        if response.json()["result"]['status'] != "Ok":
            print(datetime.datetime.now(), f"Request failed. Status code: {response.status_code}")
            print(json_data)
            return
        
        if response.json()["result"]['status'] == "Ok":
            print(datetime.datetime.now(), f"Se actualizó la ruta de la pieza correctamente: {response.status_code}")
            print(json_data)
            archivo['success'] = True
            return
        
    except Exception as e:
        print(datetime.datetime.now(), f"Request failed. Error status: {str(e)}")
        messagebox.showerror("Error en el envío", f"Se modificó el nombre pero no se pudo actualizar la URL del archivo {archivo['name']} en Odoo. Por favor, actualícela manualmente.")
        print(json_data)
        error = True
        return


    return
        
#main donde sucede procesamiento de la carpeta. 
#Se procesan las piezas y el ensamble
#Se detectan errores en el #* procesamiento
def procesamiento(input_folder_path):

    global ensamble
    global piezas
    piezas = []
    ensamble = {}

    #global variables
    global error
    global folder_path
    folder_path = input_folder_path
    global swApp
    file_names = os.listdir(folder_path)
    global sldasm_files
    global sldprt_files    
    sldprt_files = [file_name for file_name in file_names if file_name.endswith('.SLDPRT')]
    sldasm_files = [file_name for file_name in file_names if file_name.endswith('.SLDASM')]
    
    # check and open SolidWorks
    
    # Connect to an existing SolidWorks instance or create a new one if not available
    try:
        swApp = win32com.client.GetObject("SldWorks.Application")
    except:
        global error
        swApp = win32com.client.Dispatch("SldWorks.Application")
    
           
    #procesar cada pieza sldprt y guardarla en la lista de piezas
    for sldprt_file in sldprt_files:
        if not sldprt_file.startswith("~") and not sldprt_file.startswith("$"):
            procesar_pieza(sldprt_file, folder_path)
            if error == True:
                return
            print("Pieza analizada")

    #check if there is a sldasm file and process it
    if sldasm_files:

        procesar_ensamble(sldasm_files, folder_path)
        if error == True:
            return
        print("Ensamble analizado")
    
    return

#*################### ACÁ CAMBIAR A OTRA FUNCIÓN Y MOSTRAR LO PRE ANALIZADO ANTES DE ENVIAR CON UN BOTÓN

# Función para enviar las piezas a Odoo
# Si llegó hasta acá quiede decir que pudo encontrar todos los datos necesarios
def envio():

    global ensamble
    global piezas

    #global variables
    global error
    global folder_path
    global swApp
    global sldasm_files
    global sldprt_files
    global dont_replace
    global new_folder_path

    #enviar cada pieza a odoo
    # si alguna está duplicada no va a ser enviada
        # pero además, intentará obtener el codigo de producto de la pieza duplicada 
        # para poder agregarla a la lista de materiales del ensamble
    for pieza in piezas:
        print(datetime.datetime.now(), pieza['name'])
        enviar_pieza(pieza)
        if error == True:
            return
        print("Pieza enviada")
    
    #check if there is a sldasm file and send it to odoo
    #
    if sldasm_files:

        enviar_ensamble()
        if error == True:
            return
        print("Ensamble enviado")


    # once every product is sent, we have to rename the main folder, and then update every url
    if not any(word.startswith('W') and word[1:].isdigit() for word in folder_path.split(os.sep)[-1].split()):
        if sldasm_files and ensamble["default_code"] is not None:
            #rename with the name of the assembly
            ensamble_name = ensamble["default_code"]
            try:
                new_folder_path = " ".join([folder_path, ensamble_name])
                os.rename(folder_path, new_folder_path)
                folder_path = new_folder_path
                print(f'Renamed folder to {folder_path}')
            except Exception as e:
                messagebox.showerror("Error al renombrar archivo", f"No se pudo renombrar la carpeta{folder_path}. Error: {str(e)}")
                print(datetime.datetime.now(), f"Error al renombrar la carpeta{folder_path}. Error: {str(e)}")
                return

        else:
            #rename with the name of the first part in the list
            try:
                primer_pieza = piezas[0]
                pieza_name = primer_pieza["default_code"]
                new_folder_path = " ".join([folder_path, pieza_name])
                os.rename(folder_path, new_folder_path)
                folder_path = new_folder_path
                print(f'Renamed folder to {folder_path}')
            except Exception as e:
                messagebox.showerror("Error al renombrar archivo", f"No se pudo renombrar la carpeta{folder_path}. Error: {str(e)}")
                print(datetime.datetime.now(), f"Error al renombrar la carpeta{folder_path}. Error: {str(e)}")
                return

    if sldasm_files:
        update_url(ensamble)
        if error == True:
            return
        print("Ensamble URL updated")
    for pieza in piezas:
        update_url(pieza)
        if error == True:
            return
        print("Pieza URL updated")








    #finish#################################
    # Show a message box if there are files that were not replaced   
    if dont_replace:

        # Convert the array into a nicely formatted string
        dont_replace_str = ', '.join(dont_replace)

        if len(dont_replace) == 1:
            messagebox.showinfo("SolidWorks", f"Proceso finalizado. El archivo {dont_replace_str}, no ha sido enviado al sistema ya que ha sido cargado previamente.")  

        else:
            messagebox.showinfo("SolidWorks", f"Proceso finalizado. Los archivos {dont_replace_str}, no han sido enviados al sistema ya que han sido cargados previamente.")  
    else:
        messagebox.showinfo("SolidWorks", "Proceso finalizado.")
    #print a message in the console
    print(datetime.datetime.now(), f"Proceso finalizado. Las piezas {dont_replace} no han sido enviadas al sistema ya que ya han sido cargadas previamente")

    # clean the global variables
    piezas = []
    ensamble = {}
    dont_replace = []
    error = False
    
    sldprt_files = []
    sldasm_files = []
    swApp = None
    new_folder_path = ""
    folder_path = ""

    return

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

        #self.show_results_window()

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
                messagebox.showinfo("Envío de archivos", "Hay más de un ensamblaje en la carpeta, envíe sólo uno.")

            elif sldprt_files:
                
                time.sleep(1)
            
                # Call the function from another_module with the file path as a parameter
                procesamiento(folder_path)

                time.sleep(1)

                # Show a new screen with the results of the processing
                self.show_results_window()

                time.sleep(1)
                #finish program
                self.destroy
                time.sleep(1)
                return

            else:
                self.drop_area.config(text=f"Ingrese una carpeta con piezas de SolidWorks con sólo un ensamblaje. {folder_path}")
        else:
            self.drop_area.config(text=f"Ingrese una carpeta con piezas de SolidWorks para continuar. {folder_path} no es una carpeta.")

    def show_results_window(self):
        results_window = tk.Toplevel(self)
        results_window.title("Resultados del procesamiento")
        results_window.geometry("800x500")

        # Set the background color
        results_window.configure(bg="white")

        # Add window icon
        icon_path = r".\resources\metalux-logo.ico"
        results_window.iconbitmap(icon_path)
        

        # Use a modern font and color scheme
        header_font = ("Roboto", 12, "bold")
        text_font = ("Roboto", 9)
        header_color = "#CF000A"
        text_color = "#202020"
        bg_color = "white"

        # Create a dictionary for translations
        translations = {
            "name": "Nombre",
            "default_code": "Código",
            "quantity": "Cantidad",
            "product_tag_ids": "Tipo de producto",
            "weight": "Peso neto",
            "gross_weight": "Peso bruto",
            "volume": "Volumen",
            "superficie": "Superficie",
            "broad": "Ancho",
            "long": "Largo",
            "sheet_type": "Tipo de chapa",
            "thickness": "Espesor",
            "sale_ok": "Disponible para la venta",
            "purchase_ok": "Disponible para la compra",
            "product_route": "Link del producto",
            "bill_of_materials": "Lista de materiales",
            "surface": "Superficie"
        }

        # Display results header
        results_label = tk.Label(results_window, text="Resultados del procesamiento", font=header_font, fg=text_color, bg=bg_color)
        results_label.pack(pady=10)

        # Create a canvas and a vertical scrollbar
        canvas = tk.Canvas(results_window, bg=bg_color)
        scrollbar = tk.Scrollbar(results_window, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=bg_color)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Enable mouse wheel scrolling
        def _on_mouse_wheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        results_window.bind_all("<MouseWheel>", _on_mouse_wheel)

        # Frame for assembly information
        if ensamble:
            ensamble_frame = tk.Frame(scrollable_frame, bg=bg_color)
            ensamble_frame.pack(padx=20, pady=10, anchor="w")

            ensamble_label = tk.Label(ensamble_frame, text="Ensamble procesado", font=header_font, fg=header_color, bg=bg_color)
            ensamble_label.grid(row=0, column=0, sticky="w", columnspan=2)

            ensamble_table = tk.Frame(ensamble_frame, bg=bg_color)
            ensamble_table.grid(row=1, column=0, columnspan=2, sticky="w")

            for row, (key, value) in enumerate(ensamble.items()):
                translated_key = translations.get(key, key)  # Use translation if available, otherwise use the original key
                key_label = tk.Label(ensamble_table, text=f"{translated_key}:", font=text_font, fg=text_color, bg=bg_color)
                key_label.grid(row=row, column=0, sticky="w", padx=(10, 5))
                value_label = tk.Label(ensamble_table, text=value, font=text_font, fg=text_color, bg=bg_color)
                value_label.grid(row=row, column=1, sticky="w")

        # Frame for parts information
        if piezas:
            piezas_frame = tk.Frame(scrollable_frame, bg=bg_color)
            piezas_frame.pack(padx=0, pady=10, anchor="w")

            piezas_label = tk.Label(piezas_frame, text="Piezas procesadas:", font=header_font, fg=header_color, bg=bg_color)
            piezas_label.grid(row=0, column=0, sticky="w", columnspan=2)

            for index, pieza in enumerate(piezas):
                if index > 0:
                    # Add a divider between each pieza
                    divider = tk.Frame(piezas_frame, height=1, bg=text_color)
                    divider.grid(row=index*2, column=0, columnspan=2, sticky="we", pady=(10, 0))
                
                pieza_label = tk.Label(piezas_frame, text=f"Pieza {index + 1}:", font=header_font, fg=text_color, bg=bg_color)
                pieza_label.grid(row=index*2 + 1, column=0, sticky="w", padx=(20, 5))

                pieza_table = tk.Frame(piezas_frame, bg=bg_color)
                pieza_table.grid(row=index*2 + 1, column=1, sticky="w")

                for row, (key, value) in enumerate(pieza.items()):
                    translated_key = translations.get(key, key)  # Use translation if available, otherwise use the original key
                    key_label = tk.Label(pieza_table, text=f"{translated_key}:", font=text_font, fg=text_color, bg=bg_color)
                    key_label.grid(row=row, column=0, sticky="w", padx=(10, 5))
                    value_label = tk.Label(pieza_table, text=value, font=text_font, fg=text_color, bg=bg_color)
                    value_label.grid(row=row, column=1, sticky="w")


        # Frame for buttons
        button_frame = tk.Frame(scrollable_frame, bg=bg_color)
        button_frame.pack(pady=20)

        # Confirmation button
        confirm_button = tk.Button(button_frame, text="Confirmar y enviar", command=lambda: [envio(), results_window.destroy()], bg=text_color, fg='white', font=text_font)
        confirm_button.grid(row=1, column=0, padx=10)

        # Cancel button
        cancel_button = tk.Button(button_frame, text="Cancelar", command=results_window.destroy, bg=text_color, fg='white', font=text_font)
        cancel_button.grid(row=1, column=1, padx=10)

        # Enable text selection
        for widget in scrollable_frame.winfo_children():
            widget.bind("<Button-1>", lambda e: widget.focus_set())
        
        results_window.lift()  # Brings window to the front

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

#"""
if __name__ == "__main__":
    app = SimpleGUI()
    app.mainloop()
    log_file.close()  
