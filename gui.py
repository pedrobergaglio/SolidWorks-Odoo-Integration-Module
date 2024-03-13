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


# Specify your macro name and part file path
macro_name = r".\Envío de piezas a Odoo\main.swp"
# "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe" "C:\Users\Usuario\Downloads\08978 Puerta Tablero.SLDPRT"

sldprt_files = []
sldasm_files = []
swApp = None

ensamble = {}
piezas = []

new_folder = ""
folder_path = ""

#importar referencias
espesores = pd.read_excel(r".\resources\espesores.xlsx")
insumos_piezas = pd.read_excel(r".\resources\insumos-piezas.xlsx")
peso_especifico = pd.read_excel(r".\resources\peso-especifico.xlsx")

create_url = "http://ec2-3-15-193-242.us-east-2.compute.amazonaws.com:8069/itec-api/create/product"
update_url = "http://ec2-3-15-193-242.us-east-2.compute.amazonaws.com:8069/itec-api/update/product"

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

        #make the file path url
        sldasm_file_path_url = "file:///" + sldasm_file_path.replace(" ", "%20")
        
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
        volumen = get_text_file_content("Volumen").strip().replace(",", ".")
        superficie = get_text_file_content("Superficie").strip().replace(",", ".")

        #obtener tag_id a partir del nombre del archivo
        #calcular masa a partir del peso específico
        material = sldasm_files[0].split(" ")[0]
        try:
            material_tag = peso_especifico.loc[peso_especifico["REFERENCIA"] == material, "TAG"].item()
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
        superficie = 0
        ids = []
        try:
            for pieza in piezas:
                net_weight += pieza["weight"]
                gross_weight += pieza["gross_weight"]
                superficie += float(pieza["superficie"])
                
                ids.append({
                     "default_code": pieza["id"],
                     "quantity": pieza["quantity"]})
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error procesando las piezas: {str(e)}")
        
        #save everything in a dictionary
            
        ensamble = {
            "name": " ".join(sldasm_files[0].split()[1:]).split(".")[0],
            "product_tag_ids": "Conjunto",
            "weight": net_weight,
            "gross_weight": gross_weight,
            "volume": volumen,
            "surface": superficie,
            "categ_id": material_tag,
            "sale_ok": "true",
            "purchase_ok": "false",
            #"tracking": "N° de CNC",
            "product_route": sldasm_file_path_url,
            "bill_of_materials": ids
        }

def process_sldprt(sldprt_file, folder_path):

    #escribir la ruta en el archivo input
        path_file = r".\Envío de piezas a Odoo\Ruta.txt"
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
            if error_text:
                messagebox.showerror("SolidWorks Error", error_text)
                return

        #recopilar los datos guardados
        volumen = get_text_file_content("Volumen").strip().replace(",", ".")
        superficie = get_text_file_content("Superficie").strip().replace(",", ".")
        ancho = get_text_file_content("Ancho").strip().replace(",", ".")
        largo = get_text_file_content("Largo").strip().replace(",", ".")
        espesor = get_text_file_content("Grosor").strip().replace(",", ".")

        #calcular masa a partir del peso específico
        material = sldprt_file.split(" ")[0]
        try:
            peso_especifico_value = float(peso_especifico.loc[peso_especifico["REFERENCIA"] == material, "VALOR"].item())

            if peso_especifico_value == 0 or peso_especifico_value == None:
                messagebox.showerror("Error", f"Error al encontrar el peso específico del material en el archivo {sldprt_file}, por favor verifique que tiene asignado un valor correcto en VALOR.")
                return
            
            peso_especifico_value = peso_especifico_value / 1000000

            net_weight = float(volumen) * peso_especifico_value
            gross_weight = float(espesor) * float(ancho) * float(largo) * peso_especifico_value

            
        except Exception as e:
            messagebox.showerror("Error", f"Error al encontrar la referencia al material en el archivo {sldprt_file}, por favor verifique que la referencia es correcta: {str(e)}")
            return

        try:
            material_tag = peso_especifico.loc[peso_especifico["REFERENCIA"] == material, "TAG"].item()
            

            if material_tag == "" or material_tag == None:
                messagebox.showerror("Error", f"Error al encontrar el tag del material para archivo {sldprt_file}, por favor verifique que tiene asignado un valor correcto en TAG.")
                return
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al encontrar la referencia al material en el archivo {sldprt_file}, por favor verifique que la referencia es correcta: {str(e)}")
            return

        #take quantity of pieces
        quantity = 1
        file_name_parts = sldprt_file.split()
        if len(file_name_parts) > 1:
            second_word = file_name_parts[1]
            if second_word.isdigit():
                quantity = int(second_word)
                sldprt_file = " ".join([file_name_parts[0]] + sldprt_file.split()[2:])
                print(sldprt_file)

        #ordenar los valores
        ancho, largo, espesor = ordenar_valores(ancho, largo, espesor)

        #calcular espesor
        espesor_values = espesores['ESPESOR'].values
        # Find the closest value in espesor_values to espesor
        espesor_found = min(espesor_values, key=lambda x:abs(float(x)-float(espesor)))
        
        #get the value in col STRING in the same row as espesor_found
        espesor_string = espesores.loc[espesores['ESPESOR'] == espesor_found, 'STRING'].item()

        #seleccion de insumo para la pieza
        insumo = insumos_piezas.loc[(insumos_piezas["ESPESOR"] == espesor_found) & (insumos_piezas["MATERIAL"] == material), "INSUMO"].item()
        
        #save everything in a dictionary
        global piezas
        volumen = float(volumen)/1000000

        pieza = {
            "name": " ".join(sldprt_file.split()[1:]).split(".")[0],
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

        #print(pieza)

def ensamble_odoo(ensamble, folder_path):
    #print("Ensamble: ", file)
    #print(masa, "Kg", volumen, "mm3", superficie, "mm2")

    #send request to odoo
    global create_url
    
    #do request

    # Convert data to JSON format
    json_data = json.dumps(ensamble)
    print(json_data)

    # Send POST request
    response = requests.get(create_url, data=json_data)

    try:
        if response.json()["result"]['status'] != 200:
            error_message = f"Envío de datos fallido. Mensaje de error: {response.json()['result']['message']}"
            messagebox.showerror("Data Sending Error", error_message)
            return
    except:
        error_message = f"Envío de datos fallido. Estado de error: {response}"
        messagebox.showerror("Data Sending Error", error_message)
        return

    # Get response data
    #response_data = response.json()
    ensamble["id"] = response.json()["result"]["default_code"]

    # Split the folder path
    folder_path_parts = folder_path.split(os.sep)
    
    # Edit the last part of the folder path
    folder_path_parts[-1] = ensamble["id"]+ " " + folder_path_parts[-1]
    
    # Join the parts back together
    global new_folder
    new_folder = folder_path_parts[-1]

    os.rename(folder_path, new_folder)

def pieza_odoo(pieza):

    #send request to odoo

    global create_url

    #take the col 'quantity' of pieza
    _ = pieza.pop('quantity')
    
    #do request
    params_pieza = {"params": pieza}
    #add accept header
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    # Convert data to JSON format
    json_data = json.dumps(params_pieza)
    print(json_data)

    # Send POST request
    response = requests.get(create_url, data=json_data, headers=headers)

    # Check response
    try:
        if response.json()["result"]['status'] != 200:
            error_message = f"Envío de datos fallido. Mensaje de error: {response.json()['result']['message']}"
            messagebox.showerror("Data Sending Error", error_message)
            return
    except:
        error_message = f"Envío de datos fallido. Estado de error: {response}"
        messagebox.showerror("Data Sending Error", error_message)
        return

    # Get response data
    #response_data = response.json()
    pieza["id"] = response.json()["result"]["default_code"]

def update_url(producto):
    #file_url = new_folder + "/" + producto["name"] + ".SLDPRT"
    #send request to odoo
    global update_url
    
    #generar url
    global new_folder
    route_parts = producto["product_route"].split(os.sep)
    route_parts[-1] = new_folder
    producto["product_route"] = os.sep.join(route_parts)

    # Convert data to JSON format
    json_data = json.dumps(producto)

    # Send POST request
    response = requests.get(update_url, data=json_data)

    # Check response
    try:
        if response.json()["result"]['status'] != 200:
            error_message = f"Envío de datos fallido. Mensaje de error: {response.json()['result']['message']}"
            messagebox.showerror("Data Sending Error", error_message)
            return
    except:
        error_message = f"Envío de datos fallido. Estado de error: {response}"
        messagebox.showerror("Data Sending Error", error_message)
        return

#main donde inicia el procesamiento de la carpeta
def folder(input_folder_path):

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
           
    #procesar cada pieza sldprt
    for sldprt_file in sldprt_files:
        process_sldprt(sldprt_file, folder_path)

    #for pieza in piezas:
     #   pieza_odoo(pieza)
    
    #return
    
    #check if there is a sldasm file
    if sldasm_files:
        process_sldasm(sldasm_files, folder_path)
        ensamble_odoo(ensamble, folder_path)
        update_url(ensamble)

        #sólo si hay encamble se va a modificar la carpeta y la ruta
        for pieza in piezas:
            update_url(pieza)
        
        

    print("")
    print("")
    print("")

    print(ensamble)
    print(piezas)

    #finish program
    messagebox.showinfo("SolidWorks", "Proceso finalizado.")
    print("Proceso finalizado.")
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

    

    def on_drop(self, event):
        folder_path = event.data[1:-1]
        
        if os.path.isdir(folder_path):
            
            file_names = os.listdir(folder_path)
            
            sldprt_files = [file_name for file_name in file_names if file_name.endswith('.SLDPRT') and not file_name.startswith('~$')]
            sldasm_files = [file_name for file_name in file_names if file_name.endswith('.SLDASM') and not file_name.startswith('~$')]
            #if there are more than one sldasm file, show a message error
            if len(sldasm_files) > 1:
                self.drop_area.config(text="Hay más de un ensamblaje en la carpeta, envíe sólo uno.")
            elif sldprt_files:
                self.drop_area.config(text="Procesando carpeta...")
                time.sleep(1)
            
                # Call the function from another_module with the file path as a parameter
                folder(folder_path)
                #finish program
                self.destroy
                return

            else:
                self.drop_area.config(text="Ingrese una carpeta con piezas de SolidWorks.")
        else:
            self.drop_area.config(text="Ingrese una carpeta con piezas de SolidWorks.")

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

"""if __name__ == "__main__":
    app = SimpleGUI()
    app.mainloop()
    """


folder(r"C:\Users\Usuario\Downloads\04955 GAB-PEX-11")

data= {"params": 
{"name": "03619 GAB-PEX-11-B V2", 
"default_code": 550167, 
"product_tag_ids": "Piezas", 
"weight": 1.0420084037, 
"gross_weight": 1.1209409240831998, 
"volume": 0.13341977, 
"superficie": "301092.89", 
"broad": 251.52, 
"long": 634.04, 
"categ_id": "Chapa Galvanizada SAE 1010", 
"thickness": 0.9, 
"sale_ok": "true", 
"purchase_ok": "false", 
"product_route": "file:///C:/Users/Usuario/Downloads/04955%20GAB-PEX-11/G%2004956%20GAB-PEX-11-B%20V2.2%20CUERPO.SLDPRT", 
"bill_of_materials": [{
  "default_code": 20013, 
  "product_qty": 1.1209409240831998
}]}}