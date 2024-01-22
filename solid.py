import os
import win32com.client

def run_solidworks_macro(swApp, macro_name, part_file_path):
    try:
        # Connect to SolidWorks
        swApp.Visible = True

        # Open the file
        swModel = swApp.OpenDoc(part_file_path, 1)  # 1 = swDocumentPart

        # Run your VBA macro
        macro_full_path = os.path.join(os.path.expanduser("~"), "AppData\\Roaming\\SolidWorks\\SolidWorks 2019\\macros", macro_name)
        swApp.RunMacro(macro_full_path, "macro01", "macro1")

        # Close the SolidWorks document
        #swApp.CloseDoc(swModel.GetTitle())

    except Exception as e:
        print(f"Error: {str(e)}")

def folder(folder_path):
    # Connect to an existing SolidWorks instance or create a new one if not available
    try:
        swApp = win32com.client.GetObject("SldWorks.Application")
    except:
        swApp = win32com.client.Dispatch("SldWorks.Application")

    # Specify your macro name and part file path
    macro_name = r"C:\Users\Usuario\Desktop\Envío de piezas a Odoo\macro0.swp"
    part_file_path = r"C:\Users\Usuario\Downloads\08978 Puerta Tablero.SLDPRT"

    # Run the SolidWorks macro without starting a new session
    #run_solidworks_macro(swApp, macro_name, part_file_path)

    file_names = os.listdir(folder_path)
            
    sldprt_files = [file_name for file_name in file_names if file_name.endswith('.SLDPRT')]
    sldasm_file = [file_name for file_name in file_names if file_name.endswith('.SLDASM')]

    #check if there is a sldasm file
    if sldasm_files:
        
        #escribir la ruta en el archivo input

        #abrir la pieza oculta

        #correr el macros

        #verificar el archivo de errores
        
        #cerrar la pieza si no estaba abierta

        #recopilar los datos guardados

        #generar url ruta

        #enviar los datos a odoo
        #no enviar grosor para los ensambles

        #renombrar los archivos, el ensamble queda así?
            #luego de recibir el codigo del ensamble, renombrarlo y actualizar la ruta