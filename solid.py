import os
import win32com.client

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

"""def check_and_open_solidworks_file(swApp, file_path, file_type):
    try:
        # Check if the document is already open
        is_open = swApp.IsOpened(file_path)
        is_open, model = swApp.IsDocument(file_path)

        if not is_open:
            # Document is not open, open it
            swApp.OpenDoc(file_path, file_type)  # Specify the appropriate file type (e.g., swDocumentPart)

            # Optionally, hide the SolidWorks window
            swApp.Visible = False

    except Exception as e:
        print(f"Error: {str(e)}")"""
        
def get_open_documents(swApp):
    try:
        # Access the Documents collection
        documents = swApp.Documents

        # Iterate through the open documents
        for doc in documents:
            # Print or process each open document
            print("Document Title:", doc.GetTitle())

    except Exception as e:
        print(f"Error: {str(e)}")

def folder(folder_path):
    # Connect to an existing SolidWorks instance or create a new one if not available
    try:
        swApp = win32com.client.GetObject("SldWorks.Application")
    except:
        swApp = win32com.client.Dispatch("SldWorks.Application")

    # Specify your macro name and part file path
    macro_name = r".\Envío de piezas a Odoo\main.swp"
    part_file_path = r"C:\Users\Usuario\Downloads\08978 Puerta Tablero.SLDPRT"
    # "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe" "C:\Users\Usuario\Downloads\08978 Puerta Tablero.SLDPRT"

    file_names = os.listdir(folder_path)
            
    sldprt_files = [file_name for file_name in file_names if file_name.endswith('.SLDPRT')]
    sldasm_files = [file_name for file_name in file_names if file_name.endswith('.SLDASM')]

    #check if there is a sldasm file
    if sldasm_files:
        
        #escribir la ruta en el archivo input
        path_file = r".\Envío de piezas a Odoo\Ruta.txt"
        sldasm_file_path = os.path.join(folder_path, sldasm_files[0])
        sldasm_file_path = sldasm_file_path.replace("\\", "/")
        sldasm_file_path_url = sldasm_file_path.replace(" ", "%20")
        sldasm_file_path_url = "file:///" + sldasm_file_path
        sldasm_file_path = r"C:\Users\Usuario\Downloads\04955 GAB-PEX-11\04955 GAB-PEX-11-B V2.2 ENSAMBLAJE.SLDASM"

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

        #verificar el archivo de errores
        
        #cerrar la pieza si no estaba abierta

        #recopilar los datos guardados

        #generar url ruta

        #enviar los datos a odoo
        #no enviar grosor para los ensambles

        #renombrar los archivos, el ensamble queda así?
            #luego de recibir el codigo del ensamble, renombrarlo y actualizar la ruta

folder(r"C:\Users\Usuario\Downloads\04955 GAB-PEX-11")

