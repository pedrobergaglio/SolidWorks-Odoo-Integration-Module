import time
import os
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image, ImageTk
import socket
import tkinter.messagebox as messagebox
import solid

class SimpleGUI(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("METALUX - SOLIDWORKS A ODOO")
        self.geometry("600x300")

        # Set the background color
        self.configure(bg="white")

        # Add window icon
        icon_path = r"C:\Users\Usuario\Downloads\metalux-logo.ico"
        self.iconbitmap(icon_path)

        # Company logo label
        logo_path = r"C:\Users\Usuario\Downloads\metalux brand.png"
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
                solid.folder(folder_path)
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

if __name__ == "__main__":
    app = SimpleGUI()
    app.mainloop()
    
 