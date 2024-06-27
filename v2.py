import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tqdm import tqdm
import win32com.client
from datetime import datetime
import subprocess

# Detectar la ruta donde está el ejecutable o el script
application_path = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS

# Lista global para almacenar las carpetas seleccionadas
selected_folders = []

# Ruta del archivo de plantilla
template_path = ""

# Variable global para verificar conexión con Photoshop
photoshop_connected = False

# Función para verificar y conectar con Photoshop
def connect_to_photoshop():
    global photoshop_connected
    try:
        win32com.client.Dispatch("Photoshop.Application")
        photoshop_connected = True
        messagebox.showinfo("Conexión establecida", "Se ha establecido conexión con Adobe Photoshop.")
    except Exception as e:
        print(f"No se pudo conectar con Photoshop: {e}")
        messagebox.showerror("Error", f"No se pudo conectar con Photoshop: {e}")

# Función para seleccionar carpetas
def select_folders():
    if not photoshop_connected:
        messagebox.showerror("Error", "Primero debe establecer conexión con Adobe Photoshop.")
        return

    folder_path = filedialog.askdirectory()
    if folder_path:
        selected_folders.append(folder_path)
        folder_list.insert(tk.END, folder_path)

# Función para eliminar la carpeta seleccionada de la lista
def remove_selected_folder():
    selected_index = folder_list.curselection()
    if selected_index:
        selected_folders.pop(selected_index[0])
        folder_list.delete(selected_index)

# Función para seleccionar el archivo de plantilla
def select_template():
    if not photoshop_connected:
        messagebox.showerror("Error", "Primero debe establecer conexión con Adobe Photoshop.")
        return

    global template_path
    template_path = filedialog.askopenfilename(filetypes=[("PSD files", "*.psd")])
    if template_path:
        process_all_folders()
    else:
        messagebox.showerror("Error", "No se seleccionó el archivo de plantilla")

# Función para procesar todas las carpetas seleccionadas
def process_all_folders():
    if not photoshop_connected:
        messagebox.showerror("Error", "Primero debe establecer conexión con Adobe Photoshop.")
        return

    if selected_folders:
        # Crear carpeta "Robot Edición" con fecha y hora actuales
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        root_output_folder = os.path.join(os.getcwd(), f"Robot Edición {timestamp}")
        os.makedirs(root_output_folder, exist_ok=True)
        
        for folder_path in selected_folders:
            process_images(folder_path, root_output_folder)
        
        # Mostrar mensaje de finalización
        messagebox.showinfo("Proceso completado", f"Todas las imágenes han sido procesadas y guardadas en {root_output_folder}")
        
        # Abrir carpeta de salida
        open_folder(root_output_folder)
    else:
        messagebox.showerror("Error", "No se han seleccionado carpetas para procesar")

# Función para procesar imágenes con plantilla de Photoshop
def process_image_with_template(image_path, output_folder, photoshop, output_format, folder_name, text_content=None, text_layer_name=None, resize_dim=(1400, 1400)):
    try:
        # Abrir la plantilla
        doc = photoshop.Open(template_path)
        
        # Insertar imagen en la plantilla
        inserted_layer = doc.ArtLayers.Add()
        placed_doc = photoshop.Open(image_path)
        placed_doc.Selection.SelectAll()
        placed_doc.Selection.Copy()
        placed_doc.Close(2)  # Cerrar la imagen original
        doc.Paste()

        # Ajustar el tamaño de la capa insertada para que se ajuste al documento
        inserted_layer = doc.ActiveLayer
        inserted_layer.Resize(100, 100)  # Cambia el porcentaje si es necesario
        
        # Cambiar el contenido de la capa de texto si se proporciona
        if text_content and text_layer_name:
            try:
                text_layer = doc.ArtLayers[text_layer_name]
                text_layer.TextItem.contents = text_content
            except Exception as e:
                print(f"No se pudo actualizar el texto en la capa {text_layer_name}: {e}")

        # Renombrar el archivo según las reglas especificadas
        new_file_name = rename_file(os.path.basename(image_path), folder_name, output_format)
        output_path = os.path.join(output_folder, new_file_name)

        # Guardar el resultado en el formato seleccionado
        if output_format in ['jpg', 'jpeg']:
            options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
            options.Format = 6  # JPEG
            options.Quality = 100  # Valor de 0 a 100
            doc.Export(ExportIn=output_path, ExportAs=2, Options=options)
        elif output_format == 'png':
            options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
            options.Format = 13  # PNG
            options.PNG8 = False  # PNG-24
            doc.Export(ExportIn=output_path, ExportAs=2, Options=options)
        elif output_format == 'psd':
            doc.SaveAs(output_path)
        
        doc.Close(2)  # Cerrar el documento y no guardar los cambios en la plantilla

        print(f"Imagen guardada en: {output_path}")

    except Exception as e:
        print(f"Error {e} en {image_path}")

# Función para renombrar archivos
def rename_file(original_name, folder_name, output_format):
    # Separar el nombre del archivo por "_"
    parts = original_name.split("_")
    
    # Extraer el número y construir el nuevo nombre
    if len(parts) > 2:
        number = parts[1]
        new_name = f"{folder_name}-{number}.{output_format}"
        return new_name
    else:
        return original_name

# Función para procesar imágenes
def process_images(input_folder, root_output_folder):
    try:
        photoshop = win32com.client.Dispatch("Photoshop.Application")
    except Exception as e:
        print(f"No se pudo iniciar Photoshop: {e}")
        messagebox.showerror("Error", f"No se pudo iniciar Photoshop: {e}")
        return

    # Crear carpeta de salida con el mismo nombre que la carpeta de entrada dentro de "Robot Edición"
    output_folder_name = os.path.basename(input_folder)
    output_folder = os.path.join(root_output_folder, output_folder_name)
    os.makedirs(output_folder, exist_ok=True)

    # Listar imágenes en la carpeta de entrada
    image_files = [x for x in os.listdir(input_folder) if x.lower().endswith(('png', 'jpg', 'jpeg', 'bmp', 'gif'))]

    # Obtener formato de salida seleccionado
    output_format = format_var.get()

    # Procesar imágenes con Photoshop
    for index, image_file in enumerate(tqdm(image_files, desc=f"Procesando imágenes en {input_folder}"), start=1):
        input_path = os.path.join(input_folder, image_file)
        try:
            process_image_with_template(input_path, output_folder, photoshop, output_format, output_folder_name, text_content="Ejemplo de texto", text_layer_name="Facts")
        except Exception as e:
            print(f"No se pudo procesar la imagen {input_path}: {e}")

# Función para abrir la carpeta
def open_folder(path):
    if os.name == 'nt':  # Para Windows
        os.startfile(path)
    elif os.name == 'posix':  # Para macOS y Linux
        subprocess.call(['open' if sys.platform == 'darwin' else 'xdg-open', path])

# Configurar interfaz gráfica
root = tk.Tk()
root.title("Procesador de Imágenes")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(padx=10, pady=10)

label = tk.Label(frame, text="Conectar con Adobe Photoshop:")
label.grid(row=0, column=0, padx=5, pady=5)

connect_button = tk.Button(frame, text="Conectar", command=connect_to_photoshop)
connect_button.grid(row=0, column=1, padx=5, pady=5)

label = tk.Label(frame, text="Ingrese la ruta de las carpetas con las imágenes y presione 'Seleccionar carpeta':")
label.grid(row=1, column=0, padx=5, pady=5)

folder_list = tk.Listbox(frame, width=50, height=10)
folder_list.grid(row=2, column=0, padx=5, pady=5, columnspan=2)

remove_button = tk.Button(frame, text="Eliminar carpeta seleccionada", command=remove_selected_folder)
remove_button.grid(row=3, column=1, padx=5, pady=5)

select_folders_button = tk.Button(frame, text="Seleccionar carpeta", command=select_folders)
select_folders_button.grid(row=4, column=0, columnspan=2, pady=5)

process_button = tk.Button(frame, text="Seleccionar plantilla y procesar carpetas", command=select_template)
process_button.grid(row=5, column=0, columnspan=2, pady=10)

label_format = tk.Label(frame, text="Seleccione el formato de salida:")
label_format.grid(row=6, column=0, padx=5, pady=5)

format_var = tk.StringVar(value="jpg")
format_combobox = ttk.Combobox(frame, textvariable=format_var, values=["jpg", "jpeg", "png", "psd"], state="readonly")
format_combobox.grid(row=6, column=1, padx=5, pady=5)

root.mainloop()
