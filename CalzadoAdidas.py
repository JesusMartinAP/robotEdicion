import os
import sys
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tqdm import tqdm
import win32com.client
from datetime import datetime
import cv2  # OpenCV

application_path = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS

selected_folders = []
template_path = ""

def drop(event):
    paths = root.tk.splitlist(event.data)
    for path in paths:
        if os.path.isdir(path):
            selected_folders.append(path)
            folder_list.insert(tk.END, path)

def remove_selected_folder():
    selected_index = folder_list.curselection()
    if selected_index:
        selected_folders.pop(selected_index[0])
        folder_list.delete(selected_index)

def select_folders():
    folder_path = filedialog.askdirectory()
    if folder_path:
        selected_folders.append(folder_path)
        folder_list.insert(tk.END, folder_path)

def select_template():
    global template_path
    template_path = filedialog.askopenfilename(filetypes=[("PSD files", "*.psd")])
    if template_path:
        process_all_folders()
    else:
        messagebox.showerror("Error", "No se seleccionó el archivo de plantilla")

def process_all_folders():
    if selected_folders:
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        root_output_folder = os.path.join(os.getcwd(), f"Robot Edición {timestamp}")
        os.makedirs(root_output_folder, exist_ok=True)
        
        for folder_path in selected_folders:
            process_images(folder_path, root_output_folder)
        
        messagebox.showinfo("Proceso completado", f"Todas las imágenes han sido procesadas y guardadas en {root_output_folder}")
        open_folder(root_output_folder)
    else:
        messagebox.showerror("Error", "No se han seleccionado carpetas para procesar")

def process_image_with_template(image_path, output_folder, photoshop, output_format, folder_name, text_content=None, text_layer_name=None, resize_dim=(1400, 1400)):
    try:
        # Asegurarse de que el archivo de imagen se recorta correctamente antes de la plantilla
        if "_10_" in image_path:
            recortada_path = process_image_10_with_opencv(image_path)
            image_path = recortada_path

        time.sleep(1)
        doc = photoshop.Open(template_path)
        inserted_layer = doc.ArtLayers.Add()
        placed_doc = photoshop.Open(image_path)
        placed_doc.Selection.SelectAll()
        placed_doc.Selection.Copy()
        placed_doc.Close(2)
        doc.Paste()

        inserted_layer = doc.ActiveLayer
        inserted_layer.Resize(100, 100)
        
        if text_content and text_layer_name:
            try:
                text_layer = doc.ArtLayers[text_layer_name]
                text_layer.TextItem.contents = text_content
            except Exception as e:
                print(f"No se pudo actualizar el texto en la capa {text_layer_name}: {e}")

        new_file_name = rename_file(os.path.basename(image_path), folder_name, output_format)
        output_path = os.path.join(output_folder, new_file_name)

        if output_format in ['jpg', 'jpeg']:
            options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
            options.Format = 6
            options.Quality = 100
            doc.Export(ExportIn=output_path, ExportAs=2, Options=options)
        elif output_format == 'png':
            options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
            options.Format = 13
            options.PNG8 = False
            doc.Export(ExportIn=output_path, ExportAs=2, Options=options)
        elif output_format == 'psd':
            doc.SaveAs(output_path)
        
        doc.Close(2)
        print(f"Imagen guardada en: {output_path}")

    except Exception as e:
        print(f"Error {e} en {image_path}")

def process_image_10_with_opencv(image_path):
    try:
        img = cv2.imread(image_path)
        height, width = img.shape[:2]
        half_height = height // 2

        # Recortar la mitad superior
        img_cropped = img[:half_height, :]

        # Guardar la imagen recortada
        cropped_image_path = image_path.replace(".jpg", "_cropped.jpg")
        cv2.imwrite(cropped_image_path, img_cropped)
        return cropped_image_path

    except Exception as e:
        print(f"Error al recortar la imagen 10 con OpenCV: {e}")
        return image_path

def rename_file(original_name, folder_name, output_format):
    parts = original_name.split("_")
    if len(parts) > 2:
        number = parts[1]
        new_name = f"{folder_name}-{number}.{output_format}"
        return new_name
    else:
        return original_name

def process_images(input_folder, root_output_folder):
    photoshop = win32com.client.Dispatch("Photoshop.Application")

    output_folder_name = os.path.basename(input_folder)
    output_folder = os.path.join(root_output_folder, output_folder_name)
    os.makedirs(output_folder, exist_ok=True)

    image_files = [x for x in os.listdir(input_folder) if x.lower().endswith(('png', 'jpg', 'jpeg', 'bmp', 'gif'))]

    output_format = format_var.get()

    for index, image_file in enumerate(tqdm(image_files, desc=f"Procesando imágenes en {input_folder}"), start=1):
        input_path = os.path.join(input_folder, image_file)
        try:
            process_image_with_template(input_path, output_folder, photoshop, output_format, output_folder_name, text_content="Ejemplo de texto", text_layer_name="Facts")
        except Exception as e:
            print(f"No se pudo procesar la imagen {input_path}: {e}")
        time.sleep(1)  # Añadir un pequeño retraso entre cada procesamiento de imagen

def open_folder(path):
    if os.name == 'nt':
        os.startfile(path)
    elif os.name == 'posix':
        subprocess.call(['open' if sys.platform == 'darwin' else 'xdg-open', path])

root = TkinterDnD.Tk()
root.title("Procesador de Imágenes")

format_var = tk.StringVar(value="jpg")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(padx=10, pady=10)

label = tk.Label(frame, text="Arrastra y suelta las carpetas con las imágenes aquí:")
label.grid(row=0, column=0, padx=5, pady=5)

folder_list = tk.Listbox(frame, width=50, height=10)
folder_list.grid(row=1, column=0, padx=5, pady=5, columnspan=3)
folder_list.drop_target_register(DND_FILES)
folder_list.dnd_bind('<<Drop>>', drop)

remove_button = tk.Button(frame, text="Eliminar carpeta seleccionada", command=remove_selected_folder)
remove_button.grid(row=2, column=0, padx=5, pady=5)

add_button = tk.Button(frame, text="Agregar carpetas", command=select_folders)
add_button.grid(row=2, column=1, padx=5, pady=5)

template_button = tk.Button(frame, text="Seleccionar plantilla PSD", command=select_template)
template_button.grid(row=3, column=0, padx=5, pady=5)

format_label = tk.Label(frame, text="Formato de salida:")
format_label.grid(row=4, column=0, padx=5, pady=5)

format_options = ttk.Combobox(frame, textvariable=format_var, values=["jpg", "png", "psd"])
format_options.grid(row=4, column=1, padx=5, pady=5)

root.mainloop()
