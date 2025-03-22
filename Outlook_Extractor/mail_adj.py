"""
Extraer Archivos Adjuntos y URLs de Outlook
Autor: Carlos Cusi
Versión: 1.0
Descripción: Este script permite extraer archivos adjuntos y URLs de correos electrónicos
             en una carpeta de Outlook, filtrando por fechas y tipos de archivo.
"""
import os
import re
import time
import threading
import gc
from datetime import datetime, timedelta
from tkinter import Tk, Button, Label, Entry, messagebox, ttk
from tkcalendar import DateEntry
import win32com.client
import pythoncom

# Diccionario de categorías de archivo
TIPOS_DE_ARCHIVO = {
    "Excel": ["xls", "xlsx", "xlsm", "csv", "ods"],
    "Documentos": ["doc", "docx", "txt", "odt", "rtf"],
    "PDFs": ["pdf"],
    "Presentaciones": ["ppt", "pptx", "pps", "ppsx", "odp"],
    "Imágenes": ["jpg", "jpeg", "png", "gif", "bmp"],
    "Comprimidos": ["zip", "rar", "7z", "tar"],
    "Solo URLs": [],
    "Todas": []
}

REDES_SOCIALES = ["facebook.com", "twitter.com", "instagram.com", "linkedin.com", "tiktok.com", "pinterest.com", "snapchat.com", "reddit.com"]
URL_REGEX = re.compile(r"http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+")
MIN_IMAGE_SIZE = 20480  # 20 KB en bytes

def chunks(iterable, batch_size=50):
    batch = []
    for item in iterable:
        batch.append(item)
        if len(batch) == batch_size:
            yield batch
            batch = []
    if batch:
        yield batch

def process_folder(folder, search_term, file_type, save_folder, start_date, end_date, progress_callback=None):
    processed_items = 0
    errors = []
    urls_file_path = os.path.join(save_folder, "urls.txt")
    allowed_extensions = TIPOS_DE_ARCHIVO.get(file_type, [])
    url_count = 0
    urls_extracted = []

    items = (item for item in folder.Items if item.Class == 43 and start_date <= item.ReceivedTime.replace(tzinfo=None) <= end_date)
    for batch in chunks(items):
        for item in batch:
            try:
                if file_type == "Solo URLs" or file_type == "Todas":
                    urls = URL_REGEX.findall(str(item.Body or ""))
                    for url in urls[:50]:
                        if not any(rs in url.lower() for rs in REDES_SOCIALES):
                            if not search_term or search_term.lower() in url.lower():
                                urls_extracted.append(url)
                                url_count += 1
                if file_type != "Solo URLs":
                    for attachment in item.Attachments:
                        file_name = attachment.FileName.lower()
                        search_term_lower = search_term.lower() if search_term else ""
                        name_matches = not search_term or search_term_lower in file_name
                        extension_matches = any(file_name.endswith(f".{ext}") for ext in allowed_extensions) if allowed_extensions else True

                        # Filtrar imágenes menores a 20 KB
                        is_image = any(file_name.endswith(f".{ext}") for ext in TIPOS_DE_ARCHIVO["Imágenes"])
                        size_ok = not is_image or attachment.Size >= MIN_IMAGE_SIZE  # Guardar si no es imagen o si es >= 20 KB

                        if name_matches and extension_matches and size_ok:
                            email_date = item.ReceivedTime.replace(tzinfo=None).strftime("%d_%m")
                            base_file_name = os.path.splitext(file_name)[0]
                            file_ext = os.path.splitext(file_name)[1]
                            new_file_name = f"{base_file_name} ({email_date}){file_ext}"
                            file_path = os.path.join(save_folder, new_file_name)
                            counter = 1
                            while os.path.exists(file_path):
                                new_file_name = f"{base_file_name} ({email_date}) ({counter}){file_ext}"
                                file_path = os.path.join(save_folder, new_file_name)
                                counter += 1
                            attachment.SaveAsFile(file_path)
                            processed_items += 1
                del item
            except Exception as e:
                errors.append(f"Error en '{item.Subject}': {e}")
            gc.collect()
            if progress_callback:
                progress_callback(processed_items)

    if urls_extracted:
        with open(urls_file_path, "a", encoding="utf-8") as urls_file:
            fecha = datetime.now().strftime("%d/%m/%y")
            urls_file.write(f"{fecha} {url_count} elemento{'s' if url_count != 1 else ''} obtenido{'s' if url_count != 1 else ''}\n")
            for url in urls_extracted:
                urls_file.write(f"{url}\n")
            urls_file.write("\n")

    return processed_items, errors

def count_items(folder, start_date, end_date):
    return sum(1 for item in folder.Items if item.Class == 43 and start_date <= item.ReceivedTime.replace(tzinfo=None) <= end_date)

def start_extraction():
    button_extract.config(state="disabled")
    cal_start_date.config(state="disabled")
    cal_end_date.config(state="disabled")
    entry_search_term.config(state="disabled")
    combo_file_type.config(state="disabled")
    button_help.config(state="disabled")
    label_processing.config(state="normal")
    label_processing.config(text="Procesando...")

    start_date = cal_start_date.get_date()
    end_date = cal_end_date.get_date()
    search_term = entry_search_term.get()
    file_type = combo_file_type.get()

    start_date = datetime.combine(start_date, datetime.min.time())
    end_date = datetime.combine(end_date, datetime.max.time())

    def extraction_thread():
        processed_emails = 0
        processed_files = 0
        save_folder = None
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            time.sleep(2)  # Espera 2 segundos para estabilizar
            selected_folder = outlook.PickFolder()
            if not selected_folder:
                root.after(0, lambda: messagebox.showwarning("Advertencia", "No se seleccionó ninguna carpeta. Intenta de nuevo."))
                root.after(0, restore_controls)
                return

            save_folder = os.path.join(os.path.expanduser("~"), "Desktop", "Adjuntos Extraidos")
            os.makedirs(save_folder, exist_ok=True)
            try:
                _ = selected_folder.Items
            except AttributeError as e:
                root.after(0, lambda: messagebox.showerror("Error", f"La carpeta seleccionada '{selected_folder.Name}' no contiene elementos válidos: {str(e)}. Selecciona otra carpeta."))
                root.after(0, restore_controls)
                return

            total_items = count_items(selected_folder, start_date, end_date)
            processed_emails = total_items
            progress_bar["maximum"] = total_items
            start_time = time.time()
            processed_files, errors = process_folder(selected_folder, search_term, file_type, 
                                                     save_folder, start_date, end_date, 
                                                     lambda x: progress_bar.config(value=x))
            end_time = time.time()
            elapsed_time = str(timedelta(seconds=int(end_time - start_time)))
            progress_bar.config(value=total_items)
            root.update_idletasks()
            root.after(0, restore_controls)
            root.after(0, lambda: messagebox.showinfo("Extracción Completada", 
                                f"📧 Correos revisados: {processed_emails}\n"
                                f"📁 Archivos descargados: {processed_files}\n"
                                f"⏱️ Tiempo usado: {elapsed_time}\n"
                                f"📂 Carpeta: {save_folder}\n"
                                f"URLs guardadas en: {os.path.join(save_folder, 'urls.txt')}"))
            if errors:
                root.after(0, lambda: messagebox.showwarning("Advertencia", f"Se encontraron errores durante la extracción:\n" + "\n".join(errors[:5])))
            os.startfile(save_folder)
        except Exception as e:
            root.after(0, lambda e=e: messagebox.showerror("Error", f"No se pudo conectar con Outlook: {str(e)}. Asegúrate de que Outlook esté instalado y funcionando correctamente."))
            root.after(0, restore_controls)
        finally:
            pythoncom.CoUninitialize()
            gc.collect()

    def restore_controls():
        button_extract.config(state="normal")
        cal_start_date.config(state="normal")
        cal_end_date.config(state="normal")
        entry_search_term.config(state="normal")
        combo_file_type.config(state="normal")
        button_help.config(state="normal")
        label_processing.config(state="disabled")
        progress_bar.config(value=0)
        label_processing.config(text="Procesado")

    threading.Thread(target=extraction_thread, daemon=True).start()

def show_help():
    messagebox.showinfo("?", 
                        "Extraer Archivos y URLs de Outlook\n"
                        "Autor: Carlos Cusi\n"
                        "Versión: 1.0\n"
                        "Fecha: Marzo 2025\n\n"
                        "Este programa extrae archivos adjuntos y URLs de correos en Outlook.\n"
                        "1. Configura las fechas y el tipo de archivo.\n"
                        "2. Opcionalmente, ingresa un término de búsqueda.\n"
                        "3. Haz clic en 'Extraer Archivos y URLs' para seleccionar una carpeta y comenzar.\n"
                        "Los archivos se guardan en 'Adjuntos Extraidos' en el escritorio, y las URLs en 'urls.txt'.\n"
                        "Nota: Las imágenes menores a 20 KB (e.g., íconos de firmas) no se capturan.")

root = Tk()
root.title("Extraer Archivos Adjuntos y URLs de Outlook")
root.configure(bg="#2b2b2b")

style = ttk.Style()
style.theme_use("default")
style.configure("blue.Horizontal.TProgressbar", background="#0078D7", troughcolor="#2b2b2b", bordercolor="#2b2b2b")
style.configure("TCombobox", fieldbackground="#3a3a3a", background="#3a3a3a", foreground="#d3d3d3", arrowcolor="#d3d3d3")

label_start_date = Label(root, text="Fecha de inicio:", bg="#2b2b2b", fg="#d3d3d3")
label_start_date.grid(row=0, column=0, padx=10, pady=5)
cal_start_date = DateEntry(root, date_pattern="dd/mm/yy", background="#2a2a2a", foreground="white", borderwidth=0)
cal_start_date.grid(row=0, column=1, padx=10, pady=5)

label_end_date = Label(root, text="Fecha de fin:", bg="#2b2b2b", fg="#d3d3d3")
label_end_date.grid(row=1, column=0, padx=10, pady=5)
cal_end_date = DateEntry(root, date_pattern="dd/mm/yy", background="#2a2a2a", foreground="white", borderwidth=0)
cal_end_date.grid(row=1, column=1, padx=10, pady=5)

label_search_term = Label(root, text="Nombre contiene (opcional):", bg="#2b2b2b", fg="#d3d3d3")
label_search_term.grid(row=2, column=0, padx=10, pady=5)
entry_search_term = Entry(root, bg="#3a3a3a", fg="#d3d3d3", insertbackground="#d3d3d3")
entry_search_term.grid(row=2, column=1, padx=10, pady=5)

label_file_type = Label(root, text="Tipo de archivo:", bg="#2b2b2b", fg="#d3d3d3")
label_file_type.grid(row=3, column=0, padx=10, pady=5)
combo_file_type = ttk.Combobox(root, values=["Todas", "Solo URLs", "Excel", "Documentos", "PDFs", "Presentaciones", "Imágenes", "Comprimidos"], style="TCombobox")
combo_file_type.grid(row=3, column=1, padx=10, pady=5)
combo_file_type.set("Todas")

button_extract = Button(root, text="Extraer Archivos y URLs", command=start_extraction, 
                        bg="#1f6aa8", fg="#ffffff", activebackground="#2a85d6", relief="flat")
button_extract.grid(row=4, column=0, columnspan=2, pady=10)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate", style="blue.Horizontal.TProgressbar")
progress_bar.grid(row=5, column=0, columnspan=2, pady=10)

label_processing = Label(root, text="Procesando...", bg="#2b2b2b", fg="#d3d3d3", state="disabled")
label_processing.grid(row=6, column=0, columnspan=2, pady=5)

button_help = Button(root, text="?", command=show_help, 
                     bg="#1f6aa8", fg="#ffffff", activebackground="#2a85d6", relief="flat", width=3)
button_help.grid(row=7, column=0, padx=10, pady=5, sticky="sw")

label_version = Label(root, text="V1.0", bg="#2b2b2b", fg="#a9a9a9")
label_version.grid(row=7, column=1, sticky="se", padx=10, pady=5)

root.mainloop()