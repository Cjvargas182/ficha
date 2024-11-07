import tkinter as tk
from tkinter import messagebox
import qrcode
from PIL import Image, ImageTk
import openpyxl
from openpyxl import Workbook
import os
from datetime import datetime
import urllib.parse

def guardar_en_excel(nombre, area, rfc, f_nacimiento, enfermedades, medicamentos, alergias_a, alergias_f, g_sanguineo, c_nombre1, parentesco1, telefono1, c_nombre2, parentesco2, telefono2):
    archivo_excel = 'FICHA_TRABAJADOR.xlsx'

    if not os.path.exists(archivo_excel):
        wb = Workbook()
        ws = wb.active
        ws.title = "FICHA_TRABAJADOR"
        ws.append(["NOMBRE", "AREA", "RFC", "F_NACIMIENTO", "ENFERMEDADES", "MEDICAMENTOS", "ALERGIAS_A", "ALERGIAS_F", "G_SANGUINEO", "C_NOMBRE1", "PARENTESCO1", "TELEFONO1", "C_NOMBRE2", "PARENTESCO2", "TELEFONO2"])
        wb.save(archivo_excel)

    wb = openpyxl.load_workbook(archivo_excel)
    ws = wb.active
    ws.append([nombre, area, rfc, f_nacimiento, enfermedades, medicamentos, alergias_a, alergias_f, g_sanguineo, c_nombre1, parentesco1, telefono1, c_nombre2, parentesco2, telefono2])
    wb.save(archivo_excel)

# Crear la URL que contiene los datos como parámetros
def crear_url_alerta(nombre, area, rfc,f_nacimiento,enfermedades,medicamentos,alergias_a,alergias_f,g_sanguineo,c_nombre1,parentesco1,telefono1,c_nombre2,parentesco2,telefono2):
    base_url = "https://ficha-gold.vercel.app/"  # Cambia a tu dominio
    parametros = {
        "nombre": nombre,
        "area": area,
        "rfc": rfc,
        "f_nacimiento": f_nacimiento,
        "enfermedades": enfermedades,
        "medicamentos": medicamentos,
        "alergias_a": alergias_a,
        "alergias_f": alergias_f,
        "g_sanguineo": g_sanguineo,
        "c_nombre1": c_nombre1,
        "parentesco1": parentesco1,
        "telefono1": telefono1,
        "c_nombre2": c_nombre2,
        "parentesco2": parentesco2,
        "telefono2": telefono2
        
    }
    return f"{base_url}?{urllib.parse.urlencode(parametros)}"

# Generar el código QR usando la URL con los datos
def generar_qr():
    nombre = nombre_entry.get()
    area = area_entry.get()
    rfc = rfc_entry.get()
    f_nacimiento = f_nacimiento_entry.get()
    enfermedades = enfermedades_entry.get()
    medicamentos = medicamentos_entry.get()
    alergias_a = alergias_a_entry.get()
    alergias_f = alergias_f_entry.get()
    g_sanguineo = g_sanguineo_entry.get()
    c_nombre1 = c_nombre1_entry.get()
    parentesco1 = parentesco1_entry.get()
    telefono1 = telefono1_entry.get()
    c_nombre2 = c_nombre2_entry.get()
    parentesco2 = parentesco2_entry.get()
    telefono2 = telefono2_entry.get()

    if not all([nombre, area, rfc, f_nacimiento, enfermedades, medicamentos, alergias_a, alergias_f, g_sanguineo, c_nombre1, parentesco1, telefono1, c_nombre2, parentesco2, telefono2]):
        messagebox.showwarning("Campos incompletos", "Por favor, completa todos los campos.")
        return

    url_alerta = crear_url_alerta(nombre, area, rfc)
    qr = qrcode.make(url_alerta)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_archivo_qr = f"qr_{nombre}_{timestamp}.png"
    qr.save(nombre_archivo_qr)

    img = Image.open(nombre_archivo_qr)
    img = img.resize((200, 200))
    img_tk = ImageTk.PhotoImage(img)

    qr_label.config(image=img_tk)
    qr_label.image = img_tk

    guardar_en_excel(nombre, area, rfc, f_nacimiento, enfermedades, medicamentos, alergias_a, alergias_f, g_sanguineo, c_nombre1, parentesco1, telefono1, c_nombre2, parentesco2, telefono2)

    messagebox.showinfo("Éxito", f"Código QR generado y guardado como '{nombre_archivo_qr}'. Los datos han sido guardados en 'FICHA_TRABAJADOR.xlsx'.")

ventana = tk.Tk()
ventana.title("Generador de QR de Fichas")
ventana.geometry("400x850")

labels = [
    "Nombre", "Área", "RFC", "Fecha de nacimiento", "Enfermedades", "Medicamentos", "Alergias alimenticias", "Alergias farmacológicas",
    "Grupo sanguíneo", "Contacto 1 Nombre", "Contacto 1 Parentesco", "Contacto 1 Teléfono",
    "Contacto 2 Nombre", "Contacto 2 Parentesco", "Contacto 2 Teléfono"
]

entries = []

for i, label in enumerate(labels):
    tk.Label(ventana, text=label + ":").grid(row=i, column=0, padx=10, pady=5)
    entry = tk.Entry(ventana)
    entry.grid(row=i, column=1, padx=10, pady=5)
    entries.append(entry)

(
    nombre_entry, area_entry, rfc_entry, f_nacimiento_entry, enfermedades_entry, medicamentos_entry,
    alergias_a_entry, alergias_f_entry, g_sanguineo_entry, c_nombre1_entry, parentesco1_entry, telefono1_entry,
    c_nombre2_entry, parentesco2_entry, telefono2_entry
) = entries

generar_btn = tk.Button(ventana, text="Generar QR", command=generar_qr)
generar_btn.grid(row=len(labels), column=0, columnspan=2, padx=10, pady=10)

qr_label = tk.Label(ventana)
qr_label.grid(row=len(labels) + 1, column=0, columnspan=2, padx=10, pady=10)

ventana.mainloop()