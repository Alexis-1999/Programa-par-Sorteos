import random
import tkinter as tk
from PIL import Image, ImageTk
import os
import ctypes
import pandas as pd
import openpyxl

def cargar_participantes(archivo):
    try:
        # Validar que el archivo sea de tipo Excel
        _, extension = os.path.splitext(archivo)
        if extension.lower() not in ('.xls', '.xlsx'):
            raise ValueError("El archivo no tiene la extensión .xls o .xlsx")

        participantes_wb = openpyxl.load_workbook(archivo)
        participantes_sheet = participantes_wb.active

        # Obtener los datos de las columnas E y G a partir de la fila 3
        participantes = []
        for row in participantes_sheet.iter_rows(min_row=3, values_only=True):
            participante = {'NOMBRE': row[4], 'FACTURA': row[6]}
            participantes.append(participante)

        participantes_wb.close()
    except FileNotFoundError:
        print(f"Error: El archivo '{archivo}' no se encuentra.")
        participantes = []
    except ValueError as ve:
        print(f"Error: {ve}")
        participantes = []
    except Exception as e:
        print(f"Error al cargar participantes: {str(e)}")
        participantes = []

    return participantes

def guardar_ganador(ganador, archivo_ganadores):
    with open(archivo_ganadores, 'a') as file:
        file.write(f"{ganador['NOMBRE']},{ganador['FACTURA']}\n")

def mostrar_resultado_ganador(ganador, archivo_ganadores, fondo_path=None):
    ventana_resultado = tk.Tk()
    ventana_resultado.title("Resultado del Sorteo")

    # Cambiar el icono de la ventana
    icono_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icon.ico')
    if os.path.exists(icono_path):
        ventana_resultado.iconbitmap(default=icono_path)

    if fondo_path and os.path.exists(fondo_path):
        fondo_img = Image.open(fondo_path)
        fondo_photo = ImageTk.PhotoImage(fondo_img)
        fondo_label = tk.Label(ventana_resultado, image=fondo_photo)
        fondo_label.image = fondo_photo
        fondo_label.pack(fill=tk.BOTH, expand=True)

    etiqueta_ganador = tk.Label(ventana_resultado, text=f"¡Ganador: {ganador['NOMBRE']} con número de factura: {ganador['FACTURA']}!", font=("Arial", 20), fg="black")
    etiqueta_ganador.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    guardar_ganador(ganador, archivo_ganadores)

    ventana_resultado.update_idletasks()
    width = ventana_resultado.winfo_width()
    height = ventana_resultado.winfo_height()
    x = (ventana_resultado.winfo_screenwidth() - width) // 2
    y = (ventana_resultado.winfo_screenheight() - height) // 2
    ventana_resultado.geometry('+{}+{}'.format(x, y))

    ventana_resultado.mainloop()

def main():
    script_dir = os.path.dirname(os.path.realpath(__file__))
    archivo_participantes = os.path.join(script_dir, 'participantes.xlsx')
    archivo_ganadores = os.path.join(script_dir, 'ganadores.txt')
    fondo_path = os.path.join(script_dir, 'fondo.jpeg')

    participantes = cargar_participantes(archivo_participantes)

    if not participantes:
        print("No hay participantes o hay un error en la carga.")
        return

    ganador = random.choice(participantes)

    mostrar_resultado_ganador(ganador, archivo_ganadores, fondo_path)

if __name__ == "__main__":
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    main()
