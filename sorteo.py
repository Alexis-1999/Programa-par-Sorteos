import random
import tkinter as tk
from PIL import Image, ImageTk
import os
import openpyxl

def cargar_participantes(archivo, factura_col, nombre_col):
    try:
        _, extension = os.path.splitext(archivo)
        if extension.lower() not in ('.xls', '.xlsx'):
            raise ValueError("El archivo no tiene la extensión .xls o .xlsx")

        participantes_wb = openpyxl.load_workbook(archivo)
        participantes_sheet = participantes_wb.active

        participantes = []
        for row in participantes_sheet.iter_rows(min_row=3, values_only=True):
            participante = {'FACTURA': row[factura_col], 'NOMBRE': row[nombre_col]}
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
    ventana_resultado = tk.Toplevel()
    ventana_resultado.title("Resultado del Sorteo")

    if fondo_path and os.path.exists(fondo_path):
        fondo_img = Image.open(fondo_path)
        fondo_photo = ImageTk.PhotoImage(fondo_img)
        fondo_label = tk.Label(ventana_resultado, image=fondo_photo)
        fondo_label.image = fondo_photo
        fondo_label.pack(fill=tk.BOTH, expand=True)

    etiqueta_sorteo = tk.Label(ventana_resultado, text="¡Resultado del Sorteo!", font=("Arial", 20), fg="red")
    etiqueta_sorteo.place(relx=0.5, rely=0.2, anchor=tk.CENTER)

    etiqueta_ganador_texto = f"¡El ganador es {ganador['NOMBRE']} con el número de factura {ganador['FACTURA']}!"
    etiqueta_ganador = tk.Label(ventana_resultado, text=etiqueta_ganador_texto, font=("Arial", 15), fg="black", wraplength=400)
    etiqueta_ganador.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    guardar_ganador(ganador, archivo_ganadores)

    ventana_resultado.mainloop()

def cargar_participantes_sucursal(script_dir, sucursal):
    archivo = ""
    factura_col = 0
    nombre_col = 0

    if sucursal == 1:
        archivo = os.path.join(script_dir, 'FACTURASNEMBYAL25.xlsx')
        factura_col = 4  # Columna E
        nombre_col = 5   # Columna F
    elif sucursal == 2:
        archivo = os.path.join(script_dir, 'FACTURASSANLOAL25.xlsx')
        factura_col = 4  # Columna E
        nombre_col = 5   # Columna F
    elif sucursal == 3:
        archivo = os.path.join(script_dir, 'KM6AL25.xlsx')
        factura_col = 3  # Columna E
        nombre_col = 4   # Columna F

    participantes = cargar_participantes(archivo, factura_col, nombre_col)
    return participantes

def main():
    script_dir = os.path.dirname(os.path.realpath(__file__))
    archivo_ganadores = os.path.join(script_dir, 'ganadores.txt')
    fondo_path = os.path.join(script_dir, 'fondo.jpeg')

    def sortear_y_mostrar_resultado(sucursal):
        participantes_sucursal = cargar_participantes_sucursal(script_dir, sucursal)
        if not participantes_sucursal:
            print(f"No hay participantes o hay un error en la carga de la sucursal {sucursal}.")
            return

        ganador = random.choice(participantes_sucursal)
        mostrar_resultado_ganador(ganador, archivo_ganadores, fondo_path)

    def cargar_y_sortear_sucursal1():
        sortear_y_mostrar_resultado(1)

    def cargar_y_sortear_sucursal2():
        sortear_y_mostrar_resultado(2)

    def cargar_y_sortear_sucursal3():
        sortear_y_mostrar_resultado(3)

    ventana_principal = tk.Tk()
    ventana_principal.title("Sorteo por Sucursal")

    boton_sucursal1 = tk.Button(ventana_principal, text="Sucursal 1", command=cargar_y_sortear_sucursal1)
    boton_sucursal1.pack(side=tk.LEFT, padx=10)

    boton_sucursal2 = tk.Button(ventana_principal, text="Sucursal 2", command=cargar_y_sortear_sucursal2)
    boton_sucursal2.pack(side=tk.LEFT, padx=10)

    boton_sucursal3 = tk.Button(ventana_principal, text="Sucursal 3", command=cargar_y_sortear_sucursal3)
    boton_sucursal3.pack(side=tk.LEFT, padx=10)

    ventana_principal.mainloop()

if __name__ == "__main__":
    main()
