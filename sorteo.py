import ctypes
import tkinter.ttk as ttk
import random
import tkinter as tk
from PIL import Image, ImageTk
import os
import openpyxl
import threading

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

def centrar_ventana(ventana, fondo_img):
    ventana.update_idletasks()
    width = ventana.winfo_screenwidth()
    height = ventana.winfo_screenheight()
    x = (width - fondo_img.width) // 2
    y = (height - fondo_img.height) // 2
    ventana.geometry('{}x{}+{}+{}'.format(fondo_img.width, fondo_img.height, x, y))

def mostrar_resultado_ganador(ventana_principal, ganador, archivo_ganadores, fondo_path=None, loading_label=None):
    ventana_resultado = tk.Toplevel()
    ventana_resultado.title("Resultado del Sorteo")

    # Agregar el icono a la ventana
    ruta_icono = "icon.ico"  # Reemplazar con la ruta correcta del archivo .ico
    if os.path.exists(ruta_icono):
        ventana_resultado.iconbitmap(default=ruta_icono)

    fondo_img = None
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

    # Crear un estilo
    style = ttk.Style()
    style.configure("TButton",
                font=("Arial", 12),
                padding=10,
                foreground="black",
                background="red",
                border= "white 8px groove",
                )

    # Crear el botón con el estilo
    boton_cerrar = ttk.Button(ventana_resultado, text="Cerrar", command=ventana_resultado.destroy, style="TButton")
    boton_cerrar.place(relx=0.5, rely=0.8, anchor=tk.CENTER)

    if fondo_img:
        centrar_ventana(ventana_resultado, fondo_img)

    if loading_label:
        loading_label.config(text="")
    
    # Mostrar la ventana principal nuevamente
    ventana_principal.deiconify()

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
        factura_col = 4  # Columna E
        nombre_col = 5   # Columna F

    participantes = cargar_participantes(archivo, factura_col, nombre_col)
    return participantes

def main():
    script_dir = os.path.dirname(os.path.realpath(__file__))
    archivo_ganadores = os.path.join(script_dir, 'ganadores.txt')
    fondo_path = os.path.join(script_dir, 'fondo.jpeg')

    def sortear_y_mostrar_resultado(sucursal, loading_label):
        loading_label.config(text="Cargando participantes, espere...")

        participantes_sucursal = cargar_participantes_sucursal(script_dir, sucursal)
        if not participantes_sucursal:
            loading_label.config(text=f"No hay participantes o hay un error en la carga de la sucursal {sucursal}.")
            return

        ganador = random.choice(participantes_sucursal)
        
        ventana_resultado = tk.Toplevel()
        ventana_resultado.title("Resultado del Sorteo")
        ventana_resultado.withdraw()  # Ocultar la ventana de resultado

        ventana_principal.withdraw()  # Ocultar la ventana principal
        loading_label.config(text="Mostrando resultado...")

        # Mostrar el resultado sin esperar
        mostrar_resultado_ganador(ventana_principal, ganador, archivo_ganadores, fondo_path, loading_label)

    def cargar_y_sortear_sucursal(sucursal, loading_label):
        threading.Thread(target=lambda: sortear_y_mostrar_resultado(sucursal, loading_label)).start()

    ventana_principal = tk.Tk()
    ventana_principal.title("Sorteo por Sucursal")

    loading_label = tk.Label(ventana_principal, text="", font=("Arial", 12), fg="blue")
    loading_label.pack(pady=10)

    # Utilizar grid para organizar los botones de manera más flexible
    botones_frame = tk.Frame(ventana_principal)
    botones_frame.pack(pady=10)

    boton_sucursal1 = tk.Button(botones_frame, text="Sucursal Ñemby", command=lambda: cargar_y_sortear_sucursal(1, loading_label), padx=20, pady=10)
    boton_sucursal1.grid(row=0, column=0, padx=10)

    boton_sucursal2 = tk.Button(botones_frame, text="Sucursal San Lorenzo", command=lambda: cargar_y_sortear_sucursal(2, loading_label), padx=20, pady=10)
    boton_sucursal2.grid(row=0, column=1, padx=10)

    boton_sucursal3 = tk.Button(botones_frame, text="Sucursal KM6", command=lambda: cargar_y_sortear_sucursal(3, loading_label), padx=20, pady=10)
    boton_sucursal3.grid(row=0, column=2, padx=10)

    ventana_principal.mainloop()

if __name__ == "__main__":
    # Ocultar la ventana de la consola
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    main()
