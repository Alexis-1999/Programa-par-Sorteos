import tkinter.ttk as ttk
from PIL import Image, ImageTk
import random
import tkinter as tk
import os
import openpyxl
import threading
import time

def cargar_participantes(archivo, factura_col, nombre_col):
    try:
        _, extension = os.path.splitext(archivo)
        if extension.lower() not in ('.xls', '.xlsx'):
            raise ValueError("El archivo no tiene la extensión .xls o .xlsx")

        participantes_wb = openpyxl.load_workbook(archivo)
        participantes_sheet = participantes_wb.active

        participantes = []
        for row in participantes_sheet.iter_rows(min_row=3, values_only=True):
            nombre = row[nombre_col]
            factura = row[factura_col]

            # Verificar si el nombre es "SIN NOMBRE" y omitirlo
            if nombre != "SIN NOMBRE":
                participante = {'FACTURA': factura, 'NOMBRE': nombre}
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

def mostrar_resultado_ganador(ventana_principal, ganador, archivo_ganadores, fondo_path=None, loading_label=None, ventana_width=800, ventana_height=600):
    ventana_principal.withdraw()  # Ocultar la ventana principal temporalmente

    ventana_resultado = tk.Toplevel()
    ventana_resultado.title("Resultado del Sorteo")

    ruta_icono = "icon.ico"
    if os.path.exists(ruta_icono):
        ventana_resultado.iconbitmap(default=ruta_icono)

    if fondo_path and os.path.exists(fondo_path):
        fondo_img = Image.open(fondo_path)
        fondo_width, fondo_height = fondo_img.size

        fondo_photo = ImageTk.PhotoImage(fondo_img)

        fondo_label = tk.Label(ventana_resultado, image=fondo_photo)
        fondo_label.image = fondo_photo
        fondo_label.place(x=0, y=0, relwidth=1, relheight=1)  # Asegurar que la imagen de fondo cubra toda la ventana

        etiqueta_sorteo = tk.Label(ventana_resultado, text="¡Resultado del Sorteo!", font=("Arial", 20), fg="red")
        etiqueta_sorteo.place(relx=0.5, rely=0.1, anchor=tk.CENTER)  # Centrar el texto del sorteo en la parte superior de la ventana

        marco_central = tk.Frame(ventana_resultado)
        marco_central.place(relx=0.5, rely=0.4, anchor=tk.CENTER)  # Centrar el marco en el medio de la ventana

    if ganador['NOMBRE'] != "SIN NOMBRE":
        etiqueta_ganador_texto = f"¡El ganador es {ganador['NOMBRE']} con el número de factura {ganador['FACTURA']}!"
    else:
        etiqueta_ganador_texto = f"¡El ganador es el número de factura {ganador['FACTURA']}!"

    etiqueta_ganador = tk.Text(marco_central, font=("Arial", 13), fg="black", wrap="word", height=2, width=40)
    etiqueta_ganador.insert(tk.END, etiqueta_ganador_texto)
    etiqueta_ganador.configure(state='disabled', bd=0, highlightthickness=0)
    etiqueta_ganador.pack()


    # Establecer el fondo transparente y eliminar el borde y el resaltado
    etiqueta_ganador.config(bd=0, highlightthickness=0)

    guardar_ganador(ganador, archivo_ganadores)

    style = ttk.Style()
    style.configure("TButton",
                    font=("Arial", 12),
                    padding=10,
                    foreground="black",
                    background="red",
                    border="white 8px groove",
                    )

    boton_cerrar = ttk.Button(marco_central, text="Cerrar", command=ventana_resultado.destroy, style="TButton")
    boton_cerrar.pack()

    ventana_resultado.geometry(f"{ventana_width}x{ventana_height}")  # Establecer tamaño de la ventana según los parámetros proporcionados

    if loading_label:
        loading_label.config(text="")

    # Centrar la ventana en la pantalla
    ventana_resultado.update_idletasks()
    width = ventana_resultado.winfo_width()
    height = ventana_resultado.winfo_height()
    x = (ventana_resultado.winfo_screenwidth() // 2) - (width // 2)
    y = (ventana_resultado.winfo_screenheight() // 2) - (height // 2)
    ventana_resultado.geometry(f"+{x}+{y}")

    ventana_principal.deiconify()  # Mostrar nuevamente la ventana principal



def cargar_participantes_sucursal(script_dir, sucursal):
    archivo = ""
    factura_col = 0
    nombre_col = 0

    if sucursal == 1:
        archivo = os.path.abspath('FACTURASÑEMBYF.xlsx')
        factura_col = 4
        nombre_col = 5
    elif sucursal == 2:
        archivo = os.path.abspath('SANLOFACTURAS.xlsx')
        factura_col = 4
        nombre_col = 5
    elif sucursal == 3:
        archivo = os.path.abspath('FACTURAS07KM6.xlsx')
        factura_col = 4
        nombre_col = 5

    participantes = cargar_participantes(archivo, factura_col, nombre_col)
    return participantes

def main():
    script_dir = os.path.dirname(os.path.realpath(__file__))
    archivo_ganadores = os.path.abspath('ganadores.txt')
    fondo_path = os.path.abspath('fondo.jpeg')

    def sortear_y_mostrar_resultado(sucursal, loading_label):
        loading_label.config(text="SORTEANDO GANADOR DE LA PROMO...")

        participantes_sucursal = cargar_participantes_sucursal(script_dir, sucursal)
        if not participantes_sucursal:
            loading_label.config(text=f"No hay participantes o hay un error en la carga de la sucursal {sucursal}.")
            return

        time.sleep(10)  # Pausa de 10 segundos

        ganador = random.choice(participantes_sucursal)
        mostrar_resultado_ganador(ventana_principal, ganador, archivo_ganadores, fondo_path, loading_label, ventana_width=800, ventana_height=800)

    def cargar_y_sortear_sucursal(sucursal, loading_label):
        threading.Thread(target=lambda: sortear_y_mostrar_resultado(sucursal, loading_label)).start()

    ventana_principal = tk.Tk()
    ventana_principal.title("Sorteo por Sucursal")

    # Añadir el siguiente bloque para establecer el icono
    ruta_icono = "icon.ico"  # Cambia esto con la ruta de tu icono
    if os.path.exists(ruta_icono):
        ventana_principal.iconbitmap(default=ruta_icono)
        ventana_principal.wm_iconbitmap(ruta_icono)  # Establece el ícono para la barra de tareas

    loading_label = tk.Label(ventana_principal, text="", font=("Arial", 12), fg="blue")
    loading_label.pack(pady=10)

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
    import ctypes
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    main()