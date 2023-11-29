import tkinter as tk
from tkinter import ttk, PhotoImage
import subprocess

def abrir_modulo(nombre_modulo):
    if hasattr(subprocess, 'STARTF_USESHOWWINDOW'):  # Verifica si es Windows
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        subprocess.Popen(["python", "-m", nombre_modulo], startupinfo=startupinfo)
    else:
        subprocess.Popen(["python", "-m", nombre_modulo])
        
def mostrar_instrucciones(accion):
    popup = tk.Toplevel(ventana)
    popup.title(f"Instrucciones - {accion}")

    instrucciones_texto = f"Informacion sobre {accion}"
    label_texto = ttk.Label(popup, text=instrucciones_texto, font=('Calibri', 18, 'bold'), foreground='#2196F3')
    label_texto.pack(padx=20, pady=10)

    frame_imagenes = ttk.Frame(popup)
    frame_imagenes.pack(pady=20)

    if accion == "Crear":
        imagen_path = "Crear.png"
        imagen = PhotoImage(file=imagen_path)
        imagen = imagen.subsample(1, 1)

        label_imagen = ttk.Label(frame_imagenes, image=imagen)
        label_imagen.image = imagen
        label_imagen.pack(side=tk.LEFT, padx=20, pady=20)

    elif accion == "Abrir":
        imagen_path_1 = "segunda.png"
        imagen_path_2 = "primera.png"

        imagen_1 = PhotoImage(file=imagen_path_1)
        imagen_1 = imagen_1.subsample(1, 1)

        label_imagen_1 = ttk.Label(frame_imagenes, image=imagen_1)
        label_imagen_1.image = imagen_1
        label_imagen_1.pack(side=tk.LEFT, padx=20, pady=20)

        imagen_2 = PhotoImage(file=imagen_path_2)
        imagen_2 = imagen_2.subsample(1, 1)

        label_imagen_2 = ttk.Label(frame_imagenes, image=imagen_2)
        label_imagen_2.image = imagen_2
        label_imagen_2.pack(side=tk.LEFT, padx=20, pady=20)
        
    elif accion == "Acerca":
        instrucciones_texto = "Nombre de la aplicación: CaEx\nVersión: 1.0.0\n\nDescripción:\nEs una App para crear y editar Archivos Excel. Tiene por defecto un Formulario donde puedo agregar datos y cuando este listo, se creara un archivo Excel.\nTiene dos interfaces, la primera es el Formulario donde puedo agregar datos, guardarlos y cuando ya quiera crear el Excel, se podra crear y en la misma \ncarpeta donde esta el Ejecutable se guardaran los Excel. La segunda te permite elegir un archivo de Excel y te mostrara los nombres que tendran las \ncolumnas y cajas de texto donde se podra ingresar datos,luego se guardan los datos donde se mostraran los datos en nuevas filas. Estas dos interfaces \nestán manejadas por otra donde tiene los botones llama a su respectiva ventana.\n\nDesarrollado por: Yani Isaias\n\nCaracterísticas principales:\n- Guarda y crea archivos Excel.\n- Poder elegir archivos Excel y agregarle nuevos datos.\n- Cambiarle el nombre a los campos donde se crean los Excel.\n\nPara obtener más información, visite nuestro sitio web: [https://www.instagram.com/isaiassss.04/]"
        label_texto = ttk.Label(popup, text=instrucciones_texto, font=('Calibri', 14, 'bold'), foreground='#000')
        label_texto.pack(padx=20, pady=10)

    boton_cerrar = tk.Button(popup, text="Cerrar", width=25, bg='#FF4500', font=('Calibri', 12, 'bold'), fg='black', command=popup.destroy)
    boton_cerrar.pack(pady=10)
    
def mostrar_mensaje_inicial():
    mensaje = "          ¡Bienvenido!\n Las instrucciones están\n  en el menú 'AYUDA'"
    
    mensaje_popup = tk.Toplevel(ventana)
    mensaje_popup.title("Guia")
    
    frame = ttk.Frame(mensaje_popup, padding=(20, 20))
    frame.grid(row=0, column=0)
    
    estilo_texto = ttk.Style()
    estilo_texto.configure("EstiloMensaje.TLabel", font=('Calibri', 14, 'bold'), foreground='#2196F3')
    estilo_texto.configure("EstiloMensaje.TFrame")
    
    label = ttk.Label(frame, text=mensaje, style="EstiloMensaje.TLabel", wraplength=300, anchor="center")
    label.grid(row=0, column=0, pady=(0, 10))
    
    boton_cerrar = tk.Button(frame, text="Cerrar", width=25, bg='#FF4500', font=('Calibri', 12, 'bold'), fg='black', command=mensaje_popup.destroy)
    boton_cerrar.grid(row=1, column=0, pady=(0, 10))


ventana = tk.Tk()
ventana.geometry('400x300')
ventana.title("Formularios de Excel")
image = PhotoImage(file="images.png")

label = ttk.Label(ventana, image=image)
label.pack(pady=10)

ventana.after(100, mostrar_mensaje_inicial)

barra_navegacion = tk.Menu(ventana)
ventana.config(menu=barra_navegacion)

menu_ayuda = tk.Menu(barra_navegacion, tearoff=0)
barra_navegacion.add_cascade(label="AYUDA", menu=menu_ayuda)
menu_ayuda.add_command(label="Como usar 'Crear Excel'", command=lambda: mostrar_instrucciones("Crear"))
menu_ayuda.add_command(label="Como usar 'Abrir Excel'", command=lambda: mostrar_instrucciones("Abrir"))
menu_ayuda.add_command(label="ACERCA DE'", command=lambda: mostrar_instrucciones("Acerca"))

boton_abrir = tk.Button(ventana, text="Crear Excel", width=25, bg='lightblue', font=('Calibri', 12, 'bold'), fg='black', command=lambda: abrir_modulo("formu"))
boton_abrir.pack(pady=10)

boton_abrir = tk.Button(ventana, text="Abrir Excel", width=25, bg='lightgreen', font=('Calibri', 12, 'bold'), fg='black', command=lambda: abrir_modulo("intento2"))
boton_abrir.pack(pady=10)

ventana.mainloop()
