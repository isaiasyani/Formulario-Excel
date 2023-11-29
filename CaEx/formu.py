import tkinter as tk
from tkinter import Entry, Label, Button
import pandas as pd
import configparser

ventana = tk.Tk()
ventana.title('Guardar datos en Excel')
ventana.geometry('700x600')
ventana.configure(bg='white')

frame2 = tk.Frame(ventana, bg='white')
frame2.pack()

config = configparser.ConfigParser()
config.read('config.ini')

etiquetas = config['Etiquetas']
campos = {}
datos_guardados = []
contador = 0

def guardar_datos():
    global contador
    contador += 1
    datos = {}
    for key, entry in campos.items():
        datos[etiquetas[key]] = entry.get()

    datos_guardados.append(datos)

    for entry in campos.values():
        entry.delete(0, tk.END)

    label_contador.config(text=f"Datos Guardados = {contador}")

def guardar_en_excel():
    global contador
    df = pd.DataFrame(datos_guardados)
    nom_excel = str(nombre_archivo.get() + ".xlsx")
    df.to_excel(nom_excel)
    nombre_archivo.delete(0, tk.END)
    contador = 0
    label_contador.config(text=f"Datos Guardados = {contador}")

archivo = Label(frame2, text='Ingrese Nombre del archivo', width=25, bg='white', font=('Arial', 12, 'bold'), fg='black')
archivo.grid(column=0, row=0, pady=10, padx=10)

nombre_archivo = Entry(frame2, width=30, font=('Calibri', 12), highlightbackground="#90EE90", highlightthickness=2)
nombre_archivo.grid(column=0, row=1, pady=0, padx=10)

boton_crear = Button(frame2,width=20, text='Crear Excel', font=('Arial', 12, 'bold'), bg='#90EE90', bd=2, command=guardar_en_excel)
boton_crear.grid(column=0, row=2, pady=20, padx=10, columnspan=2)

row_counter = 1
for key, value in etiquetas.items():
    label = Label(frame2, bg='white', font=('Calibri', 12, 'bold'),text=value, width=15)
    label.grid(column=2, row=row_counter, pady=20, padx=5)

    archivo2 = Label(frame2, text='Agregar Datos:', width=25, bg='white', font=('Calibri', 12, 'bold'), fg='black')
    archivo2.grid(column=3, row=0, pady=10, padx=10)

    entry = Entry(frame2, width=30, font=('Calibri', 12), highlightbackground="#87CEFA", highlightthickness=2)
    entry.grid(column=3, row=row_counter, pady=0, padx=10)

    campos[key] = entry
    row_counter += 1

boton_guardar = Button(frame2, width=20, text='Guardar', font=('Calibri', 12, 'bold'), bg='#87CEFA', bd=2, command=guardar_datos)
boton_guardar.grid(column=3, row=row_counter, pady=20, padx=20)

label_contador = Label(frame2, bg='white', font=('Calibri', 12, 'bold'), text=f"Datos Guardados = {contador}")
label_contador.grid(column=3, row=row_counter + 1)

ventana.mainloop()




