#Primero guardar todos los datos necesarios y luego crear el Excel
#Imortacion de las Librerias Pandas y openpyxl

from tkinter import Tk, Label, Button,Entry, Frame, END
import pandas as pd
import openpyxl

ventana = Tk()
ventana.config(bg='white')
ventana.geometry('700x650')
ventana.resizable(0,0)
ventana.title('Guardar datos en Excel')
nombre1,apellido1,edad1,correo1,telefono1 = [],[],[],[],[]

def agregar_datos():
	global nombre1, apellido1, dni1, correo1, telefono1

	nombre1.append(ingresa_nombre.get())
	apellido1.append(ingresa_apellido.get())
	edad1.append(ingresa_edad.get())
	correo1.append(ingresa_correo.get())
	telefono1.append(ingresa_telefono.get())

	ingresa_nombre.delete(0,END)
	ingresa_apellido.delete(0,END)
	ingresa_edad.delete(0,END)
	ingresa_correo.delete(0,END)
	ingresa_telefono.delete(0,END)

def guardar_datos():	
	global nombre1,apellido1,edad1,correo1,telefono1
	datos = {'Nombres':nombre1,'Apellidos':apellido1, 'Edad':edad1, 'Correo':correo1, 'Telefono':telefono1 } 
	nom_excel  = str(nombre_archivo.get() + ".xlsx")	
	df = pd.DataFrame(datos,columns =  ['Nombres', 'Apellidos', 'Edad', 'Correo', 'Telefono']) 
	df.to_excel(nom_excel)
	nombre_archivo.delete(0,END)

frame1 = Frame(ventana, bg='white')
frame1.grid(column=0, row=0, sticky='nsew')
frame2 = Frame(ventana, bg='white')
frame2.grid(column=1, row=0, sticky='nsew')

archivo2 = Label(frame2, text ='Agregar Datos:', width=25, bg='white',font = ('Calibri',12, 'bold'), fg='black')
archivo2.grid(column=1, row=0, pady=10, padx= 10)

nombre = Label(frame2, text ='Nombre', width=15).grid(column=0, row=1, pady=20, padx= 10)
ingresa_nombre = Entry(frame2,  width=30, font = ('Calibri',12))
ingresa_nombre.grid(column=1, row=1)

apellido = Label(frame2, text ='Apellido', width=15).grid(column=0, row=2, pady=20, padx= 10)
ingresa_apellido = Entry(frame2, width=30, font = ('Calibri',12))
ingresa_apellido.grid(column=1, row=2)

edad = Label(frame2, text ='Edad', width=15).grid(column=0, row=3, pady=20, padx= 10)
ingresa_edad = Entry(frame2,  width=30, font = ('Calibri',12))
ingresa_edad.grid(column=1, row=3)

correo = Label(frame2, text ='Gmail', width=15).grid(column=0, row=4, pady=20, padx= 10)
ingresa_correo = Entry(frame2,  width=30, font = ('Calibri',12))
ingresa_correo.grid(column=1, row=4)

telefono = Label(frame2, text ='Telefono', width=15).grid(column=0, row=5, pady=20, padx= 10)
ingresa_telefono = Entry(frame2, width=30, font = ('Calibri',12))
ingresa_telefono.grid(column=1, row=5)

agregar = Button(frame2, width=20, font = ('Calibri',12, 'bold'), text='Guardar', bg='#87CEFA',bd=2, command =guardar_datos)
agregar.grid(column=1, row=6, pady=20, padx= 10)
            #columnspan=2

archivo = Label(frame1, text ='Ingrese Nombre del archivo', width=25, bg='white',font = ('Arial',12, 'bold'), fg='black')
archivo.grid(column=0, row=0, pady=10, padx= 10)

nombre_archivo = Entry(frame1, width=23, font = ('Arial',12),highlightbackground = "#90EE90", highlightthickness=2)
nombre_archivo.grid(column=0, row=1, pady=0, padx= 10)

guardar = Button(frame1, width=20, font = ('Arial',12, 'bold'), text='Crear Excel', bg='#90EE90',bd=2, command =guardar_datos)
guardar.grid(column=0, row=2, pady=20, padx= 10)

ventana.mainloop()
