import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import tkinter
import pandas as pd
import re
from functions_GCA.realizarFactura import realizarFactura
#from realizarFactura import realizarFactura
from selenium import webdriver
import time


def getAFIPData_fromExcel(rowNum, filepath, processColumn):
    if processColumn == "A":
        try:
            df = pd.read_excel(filepath)
            desde = (df[df.columns[0]][rowNum-1].replace('-', '/'))
            
            return desde

        except FileNotFoundError:
            messagebox.showwarning("Aviso", "No se puede encontrar el archivo.")
    
    if processColumn == "B":
        try:
            df = pd.read_excel(filepath)
            # matchPattern = re.compile(r'TRANSFER\.\s*(\d{11})')
            matchPattern = re.compile(r'(?<!\d)(\d{10}|\d{11}|\d{12})(?!\d)')

            return(matchPattern.findall((df[df.columns[1]][rowNum-1]))[0])

        except FileNotFoundError:
            messagebox.showwarning("Aviso", "No se puede encontrar el archivo.")
        except AttributeError:
            messagebox.showwarning("Aviso", "Hay un error, no se puede leer la celda seleecionada.\nEs posible que se tenga que corregir el número.\nAsegurese de que no cambio partes del archivo Excel descargado del Banco Ciudad (como borrar o mover el logo del Banco, etc)")
        except KeyError:
            messagebox.showwarning("Aviso", "Error en el numero de fila elegido")
    
    if processColumn == "C":
        try:
            df = pd.read_excel(filepath)
            importe = df[df.columns[2]][rowNum-1].replace("$ ", '').replace('.', '')
            sep = ','
            stripped = importe.split(sep, 1)[0]
            return stripped

        except FileNotFoundError:
            messagebox.showwarning("Aviso", "No se puede encontrar el archivo.")
        except AttributeError:
            messagebox.showwarning("Aviso", "Hay un error, no se puede leer la celda seleecionada.\nEs posible que se tenga que corregir el número.\nAsegurese de que no cambio partes del archivo Excel descargado del Banco Ciudad (como borrar o mover el logo del Banco, etc).")
        except KeyError:
            messagebox.showwarning("Aviso", "Error, no existe la celda.")



def startWindow():
    root = Tk()
    # root.geometry('780x400')
    root.wm_title("Generador de Comprobantes AFIP Automatico")
    l = Label(root, text= "Bienvenido!\nGenerador de Comprobantes AFIP Automático.", font='Helvetica 16 bold')
    l.grid(row=0, column=3)

    # Browse file
    factura_FilePath = StringVar()
    
    label_xlsxFile = Label(root, textvariable= factura_FilePath)
    factura_FilePath.set("Aqui se mostrara el camino del archivo elegido.")
    label_xlsxFile.grid(row=2, column=3)
    def browsefunc():
        filename = filedialog.askopenfilename(filetypes=(("xlsx files","*.xlsx"),("All files","*.*")))
        factura_FilePath.set(filename) # setting factura_FilePath

        if factura_FilePath.get() != "":
            root.update_idletasks()
            callback()
            entryDESDE.grid(row=6, column=3)
            labelPD.grid(row=5, column=3)
            entryHASTA.grid(row=8, column=3)
            labelFF.grid(row=7, column=3)
            entryEmpresa.grid(row=10, column=3)
            labelEmp.grid(row=9, column=3)
            entryServicio.grid(row=12, column=3)
            labelS.grid(row=11, column=3)
        else:
            factura_FilePath.set("Aqui se mostrara el camino del archivo elegido.")
            entryDESDE.grid_forget()
            labelPD.grid_forget()
            entryHASTA.grid_forget()
            labelFF.grid_forget()
            entryEmpresa.grid_forget()
            labelEmp.grid_forget()
            entryServicio.grid_forget()
            labelS.grid_forget()

    # Variables
    stringHASTA = tk.StringVar(root)
    stringEmpresa = tk.StringVar(root)
    stringServicio = tk.StringVar(root)
    numRowExcel = tk.IntVar(root)
    numRowExcel.set(6)
    periodoFacturadoDESDE = tk.StringVar(root)
    Cuil = StringVar(root)
    password = StringVar(root)

    # "Empezar" button
    list_Of_InputValues = []
    def getvalue(lst):
        try:
            lst.extend([Cuil.get(), password.get(), stringEmpresa.get(), stringHASTA.get(), periodoFacturadoDESDE.get(), getAFIPData_fromExcel(numRowExcel.get(), factura_FilePath.get(), 'B'), getAFIPData_fromExcel(numRowExcel.get(), factura_FilePath.get(), 'C'), stringServicio.get()])
            if (lst == []) or ('' in lst):
                lst.clear()
                messagebox.showwarning("Aviso", "Debe completar las casillas antes de Empezar.")
            else:
                driver = webdriver.Safari()
                lst.insert(0, driver)
                root.withdraw()
                root.update()
                realizarFactura(*lst)
                root.deiconify()
                root.update()
                numRowExcel.set(numRowExcel.get() + 1)
                lst.clear()
                root.update()

        except TclError:
            messagebox.showwarning("Aviso", "Es probable que no haya puesto un numero entero en Cantidad de Facturas a realizar.")


    # Handling grey default text
    def handle_focus_in(_):
        entryDESDE.delete(0, tk.END)
        entryDESDE.config(fg='black')
        labelAutoFillmsg.grid_forget()

    def handle_focus_out(_):
        if periodoFacturadoDESDE.get() == "":
            entryDESDE.delete(0, tk.END)
            entryDESDE.config(fg='black')
            entryDESDE.insert(0, periodoFacturadoDESDE.get())
            root.update_idletasks()


    def callback(*args):
        try:
            entryDESDE.delete(0, tk.END)
            entryDESDE.config(fg='black')
            periodoFacturadoDESDE.set(getAFIPData_fromExcel(numRowExcel.get(), factura_FilePath.get(), 'A'))
            labelAutoFillmsg.grid(row=6, column=4)
            # entryDESDE.insert(0, periodoFacturadoDESDE.get())
            root.update_idletasks()
        
        except AttributeError:
            entryDESDE.delete(0, tk.END)
            entryDESDE.config(fg='grey')
            entryDESDE.insert(0, "Corregir el número de fila")
            # messagebox.showwarning("Aviso", "Hay un error con el numero de fila puesto.\nEs posible que se tenga que corregir el número.\nAsegurese de que no cambio partes del archivo Excel descargado del Banco Ciudad (como borrar o mover el logo del Banco, etc)")
        except KeyError:
            entryDESDE.delete(0, tk.END)
            entryDESDE.config(fg='grey')
            entryDESDE.insert(0, "Error en numero de fila elegido")
            # messagebox.showwarning("Aviso", "Error en el numero de fila elegido")
        except tkinter.TclError as e:  # as e syntax added in ~python2.5
            if "expected floating-point number but got" in str(e):
                entryDESDE.delete(0, tk.END)
                entryDESDE.config(fg='grey')
                entryDESDE.insert(0, "Poner numero de la primer fila")

    getFileButton = tk.Button(root, text='Buscar Archivo', height = 1, width = 12, command= browsefunc).grid(row=3, column=3)
    entryExcelRow = Entry(root,textvariable = numRowExcel, width=4,fg="blue",bd=3,selectbackground='violet')
    entryExcelRow.grid(row=5, column=7)
    labelExcelRow = Label(root, text= "Primer fila del\nExcel con factura a\nrealizar:", font='Helvetica 14 underline').grid(row=5, column=6)
    CuilLabel = Label(root, text="CUIL/CUIT:").grid(row=5, column=0)
    CuilEntry = Entry(root, textvariable=Cuil, width=14).grid(row=5, column=1)
    passwordLabel = Label(root, text="Clave:").grid(row=6, column=0)
    passwordEntry = Entry(root, textvariable=password, show='*', width=14).grid(row=6, column=1)

    buttonStart = tk.Button(root, text='Empezar', font='Helvetica 16', height = 1, width = 10, command= lambda arg1 = list_Of_InputValues: getvalue(arg1)).grid(row=13, column=3)
    
    entryDESDE = Entry(root,textvariable = periodoFacturadoDESDE, width=22,fg="blue",bd=3,selectbackground='violet')
    labelPD = Label(root, text= "Periodo Facturado \"Desde\":", font='Helvetica 14 underline')

    entryHASTA = Entry(root,textvariable = stringHASTA,width=22,fg="blue",bd=3,selectbackground='violet')
    labelFF = Label(root, text= "Fecha de Facturación/\nFecha del Comprobante/\nPeriodo Facturado \"Hasta\":", font='Helvetica 14 underline')

    entryEmpresa = Entry(root,textvariable = stringEmpresa,width=24,fg="blue",bd=3,selectbackground='violet')
    labelEmp = Label(root, text= "Empresa:", font='Helvetica 14 underline')

    entryServicio = Entry(root, textvariable = stringServicio, width=22,fg="blue",bd=3,selectbackground='violet')
    labelS = Label(root, text= "Servicio:", font='Helvetica 14 underline')

    labelAutoFillmsg = Label(root, text= "Se tomara automaticamente\nlo que haya en la celda\nPeriodo Facturado \"Desde\"\nsi no lo cambia.", fg="grey", borderwidth=2, relief="ridge", font='Helvetica 10 bold')
    
    entryDESDE.bind("<FocusIn>", handle_focus_in)
    entryDESDE.bind("<FocusOut>", handle_focus_out)
    
    numRowExcel.trace_add('write', callback)

    root.mainloop()

    return list_Of_InputValues


if __name__ == '__main__':
    print(startWindow())
    # print(againWindow("test","test","test","test","test",8))
