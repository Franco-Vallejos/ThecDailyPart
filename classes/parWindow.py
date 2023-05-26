import tkinter as tk
from tkinter import END
from tkinter import messagebox
import os


class parWindow(tk.Toplevel):
    def ___init___(self, master = None):
        super().__init__(master)
        self.master = master
        self.grab_set()
        self.geometry("200x150")
        self.title("Parametros")

    def initParWindow(self, tecRx = None, tecSala = None, tecReava = None, tecTotales = None, master = None):

        self.textBox_Rx = tk.Entry(self, width = 4, bg = "light grey")
        self.textBox_Sala = tk.Entry(self, width = 4, bg = "light grey")
        self.textBox_ReavaTx = tk.Entry(self, width = 4, bg = "light grey")
        self.textBox_TotalTech = tk.Entry(self, width = 4, bg = "light grey")

        self.textoRx = tk.Label(self, text = "Tecnicos Rx: ", bg = '#2B9EFF', font = ("Helvetica", 8, "bold"))
        self.textoSala = tk.Label(self, text = "Tecnicos Sala: ", bg = '#2B9EFF', font = ("Helvetica", 8, "bold"))
        self.textoReava = tk.Label(self, text = "Tecnicos Reava/tx: ", bg = '#2B9EFF', font = ("Helvetica", 8, "bold"))
        self.textoTectT = tk.Label(self, text = "Tecnicos Totales: ", bg = '#2B9EFF', font = ("Helvetica", 8, "bold"))

        if tecRx:
            self.textBox_Rx.delete(0 , END)
            self.textBox_Rx.insert(0 , tecRx)

        if tecReava:
            self.textBox_ReavaTx.delete(0 , END)
            self.textBox_ReavaTx.insert(0 , tecReava)
        
        if tecSala:
            self.textBox_Sala.delete(0 , END)
            self.textBox_Sala.insert(0 , tecSala)

        if tecTotales:
            self.textBox_TotalTech.delete(0 , END)
            self.textBox_TotalTech.insert(0 , tecTotales)

        self.b_guardarParametros = tk.Button(self, text = 'Guardar', bg = "light grey", command = lambda :  self.guardarParametros(master))

    def placeParWindow(self):

        self.textoRx.grid(row = 1, column = 0, ipady = 5, sticky = tk.E)
        self.textoSala.grid(row = 2, column = 0, ipady = 5, sticky = tk.E)
        self.textoReava.grid(row = 3, column = 0, ipady = 5, sticky = tk.E)
        self.textoTectT.grid(row = 4, column = 0, ipady = 5, sticky = tk.E)

        self.textBox_Rx.grid(row = 1, column = 1, sticky = tk.W)
        self.textBox_Sala.grid(row = 2, column = 1, sticky = tk.W)
        self.textBox_ReavaTx.grid(row = 3, column = 1, sticky = tk.W)
        self.textBox_TotalTech.grid(row = 4, column = 1, sticky = tk.W)

        self.b_guardarParametros.grid(row = 5, column = 1, sticky = tk.E)

    def guardarParametros(self, master):
        tecRx = self.textBox_Rx.get()
        tecSala = self.textBox_Sala.get()
        tecReava = self.textBox_ReavaTx.get()
        tecTotales = self.textBox_TotalTech.get()
            
        try:
            os.remove("parameters.txt")
        except:
            messagebox.showerror('Error', 'Error al guardar Parametros (verifique no tener abierto el archivo "parameters"')
            return

        with open("parameters.txt", "wt") as archivo:
            archivo.write('Rx: ' + tecRx + '\n')
            archivo.write('Sala: ' + tecSala + '\n')
            archivo.write('Reava/Tx: ' + tecReava + '\n')
            archivo.write('Totales: ' + tecTotales + '\n')

        master.openParameters = None 
        self.destroy()
    
