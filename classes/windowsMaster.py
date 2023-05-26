import tkinter as tk
from tkinter.filedialog import askopenfile
from tkinter import messagebox
from tkinter import filedialog
from tkinter import END
from classes.excelWoorkBook import *
from tkinter import Listbox


def closeWindow(window, master):
    print("Cerrandose")
    window.destroy()
    window.grab_release()
    

class windowsMaster(tk.Frame):
    def __init__(self, master = None, eanaIcon = None):
        super().__init__(master)
        self.master = master
        self.grid(row = 0)
        self.workBook = excelWorkBook()
        self.eanaIcon = eanaIcon

    def initInterface(self, bigEanaImage = None, MalvinasImage = None, airplaneImagenRuteL = None, airplaneImagenRuteR = None):
        self['bg'] = '#2B9EFF'

        if bigEanaImage:
            self.eanaImage = tk.Label(self, image = bigEanaImage, bg = '#2B9EFF')
        if MalvinasImage:
            self.MalvinasImageVent = tk.Label(self, image = MalvinasImage, bg = '#2B9EFF')
        if airplaneImagenRuteL:
            self.airplaneImagenRuteVentL = tk.Label(self, image = airplaneImagenRuteL, bg = '#2B9EFF')
        if airplaneImagenRuteR:
            self.airplaneImagenRuteVentR = tk.Label(self, image = airplaneImagenRuteR, bg = '#2B9EFF')

        self.text_load = tk.Label(self, text = "Carge el archivo con la lista de turno", bg = '#2B9EFF', font = ("Helvetica", 10, "bold"))
        self.textBox_rute = tk.Entry(self, width = 67, bg = "light grey")
        self.buttom_archive = tk.Button(self, text = "...", command = self.getDic)
        self.buttom_load = tk.Button(self, text = "cargar", command = self.load)
        self.text_Sheets = tk.Label(self, text= "Seleccione la Hoja del Excel correspondiente", bg = '#2B9EFF', font = ("Helvetica", 10, "bold"))
        self.listSheet = Listbox(self, selectmode = "browse", width=20, bg = "light grey")
        self.text_period = tk.Label(self, text = "Ingrese el periodo:   Desde            Hasta", bg = '#2B9EFF', font = ("Helvetica", 10, "bold"))
        self.textBox_from = tk.Entry(self, width = 4, bg = "light grey")
        self.textBox_to = tk.Entry(self, width = 4, bg = "light grey")
        self.text_default = tk.Label (self, text = "(por Defecto inicio - fin)", bg = '#2B9EFF', font = ("Helvetica", 8, "bold"))
        self.buttom_save = tk.Button(self, text= "Guardar", command = self.save)
        self.text_rationing = tk.Label(self, text = " ", bg = '#2B9EFF', font = ("Helvetica", 8, "bold"))

    def placeInterface(self):

        self.eanaImage.grid(columnspan=4, column=0, row=0)
        self.text_load.grid(row = 1, column = 0, sticky= tk.W)
        self.textBox_rute.place(x = 0, y=70)
        self.buttom_archive.grid(row = 2, column = 2)
        self.buttom_load.place(x = 442, y = 66)
        self.text_Sheets.grid(row = 3, column = 0, sticky = tk.W, ipady = 5)
        self.listSheet.grid(row = 4, column=0)
        self.MalvinasImageVent.place(x = 300, y = 130)
        self.airplaneImagenRuteVentL.place(x = 32, y = 155)
        self.airplaneImagenRuteVentR.place(x = 100, y = 155)
        self.text_period.grid(row = 5, column = 0, sticky = tk.W, pady = 7)
        self.textBox_from.place(x = 180, y =296)
        self.textBox_to.place(x = 262, y =296)
        self.text_default.place(x = 300, y = 296)
        self.buttom_save.place(x=440, y = 316)
        self.text_rationing.grid(row = 6, column = 0, sticky= tk.E)

    def getDic(self):
        path = askopenfile(title = "seleccione el archivo..", mode = 'r', filetypes = [('Microsoft Excel Worksheet', '*.xlsx')])

        if not path:
            return None

        self.textBox_rute.delete("0", "end")
        self.textBox_rute.insert(0, path.name)

        if not self.workBook.initWorkBook(path.name):
            messagebox.showerror("Error de Apertura", "Verifique la existencia del archivo, y si esta cerrado")
            return

        self.showSheets() 

    def load(self):
        if not self.textBox_rute.get():
            messagebox.showerror("Error de Apertura", "No ingreso un archivo")
            return
        
        if not self.workBook.initWorkBook(self.textBox_rute.get()):
            messagebox.showerror("Error de Apertura", "Verifique la existencia del archivo, y si esta cerrado")
            return

        self.showSheets()

    def showSheets(self):
        sheets = self.workBook.getSheets()
        self.listSheet.delete("0", "end")
        for each_item in range(len(sheets)):
            self.listSheet.insert(END, sheets[each_item])
            self.listSheet.itemconfig(each_item, bg = "#B9FEFF" if each_item % 2 == 0 else "#72EAFF")

    def save(self):
        try:
            sheet = self.listSheet.get(self.listSheet.curselection())
        except:
            messagebox.showerror("Error de Seleccion","No selecciono ninguna hoja valida")
            return

        self.workBook.initSheet(sheet)
        pathFileSave = filedialog.askdirectory()
        
        self.workBook.process(pathFileSave, self.textBox_from.get(), self.textBox_to.get())
