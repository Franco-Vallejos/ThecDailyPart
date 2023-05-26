import tkinter as tk
from PIL import ImageTk
import PIL as pil
from classes.windowsMaster import *
from classes.excelWoorkBook import *

app = tk.Tk()

app.title('RACIONAMIENTO - by Valle')
app.geometry('505x345')
app.resizable (0,0)
eanaIcon = pil.Image.open("resourse/image/Eana.ico")
eanaIconRebuilded = ImageTk.PhotoImage(eanaIcon)
app.wm_iconphoto(False, eanaIconRebuilded)

window = windowsMaster(master = app, eanaIcon  = eanaIconRebuilded)
bigEanaImage = tk.PhotoImage(file = "resourse/image/EanaLogo.jpg")
MalvinasImage= tk.PhotoImage(file = "resourse/image/airplane.png")
airplaneImagenRuteL = tk.PhotoImage(file = "resourse/image/airplaneRuteL.png")
airplaneImagenRuteR = tk.PhotoImage(file = "resourse/image/airplaneRuteR.png")

window.initInterface(bigEanaImage, MalvinasImage, airplaneImagenRuteL, airplaneImagenRuteR)
window.placeInterface()

window.mainloop()