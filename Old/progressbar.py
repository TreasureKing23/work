import time
from tkinter import Tk, Toplevel, Label
from tkinter.ttk import Progressbar

root = Tk()
root.withdraw()

win = Toplevel()
win.title("Progress bar test")

Label(win, text="Processing...").pack(padx=10, pady=10)
pb = Progressbar(win, length=300,mode="determinate", maximum=10)
pb.pack(padx=10, pady=10)

for i in range(10):
    pb["value"] = i+1
    win.update()
    time.sleep(0.3)

win.destroy()

