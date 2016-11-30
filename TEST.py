from tkinter import Tk, Checkbutton, Label
from tkinter import StringVar, IntVar

root = Tk()

text = StringVar()
text.set('old')
status = IntVar()

def change():
    if status.get() == 1:   # if clicked
        text.set('new')
    else:
        text.set('old')


cb = Checkbutton(root, variable=status, command=change)
lb = Label(root, textvariable=text)
cb.pack()
lb.pack()
root.mainloop()