from tkinter import *

root = Tk()

button = Button(root, text="Click me!")
img = PhotoImage(file="/logo.ico") # make sure to add "/" not "\"
button.config(image=img)
button.pack() # Displaying the button

root.mainloop()