# -*- coding: cp1252 -*- 
import json
import tkinter as tk
from tkinter import Button, filedialog

root = tk.Tk()
root.geometry("500x500")
root.title("Datei Import/Export")
label = tk.Label(root, text="Importieren Sie Ihre Datei", font=('Arial', 16))
label.pack(padx=20, pady=20)
def select_path():
    file_path = filedialog.askopenfilename()
    with open(file_path, 'r', encoding='utf-8') as f:
        json_data = json.load(f)
    return json_data

promptButton = tk.Button(root, text="Datei auswählen", font=("Arial", 16), command=select_path)
promptButton.pack(padx=50, pady=50)

buttonFrame = tk.Frame(root)
buttonFrame.columnconfigure(2, weight=5)
buttonFrame.columnconfigure(3, weight=5)
buttonFrame.columnconfigure(4, weight=5)
buttonFrame.columnconfigure(5, weight=5)

btn1 = tk.Button(buttonFrame, text="json -> excel", font=('Arial', 18), height=2, width=15)
btn1.grid(row=4, column=2, sticky=tk.W+tk.E)

btn2 = tk.Button(buttonFrame, text="excel -> json", font=('Arial', 18), height=2, width=15)
btn2.grid(row=4, column=5, sticky=tk.W+tk.E)

buttonFrame.pack()


root.mainloop()