# -*- coding: cp1252 -*- 
import select
import pandas as pd
import json
import xlsxwriter
from tkinter import filedialog, Button, Tk, StringVar
import tkinter as tk

root = tk.Tk()
json_data_var = tk.StringVar()
root.geometry("500x500")
root.title("Datei Import/Export")
label = tk.Label(root, text="Importieren Sie Ihre Datei", font=('Arial', 16))
label.pack(padx=20, pady=20)

def select_path():
    file_path = filedialog.askopenfilename()
    with open(file_path, 'r', encoding='utf-8') as f:
        json_data = json.load(f)
        json_data_var.set(json_data)
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

json_data = json_data_var.get()
print(json_data, type(json_data))


fields_df = pd.DataFrame([
    {'Typ': fdf['type'], 'Name': fdf['name'], 'Datentyp': fdf['dataType']}
    for fdf in json_data['fields']
    ])

search_lists_df = pd.DataFrame([
    {'Name': sl['name'], 'Wert': value} 
    for sl in json_data['searchLists'] 
    for value in sl['values']
])

rules_records = []
for rule in json_data['rules']:
    for criterion in rule['criteria']:
        record = {
            'Aktiv': rule['isActive'],
            'Name': rule['name'],
            'Ergebnis': rule['result'],
            'Kriterientyp': criterion['type'],
            'Kriterienfeld': criterion['field']
        }
        if criterion['type'] == 'comparison':
            record['Operator'] = criterion['operator']
            record['Wert'] = criterion['value']
        elif criterion['type'] == 'textsearch':
            record['Suchliste'] = criterion['searchList']
        elif criterion['field'] == 'ERP_BRUTTO_BETRAG':
            record['Von'] = criterion['lowerLimit']
            record['Bis'] = criterion['upperLimit']
        
        
        rules_records.append(record)

rules_df = pd.DataFrame(rules_records)

excel_path = r"C:\Users\b.granderath\OneDrive - Wünsche Group\Desktop\Formatting_json_to_excel\Export_Beispiele\mapped_daten.xlsx"
with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
    fields_df.to_excel(writer, sheet_name='Felder', index=False)
    rules_df.to_excel(writer, sheet_name='Regeln', index=False)
    search_lists_df.to_excel(writer, sheet_name='Suchlisten', index=False)

excel_path
