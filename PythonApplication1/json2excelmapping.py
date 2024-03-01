# -*- coding: cp1252 -*- 
import pandas as pd
import json
from tkinter import *
from tkinter import filedialog
import codecs


def openFile():
    with codecs.open(filedialog.askopenfilename(), 'r', 'utf-8') as f:
        json_data = json.load(f)
    print(json_data)
    return json_data

json_data = openFile()
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
