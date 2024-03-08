import pandas as pd
import json
from tkinter import filedialog, messagebox
import tkinter as tk
import os

class ConversionToolGUI:
    def __init__(self, root):
        self.json_data = None
        self.excel_path = None
        self.root = root
        self.displayFilePath = None
        self.setup_ui()
        

    def setup_ui(self):
        self.root.title('Conversion Tool')
        self.root.geometry('700x300')
        
        promptButton = tk.Button(self.root, text="Datei auswählen", font=("Arial", 16), command=self.read_path)
        buttonFrame = tk.Frame(self.root)
        
        self.displayFilePath = tk.Label(self.root, text="Keine Datei ausgewählt")
        
        buttonFrame.columnconfigure(2, weight=5)
        buttonFrame.columnconfigure(3, weight=5)
        buttonFrame.columnconfigure(4, weight=5)
        buttonFrame.columnconfigure(5, weight=5)

        btn1 = tk.Button(buttonFrame, text="json -> excel", font=('Arial', 18), height=2, width=15, command=self.convert_to_excel)
        btn1.grid(row=4, column=2, sticky=tk.W+tk.E)

        btn2 = tk.Button(buttonFrame, text="excel -> json", font=('Arial', 18), height=2, width=15, command=self.convert_to_json)
        btn2.grid(row=4, column=5, sticky=tk.W+tk.E)
        
        promptButton.pack(padx=50, pady=50)
        buttonFrame.pack()
        self.displayFilePath.pack()

    def read_path(self):
        file_path = filedialog.askopenfilename()
        if file_path.endswith('.json'):
            with open(file_path, 'r', encoding='utf-8') as f:
                self.json_data = json.load(f)
            print(self.json_data)
        elif file_path.endswith('.xlsx'):
            self.excel_path = file_path
        else:
            messagebox.showerror(title="Fehler", message='Datei nicht als json oder xlsx erkannt')
        self.displayFilePath.config(text=f"{file_path}")

    def convert_to_excel(self):
        if self.json_data:
            fields_df = pd.DataFrame([
                {'Typ': fdf['type'], 'Name': fdf['name'], 'Datentyp': fdf['dataType']}
                for fdf in self.json_data['fields']
            ])

            search_lists_df = pd.DataFrame([
                {'Name': sl['name'], 'Wert': value} 
                for sl in self.json_data['searchLists'] 
                for value in sl['values']
            ])

            rules_records = []
            for rule in self.json_data['rules']:
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

            excel_path = filedialog.asksaveasfilename(initialfile = "mapped_excel_data.xlsx", defaultextension=".xlsx", filetypes=(("Excel-Tabelle", "*.xlsx"), ("Alle Dateien", "*.*")))
            with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                fields_df.to_excel(writer, sheet_name='Felder', index=False)
                rules_df.to_excel(writer, sheet_name='Regeln', index=False)
                search_lists_df.to_excel(writer, sheet_name='Suchlisten', index=False)
            messagebox.showinfo(title="Erfolg!", message=f"Datei erfolgreich zu {excel_path} geschrieben!")
        
        elif self.json_data == None:
            messagebox.showerror(title="Fehler", message="Keine Valide Json Datei ausgewählt")
            
    def convert_to_json(self):
        if self.excel_path:
            fields_df = pd.read_excel(self.excel_path, sheet_name='Felder')
            rules_df = pd.read_excel(self.excel_path, sheet_name='Regeln')
            search_lists_df = pd.read_excel(self.excel_path, sheet_name='Suchlisten')
            print(search_lists_df)
            json_data = {
                    'fields': [],
                    'searchLists': [],
                    'rules': []
                }
            if fields_df.empty == False:
                for name, group in fields_df.groupby('Name'):
                    json_data['fields'].append({
                        'type': group['Typ'].iloc[0],
                        'name': name,
                        'dataType': group['Datentyp'].iloc[0]
                    })
            if rules_df.empty == False:
                rules_order = {}
                for index, row in rules_df.iterrows():
                    rule_name = row['Name']
                    if rule_name not in rules_order:
                        rules_order[rule_name] = {
                            'index': len(json_data['rules']),
                            'criteria': []
                        }
                        json_data['rules'].append({
                            'isActive': row['Aktiv'],
                            'name': rule_name,
                            'result': row['Ergebnis'],
                            'criteria': rules_order[rule_name]['criteria']
                        })
                    criterion = {'type': row['Kriterientyp'], 'field': row['Kriterienfeld']}
                    if pd.notnull(row.get('Operator')):
                        criterion['operator'] = row['Operator']
                        criterion['value'] = str(int(row['Wert']))
                    if pd.notnull(row.get('Suchliste')):
                        criterion['searchList'] = row['Suchliste']
                    if pd.notnull(row.get('Von')):
                        criterion['lowerLimit'] = str(row['Von'])
                        criterion['upperLimit'] = str(row['Bis'])
                    rules_order[rule_name]['criteria'].append(criterion)
                    
            if search_lists_df.empty == False:
                search_lists_order = {}
                for index, row in search_lists_df.iterrows():
                    name = row['Name']
                    if name not in search_lists_order:
                        search_lists_order[name] = len(search_lists_order)
                        json_data['searchLists'].append({'name': name, 'values': []})
                    json_data['searchLists'][search_lists_order[name]]['values'].append(row['Wert'])
    
            json_path = filedialog.asksaveasfilename(initialfile = "mapped_json_data.json", defaultextension=".json", filetypes=(("Json", "*.json"), ("Alle Dateien", "*.*")))
            with open(json_path, 'w', encoding='utf8') as json_file:
                json.dump(json_data, json_file, indent=4, default=bool, ensure_ascii=False)
            messagebox.showinfo(title="Erfolg!", message=f"Datei erfolgreich zu {json_path} geschrieben!")                
            
        else:
            messagebox.showerror(title="Fehler", message="Keine Valide Excel Datei ausgewählt")
            
    
root = tk.Tk()
ConversionToolGUI(root)
root.mainloop()