import pandas as pd
import json
from tkinter import filedialog, messagebox
import tkinter as tk

class json2excel:
    def __init__(self, root):
        self.json_data = None
        self.root = root
        self.setup_ui()
        

    def setup_ui(self):
        promptButton = tk.Button(self.root, text="Datei auswählen", font=("Arial", 16), command=self.read_path)
        promptButton.pack(padx=50, pady=50)
        buttonFrame = tk.Frame(self.root)
        
        buttonFrame.columnconfigure(2, weight=5)
        buttonFrame.columnconfigure(3, weight=5)
        buttonFrame.columnconfigure(4, weight=5)
        buttonFrame.columnconfigure(5, weight=5)

        btn1 = tk.Button(buttonFrame, text="json -> excel", font=('Arial', 18), height=2, width=15, command=self.convert_to_excel)
        btn1.grid(row=4, column=2, sticky=tk.W+tk.E)

        btn2 = tk.Button(buttonFrame, text="excel -> json", font=('Arial', 18), height=2, width=15)
        btn2.grid(row=4, column=5, sticky=tk.W+tk.E)

        buttonFrame.pack()

    
    def read_path(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            with open(file_path, 'r', encoding='utf-8') as f:
                self.json_data = json.load(f)
                display_file_path = tk.Label(self.root, text=f"{file_path}")
                display_file_path.pack()

            
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

            excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
            with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
                fields_df.to_excel(writer, sheet_name='Felder', index=False)
                rules_df.to_excel(writer, sheet_name='Regeln', index=False)
                search_lists_df.to_excel(writer, sheet_name='Suchlisten', index=False)
            messagebox.showinfo(title="Erfolg!", message=f"Datei erfolgreich zu {excel_path} geschrieben!")

        else:
            messagebox.showerror(title="Kein Erfolg", message="Wir können leider keine Luft konvertieren")

            
root = tk.Tk()
app = json2excel(root)
root.mainloop()