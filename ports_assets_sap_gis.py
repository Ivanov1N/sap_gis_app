import tkinter as tk
from tkinter import filedialog
import pandas as pd
import ttkbootstrap as ttkb

# Creating application window
class AppWindow(ttkb.Window):
    def __init__(self):
        super().__init__(self, themename='superhero')
        self.title('Find Missing BPI Co. Assets (SAP and GIS/DMA)')
        self.geometry('800x500')
        self.resizable(False, False)
        #self.iconbitmap('app_icon.ico')
        load_frame = LoadData(self)
        ttkb.Separator(self, orient='horizontal').pack(fill='x')
        results_frame = Results(self)
        Compare(self, load_frame, results_frame)
        
        self.mainloop()
# Creating frame for data loading
class LoadData(ttkb.Frame):
    def __init__(self, parent):
        super().__init__(master=parent)
        # SAP table widgets
        self.sap_label = ttkb.Label(self, text='Load SAP table:')
        self.sap_label.grid(row=0, column=0, sticky='w', padx=20)
        self.sap_browse_button = ttkb.Button(
            self,
            text='Browse...',
            command=self.load_sap_table
        )
        self.sap_browse_button.grid(row=0, column=1, sticky='w')
        self.sap_entry = ttkb.Label(self, width=40, background='white')
        self.sap_entry.grid(row=0, column=2, sticky='w', padx=20)
        self.sap_check = ttkb.Label(self, font='Arial 12 bold',width=10)
        self.sap_check.grid(row=0, column=3, sticky='w')

        # GIS/DMA table widgets
        self.gis_label = ttkb.Label(self, text='Load GIS table:')
        self.gis_label.grid(row=1, column=0, sticky='w', padx=20)
        self.gis_browse_button = ttkb.Button(
            self,
            text='Browse...',
            command=self.load_gis_table
        )
        self.gis_browse_button.grid(row=1, column=1, sticky='w', pady=20)
        self.gis_entry = ttkb.Label(self, width=40, background='white')
        self.gis_entry.grid(row=1, column=2, sticky='w', padx=20)
        self.gis_check = ttkb.Label(self, font='Arial 12 bold')
        self.gis_check.grid(row=1, column=3, sticky='w')

        # Generating full SAP code widgets
        self.validate_label = ttkb.Label(self, text='Generate full SAP code:')
        self.validate_label.grid(row=2, column=0, sticky='w', padx=20)
        self.validate_button = ttkb.Button(self, text='Generate', command=self.add_sap_full)
        self.validate_button.grid(row=2, column=1, pady=10, sticky='w')
        self.check_symbol = ttkb.Label(self, text='!', font='Arial 12 bold', foreground='red')
        self.check_symbol.grid(row=2, column=2, sticky='w', padx=10)
        
        self.pack(pady=20)

    # Preparing columns with NaN data
    def fields_fill_na(self, table, main_field, secondary_field, n):
        table[main_field] = table[main_field].fillna(n).astype(int)
        table[secondary_field] = table[secondary_field].fillna(n).astype(int)

    # Loading and validating SAP table
    def load_sap_table(self):
        t = filedialog.askopenfilename(filetypes=[('Excel files', '.xlsx')])
        if len(t) == 0:
            pass
        else:
            global sap_xls
            sap_xls = pd.read_excel(t)
            self.sap_entry.configure(text=t[t.rfind('/')+1:], foreground='blue')
            self.check_symbol.configure(text='!', foreground='red', font='Arial 10 bold')
            if 'Основно средство' and 'Подномер' in sap_xls.columns:
                self.sap_check.configure(text='OK', foreground='green')
            else: self.sap_check.configure(text='X', foreground='red')
    
    #Loading and validating  GIS table
    def load_gis_table(self):
        t = filedialog.askopenfilename(filetypes=[('Excel files', '.xlsx')])
        if len(t) == 0:
            pass
        else:
            global gis_xls
            gis_xls = pd.read_excel(t)
            self.gis_entry.configure(text=t[t.rfind('/')+1:], foreground='blue')
            self.check_symbol.configure(text='!', foreground='red', font='Arial 12 bold')
            if 'САП номер' and 'САП подномер' in gis_xls.columns:
                self.gis_check.configure(text = 'OK', foreground = 'green')
            else: self.gis_check.configure(text = 'X', foreground = 'red')

    # Generating full SAP code
    def add_sap_full(self):
        try:
            global gis_xls, sap_xls
            self.fields_fill_na(sap_xls, "Основно средство", "Подномер", 999999999)
            sap_xls["sap_full"] = sap_xls["Основно средство"].astype(str) + sap_xls["Подномер"].astype(str)
            self.fields_fill_na(gis_xls, "САП номер", "САП подномер", 999999998)
            gis_xls["sap_full"] = gis_xls["САП номер"].astype(str) + gis_xls["САП подномер"].astype(str)
            if 'sap_full' in sap_xls.columns and 'sap_full' in gis_xls.columns:
                self.check_symbol.configure(text='OK', foreground='green', font='Arial 12 bold')
            else:
                pass
        except (NameError, KeyError) as error:
            self.check_symbol.configure(text='Please check the tables again!', font='Arial 12 bold')

# Creating action frame
class Compare(ttkb.Frame):
    def __init__(self, parent, load_data, results):
        super().__init__(master=parent)
        self.results = results
        self.load_data = load_data

        # Generating widgets for table comparison
        self.compare_label = ttkb.Label(self, text='Type of comparison:')
        self.compare_label.grid(row=4, column=0, sticky='w')
        self.radvar = tk.IntVar()
        self.compare_type1 = ttkb.Radiobutton(self, text='SAP assets missing in GIS/DMA', variable=self.radvar, value=1)
        self.compare_type1.grid(row=5, column=0, sticky='w', padx=20, pady=10)
        self.compare_type2 = ttkb.Radiobutton(self, text='GIS/DMA assets missing in SAP', variable=self.radvar, value=2)
        self.compare_type2.grid(row=6, column=0, sticky='w', padx=20)
        self.start_button = ttkb.Button(self, width=15, text='START/SAVE', command=self.compare_func)
        self.start_button.grid(row=7, column=1, padx=30)
        self.restart_button = ttkb.Button(self, width=15, text='RESET', command=self.restart_func)
        self.restart_button.grid(row=7, column=2, pady=30)

        self.pack(side='left', expand=True, fill='both')

    # Looking for missing assets
    def compare_func(self):
        try:
            global sap_xls, gis_xls
            if self.radvar.get() == 2:
                g = gis_xls[~gis_xls.sap_full.isin(sap_xls.sap_full)]
                if g.shape[0] == 0:
                    self.results.result_label.configure(text='* No missing assets')
                    pass
                else:
                    new_filename = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files')])
                    g.to_excel(new_filename)
                    self.results.result_label.configure(text=('* Number of missing assets - ' + str(g.shape[0])))
            elif self.radvar.get() == 1:
                g = sap_xls[~sap_xls.sap_full.isin(gis_xls.sap_full)]
                if g.shape[0] == 0:
                    self.results.result_label.configure(text='* No missing assets')
                    pass
                else:
                    new_filename = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files')])
                    g.to_excel(new_filename)
                    self.results.result_label.configure(text=('* Number of missing assets - ' + str(g.shape[0])))
            else:
                pass
        except (NameError, AttributeError) as error:
            self.results.result_label.configure(text='Missing tables or full SAP code!')
        
        except ValueError:
            self.results.result_label.configure(text='Press again "START/SAVE" to save the table!')

    # Reseting the app
    def restart_func(self):
        self.load_data.sap_entry.configure(text='')
        self.load_data.gis_entry.configure(text='')
        self.load_data.check_symbol.configure(text='!', foreground='red')
        self.load_data.sap_check.configure(text='')
        self.load_data.gis_check.configure(text='')
        self.results.result_label.configure(text='...')
        self.radvar.set(0)
        if 'sap_xls' in globals():
            del globals()['sap_xls']
        if 'gis_xls' in globals():
            del globals()['gis_xls']

# Creating frame for showing results
class Results(ttkb.LabelFrame):
    def __init__(self, parent):
        super().__init__(master=parent)
        self.configure(text='Results')
        self.result_label = ttkb.Label(self, text='...')
        self.result_label.grid(row=9, column=0, sticky='w', ipadx=20)

        self.pack(side='bottom',expand=True, fill='both')

if __name__ == "__main__":
    AppWindow()
