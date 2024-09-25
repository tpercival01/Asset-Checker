import pandas as pd
import tkinter as tk
from pathlib import Path

p = Path(__file__).with_name('assets_dummy_data.xlsb')

jamf_df = pd.read_excel(p, sheet_name='jamf')
intune_df = pd.read_excel(p, sheet_name='intune')
jamf_other_df = pd.read_excel(p, sheet_name='jamf_other')
landlines_df = pd.read_excel(p, sheet_name='landlines')
assets_sn_df = pd.read_excel(p, sheet_name='assets_SN')
jamf_names_list = jamf_df["Full Name"].astype(str).to_list()
intune_names_list = intune_df["Primary user display name"].astype(str).tolist()
user_names_list = jamf_names_list + intune_names_list
user_names_list.sort()

import tkinter as tk

class App(tk.Tk):
    def __init__(self):
        self.entries_dict = {}
        self.grid_entries = []
        self.count = 0
        tk.Tk.__init__(self)

        self.title('CAA SD - Asset Check')
        self.results = []

        title_label = tk.Label(self, text="Service Desk Asset Checker", font=('helvetica 18 bold'))
        title_label.grid(row=0, columnspan=5)

        instruction_label = tk.Label(self, text="Enter username in the field below: ")
        instruction_label.grid(row=1 ,columnspan=5)

        search_button = tk.Button(self, text="Find assets", command=self.create_grid)
        search_button.grid(row=4, columnspan=5, pady=20)

        # Put the filter in a frame at the top spanning across the columns.
        self.frame = tk.Frame(self)
        self.frame.grid(row=2, columnspan=5)

        # Put the filter label and entry box in the frame.
        tk.Label(self.frame, text='Name:').pack(side='left')

        self.filter_box = tk.Entry(self.frame)
        self.filter_box.pack(side='left', fill='x', expand=True)

        # A listbox with scrollbars.

        self.listbox = tk.Listbox(self)
        self.listbox.grid(row=3, columnspan=5)

        # The current filter. Setting it to None initially forces the first update.
        self.curr_filter = None

        # All of the items for the listbox.
        self.items = user_names_list

        # The initial update.
        self.on_tick()

    def on_tick(self):
        if self.filter_box.get().lower() != self.curr_filter:
            # The contents of the filter box has changed.
            self.curr_filter = self.filter_box.get().lower()

            # Refresh the listbox.
            self.listbox.delete(0, 'end')

            for item in self.items:
                if self.curr_filter.lower() in item.lower():
                    self.listbox.insert('end', item.lower())

        self.after(250, self.on_tick)
    
    def find_assets_jamf(self):
        name_var = self.listbox.get(tk.ANCHOR)
        try:
            found_asset = jamf_df.loc[jamf_df['Full Name'].str.lower() == name_var.lower()] 
            for index, row in found_asset.iterrows():
                asset_list = row.to_list()
                temp = ["JAMF", "COMPUTER", asset_list[2], asset_list[3], asset_list[4]]
                self.results.append(temp)
        except ValueError as e:
            print(e)
        
    def find_assets_intune(self):
        # search intune
        name_var = self.listbox.get(tk.ANCHOR)
        try:
            found_asset = intune_df.loc[intune_df['Primary user display name'].str.lower() == name_var.lower()]
            for index, row in found_asset.iterrows():
                asset_list = row.to_list()
                temp = ["INTUNE", "COMPUTER", asset_list[2], asset_list[3], asset_list[4]]
                self.results.append(temp)
        except ValueError as e:
            print(e)
    
    def find_assets_jamf_other(self):
        # search jamf other
        name_var = self.listbox.get(tk.ANCHOR)
        try:
            found_asset = jamf_other_df.loc[jamf_other_df['Full Name'].str.lower() == name_var.lower()]
            for index, row in found_asset.iterrows():
                asset_list = row.to_list()
                temp = ["JAMF OTHER", "iPhone / iPad", asset_list[2], asset_list[3], asset_list[4]]
                self.results.append(temp)
        except ValueError as e:
            print(e)

    def find_assets_landlines(self):
        # search landlines
        name_var = self.listbox.get(tk.ANCHOR)
        try:
            found_asset = landlines_df.loc[landlines_df['Line Description 1'].str.lower() == name_var.lower()]
            for index, row in found_asset.iterrows():
                asset_list = row.to_list()
                temp = ["LANDLINES", "PHONE", asset_list[2], asset_list[3], asset_list[4]]
                self.results.append(temp)
        except ValueError as e:
            print(e)
    
    def find_assets_service_now(self):
        # search service now
        name_var = self.listbox.get(tk.ANCHOR)
        print(name_var)
        try:
            found_asset = assets_sn_df.loc[assets_sn_df['assigned_to'].str.lower() == name_var.lower()]
            for index, row in found_asset.iterrows():
                asset_list = row.to_list()
                temp = ["SERVICE NOW ASSET", "COMPUTER", asset_list[2], asset_list[3], asset_list[4]]
                self.results.append(temp)
        except ValueError as e:
            print(e)

    def clear_grid(self):
        print("clearing...")
        for entry in self.grid_entries:
            entry.delete(0, 'end')
        self.results.clear()
        

    def create_grid(self):
        self.clear_grid()

        self.find_assets_jamf()
        self.find_assets_intune()
        self.find_assets_jamf_other()
        self.find_assets_service_now()
        self.find_assets_landlines()

        grid_labels = ["Database", "Device type", "Users name", "Serial Number", "Description"]
        for i,j in enumerate(grid_labels):
            temp = tk.Label(self, text=grid_labels[i], font=('helvetica 18 bold'))
            temp.grid(row=5, column=i)

        for i in range(len(self.results)):
            row_entries = []
            for j in range(5):
                if i * 5 + j < len(self.grid_entries): 
                    temp = self.grid_entries[i * 5 + j]
                    temp.delete(0, 'end')  
                    temp.insert(0, self.results[i][j])
                else:
                    temp = tk.Entry(self, width=30, justify="left")
                    temp.grid(row=6 + i, column=j)
                    temp.insert(0, self.results[i][j])
                    self.grid_entries.append(temp)
                row_entries.append(temp)

App().mainloop()