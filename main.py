#  This script is used to analyse the brepla data


import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import xlsxwriter
import tkinter as tk
from tkinter import filedialog
import os


# define a class to hold individual data
class Measurement:
    def __init__(self, file_loc):

        self.file_location = file_loc  # Working folder
        self.name = ""  # Sample Name
        self.data = pd.DataFrame()  # Placeholder for the data
        self.results = pd.DataFrame(data=np.zeros([1,7]), columns=["Name", "60", "std_60", "Shore", "std_Shore", "CA", "std_CA"])  # Placeholder for the averaged data

    def read_excel(self):
        # Read Excel Data
        self.data = pd.read_excel(self.file_location)

        columns = ["Index", "Name", "20", "60", "85", "Shore", "CA"]
        columns_short = ["Index", "Name", "60", "Shore", "CA"]

        if len(self.data.columns) == 3:
            self.data.columns = columns_short[0:3]
        elif len(self.data.columns) == 4:
            self.data.columns = columns_short[0:4]
        elif len(self.data.columns) == 5:
            self.data.columns = columns_short[0:5]
        elif len(self.data.columns) == 6:
            self.data.columns = columns[0:6]
        elif len(self.data.columns) == 7:
            self.data.columns = columns[0:7]
        else:
            print("Error")

        # print(self.data)

        return

    def calculate_averages(self):

        self.name = str.rsplit(self.file_location, sep="\\")[1]
        print(self.name)
        self.name = str.split(self.name, sep=".")[0]
        self.results = self.results.astype({"Name": str})
        self.results.loc[0, "Name"] = str(self.name)

        try:
            data_60 = self.data["60"]

            self.results.loc[0, "60"] = data_60.mean()
            self.results.loc[0, "std_60"] = data_60.std()

        except KeyError:
            print("No data for gloss 60° found")

        try:
            data_Shore = self.data["Shore"]

            self.results.loc[0, "Shore"] = data_Shore.mean()
            self.results.loc[0, "std_Shore"] = data_Shore.std()

        except KeyError:
            print("No data for Shore D hardness found")

        try:
            data_CA = self.data["CA"]

            self.results.loc[0, "CA"] = data_CA.mean()
            self.results.loc[0, "std_CA"] = data_CA.std()

        except KeyError:
            print("No data for Shore D hardness found")

        return

def scan_folder():
    #Scan folder for excel files

    root = tk.Tk()
    root.withdraw()  # Hide the root window

    folder_selected = filedialog.askdirectory(title="Select Folder Containing Excel Files")

    files_list = []

    if folder_selected:
        excel_files = [file for file in os.listdir(folder_selected) if file.endswith(".xlsx") or file.endswith(".xls")]
        if excel_files:
            for file in excel_files:
                files_list.append(os.path.join(folder_selected, file))
        else:
            print("No Excel files found in the selected folder.")
    else:
        print("No folder selected.")

    files_tuple = tuple(files_list)
    return files_tuple

def save_excel(to_excel):
    filepath_export = filedialog.asksaveasfilename(defaultextension='xlsx', filetypes=[('Excel File', '.xlsx')])
    with pd.ExcelWriter(filepath_export, engine='xlsxwriter',
                        engine_kwargs={'options': {'strings_to_numbers': True}}) as writer:
        to_excel.to_excel(writer, sheet_name='Results')


excel_files_tuple = scan_folder()
measurements = []
end_results = pd.DataFrame(columns=["Name", "60", "std_60", "Shore", "std_Shore", "CA", "std_CA"])

for file in excel_files_tuple:
    measurements.append(Measurement(file))

for result in measurements:
    result.read_excel()
    result.calculate_averages()
    end_results = pd.concat([end_results, result.results])

save_excel(end_results)
