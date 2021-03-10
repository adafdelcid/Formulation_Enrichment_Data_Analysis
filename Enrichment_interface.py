'''
Author: Ada Del Cid
GitHub: @adafdelcid
Jan.2021

Enrichment_interface: Graphical user interface (GUI) for CSV2Excel_Functionalized.py
'''
from tkinter import filedialog, Tk, StringVar, Label, Button, Entry, OptionMenu
from os import path
import sys

import CSV2Excel_Functionalized

class MyGUI: # pylint: disable=too-many-instance-attributes
    '''
    GUI for enrichment analysis
    '''

    def __init__(self, master):
        '''
        Saves basic GUI buttons and data entries
        '''

        self.master = master
        master.geometry("600x500")
        master.title("Enrichment Analysis Tool")

        # string variables from user input
        self.fsp = StringVar() # Formulation Sheet file Path
        self.ncp = StringVar() # Normalized Counts file Path
        self.sc = StringVar()  # list of Sorted Cells
        self.dfp = StringVar() # Destination Folder Path
        self.tbp = StringVar() # Top/Bottom Percent
        self.op = StringVar()  # Outliers Percentile
        self.opt = StringVar() # Average and sort by
        self.org = StringVar() # Organ
        self.ct = StringVar()  # Cell Type

        Label(master, text="Enrichment Analysis", relief="solid", font=("arial", 16,\
            "bold")).pack()

        Label(master, text="Formulation Sheet File Path", font=("arial", 12,\
            "bold")).place(x=20, y=50)
        Button(master, text="Formulation sheet file", width=20, fg="green", font=("arial",\
            16), command=self.open_excel_file).place(x=370, y=48)

        Label(master, text="Normalized Counts CSV File Path", font=("arial", 12,\
            "bold")).place(x=20, y=82)
        Button(master, text="Normalized counts file", width=20, fg="green", font=("arial",\
            16), command=self.open_csv_file).place(x=370, y=80)

        Label(master, text="List of Sorted Cells (separate by commas)", font=("arial", 12,\
            "bold")).place(x=20, y=114)
        Entry(master, textvariable=self.sc).place(x=370, y=112)

        Label(master, text="Destination Folder Path", font=("arial", 12, "bold")).place(\
            x=20, y=146)
        Entry(master, textvariable=self.dfp).place(x=370, y=144)

        Label(master, text="Top/Bottom Percent", font=("arial", 12, "bold")).place(x=20,\
            y=178)
        Entry(master, textvariable=self.tbp).place(x=370, y=176)

        Label(master, text="(OPTIONAL: Default = 99.9) Outliers Percentile", font=("arial",\
            12, "bold")).place(x=20, y=212)
        Entry(master, textvariable=self.op).place(x=370, y=210)

        # Pick how user would like data sorted by
        Label(master, text="Sort by", font=("arial", 12, "bold")).place(x=20,\
            y=244)
        options = ["Organ average","Cell type average", "Average of all samples"]

        org_droplist = OptionMenu(master, self.opt, *options)
        self.opt.set("Sorting method")
        org_droplist.config(width=15)
        org_droplist.place(x=370, y=246)

        Label(master, text="Select organ and/or cell type", font=("arial", 12, "bold")).place(x=20,\
            y=276)

        # drop down menu
        organ = ["Liver", "Lung", "Spleen", "Heart", "Kidney", "Pancreas", "Marrow", "Muscle",\
        "Brain", "Lymph Node", "Thymus"]
        cell_type = ["Hepatocytes", "Endothelial", "Kupffer", "Other Immune", "Dendritic",\
        "B cells", "T cells", "Macrophages", "Epithelial", "Hematopoetic Stem Cells",\
        "Fibroblasts", "Satellite Cells", "Other"]

        org_droplist = OptionMenu(master, self.org, *organ)
        self.org.set("Select Organ")
        org_droplist.config(width=15)
        org_droplist.place(x=220, y=278)

        ct_droplist = OptionMenu(master, self.ct, *cell_type)
        self.ct.set("Select Cell Type")
        ct_droplist.config(width=15)
        ct_droplist.place(x=400, y=278)

        Button(master, text="ENTER", width=16, fg="blue", font=("arial", 16),\
            command=self.enrichment_analysis).place(x=150, y=310)
        Button(master, text="CANCEL", width=16, fg="blue", font=("arial", 16),\
            command=exit1).place(x=300, y=310)

    def open_excel_file(self):
        '''
        To open a file searcher and select a file
        '''
        self.fsp = filedialog.askopenfilename()

    def open_csv_file(self):
        '''
        To open a file searcher and csv file
        '''
        self.ncp = filedialog.askopenfilename()

    def enrichment_analysis(self): # pylint: disable=too-many-branches
    # pylint: disable=too-many-statements
        '''
        Checks for any entry errors, returns list of errors or runs the enrichment analysis
        '''
        errors = False

        temp_opt = self.opt.get()
        temp_org = self.org.get()
        temp_ct = self.ct.get()

        if temp_opt == "Sorting method":
            color0 = "red"
            errors = True
            temp_org = ""
            temp_ct = ""
        elif temp_opt == "Average of all samples":
            color0 = "white"
            temp_org = ""
            temp_ct = ""
        elif temp_org == "Select Organ":
            color0 = "red"
            errors = True
            temp_ct = ""
        elif temp_opt == "Cell type average" and temp_ct == "Select Cell Type":
            color0 = "red"
            errors = True
            temp_org = ""
        else:
            color0 = "white"

        cell_types = string_to_list(self.sc.get()) # list of cell types
        sort_by = get_cell_type(temp_opt, temp_org, temp_ct) # cell type to base analysis
        print(sort_by)
        print(cell_types)
        fold_path = self.dfp.get() # Destination Folder Path
        percent = self.tbp.get() # Top/Bottom Percent
        percentile = self.op.get()  # Outliers Percentile

        # check for errors with formulation sheet
        if ".xlsx" in self.fsp and path_exists(self.fsp):
            color1 = "white"
        else:
            color1 = "red"
            errors = True

        Label(self.master, text="Invalid formulation sheet file path!", fg=color1,\
            font=("arial", 12, "bold")).place(x=20, y=350)

        # check for errors with normalized counts csv file
        if ".csv" in self.ncp and path_exists(self.ncp):
            color2 = "white"
        else:
            color2 = "red"
            errors = True

        Label(self.master, text="Invalid normalized counts file path!", fg=color2,\
            font=("arial", 12, "bold")).place(x=20, y=370)

        # check for erros with destination folder
        if path_exists(fold_path):
            color3 = "white"
        else:
            color3 = "red"
            errors = True

        Label(self.master, text="Invalid destination folder!", fg=color3,\
            font=("arial", 12, "bold")).place(x=20, y=390)

        # check for errors with top/bottom percent
        try:
            color4 = "white"
            percent = float(percent)
        except ValueError:
            color4 = "red"
            errors = True

        Label(self.master,\
            text="Invalid top/bottom percent, enter a value between 0.1-99.9!",\
            fg=color4, font=("arial", 12, "bold")).place(x=20, y=410)

        # check for errors with outliers percentile
        if percentile != "":
            try:
                color5 = "white"
                percentile = float(percentile)
            except ValueError:
                color5 = "red"
                errors = True
        else:
            color5 = "white"
            percentile = 99.9

        Label(self.master,\
            text="Invalid outlier percentile, enter a value between 0.1-99.9 or leave blank!",\
            fg=color5, font=("arial", 12, "bold")).place(x=20, y=430)

        # check for sort_by values to check for error here
        if sort_by == "AVG":
            color6 = "white"
        elif len(sort_by) == 1:
            check = False
            for item in cell_types:
                if item[0] == sort_by:
                    check = True
                    color6 = "white"
                    break
            if not check:
                color6 = "red"
                errors = True
        elif sort_by not in cell_types:
            color6 = "red"
            errors = True
        else:
            color6 = "white"

        Label(self.master,\
            text="Invalid cell type to sort by, not in list of sorted cells",\
            fg=color6, font=("arial", 12, "bold")).place(x=20, y=450)

        Label(self.master,\
            text="Missing sort by method, organ and/or cell type",\
            fg=color0, font=("arial", 12, "bold")).place(x=20, y=470)

        if not errors:
            CSV2Excel_Functionalized.run_enrichment_analysis(fold_path, self.fsp, self.ncp,
                cell_types, percent, sort_by, percentile)
            print("Enrichment analysis performed!")
            exit1()
        else:
            self.master.mainloop()

def get_cell_type(option, organ="", cell_type=""):
    '''
    Returns acronym of organ and cell type
    '''
    sort_by_ct = ""

    organ_dict = {"Liver":"V", "Lung":"L", "Spleen":"S", "Heart":"H", "Kidney":"K",\
    "Pancreas":"P", "Marrow":"M", "Muscle":"U", "Brain":"B", "Lymph Node":"N",\
    "Thymus":"T"}
    cell_type_dict = {"Hepatocytes":"H", "Endothelial":"E", "Kupffer":"K", "Other Immune":"I",\
    "Dendritic":"D", "B cells":"B", "T cells":"T", "Macrophages":"M", "Epithelial":"EP",\
    "Hematopoetic Stem Cells":"HSC", "Fibroblasts":"F", "Satellite Cells":"SC", "Other":"O"}

    if option == "Organ average":
        try:
            sort_by_ct = organ_dict[organ]
        except KeyError:
            sort_by_ct = ""
    elif option == "Cell type average":
        try:
            sort_by_ct = organ_dict[organ] + cell_type_dict[cell_type]
        except KeyError:
            sort_by_ct = ""
    elif option == "Average of all samples":
        sort_by_ct = "AVG"
    else:
        sort_by_ct = ""

    return sort_by_ct

def exit1():
    '''
    exit and close GUI
    '''
    sys.exit()

def string_to_list(string1):
    '''
    Creates a list out of a string of items separated by commas
    '''
    string1 = remove_spaces(string1)
    list1 = list(string1.split(","))
    return list1

def remove_spaces(string1):
    '''
    remove any unnecessary spaces
    '''
    return string1.replace(" ", "")

def path_exists(path1):
    '''
    Checks if a directory path exists
    '''
    return path.exists(path1)

root = Tk()
my_gui = MyGUI(root)
root.mainloop()
