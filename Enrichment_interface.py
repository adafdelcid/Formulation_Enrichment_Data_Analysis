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

class MyGUI:
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

		Label(master, text="Analyze by Cell Type", font=("arial", 12, "bold")).place(x=20,\
			y=244)

		# drop down menu
		organ = ["Liver", "Lung", "Spleen", "Heart", "Kidney", "Pancreas", "Marrow", "Muscle",\
		"Brain", "Lymph Node", "Thymus"]
		cell_type = ["Hepatocytes", "Endothelial", "Kupffer", "Other Immune", "Dendritic",\
		"B cells", "T cells", "Macrophages", "Epithelial", "Hematopoetic Stem Cells", "Fibroblasts",\
		"Satellite Cells", "Other"]

		org_droplist = OptionMenu(master, self.org, *organ)
		self.org.set("Select Organ")
		org_droplist.config(width=15)
		org_droplist.place(x=200, y=242)

		ct_droplist = OptionMenu(master, self.ct, *cell_type)
		self.ct.set("Select Cell Type")
		ct_droplist.config(width=15)
		ct_droplist.place(x=384, y=242)

		Button(master, text="ENTER", width=16, fg="blue", font=("arial", 16),\
			command=self.enrichment_analysis).place(x=150, y=280)
		Button(master, text="CANCEL", width=16, fg="blue", font=("arial", 16),\
			command=exit1).place(x=300, y=280)

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

	def enrichment_analysis(self):
		'''
		Checks for any entry errors, returns list of errors or runs the enrichment analysis
		'''

		var1 = string_to_list(self.sc.get()) # list of cell types
		var2 = get_cell_type(self.org.get(), self.ct.get()) # cell type to base analysis
		var3 = self.dfp.get() # Destination Folder Path
		var4 = self.tbp.get() # Top/Bottom Percent
		var5 = self.op.get()  # Outliers Percentile

		errors = False

		# check for errors with formulation sheet
		if ".xlsx" in self.fsp and path_exists(self.fsp):
			color1 = "white"
		else:
			color1 = "red"
			errors = True

		Label(self.master, text="Invalid formulation sheet file path!", fg=color1,\
			font=("arial", 12, "bold")).place(x=20, y=310)

		# check for errors with normalized counts csv file
		if ".csv" in self.ncp and path_exists(self.ncp):
			color2 = "white"
		else:
			color2 = "red"
			errors = True

		Label(self.master, text="Invalid normalized counts file path!", fg=color2,\
			font=("arial", 12, "bold")).place(x=20, y=330)

		# check for erros with destination folder
		if path_exists(var3):
			color3 = "white"
		else:
			color3 = "red"
			errors = True

		Label(self.master, text="Invalid destination folder!", fg=color3,\
			font=("arial", 12, "bold")).place(x=20, y=350)

		# check for erros with top/bottom percent
		try:
			color4 = "white"
			var4 = float(var4)
		except TypeError:
			color4 = "red"
			errors = True

		Label(self.master,\
			text="Invalid top/bottom percent, enter a value between 0.1-99.9!",\
			fg=color4, font=("arial", 12, "bold")).place(x=20, y=370)

		# check for errors with outliers percentile
		if var5 != "":
			try:
				color5 = "white"
				var5 = float(var5)
			except TypeError:
				color5 = "red"
				errors = True
		else:
			color5 = "white"
			var5 = 99.9

		Label(self.master,\
			text="Invalid outlier percentile, enter a value between 0.1-99.9 or leave blank!",\
			fg=color5, font=("arial", 12, "bold")).place(x=20, y=390)

		# check for errors with cell type to sort by
		if var2 not in var1:
			color6 = "red"
			errors = True
		else:
			color6 = "white"

		Label(self.master,\
			text="Invalid cell type to sort by, not in list of sorted cells",\
			fg=color6, font=("arial", 12, "bold")).place(x=20, y=410)

		if not errors:
			CSV2Excel_Functionalized.run_enrichment_analysis(var3, self.fsp, self.ncp, var1,\
				var4, var2, var5)
			print("Enrichment analysis performed!")
			exit1()
		else:
			self.master.mainloop()

def get_cell_type(organ, cell_type):
	'''
	Returns acronym of organ and cell type
	'''

	organ_dict = {"Liver":"V", "Lung":"L", "Spleen":"S", "Heart":"H", "Kidney":"K",\
	"Pancreas":"P", "Marrow":"M", "Muscle":"U", "Brain":"B", "Lymph Node":"N",\
	"Thymus":"T"}
	cell_type_dict = {"Hepatocytes":"H", "Endothelial":"E", "Kupffer":"K", "Other Immune":"I",\
	"Dendritic":"D", "B cells":"B", "T cells":"T", "Macrophages":"M", "Epithelial":"EP",\
	"Hematopoetic Stem Cells":"HSC", "Fibroblasts":"F", "Satellite Cells":"SC", "Other":"O"}
	return organ_dict[organ] + cell_type_dict[cell_type]

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
