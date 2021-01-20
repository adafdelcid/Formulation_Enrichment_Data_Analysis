'''
Author: Ada Del Cid 
GitHub: @adafdelcid
Jan.2021

Enrichment_interface: Graphical user interface (GUI) for CSV2Excel_Functionalized.py
'''
from tkinter import *
import os.path
from os import path
import CSV2Excel_Functionalized

root = Tk()
root.geometry("600x500")
root.title("Enrichment Analysis Tool")

# string variables from user input
fsp = StringVar() # Formulation Sheet file Path 
ncp = StringVar() # Normalized Counts file Path
sc = StringVar()  # list of Sorted Cells
dfp = StringVar() # Destination Folder Path
tbp = StringVar() # Top/Bottom Percent
op = StringVar()  # Outliers Percentile
org = StringVar()  # Organ
ct = StringVar()  # Cell Type

def exit1():
	exit()

def enrichment_analysis():

	var1 = fsp.get() # Formulation Sheet file Path
	var2 = ncp.get() # Normalized Counts file Path
	var3 = sc.get()  # list of Sorted Cells
	var4 = dfp.get() # Destination Folder Path
	var5 = tbp.get() # Top/Bottom Percent
	var6 = op.get()  # Outliers Percentile
	var7 = org.get()  # Organ
	var8 = ct.get()  # Cell Type

	var3 = string_to_list(var3)
	var9 = get_cell_type(var7,var8)

	errors = False

	# check for errors with formulation sheet
	if ".xlsx" in var1 and path_exists(var1):
		color1 = "white"
	else:
		color1 = "red"
		errors = True
	
	label9 = Label(root,text = "Invalid formulation Sheet file path!",fg = color1, font = ("arial",12,"bold")).place(x = 20, y = 310)

	# check for errors with normalized counts csv file
	if ".csv" in var2 and path_exists(var2):
		color2 = "white"
	else:
		color2 = "red"
		errors = True

	label10 = Label(root,text = "Invalid normalized counts file path!",fg = color2, font = ("arial",12,"bold")).place(x = 20, y = 330)

	# check for erros with destination folder
	if path_exists(var4):
		color3 = "white"
	else:
		color3 = "red"
		errors = True

	label11 = Label(root,text = "Invalid destination folder!",fg = color3, font = ("arial",12,"bold")).place(x = 20, y = 350)

	# check for erros with top/bottom percent
	try:
		color4 = "white"
		var5 = float(var5)
	except:
		color4 = "red"
		errors = True

	label12 = Label(root,text = "Invalid top/bottom percent, enter a value between 0.1-99.9!",fg = color4, font = ("arial",12,"bold")).place(x = 20, y = 370)

	# check for errors with outliers percentile
	if var6 != "":
		try:
			color5 = "white"
			var6 = float(var6)
		except:
			color5 = "red"
			errors = True
	else:
		color5 = "white"
		var6 = 99.9

	label13 = Label(root,text = "Invalid outlier percentile, enter a value between 0.1-99.9 or leave blank!",fg = color5, font = ("arial",12,"bold")).place(x = 20, y = 390)

	# check for errors with cell type to sort by
	if var9 not in var3:
		color6 = "red"
		errors = True
	else:
		color6 = "white"

	label13 = Label(root,text = "Invalid cell type to sort by, not in list of sorted cells",fg = color6, font = ("arial",12,"bold")).place(x = 20, y = 410)


	if not errors:
		CSV2Excel_Functionalized.run_enrichment_analysis(var4, var1, var2, var3, var5, var9, var6)
		print("Enrichment analysis performed!")
		exit()
	else:
		root.mainloop()

def get_cell_type(organ, cell_type):
	organ_dict = {"Liver" : "V", "Lung" : "L" , "Spleen" : "S", "Heart" : "H", "Kidney" : "K", "Pancreas" : "P", "Marrow" : "M", "Muscle" : "U", "Brain" : "B", "Lymph Node" : "N", "Thymus" : "T"}
	cell_type_dict = {"Hepatocytes" : "H", "Endothelial" : "E", "Kupffer" : "K", "Other Immune" : "I", "Dendritic" : "D", "B cells" : "B", "T cells" : "T", "Macrophages" : "M", "Epithelial" : "EP", "Hematopoetic Stem Cells" : "HSC", "Fibroblasts" : "F", "Satellite Cells" : "SC", "Other" : "O"}
	return organ_dict[organ] + cell_type_dict[cell_type]


def string_to_list(string1):
	string1 = remove_spaces(string1)
	list1 = list(string1.split(","))
	return list1

def remove_spaces(string1):
	return string1.replace(" ","")

def path_exists(path1):
	return path.exists(path1)

label1 = Label(root,text = "Enrichment Analysis",  relief = "solid", font = ("arial",16,"bold")).pack()

label2 = Label(root,text = "Formulation Sheet File Path", font = ("arial",12,"bold")).place(x = 20, y = 50)
entry2 = Entry(root, textvariable=fsp).place(x = 370, y = 48)

label3 = Label(root,text = "Normalized Counts CSV File Path", font = ("arial",12,"bold")).place(x = 20, y = 82)
entry3 = Entry(root, textvariable=ncp).place(x = 370, y = 80)

label4 = Label(root,text = "List of Sorted Cells (separate by commas)", font = ("arial",12,"bold")).place(x = 20, y = 114)
entry4 = Entry(root, textvariable=sc).place(x = 370, y = 112)

label5 = Label(root,text = "Destination Folder Path", font = ("arial",12,"bold")).place(x = 20, y = 146)
entry5 = Entry(root, textvariable=dfp).place(x = 370, y = 144)

label6 = Label(root,text = "Top/Bottom Percent", font = ("arial",12,"bold")).place(x = 20, y = 178)
entry6 = Entry(root, textvariable=tbp).place(x = 370, y = 176)

label7 = Label(root,text = "(OPTIONAL: Default = 99.9) Outliers Percentile", font = ("arial",12,"bold")).place(x = 20, y = 212)
entry7 = Entry(root, textvariable=op).place(x = 370, y = 210)

label8 = Label(root,text = "Analyze by Cell Type", font = ("arial",12,"bold")).place(x = 20, y = 244)

# drop down menu
organ = ["Liver", "Lung", "Spleen", "Heart", "Kidney", "Pancreas", "Marrow", "Muscle", "Brain", "Lymph Node", "Thymus"]
cell_type = ["Hepatocytes", "Endothelial", "Kupffer", "Other Immune", "Dendritic", "B cells", "T cells", "Macrophages", "Epithelial", "Hematopoetic Stem Cells", "Fibroblasts", "Satellite Cells", "Other"]

droplist = OptionMenu(root, org, *organ)
org.set("Select Organ")
droplist.config(width = 15)
droplist.place(x = 200, y = 242)

droplist = OptionMenu(root, ct, *cell_type)
ct.set("Select Cell Type")
droplist.config(width = 15)
droplist.place(x = 384, y = 242)

b1 = Button(root, text = "ENTER", width = 16, fg = "blue", font = ("arial",16), command = enrichment_analysis).place(x = 150, y = 280)
b1 = Button(root, text = "CANCEL", width = 16, fg = "blue", font = ("arial",16), command = exit1).place(x = 300, y = 280)

root.mainloop()