import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook

'''
By Ada Del Cid 
Oct 2020
'''
#-------------------------------------------------------------------------------
#		Imports csv file with barcode readings and create spreadsheet 
#		with formulation sheets and normalized barcode readings
#-------------------------------------------------------------------------------

#Create new excel sheet
new_file_name = "/Users/adadelcid/Documents/Python_Enrichment_Project/Normalized Counts.xlsx"
wb = Workbook()
ws = wb.active
wb.save(new_file_name)

#Turn formulation sheet into data frames
formulation_sheet_filepath = "/Users/adadelcid/Documents/Python_Enrichment_Project/Formulation Sheet.xlsx"
df_formulations = pd.read_excel(formulation_sheet_filepath, sheet_name = "Formulations")

#Make spreadsheet with formulation sheet
with pd.ExcelWriter(new_file_name,engine="openpyxl",mode = "w") as writer:
	df_formulations.to_excel(writer, sheet_name = "Formulations")

#Save spreadsheet with formulations
xfile = openpyxl.load_workbook(new_file_name)
sheet = xfile["Formulations"]
sheet.delete_cols(1)
xfile.save(new_file_name)

#Read CSV file and save as data frame
csv_filepath = "/Users/adadelcid/Documents/Python_Enrichment_Project/normcounts.csv"
df = pd.read_csv(csv_filepath,sep = ',',header= 0)

#copy csv data frame to excel sheet
new_sheet_name = "Normalized Counts"
with pd.ExcelWriter(new_file_name,engine="openpyxl",mode = "a") as writer:
	df.to_excel(writer, sheet_name = new_sheet_name)

#save and edit the excel sheet (name cell B1 : "BC")
xfile = openpyxl.load_workbook(new_file_name)
sheet = xfile[new_sheet_name]
sheet["B1"] = "BC"
sheet.delete_cols(1)
xfile.save(new_file_name)

#-------------------------------------------------------------------------------
#					Calculate 99.9 percentile and remove outliers
#-------------------------------------------------------------------------------

#get row names and save barcode column
df_normalized_counts = pd.read_excel(new_file_name, sheet_name = new_sheet_name)
rows = list(df_normalized_counts.index)
barcodes = df_normalized_counts["BC"]


#calculate 99.9 percentile
percentile = 99.9
df_normalized_no_bc = df_normalized_counts.drop("BC",axis = 1) #temporarily remove BC colum
percentile_99 = np.percentile(df_normalized_no_bc.to_numpy(),percentile) #find 99.9 percentile value

#set outliers to NaN (numbers greater than 99.9 percentile)
columns = df_normalized_no_bc.columns.tolist()
for row in rows:
	for column in columns:
		if df_normalized_no_bc.at[row,column] >= percentile_99:
			df_normalized_no_bc.at[row,column] = np.nan

df_normalized_no_bc.insert(loc=0,column = "BC", value = barcodes) #add the BC column again

#-------------------------------------------------------------------------------
#					Get organized columns of data frame by cell type
#-------------------------------------------------------------------------------

#list of sorted cells
'''**********************must ask user to specify cells types*******************'''
sorted_cells = ["SB","SE","SM","ST","VE","VH","VI","VK"] 

#organize columns
organized_columns = []
for cell_type in sorted_cells:
	for column in columns:
		if cell_type in column:
			organized_columns.append(column)

count_repeats = len(columns)//len(sorted_cells)#number of repeats per cell type, should be same for all

#-------------------------------------------------------------------------------
#		Combines formulation sheet and normalized count into data frame
#-------------------------------------------------------------------------------

#inner mege of data frames around barcodes ("BC")
df_merged = df_formulations.merge(df_normalized_no_bc,on="BC")

#ordered columns
l1 = df_merged.columns.tolist()[:10] #columns up to phospholipid%
order_columns = l1 + organized_columns

#rearrange columns on df_merged
df_merged = df_merged[order_columns]

#append merged data frames onto excel spreadsheet
with pd.ExcelWriter(new_file_name,engine="openpyxl",mode = "a") as writer:
	df_merged.to_excel(writer, sheet_name = "Formulations + Norm Counts")

#edit the excel sheet
xfile = openpyxl.load_workbook(new_file_name)
sheet = xfile["Formulations + Norm Counts"]
sheet.delete_cols(1)
xfile.save(new_file_name)

#-------------------------------------------------------------------------------
#						Find average of each cell type
#-------------------------------------------------------------------------------
df_averaged = df_merged

for index in range(len(sorted_cells)):
	temporary_list = []
	for column in organized_columns:
		if sorted_cells[index] in column:
			temporary_list.append(column)
	df_averaged[sorted_cells[index]] = df_averaged[temporary_list].mean(axis=1)

order_columns = l1 + sorted_cells
df_averaged = df_averaged[order_columns]

#append merged data frames onto excel spreadsheet
with pd.ExcelWriter(new_file_name,engine="openpyxl",mode = "a") as writer:
	df_averaged.to_excel(writer, sheet_name = "Formulation Enrichment")

#edit the excel sheet
xfile = openpyxl.load_workbook(new_file_name)
sheet = xfile["Formulation Enrichment"]
sheet.delete_cols(1)
xfile.save(new_file_name)

#-------------------------------------------------------------------------------
#						Lipomer Enrichment by percent
#-------------------------------------------------------------------------------

lipomers = []

for index in range(len(df_averaged["Lipomer %"].values)-2):
	if df_averaged["Lipomer %"].values[index] not in lipomers:
		lipomers.append(df_averaged["Lipomer %"].values[index])
lipomers.sort()

lipomer_total = [0]*len(lipomers)
for bc_x in df_averaged["Lipomer %"].values:
	for index in range(len(lipomers)):
		if bc_x == lipomers[index]:
			lipomer_total[index] += 1
			break

total = sum(lipomer_total)

lipomer_percent_total = []
for each_lipomer in lipomer_total:
	lipomer_percent_total.append(round(each_lipomer/total,9))

lipomers.append("TOTAL")
lipomer_total.append(total)

#-------------------------------------------------------------------------------
#						Cholesterol Enrichment by percent
#-------------------------------------------------------------------------------

cholesterols = []

for index in range(len(df_averaged["Cholesterol %"].values)-2):
	if df_averaged["Cholesterol %"].values[index] not in cholesterols:
		cholesterols.append(df_averaged["Cholesterol %"].values[index])
cholesterols.sort()

cholesterol_total = [0]*len(cholesterols)
for bc_x in df_averaged["Cholesterol %"].values:
	for index in range(len(cholesterols)):
		if bc_x == cholesterols[index]:
			cholesterol_total[index] += 1
			break

cholesterol_percent_total = []
for each_cholesterol in cholesterol_total:
	cholesterol_percent_total.append(round(each_cholesterol/total,9))

cholesterols.append("TOTAL")
cholesterol_total.append(total)

#-------------------------------------------------------------------------------
#						PEG Enrichment by percent
#-------------------------------------------------------------------------------

pegs = []

for index in range(len(df_averaged["PEG %"].values)-2):
	if df_averaged["PEG %"].values[index] not in pegs:
		pegs.append(df_averaged["PEG %"].values[index])
pegs.sort()

peg_total = [0]*len(pegs)
for bc_x in df_averaged["PEG %"].values:
	for index in range(len(pegs)):
		if bc_x == pegs[index]:
			peg_total[index] += 1
			break

peg_percent_total = []
for each_peg in peg_total:
	peg_percent_total.append(round(each_peg/total,9))

pegs.append("TOTAL")
peg_total.append(total)

#-------------------------------------------------------------------------------
#					Phospholipid Enrichment by percent
#-------------------------------------------------------------------------------

phospholipids = []

for index in range(len(df_averaged["Phospholipid %"].values)-2):
	if df_averaged["Phospholipid %"].values[index] not in phospholipids:
		phospholipids.append(df_averaged["Phospholipid %"].values[index])
phospholipids.sort()

phospholipid_total = [0]*len(phospholipids)
for bc_x in df_averaged["Phospholipid %"].values:
	for index in range(len(phospholipids)):
		if bc_x == phospholipids[index]:
			phospholipid_total[index] += 1
			break

phospholipid_percent_total = []
for each_phospholipid in phospholipid_total:
	phospholipid_percent_total.append(round(each_phospholipid/total,9))

phospholipids.append("TOTAL")
phospholipid_total.append(total)

#-------------------------------------------------------------------------------
#							Lipomer Enrichment
#-------------------------------------------------------------------------------

lipomer_list = []

for index in range(len(df_averaged["Lipomer"].values)-2):
	if df_averaged["Lipomer"].values[index] not in lipomer_list:
		lipomer_list.append(df_averaged["Lipomer"].values[index])
lipomer_list.sort()

lipomer_list_total = [0]*len(lipomer_list)
for bc_x in df_averaged["Lipomer"].values:
	for index in range(len(lipomer_list)):
		if bc_x == lipomer_list[index]:
			lipomer_list_total[index] += 1
			break

lipomer_list_percent_total = []
for each_lipomer in lipomer_list_total:
	lipomer_list_percent_total.append(round(each_lipomer/total,9))

lipomer_list.append("TOTAL")
lipomer_list_total.append(total)

#-------------------------------------------------------------------------------
#						Cholesterol Enrichment
#-------------------------------------------------------------------------------

cholesterol_list = []

for index in range(len(df_averaged["Cholesterol"].values)-2):
	if df_averaged["Cholesterol"].values[index] not in cholesterol_list:
		cholesterol_list.append(df_averaged["Cholesterol"].values[index])
cholesterol_list.sort()

cholesterol_list_total = [0]*len(cholesterol_list)
for bc_x in df_averaged["Cholesterol"].values:
	for index in range(len(cholesterol_list)):
		if bc_x == cholesterol_list[index]:
			cholesterol_list_total[index] += 1
			break

cholesterol_list_percent_total = []
for each_cholesterol in cholesterol_list_total:
	cholesterol_list_percent_total.append(round(each_cholesterol/total,9))

cholesterol_list.append("TOTAL")
cholesterol_list_total.append(total)

#-------------------------------------------------------------------------------
#						PEG Enrichment
#-------------------------------------------------------------------------------

peg_list = []

for index in range(len(df_averaged["PEG"].values)-2):
	if df_averaged["PEG"].values[index] not in peg_list:
		peg_list.append(df_averaged["PEG"].values[index])
peg_list.sort()

peg_list_total = [0]*len(peg_list)
for bc_x in df_averaged["PEG"].values:
	for index in range(len(peg_list)):
		if bc_x == peg_list[index]:
			peg_list_total[index] += 1
			break

peg_list_percent_total = []
for each_peg in peg_list_total:
	peg_list_percent_total.append(round(each_peg/total,9))

peg_list.append("TOTAL")
peg_list_total.append(total)

#-------------------------------------------------------------------------------
#					Phospholipid Enrichment
#-------------------------------------------------------------------------------

phospholipid_list = []

for index in range(len(df_averaged["Phospholipid"].values)-2):
	if df_averaged["Phospholipid"].values[index] not in phospholipid_list:
		phospholipid_list.append(df_averaged["Phospholipid"].values[index])
phospholipid_list.sort()

phospholipid_list_total = [0]*len(phospholipid_list)
for bc_x in df_averaged["Phospholipid"].values:
	for index in range(len(phospholipid_list)):
		if bc_x == phospholipid_list[index]:
			phospholipid_list_total[index] += 1
			break

phospholipid_list_percent_total = []
for each_phospholipid in phospholipid_list_total:
	phospholipid_list_percent_total.append(round(each_phospholipid/total,9))

phospholipid_list.append("TOTAL")
phospholipid_list_total.append(total)




