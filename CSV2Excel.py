'''
Author: Ada Del Cid 
GitHub: adafdelcid
Oct 2020
'''

import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook

# -------------------------------------------------------------------------------
# Imports csv file with barcode readings and create spreadsheet 
# with formulation sheets and normalized barcode readings
# -------------------------------------------------------------------------------

# Create new excel sheet
new_file_name = "/Users/adadelcid/Documents/Python_Enrichment_Project/Normalized Counts.xlsx" #RENAME: 
wb = Workbook()
ws = wb.active
wb.save(new_file_name)

# Turn formulation sheet into data frames
formulation_sheet_filepath = "/Users/adadelcid/Documents/Python_Enrichment_Project/Formulation Sheet.xlsx"
df_formulations = pd.read_excel(formulation_sheet_filepath, sheet_name = "Formulations")

# Make spreadsheet with formulation sheet
with pd.ExcelWriter(new_file_name,engine="openpyxl",mode = "w") as writer:
	df_formulations.to_excel(writer, sheet_name = "Formulations")

# Save spreadsheet with formulations
xfile = openpyxl.load_workbook(new_file_name)
sheet = xfile["Formulations"]
sheet.delete_cols(1)
xfile.save(new_file_name)

# Read CSV file and save as data frame
csv_filepath = "/Users/adadelcid/Documents/Python_Enrichment_Project/normcounts.csv"
df = pd.read_csv(csv_filepath,sep = ',',header= 0)

# Copy csv data frame to excel sheet
new_sheet_name = "Normalized Counts"
with pd.ExcelWriter(new_file_name,engine="openpyxl",mode = "a") as writer:
	df.to_excel(writer, sheet_name = new_sheet_name)

# Save and edit the excel sheet (name cell B1 : "BC")
xfile = openpyxl.load_workbook(new_file_name)
sheet = xfile[new_sheet_name]
sheet["B1"] = "BC"
sheet.delete_cols(1)
xfile.save(new_file_name)

# -------------------------------------------------------------------------------
# Calculate 99.9 percentile and remove outliers
# -------------------------------------------------------------------------------

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
#						Sample enrichment calculations component_x
#-------------------------------------------------------------------------------

#list = [] empty list to save component x to be studied

#for component_x in each formulation:
#	if component_x not in list of saved components:
#		add component_x to list
#organize list alphabetically

#list of total enrichment for component_x = list of zeros
#for each formulation in component_x:
# 	for index in range(len(list)):
# 		if component_type == item on list at index:
# 			list[index] += 1
# 			break

# total = sum(list of total enrichment for component_x)

#list_percent_total = []
# for each_item in list of total enrichment for component_x :
# 	(round(each_lipomer/total,9)) append to list of total enrichment for component_x

# list.append("TOTAL")
# list of total enrichment for component_x .append(total)

#-------------------------------------------------------------------------------
#						Get list of components            
#-------------------------------------------------------------------------------

#start lists to save items
lipomers = []
cholesterols = []
pegs = []
phospholipids = []
lipomer_list = []
cholesterol_list = []
peg_list = []
phospholipid_list = []

for index in range(len(df_averaged["Lipomer %"].values)-2):
	if df_averaged["Lipomer %"].values[index] not in lipomers:
		lipomers.append(df_averaged["Lipomer %"].values[index])
	if df_averaged["Cholesterol %"].values[index] not in cholesterols:
		cholesterols.append(df_averaged["Cholesterol %"].values[index])
	if df_averaged["PEG %"].values[index] not in pegs:
		pegs.append(df_averaged["PEG %"].values[index])
	if df_averaged["Phospholipid %"].values[index] not in phospholipids:
		phospholipids.append(df_averaged["Phospholipid %"].values[index])
	if df_averaged["Lipomer"].values[index] not in lipomer_list:
		lipomer_list.append(df_averaged["Lipomer"].values[index])
	if df_averaged["Cholesterol"].values[index] not in cholesterol_list:
		cholesterol_list.append(df_averaged["Cholesterol"].values[index])
	if df_averaged["PEG"].values[index] not in peg_list:
		peg_list.append(df_averaged["PEG"].values[index])
	if df_averaged["Phospholipid"].values[index] not in phospholipid_list:
		phospholipid_list.append(df_averaged["Phospholipid"].values[index])

lipomers.sort()
cholesterols.sort()
pegs.sort()
phospholipids.sort()
lipomer_list.sort()
cholesterol_list.sort()
peg_list.sort()
phospholipid_list.sort()

#-------------------------------------------------------------------------------
#						Lipomer Enrichment by percent
#-------------------------------------------------------------------------------

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

lipomers.insert(0,"Lipomer")
lipomer_total.insert(0,"Total #")
lipomer_percent_total.insert(0,"% of Total")

lipomers.append("TOTAL")
lipomer_total.append(total)
lipomer_percent_total.append(1)

t_lipomers = [lipomers,lipomer_total,lipomer_percent_total]
np_temporary = np.array(t_lipomers)
np_temporary = np_temporary.T

#-------------------------------------------------------------------------------
#						Cholesterol Enrichment by percent
#-------------------------------------------------------------------------------

cholesterol_total = [0]*len(cholesterols)
for bc_x in df_averaged["Cholesterol %"].values:
	for index in range(len(cholesterols)):
		if bc_x == cholesterols[index]:
			cholesterol_total[index] += 1
			break

cholesterol_percent_total = []
for each_cholesterol in cholesterol_total:
	cholesterol_percent_total.append(round(each_cholesterol/total,9))

cholesterols.insert(0,"Cholesterol")
cholesterol_total.insert(0,"Total #")
cholesterol_percent_total.insert(0,"% of Total")

cholesterols.append("TOTAL")
cholesterol_total.append(total)
cholesterol_percent_total.append(1)

t_cholesterols = [cholesterols,cholesterol_total,cholesterol_percent_total]
np_temporary = np.array(t_cholesterols)
t_cholesterols = np_temporary.T.tolist()

#-------------------------------------------------------------------------------
#						PEG Enrichment by percent
#-------------------------------------------------------------------------------

peg_total = [0]*len(pegs)
for bc_x in df_averaged["PEG %"].values:
	for index in range(len(pegs)):
		if bc_x == pegs[index]:
			peg_total[index] += 1
			break

peg_percent_total = []
for each_peg in peg_total:
	peg_percent_total.append(round(each_peg/total,9))

pegs.insert(0,"PEG")
peg_total.insert(0,"Total #")
peg_percent_total.insert(0,"% of Total")

pegs.append("TOTAL")
peg_total.append(total)
peg_percent_total.append(1)

t_pegs = [pegs,peg_total,peg_percent_total]
np_temporary = np.array(t_pegs)
t_pegs = np_temporary.T.tolist()

#-------------------------------------------------------------------------------
#					Phospholipid Enrichment by percent
#-------------------------------------------------------------------------------

phospholipid_total = [0]*len(phospholipids)
for bc_x in df_averaged["Phospholipid %"].values:
	for index in range(len(phospholipids)):
		if bc_x == phospholipids[index]:
			phospholipid_total[index] += 1
			break

phospholipid_percent_total = []
for each_phospholipid in phospholipid_total:
	phospholipid_percent_total.append(round(each_phospholipid/total,9))

phospholipids.insert(0,"Phospholipid")
phospholipid_total.insert(0,"Total #")
phospholipid_percent_total.insert(0,"% of Total")

phospholipids.append("TOTAL")
phospholipid_total.append(total)
phospholipid_percent_total.append(1)

t_phospholipids = [phospholipids,phospholipid_total,phospholipid_percent_total]
np_temporary = np.array(t_phospholipids)
t_phospholipids = np_temporary.T.tolist()

#-------------------------------------------------------------------------------
#							Lipomer Enrichment
#-------------------------------------------------------------------------------

lipomer_list_total = [0]*len(lipomer_list)
for bc_x in df_averaged["Lipomer"].values:
	for index in range(len(lipomer_list)):
		if bc_x == lipomer_list[index]:
			lipomer_list_total[index] += 1
			break

lipomer_list_percent_total = []
for each_lipomer in lipomer_list_total:
	lipomer_list_percent_total.append(round(each_lipomer/total,9))

lipomer_list.insert(0,"Phospholipid")
lipomer_list_total.insert(0,"Total #")
lipomer_list_percent_total.insert(0,"% of Total")

lipomer_list.append("TOTAL")
lipomer_list_total.append(total)
lipomer_list_percent_total.append(1)

t_lipomer_list = [lipomer_list,lipomer_list_total,lipomer_list_percent_total]
np_temporary = np.array(t_lipomer_list)
t_lipomer_list = np_temporary.T.tolist()

#-------------------------------------------------------------------------------
#						Cholesterol Enrichment
#-------------------------------------------------------------------------------

cholesterol_list_total = [0]*len(cholesterol_list)
for bc_x in df_averaged["Cholesterol"].values:
	for index in range(len(cholesterol_list)):
		if bc_x == cholesterol_list[index]:
			cholesterol_list_total[index] += 1
			break

cholesterol_list_percent_total = []
for each_cholesterol in cholesterol_list_total:
	cholesterol_list_percent_total.append(round(each_cholesterol/total,9))

cholesterol_list.insert(0,"Phospholipid")
cholesterol_list_total.insert(0,"Total #")
cholesterol_list_percent_total.insert(0,"% of Total")

cholesterol_list.append("TOTAL")
cholesterol_list_total.append(total)
cholesterol_list_percent_total.append(1)

t_cholesterol_list = [cholesterol_list,cholesterol_list_total,cholesterol_list_percent_total]
np_temporary = np.array(t_cholesterol_list)
t_cholesterol_list = np_temporary.T.tolist()

#-------------------------------------------------------------------------------
#						PEG Enrichment
#-------------------------------------------------------------------------------

peg_list_total = [0]*len(peg_list)
for bc_x in df_averaged["PEG"].values:
	for index in range(len(peg_list)):
		if bc_x == peg_list[index]:
			peg_list_total[index] += 1
			break

peg_list_percent_total = []
for each_peg in peg_list_total:
	peg_list_percent_total.append(round(each_peg/total,9))

peg_list.insert(0,"Phospholipid")
peg_list_total.insert(0,"Total #")
peg_list_percent_total.insert(0,"% of Total")

peg_list.append("TOTAL")
peg_list_total.append(total)
peg_list_percent_total.append(1)

t_peg_list = [peg_list,peg_list_total,peg_list_percent_total]
np_temporary = np.array(t_peg_list)
t_peg_list = np_temporary.T.tolist()

#-------------------------------------------------------------------------------
#					Phospholipid Enrichment
#-------------------------------------------------------------------------------

phospholipid_list_total = [0]*len(phospholipid_list)
for bc_x in df_averaged["Phospholipid"].values:
	for index in range(len(phospholipid_list)):
		if bc_x == phospholipid_list[index]:
			phospholipid_list_total[index] += 1
			break

phospholipid_list_percent_total = []
for each_phospholipid in phospholipid_list_total:
	phospholipid_list_percent_total.append(round(each_phospholipid/total,9))

phospholipid_list.insert(0,"Phospholipid")
phospholipid_list_total.insert(0,"Total #")
phospholipid_list_percent_total.insert(0,"% of Total")

phospholipid_list.append("TOTAL")
phospholipid_list_total.append(total)
phospholipid_list_percent_total.append(1)

t_phospholipid_list = [phospholipid_list,phospholipid_list_total,phospholipid_list_percent_total]
np_temporary = np.array(t_phospholipid_list)
t_phospholipid_list = np_temporary.T.tolist()