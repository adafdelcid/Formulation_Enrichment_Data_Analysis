'''
Author: Ada Del Cid 
GitHub: @adafdelcid
Oct.2020

CSV2Excel: Converts a CSV file from sequenced samples into an excel spreadsheet. Takes in formulation sheet
with the standard formatting of the Dahlman Lab. Lastly, performs enrichment analysis by formulation
composition and cell type.
'''
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook
import math

def main():
	# create excel destination file
	destination_folder = "/Users/adadelcid/Documents/Python_Enrichment_Project/"
	destination_file = create_excel_spreadsheet(destination_folder)

	# Import formulation sheet and create dataframe
	formulations_sheet = "/Users/adadelcid/Documents/Python_Enrichment_Project/Formulation Sheet Copy.xlsx"
	df_formulations = create_df_formulation_sheet(formulations_sheet, destination_file)

	# Read CSV file and save as dataframe
	csv_filepath = "/Users/adadelcid/Documents/Python_Enrichment_Project/normcounts copy.csv"
	df_norm_counts = create_df_norm_counts(csv_filepath, destination_file)

	# Remove outliers from normalized count dataframe
	df_norm_no_outliers, sample_columns = create_df_norm_no_outliers(df_norm_counts)
	
	# Organize sample_columns by cell type
	sorted_cells = ["VH","VE","VK","VI","SB","ST","SM","SE"]
	organized_columns, count_repeats = organize_by_cell_type(sample_columns,sorted_cells)

	# Merge formulations and normalized counts data frames
	df_merged = merge_formulations_and_norm_counts(df_formulations, df_norm_counts, organized_columns, destination_file)
	
	# Average sample normalized counts by cell type
	df_averaged = average_normalized_counts(df_merged, organized_columns, sorted_cells, destination_file)

	# create enrichment tables
	create_enrichment_tables(destination_file, df_averaged)

	x_percent = 20
	enrichment_cell_type = ["SE"]
	top_bottom_enrichment(destination_file, enrichment_cell_type, df_averaged, x_percent)

def top_bottom_enrichment(destination_file, sorted_cells, df_averaged, x_percent):
	# sort normalized counts by cell type
	dict_df_sorted = sort_norm_counts(sorted_cells, df_averaged, destination_file)
	dict_df_top, dict_df_bottom = top_bottom_percent_by_cell_type(dict_df_sorted, x_percent) 

	for cell_type in sorted_cells:
		top_and_bottom_enrichment_by_cell_type(destination_file, cell_type , dict_df_top, df_averaged, "Top") 
		top_and_bottom_enrichment_by_cell_type(destination_file, cell_type , dict_df_bottom, df_averaged, "Bottom")

def top_and_bottom_enrichment_by_cell_type(destination_file, cell_type, dict_df_top_bottom, df_averaged, top_or_bottom):
	df_top_bottom_cell_type = dict_df_top_bottom[cell_type]
	create_enrichment_tables(destination_file, df_averaged, df_top_bottom_cell_type, cell_type, top_or_bottom)

def top_bottom_percent_by_cell_type(dict_df_sorted, x_percent):

	dict_df_top = {}
	dict_df_bottom = {}

	for cell_type in dict_df_sorted:
		df_sorted = dict_df_sorted[cell_type]

		df_top, df_bottom = top_and_bottom_percent(cell_type, df_sorted,x_percent)
		dict_df_top[cell_type] = df_top
		dict_df_bottom[cell_type] = df_bottom

	return dict_df_top, dict_df_bottom

def top_and_bottom_percent(cell_type, df_sorted, x_percent):
	total_LNP = len(df_sorted.index) - 2 # subtract two because of naked barcodes
	values_x_percent = math.floor(total_LNP*(x_percent/100))

	df_top = df_sorted.loc[range(0, values_x_percent)] # gets top x percent
	df_bottom = df_sorted.loc[range(total_LNP - values_x_percent, total_LNP + 2)] # gets bottom x percent

	if "NAKED1" not in df_bottom["LNP"].to_list() and "NAKED2" not in df_bottom["LNP"].to_list():
		raise NameError("Error: Naked barcodes not on bottom " + str(x_percent) + "% + 2!")

	return df_top, df_bottom

def sort_norm_counts(sorted_cells, df_averaged, destination_file):
	dict_df_sorted = {}

	for cell_type in sorted_cells:
		df_sorted = sort_norm_counts_by_cell_type(cell_type, df_averaged)
		dict_df_sorted[cell_type] = df_sorted

	return dict_df_sorted

def sort_norm_counts_by_cell_type(cell_type, df_averaged):
	
	# get data frame from "Formulation Enrichment" sheet named it: df_averaged
	df_sorted = df_averaged.sort_values(by = cell_type, ascending = False, ignore_index = True)

	return df_sorted

def create_enrichment_tables(destination_file, df_averaged, df_top_bottom_cell_type = None, cell_type = None, top_or_bottom = None):

	dict_df_components = get_all_enrichments(df_averaged, df_top_bottom_cell_type)

	current_row_1 = 1 # variable to place formulation enrichments by mole ratio
	current_row_2 = 1 # variable to place formulation enrichments by component
	enrichment_sheet = "Form Enrichment"
	if cell_type is not None:
		enrichment_sheet += " " + cell_type + " " + top_or_bottom
	with pd.ExcelWriter(destination_file, engine = "openpyxl", mode = "a") as writer:
		if df_top_bottom_cell_type is None:
			df_averaged.to_excel(writer, sheet_name = enrichment_sheet, index = False)
			off_set = len(df_averaged.columns)
		else:
			df_top_bottom_cell_type.to_excel(writer, sheet_name = enrichment_sheet, index = False)
			off_set = len(df_top_bottom_cell_type.columns)
		dict_df_components["Lipomer %"].to_excel(writer, sheet_name = enrichment_sheet, startrow = current_row_1, startcol = off_set + 2, index = False)
		dict_df_components["Lipomer"].to_excel(writer, sheet_name = enrichment_sheet, startrow = current_row_2, startcol = off_set + 6, index = False)
		current_row_1 += len(dict_df_components["Lipomer %"]) + 2
		current_row_2 += len(dict_df_components["Lipomer"]) + 2
		dict_df_components["Cholesterol %"].to_excel(writer, sheet_name = enrichment_sheet, startrow =current_row_1 ,startcol = off_set + 2, index = False)
		dict_df_components["Cholesterol"].to_excel(writer, sheet_name = enrichment_sheet, startrow=current_row_2, startcol = off_set + 6, index = False)
		current_row_1 += len(dict_df_components["Cholesterol %"]) + 2
		current_row_2 += len(dict_df_components["Cholesterol"]) + 2
		dict_df_components["PEG %"].to_excel(writer, sheet_name = enrichment_sheet, startrow =current_row_1 , startcol = off_set + 2, index = False)
		dict_df_components["PEG"].to_excel(writer, sheet_name = enrichment_sheet, startrow=current_row_2, startcol = len(df_averaged.columns)+ 6, index = False)
		current_row_1 += len(dict_df_components["PEG %"]) + 2
		current_row_2 += len(dict_df_components["PEG"]) + 2
		dict_df_components["Phospholipid %"].to_excel(writer, sheet_name = enrichment_sheet, startrow =current_row_1 , startcol = off_set + 2, index = False)
		dict_df_components["Phospholipid"].to_excel(writer, sheet_name = enrichment_sheet, startrow=current_row_2, startcol = off_set + 6, index = False)

	# Add Table labels
	# xfile = openpyxl.load_workbook(destination_file)
	# sheet = xfile[enrichment_sheet]
	# sheet["A1"] = "Formulation Enrichment Mole Ratio"
	# sheet["E1"] = "Formulation Enrichment Component"
	# xfile.save(destination_file) 

def get_all_enrichments(df_averaged, df_top_bottom_cell_type):
	dict_components = get_lists_of_components(df_averaged)
	dict_df_components = {"Lipomer %" : None, "Cholesterol %" : None, "PEG %" : None, "Phospholipid %" : None,
						"Lipomer" : None, "Cholesterol" : None, "PEG" : None, "Phospholipid" : None}

	if df_top_bottom_cell_type is None:
		for component in dict_df_components:
			dict_df_components[component] = calculate_enrichment(component, dict_components[component], df_averaged)
	else:
		for component in dict_df_components:
			dict_df_components[component] = calculate_enrichment(component, dict_components[component], df_top_bottom_cell_type)

	return dict_df_components

def calculate_enrichment(component, component_list, df_averaged):
	component_total = [0]*len(component_list)
	for bc_x in df_averaged[component].values:
		for index in range(len(component_list)):
			if bc_x == component_list[index]:
				component_total[index] += 1
				break

	total = sum(component_total)

	component_percent_total = []
	for each_component in component_total:
		component_percent_total.append(round(each_component/total,9))

	component_list.append("TOTAL")
	component_total.append(total)
	component_percent_total.append(round(sum(component_percent_total)))

	t_component_list = [component_list,	component_total,component_percent_total]
	np_temporary = np.array(t_component_list)
	np_temporary = np_temporary.T
	df_component_list = pd.DataFrame(data = np_temporary, columns = [component,"Total #", "% of Total"])

	return df_component_list	

def get_lists_of_components(df_averaged):
	'''
	average_normalized_counts : works with the "retrieve_component_list" function and returns a dictionary with all component mole ratios and types
		inputs:
				df_averaged : dataframe with averaged normalized counts by cell type
		output:
				dict_components : a dictionary containing list of all the component mole ratios and types
	'''

	dict_components = {"Lipomer %" : [], "Cholesterol %" : [], "PEG %" : [], "Phospholipid %" : [],
						"Lipomer" : [], "Cholesterol" : [], "PEG" : [], "Phospholipid" : []}

	for component in dict_components:
		dict_components[component] = retrieve_component_list(df_averaged, component)

	return dict_components


def retrieve_component_list(df_averaged, component):
	'''
	retrieve_component_list : returns a list of all the different mole ratios or types of a specific component used
		inputs:
				df_averaged : dataframe with averaged normalized counts by cell type
				component : string of the component in question
		output:
				component_list : list of all the different mole ratios or types of a component used
	'''
	component_list = []

	for index in range(len(df_averaged[component].values)-2):
		if df_averaged[component].values[index] not in component_list:
			component_list.append(df_averaged[component].values[index])

	component_list.sort()

	return component_list

def average_normalized_counts(df_merged, organized_columns, sorted_cells, destination_file):
	'''
	average_normalized_counts : creates and returns a dataframe with averaged normalized counts by cell type and appends it to excel spreadsheet
		inputs:
				df_merged : dataframe of merged formulations and normalized counts
				organized_columns : list of samples organized by cell types of sorted cells
				sorted_cells: user specified list of cells that were sorted
				destination_file : directory of the excel spreadsheet created
		output:
				df_averaged : dataframe with averaged normalized counts by cell type
	'''

	df_averaged = df_merged # copy merged dataframe

	for index in range(len(sorted_cells)): 
		temporary_list = []
		for column in organized_columns: # for each sample
			if sorted_cells[index] in column: # if sample is specific cell type
				temporary_list.append(column) # add to temporary list of samples per cell type
		df_averaged[sorted_cells[index]] = df_averaged[temporary_list].mean(axis=1) # get average of repeats of each cell type and append to dataframe

	#order columns
	l1 = df_merged.columns.tolist()[:10] # columns up to phospholipid%
	order_columns = l1 + sorted_cells # formulation columns and cell types

	# rearrange columns on df_averaged
	df_averaged = df_averaged[order_columns]

	# append merged data frames onto excel spreadsheet
	with pd.ExcelWriter(destination_file, engine="openpyxl", mode = "a") as writer:
		df_averaged.to_excel(writer, sheet_name = "Averaged Norm Counts", index = False)

	return df_averaged

def merge_formulations_and_norm_counts(df_formulations, df_norm_counts, organized_columns, destination_file):
	'''
	merge_formulations_and_norm_counts : merges formulation and norm count dataframes into single data frame and appends it to excel spreadsheet
		inputs:
				df_formulations : formulations datasheet
				df_norm_counts : data frame of normalized counts
				organized_columns : list of samples organized by cell types of sorted cells
				destination_file : directory of the excel spreadsheet created
		output:
				df_merged : dataframe of merged formulations and normalized counts
	'''

	# inner merge of data frames around barcodes ("BC")
	df_merged = df_formulations.merge(df_norm_counts, on="BC")

	# ordered columns
	l1 = df_merged.columns.tolist()[:10] # columns up to phospholipid%
	order_columns = l1 + organized_columns # formulation columns and organized sample columns

	# rearrange columns on df_merged
	df_merged = df_merged[order_columns]

	# append merged data frames onto excel spreadsheet on a sheet named Formulations + Norm Counts
	with pd.ExcelWriter(destination_file, engine="openpyxl", mode = "a") as writer:
		df_merged.to_excel(writer, sheet_name = "Formulations + Norm Counts", index = False)

	return df_merged

def organize_by_cell_type(sample_columns, sorted_cells):
	'''
	organize_by_cell_type : gets data fram with normalized counts, creates a dataframe without outliers based on given percentile
		inputs:
				sample_columns :  list of the names of the columns on the dataframe of normalized counts with no outliers(names of samples)
				sorted_cells : user specified list of cells that were sorted
		output:
				organized_columns : list of samples organized by cell types of sorted cells
				count_repeats : number of repeats for each cell type
	'''

	# organize columns
	organized_columns = []
	for cell_type in sorted_cells: 
		for column in sample_columns:
			if cell_type in column: # if current cell_type is in the name of the current column (e.g: if "SB" in "AD SB102":)
				organized_columns.append(column)

	count_repeats = len(sample_columns)//len(sorted_cells) # number of repeats per cell type, should be same for all

	return organized_columns, count_repeats

def create_df_norm_no_outliers(df_norm_counts, percentile = 99.9):
	'''
	create_df_norm_no_outliers: gets data fram with normalized counts, creates a dataframe without outliers based on given percentile
		inputs:
				df_norm_counts :  data frame of normalized counts
				percentile : percentile of values accepted (default = 99.9%)
		output:
				df_norm_no_outliers : data frame with normalized counts without outliers
				sample_columns : list of the names of the columns on the dataframe (names of samples)
	'''

	# get row names and save barcode column
	rows = list(df_norm_counts.index)
	barcodes = df_norm_counts["BC"]

	# calculate given percentile
	df_norm_no_outliers = df_norm_counts.drop("BC",axis = 1) # temporarily remove barcode column
	n_at_percentile = np.percentile(df_norm_no_outliers.to_numpy(),percentile) # find 99.9 percentile value

	# set outliers to NaN (numbers greater than 99.9 percentile)
	sample_columns = df_norm_no_outliers.columns.tolist() # get columns names again, because we changed a column's name
	for row in rows:
		for column in sample_columns:
			if df_norm_no_outliers.at[row,column] >= n_at_percentile: # if value at location [row,column] if greater (outlier)
				df_norm_no_outliers.at[row,column] = np.nan # set value to NaN

	df_norm_no_outliers.insert(loc=0,column = "BC", value = barcodes) # add the BC column again

	return df_norm_no_outliers, sample_columns

def create_df_norm_counts(csv_filepath, destination_file):
	'''
	create_df_norm_counts: gets csv file path with normalized counts, creates a dataframe and appends
	it to destination_file
		inputs:
				csv_filepath : file path to csv file
				destination_file :  name of the destination excel file
		output:
				df_norm_counts : data frame with normalized counts
	'''
	new_sheet_name = "Normalized Counts"

	# Read CSV file and save as data frame
	df_norm_counts = pd.read_csv(csv_filepath,sep = ',',header= 0)

	columns = df_norm_counts.columns.tolist() # get names of columns
	df_norm_counts.rename(columns={columns[0]:"BC"}, inplace=True) #rename first column to BC for barcodes

	# Copy csv data frame to excel sheet
	with pd.ExcelWriter(destination_file,engine="openpyxl",mode = "a") as writer:
		df_norm_counts.to_excel(writer, sheet_name = new_sheet_name, index = False)

	return df_norm_counts

def create_df_formulation_sheet(formulations_sheet, destination_file):
	'''
	create_df_formulation_sheet: gets formulation sheet, creates a dataframe and appends it to destination_file
		inputs:
				formulations_sheet: file path to excel sheet of formulation sheet
				destination_file : name of the destination excel file
		output:
				df_formulations : data frame with formulations sheet
	'''

	# Turn formulation sheet into data frames
	df_formulations = pd.read_excel(formulations_sheet, sheet_name = "Formulations")

	# Make spreadsheet with formulation sheet
	with pd.ExcelWriter(destination_file,engine="openpyxl",mode = "w") as writer:
		df_formulations.to_excel(writer, sheet_name = "Formulations", index = False)

	return df_formulations

def create_excel_spreadsheet(destination_folder, file_name = "Normalized_Counts"):
	'''
	create_excel_spreadsheet: creates an excel spreadsheet
		inputs:
				destination_folder : directory of the folder where the user wants the file stored
				file_name : name of the file being created (default = "Normalized_Counts")
		output:
				destination_file : directory of the excel spreadsheet created
	'''
	
	destination_file = destination_folder + file_name + ".xlsx"
	wb = Workbook()
	wb.save(destination_file)

	return destination_file

if __name__ == "__main__":
	main()