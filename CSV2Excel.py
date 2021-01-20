'''
Author: Ada Del Cid 
GitHub: @adafdelcid
Oct.2020

CSV2Excel: Converts a CSV file from sequenced samples into an excel spreadsheet. Takes in formulation sheet
with the standard formatting of the Dahlman Lab. Lastly, performs enrichment analysis by formulation
composition and specified cell type.
'''
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook
import math

'''def main():

	# User inputs (only provide these values DO NOT change anything else on the code to prevent code from breaking)
	# Note the formulations sheet must be on a sheet with name "Formulations", otherwise error will occur
	destination_folder = "/Users/adadelcid/Documents/Python_Enrichment_Project/"
	formulations_sheet = "/Users/adadelcid/Documents/Python_Enrichment_Project/Formulation Sheet Copy.xlsx"
	csv_filepath = "/Users/adadelcid/Documents/Python_Enrichment_Project/normcounts copy.csv"
	sorted_cells = ["VH","VE","VK","VI","SB","ST","SM","SE"]
	percentile = 99.9 #optional input, for outlier removal based on percentile. If not specified, then 99.9 is used by default
	x_percent = 20
	cell_type = "SE"

	# Do not edit anything past this line
	run_enrichment_analysis(destination_folder, formulations_sheet, csv_filepath, sorted_cells, x_percent, cell_type)'''

def run_enrichment_analysis(destination_folder, formulations_sheet, csv_filepath, sorted_cells, x_percent, cell_type, percentile = 99.9):
	'''
	run_enrichment_analysis : driver function, it uses all other functions to create enrichment analysis
		inputs:
				destination_folder : user specified path to the folder where the user wants the excel file created to be saved
				formulations_sheet : user specified file path to excel spreadsheet with formulation sheet only
				csv_filepath : user specified file path to csv with normalized counts
				sorted_cells : user specified list of cells that were sorted
				x_percent : user specified integer to find top and bottom performing LNPs (0-100)
				cell_type : user specified cell type to sort by
	'''
	# check if no input (reset to 99.9)
	if percentile == 0.0:
		percentile = 99.9

	# create excel destination file
	destination_file = create_excel_spreadsheet(destination_folder)

	# Import formulation sheet and create dataframe
	df_formulations = create_df_formulation_sheet(formulations_sheet, destination_file)

	# Read CSV file and save as dataframe
	df_norm_counts = create_df_norm_counts(csv_filepath, destination_file)

	# Remove outliers from normalized count dataframe
	df_norm_no_outliers, sample_columns = create_df_norm_no_outliers(df_norm_counts, percentile)
	
	# Organize sample_columns by cell type
	organized_columns, count_repeats = organize_by_cell_type(sample_columns,sorted_cells)

	# Merge formulations and normalized counts data frames
	merge_formulations_and_norm_counts(df_formulations, df_norm_counts, organized_columns, destination_file, True) #with outliers
	df_merged = merge_formulations_and_norm_counts(df_formulations, df_norm_no_outliers, organized_columns, destination_file) #without outliers
	
	# Average sample normalized counts by cell type
	df_averaged = average_normalized_counts(df_merged, organized_columns, sorted_cells, destination_file)

	# create enrichment tables
	dict_df_components_averaged = create_enrichment_tables(destination_file, df_averaged)

	# create top and bottom x_percent enrichment tables by specified cell type
	dict_df_components_top, dict_df_components_bottom = top_bottom_enrichment(destination_file, cell_type, df_averaged, x_percent)

	# create net enrichment factor sheet
	create_net_enrichment_factor(destination_file, dict_df_components_averaged, dict_df_components_top, dict_df_components_bottom, cell_type)

	#create sheet with top/winning LNPs
	winning_LNPs(cell_type, df_averaged, x_percent, destination_file)

	
def winning_LNPs(cell_type, df_averaged, x_percent, destination_file):
	'''
	winning_LNPs: creates excel sheet with formulations and normalized counts of top performing LNPs named " Winning LNPS" + cell_type
		inputs:
				cell_type : user specified cell type to sort by
				df_averaged : dataframe with averaged normalized counts by cell type
				x_percent : user specified integer to find top and bottom performing LNPs (0-100)
				destination_file : directory of the excel spreadsheet
	'''

	df_sorted = sort_norm_counts(cell_type, df_averaged)
	df_top, df_bottom = top_and_bottom_percent(cell_type, df_sorted, x_percent)

	winning_LNP_sheet = "Winning LNPs " + cell_type
	with pd.ExcelWriter(destination_file, engine = "openpyxl", mode = "a") as writer:
		df_top.to_excel(writer, sheet_name = winning_LNP_sheet, index = False)

def create_net_enrichment_factor(destination_file, dict_df_components_averaged, dict_df_components_top, dict_df_components_bottom, cell_type):
	'''
	create_net_enrichment_factor: creates excel sheet with all enrichment analysis (averaged, top, bottom, raw enrichment and net enrichment factor) named "Net Enrichment Factors"
		inputs:
				destination_file : directory of the excel spreadsheet
				dict_df_components_averaged : dictionary with all dataframes of all enrichment calculations of df_averaged
				dict_df_components_top : dictionary with all dataframes of all enrichment calculations of df_top
				dict_df_components_bottom : dictionary with all dataframes of all enrichment calculations of df_bottom
				cell_type : user specified cell type to sort by
	'''

	dict_df_component_net_enrichment_factor, dict_raw_enrichment_top, dict_raw_enrichment_bottom = net_enrichment_factor(dict_df_components_averaged, dict_df_components_top, dict_df_components_bottom, cell_type)
	dict_df_raw_enrichment_top = dict_list_to_dict_df(dict_raw_enrichment_top, cell_type)
	dict_df_raw_enrichment_bottom = dict_list_to_dict_df(dict_raw_enrichment_bottom, cell_type)
	current_row = 1 # variable to place formulation enrichments by mole ratio
	net_enrichment_sheet = "Net Enrichment Factors " + cell_type 
	
	list_enrichments = ["Lipomer %", "Cholesterol %", "PEG %", "Phospholipid %", "Lipomer", "Cholesterol", "PEG", "Phospholipid"]

	with pd.ExcelWriter(destination_file, engine = "openpyxl", mode = "a") as writer:

		for item in list_enrichments:
			dict_df_components_averaged[item].to_excel(writer, sheet_name = net_enrichment_sheet, startrow = current_row, startcol = 0, index = False)
			dict_df_components_top[item].to_excel(writer, sheet_name = net_enrichment_sheet, startrow = current_row, startcol = 4, index = False)
			dict_df_raw_enrichment_top[item].to_excel(writer, sheet_name = net_enrichment_sheet, startrow = current_row, startcol = 8, index = False)
			dict_df_components_bottom[item].to_excel(writer, sheet_name = net_enrichment_sheet, startrow = current_row, startcol = 11, index = False)
			dict_df_raw_enrichment_bottom[item].to_excel(writer, sheet_name = net_enrichment_sheet, startrow = current_row, startcol = 15, index = False)
			dict_df_component_net_enrichment_factor[item].to_excel(writer, sheet_name = net_enrichment_sheet, startrow = current_row, startcol = 18, index = False)
			current_row += len(dict_df_components_averaged[item]) + 2

	xfile = openpyxl.load_workbook(destination_file)
	sheet = xfile[net_enrichment_sheet]
	sheet["A1"] = "Formulation Enrichment"
	sheet["E1"] = "Top"
	sheet["I1"] = "Enrichment Factor Top"
	sheet["L1"] = "Bottom"
	sheet["P1"] = "Enrichment Factor Bottom"
	sheet["S1"] = "Net Enrichment Factor"
	xfile.save(destination_file)

def dict_list_to_dict_df(dict_list,cell_type):
	'''
	dict_list_to_dict_df: converts dictionary with lists to dictionary with dataframes
		inputs:
				dict_list: dictionary containing lists
				cell_type : user specified cell type to sort by
		output:
				dict_df : dictionary with dataframes
	'''

	dict_df ={}
	for component in dict_list:
		np_temporary = np.array(dict_list[component])
		dict_df[component] = pd.DataFrame(data = np_temporary, columns = [component, cell_type])

	return dict_df

def net_enrichment_factor(dict_df_components_averaged, dict_df_components_top, dict_df_components_bottom, cell_type):
	'''
	net_enrichment_factor: creates dataframes for best and worst performing LNPs, counts and their formulations
		inputs:
				dict_df_components_averaged : dictionary with all dataframes of all enrichment calculations of df_averaged
				dict_df_components_top : dictionary with all dataframes of all enrichment calculations of df_top
				dict_df_components_bottom : dictionary with all dataframes of all enrichment calculations of df_bottom
				cell_type : user specified cell type to sort by
		output:
				dict_df_component_net_enrichment_factor : dictionary with dataframes of net enrichment factors by component type or mole ratio
				dict_raw_enrichment_factors_top : dictionary with dataframes of raw enrichment of top performing LNPs
				dict_raw_enrichment_factors_bottom : dictionary with dataframes of raw enrichment of bottom performing LNPs
	'''

	dict_component_net_enrichment_factor ={}

	dict_raw_enrichment_factors_top = raw_enrichment_factor(dict_df_components_averaged, dict_df_components_top)
	dict_raw_enrichment_factors_bottom = raw_enrichment_factor(dict_df_components_averaged, dict_df_components_bottom)

	for component in dict_raw_enrichment_factors_top:
		temporary_list = []
		for index in range(len(dict_raw_enrichment_factors_top[component])):
			enrichment_factor_row_top = dict_raw_enrichment_factors_top[component][index]
			enrichment_factor_row_bottom = dict_raw_enrichment_factors_bottom[component][index]
			
			item = [enrichment_factor_row_top[0], round(enrichment_factor_row_top[1] - enrichment_factor_row_bottom[1],9)]
			temporary_list.append(item)

		dict_component_net_enrichment_factor[component] = temporary_list
	
	dict_df_component_net_enrichment_factor = {"Lipomer %" : None, "Cholesterol %" : None, "PEG %" : None, "Phospholipid %" : None,
						"Lipomer" : None, "Cholesterol" : None, "PEG" : None, "Phospholipid" : None}

	for component in dict_component_net_enrichment_factor:
		np_temporary = np.array(dict_component_net_enrichment_factor[component])
		dict_df_component_net_enrichment_factor[component] = pd.DataFrame(data = np_temporary, columns = [component, cell_type])

	return dict_df_component_net_enrichment_factor, dict_raw_enrichment_factors_top, dict_raw_enrichment_factors_bottom

def raw_enrichment_factor(dict_df_components_averaged, dict_df_components_top_bottom):
	'''
	raw_enrichment_factor: creates dataframes for best and worst performing LNPs, counts and their formulations
		inputs:
				dict_df_components_averaged : dictionary with all dataframes of all enrichment calculations of df_averaged
				dict_df_components_top_bottom : dictionary with all dataframes of all enrichment calculations of df_top_bottom_cell_type
		output:
				dict_raw_enrichment_factors : dictionary with lists of all raw enrichment factors
	'''

	dict_components_averaged = {}
	dict_components_top_bottom = {}
	dict_raw_enrichment_factors = {}

	for component in dict_df_components_averaged:
		dict_components_averaged[component] = dict_df_components_averaged[component].values.tolist()
		dict_components_top_bottom[component] = dict_df_components_top_bottom[component].values.tolist()

		temporary_list = []
		for index in range(len(dict_components_averaged[component])):
			averaged_row = dict_components_averaged[component][index]
			top_bottom_row = dict_components_top_bottom[component][index]
			item = [averaged_row[0], round(float(top_bottom_row[2])/float(averaged_row[2]),9)]
			temporary_list.append(item)

		dict_raw_enrichment_factors[component] = temporary_list

	return dict_raw_enrichment_factors

def top_bottom_enrichment(destination_file, cell_type, df_averaged, x_percent):
	'''
	top_bottom_enrichment: creates dataframes for best and worst performing LNPs, counts and their formulations
		inputs:
				destination_file : directory of the excel spreadsheet
				cell_type : user specified cell type to sort by
				df_averaged : dataframe with averaged normalized counts by cell type
				x_percent : user specified integer to find top and bottom performing LNPs (0-100)
		output:
				dict_df_components_top : dictionary containing dataframes with enrichment analysis of top performing LNPs
				dict_df_components_bottom : dictionary containing dataframes with enrichment analysis of bottom performing LNPs
	'''

	# sort normalized counts by cell type
	df_sorted = sort_norm_counts(cell_type, df_averaged)
	df_top, df_bottom = top_and_bottom_percent(cell_type, df_sorted, x_percent) 

	dict_df_components_top = create_enrichment_tables(destination_file, df_averaged,  df_top, cell_type, "Top") 
	dict_df_components_bottom = create_enrichment_tables(destination_file, df_averaged, df_bottom, cell_type, "Bottom")

	return dict_df_components_top, dict_df_components_bottom

def top_and_bottom_percent(cell_type, df_sorted, x_percent):
	'''
	top_and_bottom_percent: creates dataframes for best and worst performing LNPs, counts and their formulations
		inputs:
				cell_type : user specified cell type to sort by
				df_sorted : dataframe with normalized counts sorted in descending order by specified cell type
				x_percent : user specified integer to find top and bottom performing LNPs (0-100)
		output:
				df_top : dataframe top performing LNPs 
				df_bottom : dataframe bottom performing LNPs
	'''

	total_LNP = len(df_sorted.index) - 2 # subtract two because of naked barcodes
	values_x_percent = math.floor(total_LNP*(x_percent/100))

	df_top = df_sorted.loc[range(0, values_x_percent)] # gets top x percent
	df_bottom = df_sorted.loc[range(total_LNP - values_x_percent, total_LNP + 2)] # gets bottom x percent

	if "NAKED1" not in df_bottom["LNP"].to_list() and "NAKED2" not in df_bottom["LNP"].to_list():
		raise NameError("Error: Naked barcodes not on bottom " + str(x_percent) + "% + 2!")

	return df_top, df_bottom

def sort_norm_counts(cell_type, df_averaged):
	'''
	sort_norm_counts: creates dataframe with normalized counts sorted in descending order by specified cell type
		inputs:
				cell_type : user specified cell type to sort by
				df_averaged : dataframe with averaged normalized counts by cell type
		output:
				df_sorted : dataframe with normalized counts sorted in descending order by specified cell type
	'''

	df_sorted = df_averaged.sort_values(by = cell_type, ascending = False, ignore_index = True)

	return df_sorted

def create_enrichment_tables(destination_file, df_averaged, df_top_bottom_cell_type = None, cell_type = None, top_or_bottom = None):
	'''
	create_enrichment_tables: creates excel sheet with formulation enrichment tables of averaged normalized counts (top or bottom performing LNPs if df_top_bottom_cell_type value is passed)
							named "Form Enrichment" (or "Form Enrichment" + cell_type + top_or_bottom if df_top_bottom_cell_type provided)
		inputs:
				destination_file : directory of the excel spreadsheet
				df_averaged : dataframe with averaged normalized counts by cell type
				df_top_bottom_cell_type : dataframe of either top or bottom performing LNPs by specified cell type (default = None)
				cell_type : user specified cell type to sort by (default = None)
				top_or_bottom : specifies if enrichment is for top or bottom performing LNPs by specified cell type(default = None)
		output:
				dict_df_components : dictionary with all data frames of all enrichment calculations of df_averaged (or df_top_bottom_cell_type if inputted)
	'''

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
			df_sorted_by_cell_type = sort_norm_counts(cell_type, df_averaged)
			df_sorted_by_cell_type.to_excel(writer, sheet_name = enrichment_sheet, index = False)
			off_set = len(df_sorted_by_cell_type.columns)

		list_enrichments = ["Lipomer %", "Cholesterol %", "PEG %", "Phospholipid %", "Lipomer", "Cholesterol", "PEG", "Phospholipid"]

		for index in range(len(list_enrichments)//2):
			dict_df_components[list_enrichments[index]].to_excel(writer, sheet_name = enrichment_sheet, startrow = current_row_1, startcol = off_set + 2, index = False)
			dict_df_components[list_enrichments[index + 4]].to_excel(writer, sheet_name = enrichment_sheet, startrow = current_row_2, startcol = off_set + 6, index = False)
			current_row_1 += len(dict_df_components[list_enrichments[index]]) + 2
			current_row_2 += len(dict_df_components[list_enrichments[index + 4]]) + 2

	return dict_df_components

def get_all_enrichments(df_averaged, df_top_bottom_cell_type):
	'''
	get_all_enrichments: calculated enrichment by component or component_ratio
		inputs:
				df_averaged : dataframe with averaged normalized counts by cell type
				df_top_bottom_cell_type : dataframe of either top or bottom performing LNPs by specified cell_type (optional input)
		output:
				dict_df_components : dictionary with all dataframes of all enrichment calculations of df_averaged (or df_top_bottom_cell_type if inputted)
	'''

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
	'''
	calculate_enrichment: calculated enrichment by component or component_ratio
		inputs:
				component : component or component ratio
				component_list : list of all component types and component ratios specified component
				df_averaged : dataframe with averaged normalized counts by cell type
		output:
				df_component_list : dataframe of enrichment for specified component
	'''

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
	get_lists_of_components : works with the "retrieve_component_list" function and returns a dictionary with all component mole ratios and component types
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
	average_normalized_counts : creates and returns a dataframe with averaged normalized counts by cell type and appends it to excel spreadsheet on a sheet named "Averaged Norm Counts"
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

def merge_formulations_and_norm_counts(df_formulations, df_norm_counts, organized_columns, destination_file, add_to_excel = False):
	'''
	merge_formulations_and_norm_counts : merges formulation and norm count dataframes into single data frame and appends it to excel spreadsheet named "Formulations + Norm Counts"
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

	if add_to_excel:
		# append merged data frames onto excel spreadsheet on a sheet named Formulations + Norm Counts, with outliers
		with pd.ExcelWriter(destination_file, engine="openpyxl", mode = "a") as writer:
			df_merged.to_excel(writer, sheet_name = "Formulations + Norm Counts", index = False)
	else: #save dataframe without outliers
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
	create_df_norm_counts: gets csv file path with normalized counts, creates a dataframe and appends it to destination_file on a sheet named "Normalized Counts"
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
	create_df_formulation_sheet : gets formulation sheet, creates a dataframe and appends it to destination_file on a sheet named " Formulations"
		inputs:
				formulations_sheet : file path to excel sheet of formulation sheet
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

def create_excel_spreadsheet(destination_folder, file_name = "Enrichment Analysis"):
	'''
	create_excel_spreadsheet: creates an excel spreadsheet
		inputs:
				destination_folder : directory of the folder where the user wants the file stored
				file_name : name of the file being created (default = "Enrichment Analysis")
		output:
				destination_file : directory of the excel spreadsheet created
	'''
	
	destination_file = destination_folder + file_name + ".xlsx"
	wb = Workbook()
	wb.save(destination_file)

	return destination_file

'''if __name__ == "__main__":
	main()'''
