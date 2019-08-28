# This file finds errors between two different excel sheets and enters the differences 
# into two different spreadsheets.

import csv
import numpy
from enum import Enum
from sets import Set
from xlwt import Workbook, Formula
import xlwt

wb = Workbook()
style = xlwt.XFStyle()
font = xlwt.Font()
font.bold = True
style.font = font

class Columns(Enum):
	n_hs_name = 0
	n_hs_date = 1
	s_hs_name = 2
	s_hs_date = 3
	n_c_name = 4
	n_c_date = 5
	s_c_name = 6
	s_c_date = 7
	n_cg_name = 8
	n_cg_date = 9
	s_cg_name = 10
	s_cg_date = 11

class Dreamer_Class(Enum):
	Boulder = 0
	Broadway = 1
	Casa = 2
	Columbine = 3
	Kalmia = 4
	Lafayette = 5
	Pioneer = 6
	Sanchez = 7

class National(Enum):
	first = 3
	last = 5
	hs_graduated = 12
	psenr_record = 16
	has_ps_cred = 20

class Salesforce(Enum):
	first = 0
	last = 1
	actual_hsgrad = 3
	actual_start = 12
	actual_colgrad = 5

# reprompts user if invalid file.
def ask():
	while True:
		try:
			class_info = raw_input("Please enter the file name: ").lower()   
			open(class_info) 
		except IOError:
			print('\n' * 5 + "Invalid file name; try again.")
			continue
		else:
			break
	return class_info

# creates all the sheets in the workbook
boulder_class = wb.add_sheet("Boulder Class")
broadway_class = wb.add_sheet("Broadway Class")
casa_class = wb.add_sheet("Casa Class")
columbine_class = wb.add_sheet("Columbine Class")
kalmia_class = wb.add_sheet("Kalmia Class")
lafayette_class = wb.add_sheet("Lafayette Class")
pioneer_class = wb.add_sheet("Pioneer Class")
sanchez_class = wb.add_sheet("Sanchez Class")

classes = [boulder_class, broadway_class, casa_class, columbine_class,
kalmia_class, lafayette_class, pioneer_class, sanchez_class]


for num in range(len(classes)):

	national = []
	salesforce = []
	n_graduated_hs = {}
	n_started_college = {}
	n_graduated_college = {}

	s_graduated_hs = {}
	s_started_college = {}
	s_graduated_college = {}

	print '\n' * 3 + "Enter the NATIONAL info for """ + str(Dreamer_Class(num)) + '\n' + "(For example: boulder_class.csv)"
	with open(ask()) as csvfile:
		readCSV = csv.reader(csvfile, delimiter = ',')
		for row in readCSV:
			if row[0] != "":
				new_row = []
				new_row.append(row[National.first.value] + " " + row[National.last.value])
				new_row.append(row[National.hs_graduated.value])
				new_row.append(row[National.psenr_record.value])
				new_row.append(row[National.has_ps_cred.value])
				if new_row[1].lower() == 'yes':
					n_graduated_hs[new_row[0]] = new_row[1]
				if new_row[2].lower() == 'yes':
					n_started_college[new_row[0]] = new_row[2]
				if new_row[3].lower() == 'yes':
					n_graduated_college[new_row[0]] = new_row[3]

	print '\n' + "Enter the SALESFORCE info for """ + str(Dreamer_Class(num)) + '\n' + "(For example: boulder_report.csv)"
	with open(ask()) as csvfile:
		readCSV = csv.reader(csvfile, delimiter = ',')
		for row in readCSV:
			 if row[0] != "":
			 	new_row = []
				new_row.append(row[Salesforce.first.value] + " " + row[Salesforce.last.value])
				new_row.append(row[Salesforce.actual_hsgrad.value])
				new_row.append(row[Salesforce.actual_start.value])
				new_row.append(row[Salesforce.actual_colgrad.value])
				if new_row[1] != '-':
					s_graduated_hs[new_row[0]] = new_row[1]
				if new_row[2] != '-':
					s_started_college[new_row[0]] = new_row[2]
				if new_row[3] != '-':
					s_graduated_college[new_row[0]] = new_row[3]

	classes[num].col(Columns.n_hs_name.value).width = 7000
	classes[num].col(Columns.s_hs_name.value).width = 7000
	classes[num].col(Columns.n_c_name.value).width = 7000
	classes[num].col(Columns.s_c_name.value).width = 7000
	classes[num].col(Columns.n_cg_name.value).width = 7000
	classes[num].col(Columns.s_cg_name.value).width = 7000

	classes[num].write(0, Columns.n_hs_name.value, "Not in Salesforce (HS Grad):", style = style)
	classes[num].write(0, Columns.s_hs_name.value, "Not in National (HS Grad):", style = style)
	classes[num].write(0, Columns.n_c_name.value, "Not in Salesforce (Started College):", style = style)
	classes[num].write(0, Columns.s_c_name.value, "Not in National (Started College):", style = style)
	classes[num].write(0, Columns.n_cg_name.value, "Not in Salesforce (College Grad):", style = style)
	classes[num].write(0, Columns.s_cg_name.value, "Not in National (College Grad):", style = style)

	row = 1
	for key in n_graduated_hs:
		if not key in s_graduated_hs:
			classes[num].write(row, Columns.n_hs_name.value, key)
			classes[num].write(row, Columns.n_hs_date.value, n_graduated_hs[key])
			row += 1
			#add name and value to things salesforce doesnt have
	classes[num].write(row + 1, Columns.n_hs_name.value, "Errors: " + str(row - 1), style = style)
	row = 1
	for key in s_graduated_hs:
		if not key in n_graduated_hs:
			classes[num].write(row, Columns.s_hs_name.value, key)
			classes[num].write(row, Columns.s_hs_date.value, s_graduated_hs[key])
			row += 1
			#add name and value to things national doesnt have
	classes[num].write(row + 1, Columns.s_hs_name.value, "Errors: " + str(row - 1), style = style)
	row = 1
	for key in n_started_college:	
		if not key in s_started_college:
			classes[num].write(row, Columns.n_c_name.value, key)
			classes[num].write(row, Columns.n_c_date.value, n_started_college[key])
			row += 1
	classes[num].write(row + 1, Columns.n_c_name.value, "Errors: " + str(row - 1), style = style)
	row = 1
	for key in s_started_college:
		if not key in n_started_college:
			classes[num].write(row, Columns.s_c_name.value, key)
			classes[num].write(row, Columns.s_c_date.value, s_started_college[key])
			row += 1
	classes[num].write(row + 1, Columns.s_c_name.value, "Errors: " + str(row - 1), style = style)	
	row = 1
	for key in n_graduated_college:
		if not key in s_graduated_college:
			classes[num].write(row, Columns.n_cg_name.value, key)
			classes[num].write(row, Columns.n_cg_date.value, n_graduated_college[key])
			row += 1
	classes[num].write(row + 1, Columns.n_cg_name.value, "Errors: " + str(row - 1), style = style)
	row = 1
	for key in s_graduated_college:
		if not key in n_graduated_college:
			classes[num].write(row, Columns.s_cg_name.value, key)
			classes[num].write(row, Columns.s_cg_date.value, s_graduated_college[key])
			row += 1
	classes[num].write(row + 1, Columns.s_cg_name.value, "Errors: " + str(row - 1), style = style)
	wb.save('Salesforce_National_Errors.xls')






