def isLetter(char):
	if(char == ',' or char == ' ' or char == '.'):
		return False
	else:
		return True

def checkTitle(str_input):
	if str_input != "Mr." and str_input != "Ms." and str_input != "Mrs.":
		return False
	else:
		return True

def checkSuffix(str_input):
	str_in = str_input
	str_lwr = str_in.lower()
	index = -1
	counter = 1
	new_str = ""
	found = False

	for x in range(len(str_in)):
		if (str_lwr[x] == 'i' and str_lwr[x + 1] == 'n' and str_lwr[x + 2] == 'c' and (isLetter(str_lwr[x - 1]) == False)):
			index = x
			found = True
			break
		elif (str_lwr[x] == 'l' and str_lwr[x + 1] == 'l' and str_lwr[x + 2] == 'c' and (isLetter(str_lwr[x - 1]) == False)):
			index = x
			found = True
			break
		elif (str_lwr[x] == 'l' and str_lwr[x + 1] == '.' and str_lwr[x + 2] == 'l' and str_lwr[x + 3] == '.'
				and str_lwr[x + 4] == 'c' and (isLetter(str_lwr[x - 1]) == False)):
			index = x
			found = True
			break

	if found == True:	
		while ((index - counter) >= 0) and (isLetter(str_in[index - counter]) == False):
			counter = counter + 1
			
		new_str = str_in[0:(index - (counter - 1))]
	else:
		new_str = str_in

	return new_str

#Openpyxl library is used for excel sheet manipulation
from openpyxl import load_workbook
from openpyxl import Workbook
import datetime
time = datetime.datetime.now()

#Loading current export excel sheet
print(" -----------------------")
print(" Welcome to Filtren v1.2 ")
print(" -----------------------\n")

found_file = False

while found_file == False:
	file_in = input("Please enter file name with extension (.xlsx): ")
	try:
	    source = load_workbook(file_in)
	    found_file = True
	except FileNotFoundError as not_found:
		print('WARNING! \"' + not_found.filename + '\" not found, please try again.\n')

sheet = source.active

#Creating new Workbook for filtered info to be saved
new_wb = Workbook(write_only=True)
wb_sheet = new_wb.create_sheet()

#Printing message to user
filter_choice = -1
valid_choice = False

while valid_choice == False:
	filter_choice = input("Please type 1 to filter for Sales, or 2 to filter for HR/Finance: ")
	if filter_choice == '1' or filter_choice == '2':
		valid_choice = True
	else:
		print("Error! Invalid option, please try again.\n")

if filter_choice == '1':
	print("[Sales] Filtering process started ...")
else:
	print("[HR/Finance] Filtering process started ...")

#Initializing variables to keep track of rows and criteria
currentRow = 1
filteredSheetRow = 1
is_empty = False
department_warning = False
status_warning = False
loop_stopped = False

#Putting header at top of file
header = ("Company Name", "Title", "First Name", "Last Name", "Email")
wb_sheet.append(header)

found_header = False

for eachRow in sheet.iter_rows():
	#Restting variables
	is_empty = False

	if found_header == False:
		for eachCol in sheet.iter_cols():
			print(eachCol)
		found_header == True

	#Parsing relevant information from export
	company_name = sheet.cell(row=currentRow, column=1).value
	email = sheet.cell(row=currentRow, column=24).value
	title = sheet.cell(row=currentRow, column=17).value
	first_name = sheet.cell(row=currentRow, column=19).value
	last_name = sheet.cell(row=currentRow, column=20).value
	department = sheet.cell(row=currentRow, column=26).value
	status = sheet.cell(row=currentRow, column=16).value

	#Only for testing purposes
	#print("[" + str(company_name) + ", " + str(email) + ", " + str(status) + ", Empty Status: " + str(is_empty) + "]")

	#Checking if fields are blank, if so then flag will be set
	if company_name is None:
		is_empty = True
	elif email is None:
		is_empty = True
	elif title is None:
		is_empty = True
	elif first_name == None:
		is_empty = True
	elif last_name == None:
		is_empty = True
	elif department == None:
		is_empty = True
	elif status == None:
		is_empty = True

	#Generating department warning
	# if (department != "Sales" and department_warning == False):
	# 	#print('\x1b[1;37;41m' + 'WARNING! This export contains other departments besides Sales, do you wish to continue?' + '\x1b[0m')
	# 	input_value = input("Type \'y\' for YES or \'n\' for NO: ")
	# 	department_warning = True
	# 	#Handling department warning response
	# 	if input_value == 'n':
	# 		print('\x1b[0;30;43m' + 'Filtering process aborted.' + '\x1b[0m')
	# 		loop_stopped = True
	# 		break

	#Generating status warning
	if (is_empty != True and status != "Mailing list" and status != "Status"):
		print('Multiple entries without type \'Mailing List\' detected and will be filtered out, do you wish to continue?')
		status_warning_response = input("Type \'y\' for YES or \'n\' for NO: ")
		status_warning = True
		#Handling status warning response
		if status_warning_response == 'n':
			print('Filtering process aborted.')
			loop_stopped = True
			break

	#Checking details in variables
	if is_empty == False:
		if checkTitle(title) == True:
			#Creating tuple to be written to file if valid
			temp_tup = (checkSuffix(company_name), title, first_name, last_name, email)

			#Making sure email address is valid
			if '@' in email:
				if title.lower() != "mr." or title.lower() != "mrs." or title.lower() != "ms." or title.lower() != "dr.":
					if filter_choice == '1':
						if department.lower() == "sales":
							wb_sheet.append(temp_tup)
					else:
						if department.lower() == "finance" or department.lower() == "hr":
							wb_sheet.append(temp_tup)

	#Increading row
	currentRow += 1

if loop_stopped == False:
	#Creating the filename

	# department_file = ""

	# if filter_choice == '1':
	# 	department_file = "Sales"
	# else:
	# 	department_file = "Finance/HR"

	#filtered_result_file_name = department_file + " Filtered Result " + str(time.day) + "-" + str(time.month) + "-" + str(time.year) + ".xlsx"
	filtered_result_file_name =  "Filtered Result " + str(time.day) + "-" + str(time.month) + "-" + str(time.year) + ".xlsx"

	#Saving the new Workbook
	new_wb.save(filtered_result_file_name)

	#Printing message to user
	print("Filtering complete.")
	print("\n\"" + filtered_result_file_name + "\" saved successfully.\n")

	while True:
		input("Press \"enter\" to exit.")
		break

else:
	err_tup = ('FILTERING WAS INTERRUPTED, DO NOT USE THIS FILE!', 'FILTERING WAS INTERRUPTED, DO NOT USE THIS FILE', 'FILTERING WAS INTERRUPTED, DO NOT USE THIS FILE!', 'FILTERING WAS INTERRUPTED, DO NOT USE THIS FILE!', 'FILTERING WAS INTERRUPTED, DO NOT USE THIS FILE!')
	wb_sheet.append(err_tup)
	new_wb.save("FILTERING ERROR, DO NOT USE.xlsx")


