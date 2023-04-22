import datetime
import pandas as pd
import simple_colors

# create a datetime object for the current date and time
now = datetime.datetime.now() 

# convert the datetime object to a string
date_string = now.strftime("%Y-%m-%d_%H%M%S") 
print (date_string)

iExcel_loc = '/Users/simonnewman/Dropbox/documents/pb/trackdays.xlsx'
oExcel_loc = '/Users/simonnewman/Dropbox/documents/pb/output/trackdays_output_'+ date_string +'.xlsx'

print(simple_colors.red('Starting: Read excel file ' + iExcel_loc))  
# Load the Excel file
excel_file = pd.read_excel(iExcel_loc)

# Select specific columns by name
selected_columns = excel_file[['Lineitem name','Name', 'Billing Name','Email', 'Phone']]

# Output the selected columns
print(selected_columns)

# Read the input file
excel_filtered = pd.read_excel(iExcel_loc, usecols=selected_columns)

# Write the output file
excel_filtered.to_excel(oExcel_loc, index=False)

print(simple_colors.red('Finished: Results written to ' + oExcel_loc  ))  

