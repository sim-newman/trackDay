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
excel_df = pd.read_excel(iExcel_loc)

# Modify values in the DataFrame
# df.loc[df['column_name'] == 'value_to_replace', 'column_name'] = 'new_value'

# Modify values in the DataFrame
# excel_df.loc[excel_df['Lineitem name'] == 'sim.newman@me.com', 'Lineitem name'] = 'sim.newman@me.com.replaced'
excel_df.loc[excel_df['Lineitem name'].str.contains('Fast'), 'Lineitem name'] = 'Fast'
excel_df.loc[excel_df['Lineitem name'].str.contains('Intermediate'), 'Lineitem name'] = 'Itermidiate'
excel_df.loc[excel_df['Lineitem name'].str.contains('Novice'), 'Lineitem name'] = 'Novice'


# Select specific columns by name
selected_columns = excel_df[['Lineitem name','Name', 'Billing Name','Email', 'Phone']]

# Output the selected columns
print(selected_columns)

# Read the input file
excel_filtered_df = pd.read_excel(iExcel_loc, usecols=selected_columns)

# Write the output file
excel_filtered_df.to_excel(oExcel_loc, index=False)

print(simple_colors.red('Finished: Results written to ' + oExcel_loc  ))  

