import datetime as dt
import pandas as pd
import simple_colors as sc

# create a datetime object for the current date and time
now = dt.datetime.now() 

# convert the datetime object to a string
date_string = now.strftime("%Y-%m-%d_%H%M%S") 
print (date_string)

iExcel_loc = '/Users/simonnewman/Dropbox/documents/pb/trackdays.xlsx'
oExcel_loc = '/Users/simonnewman/Dropbox/documents/pb/output/trackdays_output_'+ date_string +'.xlsx'

print(sc.red('Starting: Read excel file ' + iExcel_loc))  

# pick columns to be in report.
selected_columns = ['Lineitem name','Name', 'Billing Name','Email']


# Read the input file into a dataframe
excel_filtered_df = pd.read_excel(iExcel_loc, usecols=selected_columns)

# fill missing values with empty strings
excel_filtered_df['Lineitem name'].fillna('', inplace=True)
excel_filtered_df[['Date of TD', 'Group']] = excel_filtered_df['Lineitem name'].str.split('-', n=1, expand=True)


# sort by group
excel_filtered_df = excel_filtered_df.sort_values(by=['Date of TD', 'Group'], ascending=True)

# Rename columns
excel_filtered_df = excel_filtered_df.rename(columns={'Name': 'Ref'})
excel_filtered_df = excel_filtered_df.rename(columns={'Billing Name': 'Name'})

# Define the new order of columns
#new_order = ['Date of TD', 'Group', 'Ref', 'Name', 'Email', 'Lineitem name']
new_order = ['Date of TD', 'Group', 'Ref', 'Name', 'Email']


# Rearrange the columns in the DataFrame
excel_filtered_df = excel_filtered_df[new_order]

#print out the report
print (excel_filtered_df)

# Write the output file
excel_filtered_df.to_excel(oExcel_loc, index=False)

print(sc.red('Finished: Results written to ' + oExcel_loc  ))  

