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
# Load the Excel file
excel_df = pd.read_excel(iExcel_loc)

# Modify values in the DataFrame
excel_df.loc[excel_df['Lineitem name'].str.contains('Fast'), 'Lineitem name'] = 'Fast'
excel_df.loc[excel_df['Lineitem name'].str.contains('Intermediate'), 'Lineitem name'] = 'Itermidiate'
excel_df.loc[excel_df['Lineitem name'].str.contains('Novice'), 'Lineitem name'] = 'Novice'

# sort by group
excel_df = excel_df.sort_values(by='Lineitem name', ascending=False)

# rename tjhe lineitem to Group
#excel_df = excel_df.rename(columns={'Lineitem name': 'Group'})

# Select specific columns by name
selected_columns = excel_df[['Lineitem name','Name', 'Billing Name','Email', 'Phone']]



# Output the selected columns
print(selected_columns)

# Read the input file
excel_filtered_df = pd.read_excel(iExcel_loc, usecols=selected_columns)

# Modify values in the DataFrame
# excel_filtered_df.loc[excel_filtered_df['Lineitem name'].str.contains('Fast'), 'Lineitem name'] = 'Fast'
# excel_filtered_df.loc[excel_filtered_df['Lineitem name'].str.contains('Intermediate'), 'Lineitem name'] = 'Intermediate'
# excel_filtered_df.loc[excel_filtered_df['Lineitem name'].str.contains('Novice'), 'Lineitem name'] = 'Novice'

##df[['Column A', 'Column B']] = df['Column'].str.split(' ', 1, expand=True)
#df[['Column A', 'Column B']] = df['Column'].str.split(' ', n=1, expand=True)

# fill missing values with empty strings
excel_filtered_df['Lineitem name'].fillna('', inplace=True)
excel_filtered_df[['Date of TD', 'Group']] = excel_filtered_df['Lineitem name'].str.split('-', n=1, expand=True)


# sort by group
#excel_filtered_df = excel_filtered_df.sort_values(by='Lineitem name', ascending=False)

# Rename columns
excel_filtered_df = excel_filtered_df.rename(columns={'Name': 'Ref'})
#excel_filtered_df = excel_filtered_df.rename(columns={'Lineitem name': 'Group'})


# Write the output file
excel_filtered_df.to_excel(oExcel_loc, index=False)

print(sc.red('Finished: Results written to ' + oExcel_loc  ))  

