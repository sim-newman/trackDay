import datetime as dt
import pandas as pd
import simple_colors as sc


# create a datetime object for the current date and time
now = dt.datetime.now() 

# convert the datetime object to a string
date_string = now.strftime("%Y-%m-%d_%H%M%S") 

# file locations for input/output excels
iExcel_loc = '/Users/simonnewman/Dropbox/documents/pb/bookings2.xlsx'
oExcel_loc = '/Users/simonnewman/Dropbox/documents/pb/output/trackdays_output_'+ date_string +'.xlsx'

print(sc.red('Starting: Read excel file ' + iExcel_loc))  

# pick columns to be in report.
sColumns = ['Lineitem name','Created at','Name', 'Billing Name','Email','Notes']


# Read the input file into a dataframe
filtered_df = pd.read_excel(iExcel_loc, usecols=sColumns)

# Split lineitem into Date and Group
filtered_df['Lineitem name'].fillna('', inplace=True)
filtered_df[['Day', 'Group']] = filtered_df['Lineitem name'].str.split('-', n=1, expand=True)

# get year portion from the created at.
filtered_df['Created at'].fillna('', inplace=True)
filtered_df[['Year', 'month']] = filtered_df['Created at'].str.split('-', n=1, expand=True)

# sort by group
sorted_df = filtered_df.sort_values(by=['Year','Day', 'Group', 'Name'], ascending=True)

# Rename columns
sorted_df = sorted_df.rename(columns={'Name': 'Ref'})
sorted_df = sorted_df.rename(columns={'Billing Name': 'Name'})


# Define the new order of columns
new_order = ['Year','Day','Group', 'Ref', 'Name', 'Email', 'Notes']


# Rearrange the columns in the DataFrame
sorted_df = sorted_df[new_order]

# store non null names and emails into a hashmp.  
hashmap = {}
for index, row in sorted_df.iterrows():
    if str(row['Name']) != 'nan':
        hashmap[row['Email']] = row['Name']
    
# The booking system does not store names against all purchases when a person makes muliple purchases on the same order.
# The following looks for entries with no names, and then populates that with the name from the orther orders.
for index, row in sorted_df.iterrows():
    name_str = str(row['Name'])
    email  = tuple([row['Email']])
    
    if name_str == 'nan':
        try:
          print(sc.blue('We do not have a name for ' + str(row['Email']) + ' we need to auto populate with ' + str(hashmap[email[0]])))
          row['Name'] = str(hashmap[email[0]])
        except KeyError:
            print('Something went wrong! for email address lookup')

# Write the output file
sorted_df.to_excel(oExcel_loc, index=False) 

print(sc.red('Finished: Results written to ' + oExcel_loc  ))  

