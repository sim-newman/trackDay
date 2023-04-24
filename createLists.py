import datetime as dt
import pandas as pd
import simple_colors as sc

# create a datetime object for the current date and time
now = dt.datetime.now() 

# convert the datetime object to a string
date_string = now.strftime("%Y-%m-%d_%H%M%S") 
print (date_string)

#iExcel_loc = '/Users/simonnewman/Dropbox/documents/pb/trackdaysfull.xlsx'
#iExcel_loc = '/Users/simonnewman/Dropbox/documents/pb/orders_export_1.xlsx'
iExcel_loc = '/Users/simonnewman/Dropbox/documents/pb/bookings2.xlsx'

oExcel_loc = '/Users/simonnewman/Dropbox/documents/pb/output/trackdays_output_'+ date_string +'.xlsx'

print(sc.red('Starting: Read excel file ' + iExcel_loc))  

# pick columns to be in report.
sColumns = ['Lineitem name','Created at','Name', 'Billing Name','Email','Billing Phone', 'Notes']


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
sorted_df = sorted_df.rename(columns={'Billing Phone': 'Phone'})


# Define the new order of columns
#new_order = ['Date of TD', 'Group', 'Ref', 'Name', 'Email', 'Lineitem name']
new_order = ['Year','Day','Group', 'Ref', 'Name', 'Email', 'Phone','Notes']


# Rearrange the columns in the DataFrame
sorted_df = sorted_df[new_order]



#print out the report
print (sorted_df)

# Write the output file
sorted_df.to_excel(oExcel_loc, index=False)

print(sc.red('Finished: Results written to ' + oExcel_loc  ))  

