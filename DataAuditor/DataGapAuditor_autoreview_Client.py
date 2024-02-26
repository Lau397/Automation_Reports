# This script will generate an ouput Excel file for Client user.

# This Python code will review the original Data Gap Auditor files generated from the Vault, 
# and will sum up what has been found in those files.
# The code will work to create 2 separate Excel files and will loop through each dataset 
# to be worked on: AUM, Performance, Holdings, and Characteristics.

# The code will then work on that dataset and will create an output Excel file for for each user (APX and Client)
# that will contain a table summing up all the findings. It will also include
# all databases and the products/vehicles stated in the sheet_names variable.

# All the original files are put into the main folder: Manulife_DataAuditor, within it, 
# the folder contains four (4) folders for each dataset separately and within each, 
# the Data_Audit_Report Excel files for each database downloaded from the Vault
# Here's an example of main folder organization for this code: Manulife_DataAuditor/Performance/Data_Audit_Report_Alpha_Portfolio_Advisors_1_2024.xlsx

# Importing Python libraries that will be used:
import os
import glob 
import pandas as pd
from tqdm import tqdm
import time

import warnings
warnings.filterwarnings("ignore")

start_time = time.time()


# Setting dataset types:
datasets = ['AUM', 'Performance', 'Holdings', 'Characteristics']


print('Reviewing for the user: Client')

for dataset in datasets:
    # Vehicles that we're interested in are being listed here:
    # Core Fixed Income	                    12776	P73285
    # Core Plus Fixed Income	            12777	P74285 
    # Global Quality Value	                12783	P85285
    # Strategic Fixed Income	            12811	P121285
    # Strategic Fixed Income Opportunities	12812	P126285
    # US Small Cap Core	                    12823	P147285
    sheet_names = ['P73285', 'P74285', 'P85285', 'P121285', 'P126285', 'P147285']
    print('Checking dataset ', dataset)
    # Setting file path. We'll be opening first the Performance folder:
    absolute_path = "C:/Users/l.arguello/Downloads/Manulife_DataAuditor/"
    # Full file path:
    file_path = absolute_path + dataset
    # Using glob to get all the Excel file names in the selected folder, to loop through them:
    csv_files = glob.glob(os.path.join(file_path, "[!~]*.xlsx")) # [!~] to ignore temporary/opened files
    # Empty list to store file names from folder:
    file_names = []
    # Loop over the list of Excel files: 
    for f in tqdm(csv_files, desc="Loading…",ascii=False, ncols=75):        
            time.sleep(0.03) 
            # Print the location and filename 
            print('File Name:', f.split("\\")[-1]) 
            # Add each Excel file name to file_names list 
            file_names.append(f.split("\\")[-1])      
    print("Complete.")
    # Creating empty lists that will contain reviewed tables:
    final_dict = []
    nodata_ = []        
    # For loop to select the Excel file:
    for i in range(len(file_names)):
        print('Checking file name: ', file_names[i])
        # This variable will contain the first sheet in the Data Audtor (table of contents) which will be needed to fill information in the tables:
        excel_file_content = pd.read_excel(file_path+'/'+file_names[i]) 
        # For loop to select the sheet name (vehicle):
        for j in range(len(sheet_names)):
            # Will do a try and except since there are sheets that don't exist in the files, so the code doesn't crash:
            try:
                print('Checking sheet name: ', sheet_names[j])
                # Defining the Excel file to be openned and the sheet we need from the book:
                excel_file_orig = pd.read_excel(file_path+'/'+file_names[i], sheet_name=sheet_names[j])
            # If sheet is not found then let's try this so the code can continue:
            except:
                print('No sheet found for the vehicle {}'.format(sheet_names[j]))
                dict_ = {'Database': excel_file_content.iloc[4][1],            # Database name e.g. "Wilshire"
                excel_file_orig.iloc[6][1]: "No audit data generated.",        # Product/vehicle name with description of findings e.g. "Core Fixed Income Composite (P73285)"
                    } 
                nodata_.append(dict_) # Adding the respective database and vehicle name that does not exist to list
                output_df_ = pd.DataFrame(nodata_).groupby(['Database']).sum() # Grouping dataframe by database
                continue
            # Selecting the header names placed in row 7 (row in which we start to have some relevant information to gather):
            excel_file_orig.rename(columns = excel_file_orig.iloc[7], inplace= True)
            # Selecting the rows with data and resetting the index:  
            excel_file = excel_file_orig[7:][1:].set_index(['Date'], drop=True)
            # Checking data type of all columns in the file:
            excel_file.info()
            # Date column does not have the correct type, the others are mixed due to special characters being in them such as /
            # We need information from 09/2022 onwards, so I'll be turning Date column into correct type and then filter by date:
            excel_file.index = pd.to_datetime(excel_file.index)
            # Selecting data in the dataframe by the correct date:
            excel_file = excel_file[~(excel_file.index < '09/2022')]
            # Setting up the correct format for the index/Date column
            excel_file.index = excel_file.index.strftime("%m/%Y")
            # Dropping rows and columns in which all the cells contain NaN values:
            excel_file = excel_file.dropna(how='all', axis=0).dropna(how='all', axis=1)


            # Renaming the file by creating a copy of DataFrame.
            df_client = excel_file.copy()           
                
            # Creating a for loop to assign dummy variables to the Data Gap Auditor report:
            for n in range(0, df_client.shape[1]):
                    for m,p in enumerate(df_client[df_client.columns[(n)]]):
                            try:
                                if (int(float(p) >= 0)) or (int(float(p) <= 0)): 
                                    df_client[df_client.columns[n]][m] = ''
                            except:
                                continue
            # This will check for the data is not in the database/vault/matching:
            df_client = df_client.replace({r'(-?[0-9\.]+)\s*/ <NO DATA>': '',       # Data not in database
                                           r'<NO APX> \s*/ (-?[0-9\.]+)': '2',      # Data not in Vault
                                           r'(-?[0-9\.]+)\s*/ (-?[0-9\.]+)': '3',   # Data not matching
                                          }, regex=True)
            # Let's fill the NaN values for easier further processes:# "Data not in the Vault" // Client could want APX to distribute this data for them
            df_client.fillna('', inplace=True)
            # "Data not matching" // APX needs to review this data until it matches/is Complete     
            # Putting the dummy variables in a single column named 'Review':
            df_client['Review'] = df_client[df_client.columns[0:]].apply(lambda x: ''.join(x.astype(str)), axis=1)# If the cell does not contain any of this criteria above, then it's not relevant for our analysis/reviewal
            # Creating a for loop to assign the correct description for each number stated above ^:
            for m,p in enumerate(df_client['Review']):
                #if (all('2' in k for k in p)):
                #    df_client['Review'][m] = df_client['Review'][m].replace(p, 'Priority 2')    # Data not in the Vault
                #elif (all('3' in k for k in p)):
                #    df_client['Review'][m] = df_client['Review'][m].replace(p, 'Priority 3')    # Data not matching // Data needs to be reviewed
                if (any('2' in k for k in p)):
                    df_client['Review'][m] = df_client['Review'][m].replace(p, 'Priority 2')    # Data not in the Vault
                elif (any('3' in k for k in p)):
                  df_client['Review'][m] = df_client['Review'][m].replace(p, 'Priority 3')      # Data not matching // Data needs to be reviewed
                else:
                    df_client['Review'][m] = df_client['Review'][m].replace(p, 'No annotation')
            # Creating a list for each of the periods in the Review column:
            periods_2 = []       
            periods_3 = []
            # Gathering all the periods (month/year) for which each of these conditions stated above are present:
            for m,p in enumerate(zip(df_client['Review'], df_client.index)):                    
                if p[0] == 'Priority 2':
                    periods_2.append(p[1])
                elif p[0] == 'Priority 3':
                    periods_3.append(p[1])
            #0 "Complete"
            #1 "Data not in the database"
            #2 "Data not in the Vault"
            #3 "Data not matching"    
            # A description list is created to put in the final review without considering empty period lists:
            description_client = []
            if periods_2 := periods_2: description_client.append("● Priority 2: {}\n".format(list(set((periods_2)))).replace("'",'').replace('[','').replace(']',''))
            if periods_3 := periods_3: description_client.append("● Priority 3: {}\n".format((list(set(periods_3)))).replace("'",'').replace('[','').replace(']','')) 
    
    # Loading the first sheet "Table of Contents" to obtain information that can be input into the output dataframe:
    excel_file_content = pd.read_excel(file_path+'/'+file_names[i]) 
    # Building the dictionary that will contain the genetal information: database and description, to then transform it into a dataframe:
    dict = {'Database': excel_file_content.iloc[4][1],      # Database name e.g. "Wilshire"
            excel_file_orig.iloc[6][1]: description_client,        # Product/vehicle name with review description of findings e.g. "Core Fixed Income Composite (P73285)"
            }  
    # Creating a new dataframe that will sum up the findings in the Data Auditor        
    output_df = pd.DataFrame([dict])
    # Putting each description in a single line (this may duplicate the database name):
    output_df0 = output_df.explode(excel_file_orig.iloc[6][1])
    # Final dict
    final_dict.append(output_df0)
    # Transforming into a dataframe the last dictionary with the review description:
    final_dict_ = pd.concat(final_dict) 
    # final_dict_ has a numerical index, whilst output_df_ has databases as its index, so we'll arrange that:
    final_dict_.set_index("Database",drop=True, inplace=True)
    # Joining these two tables together and grouping them by database:
    review_file = pd.merge(final_dict_, output_df_, on='Database', how='outer').groupby('Database').sum()
    # Dropping unnecessary columns and replacing zero values with the description "No audit data generated":
    review_file.columns = review_file.columns.str.replace(r'_x$', '')
    review_file = review_file.drop([x for x in review_file if x.endswith('_y')], axis = 1)
    review_file = review_file.replace(0, "No audit data generated.", regex=True)
    # Sorting column names and Database names:
    review_file = review_file.reindex(sorted(review_file.columns), axis=1)
    # Sorting index alphabetically (case insensitive):
    review_file = review_file.reindex(index=(sorted(review_file.index, key=lambda s: s.lower())))
    # Making sure index name doesn't get lost:
    review_file.index.name = 'Database'
    # Output path with respective name:
    excel_output = r'C:\Users\l.arguello\Documents\Python Scripts\APX_automation_reports\output\data_auditor_review\DataAuditor_review_{}_Client.xlsx'.format(dataset)
    # Adding legend/keys table:
    legend_dict = {'Priority 2': "",
                   'Priority 3': "" }
    legend_keys = pd.DataFrame([legend_dict])
    legend_keys = legend_keys.set_axis(['Legend'], axis='index').transpose()


    with pd.ExcelWriter(excel_output, engine="xlsxwriter") as writer:
        writer.book.formats[0].set_text_wrap()  # Update global format with text_wrap
        legend_keys.to_excel(writer, startrow = 1, startcol = 1) # Export to Excel file
        review_file.to_excel(writer, startrow = 6, startcol = 1)     
    # ////////////////////////////////////////////////////////////////////////////////
    # //////////////////////////////////Extra Steps//////////////////////////////////
    # //////////////////////////////////////////////////////////////////////////////        
    ## Accessing the Pandas file and sheet to add plot:
        # Loading worksheet for some formatting:
        worksheet = writer.sheets['Sheet1']
        # Set border color for tables and set vertical alignment of text:
        file_format = writer.book.add_format()
        file_format.set_text_wrap(True)
        file_format.set_border_color('#A6A6A6')
        file_format.set_align('left')
        file_format.set_valign('vcenter')
        for col_num, value in enumerate(review_file.columns.values):    
            header_format = writer.book.add_format({'bold':True, 'fg_color': '#F2F2F2', 'border_color':'black'})
            worksheet.write(6, col_num+2, value, header_format) # Set header format in soft gray color
        worksheet.set_column('B:H', 19.86, file_format)  # Set size of column (19.86 pixels     
        # Formatting cells:
        # Create a format to use in a merged range
        merge_format1 = writer.book.add_format(
            {
                "bold": 1,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "fg_color": "#FCE4D6",
            })
        merge_format2 = writer.book.add_format(
            {
                "border": 1,
                "align": "left",
                "valign": "vcenter",
            })
      

        worksheet.merge_range("C2:F2", "Legend", merge_format1)
        worksheet.merge_range("C3:F3", "Data not in the Vault // Client could want APX to distribute this data for them", merge_format2)
        worksheet.merge_range("C4:F4", "Data not matching // Review this data until it matches/is Complete", merge_format2)     
        
        writer.close()

# This is just the time the process took to complete per dataset
timetaken = (time.time() - start_time)/60
print("Task completed in %.2f minutes" % timetaken)