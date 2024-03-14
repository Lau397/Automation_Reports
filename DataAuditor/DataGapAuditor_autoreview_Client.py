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
import re 

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
            except:
            # If sheet is not found then let's try this so the code can continue:
                print('No sheet found for the vehicle {}'.format(sheet_names[j]))
                dict_ = {'Database': excel_file_content.iloc[4][1],            # Database name e.g. "Wilshire"
                         sheet_names[j]: "No audit data generated.",        # Product/vehicle name with description of findings e.g. "Core Fixed Income Composite (P73285)"
                        } 
                nodata_.append(dict_) # Adding the respective database and vehicle name that does not exist to list
                output_df_ = pd.DataFrame(nodata_).groupby(['Database']).sum() # Grouping dataframe by database
                continue

            # Selecting the header names placed in row 7:
            excel_file_orig.rename(columns = excel_file_orig.iloc[7], inplace= True)
            # Selecting the rows with data and reseting the index:
            excel_file = excel_file_orig[7:][1:].set_index(['Date'], drop=True)    
            # We need information from 09/2022 onwards, so I'll be turning Date column into correct type and then filter by date:
            excel_file.index = pd.to_datetime(excel_file.index)
            # Selecting data in the dataframe by the correct date:
            excel_file = excel_file[~(excel_file.index < '09/2022')]
            # Setting up the correct format for the index/Date column
            excel_file.index = excel_file.index.strftime("%m/%Y")
            # Dropping rows and columns in which all the cells contain NaN values:
            excel_file = excel_file.dropna(how='all', axis=0).dropna(how='all', axis=1)
            # Removing commas and $ signs in the numerical values:
            #excel_file = excel_file.replace(, regex=True).replace('Million', '').replace('million', '')

            # Creating a for loop to assign dummy variables to the Data Gap Auditor report:
            for n in range(0, excel_file.shape[1]):
                for m,p in enumerate(excel_file[excel_file.columns[(n)]]):
                        # Avoiding code crashes using try/except:    
                    try:
                        if (int(float(p) >= 0)) or (int(float(p) <= 0)):
                                        excel_file[excel_file.columns[n]][m] = '0'  # "Complete"
                    except:
                        # APX FILE if (' / <NO DATA>') in p:
                        # APX FILE     excel_file[excel_file.columns[n]][m] = '1'
                        # APX FILE else:
                        # APX FILE    excel_file[excel_file.columns[n]][m] = ''

                        if ('<NO APX> / ') in p:
                             excel_file[excel_file.columns[n]][m] = '2'     # "Data not in the Vault" // Client could want APX to distribute this data for them
                        elif re.match(r'(-?[0-9\.]+) \s*/ (-?[0-9\.]+)', p):
                             excel_file[excel_file.columns[n]][m] = '3'     # "Data not matching" // APX needs to review this data until it matches/is Complete  
                        else:
                             excel_file[excel_file.columns[n]][m] = ''

                            #if (r'(-?[0-9\.]+)\s*/ <NO DATA>') in p:
                            #    excel_file[excel_file.columns[n]][m] = '1' # "Data not in the database" // APX needs to distribute this data
                            #elif "<NO APX> / " in p:
                            #    excel_file[excel_file.columns[n]][m] = excel_file[excel_file.columns[n]][m].replace(p, '2') # "Data not in the Vault" // Client could want APX to distribute this data for them
                            #elif " / " in p:
                            #    excel_file[excel_file.columns[n]][m] = excel_file[excel_file.columns[n]][m].replace(p, '3') # "Data not matching" // APX needs to review this data until it matches/is Complete     
                            #else:
                            #    excel_file[excel_file.columns[n]][m] = ''  # If the cell does not contain any of this criteria above, then it's not relevant for our analysis/review
            #excel_file = excel_file.replace(',','', regex=True).replace(' %','', regex=True).replace('-','', regex=True)
            #excel_file = excel_file.replace({
             #                               r'(-?[0-9\.]+) \s*/ <NO DATA>': '1',                 # Data not in database
             #                               r'[a-zA-Z0-9_.-] \s*/ <NO DATA>': '1',
             #                               '<NO DATA>': '1', 
             #                               r'<NO APX> \s*/ (-?[0-9\.]+)': '2',                 # Data not in Vault
             #                               r'<NO APX> \s*/ [a-zA-Z0-9_.-]': '2',
             #                               '<NO APX>': '2',
             #                               r'(-?[0-9\.]+) \s*/ (-?[0-9\.]+)': '3',              # Data not matching
             #                               ' / ':'0',
             #                               r"[-%()\"#/@$;:<>{}`+=~|.!?,]":''}, regex=True)

            for n in range(0, excel_file.shape[1]):
                for m,p in enumerate(excel_file[excel_file.columns[(n)]]):

                     if (p != '1') and (p != '2') and (p != '3') and (p != '0'):
                          excel_file[excel_file.columns[n]][m] = ''
                     #if not ('2' in p):
                     #     excel_file[excel_file.columns[n]][m] = ''
                     #if not ('3' in p):
                     #     excel_file[excel_file.columns[n]][m] = ''
                     #if not ('0' in p):
                     #     excel_file[excel_file.columns[n]][m] = ''

                # Let's fill the NaN values for easier further processes:
            excel_file.fillna('', inplace=True)
                # Putting the dummy variables in a single column named 'Review':
            excel_file['Review'] = excel_file[excel_file.columns[0:]].apply(lambda x: ''.join(x.astype(str)), axis=1)

            for m,p in enumerate(excel_file['Review']):


                            #if all('0' in k for k in p):
                            #    excel_file['Review'][m] = excel_file['Review'][m].replace(p, '')

                        #if any('1' in k for k in p):
                        #    excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Priority 1')          # APX needs to distribute this data

                            #elif (all('1' in k for k in p)):
                            #    excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Priority 1')     # Data not in the Vault  
                            if all('2' in k for k in p):
                                excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Priority 2')      # Client could want APX to distribute this data for them
                            elif any('2' in k for k in p) and any('0' in k for k in p):
                                excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Priority 2')      # Client could want APX to distribute this data for them

                            elif all('3' in k for k in p):
                                excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Priority 3')      # Data not matching
                            elif any('3' in k for k in p) and any('0' in k for k in p):
                                excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Priority 3')      # Data not matching
                            
                            elif any('2' in k for k in p) and any('3' in k for k in p):
                                excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Priority 2 and Priority 3')      # Data not in the Vault and not matching

                            elif all('' in k for k in p):
                                excel_file['Review'][m] = excel_file['Review'][m].replace(p, '')
                                 
                            
            excel_file.to_excel(r'C:\Users\l.arguello\Documents\Python Scripts\APX_automation_reports\output\data_auditor_review\{}_{}_sheet{}.xlsx'.format(file_names[i], dataset, sheet_names[j]))
            #periods_1 = []
            periods_2 = []       
            periods_3 = []
            periods_2_3 = []
            description = []
            for m,p in enumerate(zip(excel_file['Review'],excel_file.index)):

            #if p[0] == 'Complete':
            #    periods_0.append(p[1])

            # APX FILE     if p[0] == 'Priority 1':
            # APX FILE         periods_1.append(p[1])
            # APX FILE     elif periods_1 == periods_1:
            # APX FILE         description.append("".format((list(set(periods_1)))).replace("'",'').replace('[','').replace(']',''))    
            # APX FILE periods_1 = list(set(periods_1))

                if p[0] == 'Priority 2':
                     periods_2.append(p[1])
                elif p[0] == 'Priority 3':
                     periods_3.append(p[1])
                elif p[0] == 'Priority 2 and Priority 3':
                     periods_2_3.append(p[1])

                if periods_2 == periods_2:
                     description.append("".format((list(set(periods_2)))).replace("'",'').replace('[','').replace(']',''))    
                elif periods_3 == periods_3:
                     description.append("".format((list(set(periods_3)))).replace("'",'').replace('[','').replace(']',''))  
                elif periods_2_3 == periods_2_3:
                     description.append("".format((list(set(periods_2_3)))).replace("'",'').replace('[','').replace(']',''))
            periods_2 = list(set(periods_2))
            periods_3 = list(set(periods_3))
            periods_2_3 = list(set(periods_2_3))


                #elif p[0] == 'Priority 2':
                #    periods_2.append(p[1])
                #elif p[0] == 'Priority 3':
                #    periods_3.append(p[1])
            #0 "Complete"
            #1 "Data not in the Vault"
            #2 "Data not in the database"
            #3 "Data not matching"    
            # A description list is created to put in the final review without considering empty period lists:
            # if periods_0 := periods_0: description.append("✔ Complete for the periods: {}\n".format((list(set(periods_0)))).replace("'",'').replace('[','').replace(']',''))
            # if periods_1 := periods_1: description.append("● Priority 1: {}\n".format((list(set(periods_1)))).replace("'",'').replace('[','').replace(']',''))
            if periods_2 := periods_2: description.append("● Priority 2: {}\n".format(list(set((periods_2)))).replace("'",'').replace('[','').replace(']',''))
            if periods_3 := periods_3: description.append("● Priority 3: {}\n".format((list(set(periods_3)))).replace("'",'').replace('[','').replace(']',''))  
            if periods_2_3 := periods_2_3: description.append("● Priority 2 and Priority 3: {}\n".format((list(set(periods_2_3)))).replace("'",'').replace('[','').replace(']',''))  
            description = list(set(description))

            excel_file_content = pd.read_excel(file_path+'/'+file_names[i]) 
            # Building the dictionary to then transform it into a dataframe:
            dict = {'Database': excel_file_content.iloc[4][1],      # Database name e.g. "Wilshire"
                    excel_file_orig.iloc[6][1]: description,        # Product/vehicle name with description of findings e.g. "Core Fixed Income Composite (P73285)
                    }  
            # Creating a new dataframe that will sum up the findings in the Data Auditor        
            output_df = pd.DataFrame([dict])
            # Putting each description in a single line (this may duplicate the database name):
            output_df0 = output_df.explode(excel_file_orig.iloc[6][1])
            # Final dict
            final_dict.append(output_df0)
# ###########################
            
    # Transforming into a dataframe the last dictionary with the review description:
    final_dict_ = pd.concat(final_dict)
    final_dict_ = final_dict_.reindex(sorted(final_dict_.columns), axis=1)
    # final_dict_ has a numerical index, whilst output_df_ has databases as its index, so we'll arrange that:
    final_dict_.set_index("Database", drop=True, inplace=True)

    for col1 in output_df_.columns:
        for col2 in final_dict_.columns:
            if col1 in col2:
                output_df_.columns = final_dict_.columns    
   
        # Joining these two tables together and grouping them by database:
    #if output_df is not None:
            #final_dict = pd.DataFrame([final_dict])
            #output_df = pd.DataFrame([output_df])
    review_file = pd.merge(final_dict_, output_df_, on='Database', how='outer').groupby('Database').sum()
    #elif output_df_ is not None:
            #final_dict_ = pd.DataFrame([final_dict_])
            #output_df_ = pd.DataFrame([output_df_])
    #        review_file = pd.merge(final_dict_, output_df_, on='Database', how='outer').groupby('Database').sum()
        # Dropping unnecessary columns and replacing zero values with the description "No audit data generated":
    review_file.columns = review_file.columns.str.rstrip("_x")
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