import os
import glob 
import pandas as pd
from tqdm import tqdm
import time

import warnings
warnings.filterwarnings("ignore")

start_time = time.time()


# Allowing the user to select the dataset so we can locate the respective folder:
user_dataset = input('Enter the dataset to review from the Data Auditor: ')
print("The user would like to use the dataset: ", user_dataset)


# Vehicles that we're interested in are being listed here:

# Core Fixed Income	                    12776	P73285
# Core Plus Fixed Income	            12777	P74285 
# Global Quality Value	                12783	P85285
# Strategic Fixed Income	            12811	P121285
# Strategic Fixed Income Opportunities	12812	P126285
# US Small Cap Core	                    12823	P147285

sheet_names = ['P73285', 'P74285', 'P85285', 'P121285', 'P126285', 'P147285']


# Setting file path. We'll be opening first the Performance folder:
absolute_path = "C:/Users/l.arguello/Downloads/Manulife_DataAuditor/"

file_path = absolute_path+user_dataset

# Using glob to get all the Excel file names in the selected folder, to loop through them:

csv_files = glob.glob(os.path.join(file_path, "[!~]*.xlsx")) 
  
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
nodata_ =[]        

# For loop to select the Excel file:

for i in range(len(file_names)):
    # For loop to select the sheet name (vehicle):
    print('Checking file name: ', file_names[i])

    excel_file_content = pd.read_excel(file_path+'/'+file_names[i]) 

    for j in range(len(sheet_names)):
        
        try:
            print('Checking sheet name: ', sheet_names[j])

           # Defining the Excel file to be openned and the sheet we need from the book:
            excel_file_orig = pd.read_excel(file_path+'/'+file_names[i], sheet_name=sheet_names[j])

           # If sheet is not found then let's try this so the code can continue:
        except:
            print('No sheet found for the vehicle {}'.format(sheet_names[j]))
            dict_ = {'Database': excel_file_content.iloc[4][1],      # Database name e.g. "Wilshire"
            excel_file_orig.iloc[6][1]: "No audit data generated.",        # Product/vehicle name with description of findings e.g. "Core Fixed Income Composite (P73285)"
                } 
            nodata_.append(dict_)
            output_df_ = pd.DataFrame(nodata_).groupby(['Database']).sum()
            continue
        

        # Selecting the header names placed in row 7:
        excel_file_orig.rename(columns = excel_file_orig.iloc[7], inplace= True)

        # Selecting the rows with data and reseting the index:  
        excel_file = excel_file_orig[7:][1:].set_index(['Date'], drop=True)

        # Checking data type of all columns in the file:
        excel_file.info()
        # Date column does not have the correct type, the others are mixed due to characters being in them such as /
        
        # We need information from 09/2022 onwards, so I'll be turning Date column into correct type and then filter by date:
        excel_file.index = pd.to_datetime(excel_file.index)
        # Selecting data in the dataframe by the correct date:
        excel_file = excel_file[~(excel_file.index < '09/2022')]
        
        # Setting up the correct format for the index/Date column
        excel_file.index = excel_file.index.strftime("%m/%Y")
        
        # Dropping rows and columns in which all the cells contain NaN values:
        excel_file = excel_file.dropna(how='all', axis=0).dropna(how='all', axis=1)

        # Creating a for loop to assign dummy variables to the Data Gap Auditor report:
        for n in range(0, excel_file.shape[1]):
        
            for m,p in enumerate(excel_file[excel_file.columns[(n)]]):


                try:
                    if float(p) >= 0 or float(p) <= 0:
                    
                        excel_file[excel_file.columns[n]][m] = excel_file[excel_file.columns[n]][m].replace(p, '0') # "Complete"

                except:
                    if "<NO APX> / " in p:
                        excel_file[excel_file.columns[n]][m] = excel_file[excel_file.columns[n]][m].replace(p, '1') # "Data not in the Vault"
                    elif " / <NO DATA>" in p:
                        excel_file[excel_file.columns[n]][m] = excel_file[excel_file.columns[n]][m].replace(p, '2') # "Data not in the database"
                    elif " / " in p:
                        excel_file[excel_file.columns[n]][m] = excel_file[excel_file.columns[n]][m].replace(p, '3') # "Data not matching"  
                    else:
                        excel_file[excel_file.columns[n]][m] = excel_file[excel_file.columns[n]][m].replace(p, '')  # If the cell does not contain any of this criteria above, then it's not relevant for our analysis/reviewal


        # Let's fill the NaN values for easier further processes:
        excel_file.fillna('', inplace=True)

        # Putting the dummy variables in a single column:
        excel_file['Review'] = excel_file[excel_file.columns[0:]].apply(lambda x: ''.join(x.astype(str)), axis=1)

        # Creating a for loop to assign the correct description to each period:
        for m,p in enumerate(excel_file['Review']):

                if all('0' in k for k in p):
                    excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Complete')

                elif all('1' in k for k in p):
                    excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Data not in the Vault')

                elif all('2' in k for k in p):
                    excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Data not in the database')   

                elif all('3' in k for k in p):
                    excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Data not matching') 


        # Now we need to continue to put the other conditions:
        for m,p in enumerate(excel_file['Review']):
        
            if (('1' in p) and ('0' in p)):
                excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Data not in the Vault')

            elif (('2' in p) and ('0' in p)):
                excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Data not in the database')

            elif (('3' in p) and ('0' in p)):
                excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Data not matching')

            elif (('3' in p) and ('1' in p)):
                excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Data not in the Vault and not matching')

            elif (('2' in p) and ('1' in p)):
                excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Data not in the Vault and not in the database')

            elif (('3' in p) and ('2' in p) and ('1' in p)):
                excel_file['Review'][m] = excel_file['Review'][m].replace(p, 'Data not in the Vault, not matching and not in the database')

        
        # Creating a list for each of the periods in the Review column:
        periods_0 = []
        periods_1 = []
        periods_2 = []       
        periods_3 = []

        for m,p in enumerate(zip(excel_file['Review'],excel_file.index)):
        
            if p[0] == 'Complete':
                periods_0.append(p[1])

            elif p[0] == 'Data not in the Vault':
                periods_1.append(p[1])

            elif p[0] == 'Data not in the database':
                periods_2.append(p[1])

            elif p[0] == 'Data not matching':
                periods_3.append(p[1])


        #0 "Complete"
        #1 "Data not in the Vault"
        #2 "Data not in the database"
        #3 "Data not matching"    

        # A description list is created to put in the final review without considering empty period lists:
        description = []
        if periods_0 := periods_0: description.append("✔ Complete for the periods: {}".format((list(set(periods_0)))).replace("'",'').replace('[','').replace(']',''))
        if periods_1 := periods_1: description.append("● Data not in the Vault for the periods: {}".format((list(set(periods_1)))).replace("'",'').replace('[','').replace(']',''))
        if periods_2 := periods_2: description.append("● Data not in the database for the periods: {}".format(list(set((periods_2)))).replace("'",'').replace('[','').replace(']',''))
        if periods_3 := periods_3: description.append("● Data not matching for the periods: {}".format((list(set(periods_3)))).replace("'",'').replace('[','').replace(']',''))  


        # Loading the first sheet "Table of Contents" to obtain information that can be input into the output dataframe:
        excel_file_content = pd.read_excel(file_path+'/'+file_names[i]) 


        # Building the dictionary to then transform it into a dataframe:

        dict = {'Database': excel_file_content.iloc[4][1],      # Database name e.g. "Wilshire"
                excel_file_orig.iloc[6][1]: description,        # Product/vehicle name with description of findings e.g. "Core Fixed Income Composite (P73285)"
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
review_file = review_file.columns.str.replace(r'_x$', '')
review_file = review_file.drop([x for x in review_file if x.endswith('_y')], 1)
review_file = review_file.str.replace('0', "No audit data generated.", regex=True)

# Sorting column names and Database names:
review_file = review_file.reindex(sorted(review_file.columns), axis=1)
review_file = review_file.sort_values(by=['Database'])


print("Task completed in %.2fs seconds" % (time.time() - start_time))