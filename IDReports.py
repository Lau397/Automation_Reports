# Let's define our libraries to use:
import argparse
import os 
import pandas as pd 
import calendar
import datetime

# This code will take in the path of a CSV file that has in it compiled information with unique IDs.
# The script will drop unecessary columns and keep only those relevant IDs to manually review and to create a 
# report for. 
# The output is an Excel file with all the IDs that have a report available to view 
# regardless of their status. 

def parse_args():
    parser = argparse.ArgumentParser(description="Get the IDs of firms that have reports pending to review regardless of their status")
    parser.add_argument(
        "file_dir",
        type=str,
        help=(
            "Full path to the directory where the Excel file is located"
        ),
    )

    args = parser.parse_args()

    return args

# Opening the CSV file:
def main(file_dir):
    csv_reader = pd.read_csv(file_dir)
    csv_reader.drop(columns= ['Vehicle', 'By', 'Added', 'Firm'], inplace=True)
    csv_reader.dropna(subset = 'Id', inplace=True)
    csv_reader.reset_index(drop=True, inplace=True)
    print("Getting information in order...")
    # First, check for those IDs in 'Complete' status that don't have a report available for checking ('NaN'):
    if csv_reader['Status'].str.contains('Complete',regex=False).any():
    # Selecting 'Complete' IDs:
        complete_status = csv_reader.loc[csv_reader['Status']=='Complete']
    # Selecting 'Complete' IDs and 'NaN' Reports:
        complete_na_report = complete_status.loc[csv_reader['Report'].isna()]

    # Since we're not checking those, let's drop them from the whole df to consider only those
    # that have reports to check and are either 'Failed' or 'Complete' status:
        csv_reader.drop(complete_na_report.index, axis=0, inplace=True)

    current_time = datetime.datetime.now()

    # Checking if the path exists:
    if not os.path.exists("/revisions"):
      
        # If the output folder path doesn't exist, then, create it.
        file_dir = os.makedirs("/revisions")

    csv_reader.to_excel(file_dir + calendar.month_name[current_time.month] + str(current_time.day) + "_" + str(current_time.hour) + "_" + str(current_time.minute) + "_report.xlsx")
    print("Firm IDs information exported to Excel in Revisions folder")

if __name__ == "__main__":
    args = parse_args()
    main(args.file_dir)