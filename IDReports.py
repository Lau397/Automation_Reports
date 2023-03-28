# Let's define our libraries to use:
import argparse
import os 
import pandas as pd 
import calendar
import datetime

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

    file_dir = "C:/Users/l.arguello/Desktop/Revisions/"

    csv_reader.to_excel(file_dir + calendar.month_name[current_time.month] + str(current_time.day) + "_" + str(current_time.hour) + "_" + str(current_time.minute) + "_report.xlsx")
    print("Firm IDs information exported to Excel in Revisions folder")

if __name__ == "__main__":
    args = parse_args()
    main(args.file_dir)