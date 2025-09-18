# Imports
import multiprocessing
import time
import warnings
import os
import configparser
import configparser
from datetime import datetime

# BI Imports
import Combine
import Standardize

# Read config.ini for the Environment ("PROD" or "DEV")
config = configparser.ConfigParser()
config.read(r'C:\biautomation\config.ini')
Environment = config['ENV']['Environment']

# Hide warning message from openyxl
warnings.filterwarnings("ignore", message="Workbook contains no default style")

def Transform_Files(_ReportTypesToCombine):
    start_time = time.time()
    # Combine individual tenant files into one file
    print(f"Combining files for {_ReportTypesToCombine}...")
    combined_df = Combine.combine_files(_ReportTypesToCombine, Environment)
    # Standardize the combined file 
    print(f"Standardizing {_ReportTypesToCombine}...")
    standardized_df = Standardize.Standardize_Files(combined_df, _ReportTypesToCombine)
    # Save the Combined & Satandardized files to a csv for database
    if _ReportTypesToCombine == "BI_ARTransactions":
        output_file_name = f"BI_{_ReportTypesToCombine.replace('BI_', '')}"
    else:
        output_file_name = f"SQL_{_ReportTypesToCombine.replace('SQL_', '')}"
    # Save a "daily" version the Combined & Satandardized file to a csv for database
    if _ReportTypesToCombine == "SQL_Notes":
        todays_date = datetime.today()
        todays_date_formatted = todays_date.strftime("%m-%d-%Y")
        file_path = rf'C:\biautomation\ETL\Transform\Data_Lake\Historical\SQL_Notes\{output_file_name} {todays_date_formatted}.csv'
        # Check if the file already exists to only produce the daily once per day
        if not os.path.exists(file_path):
            standardized_df.to_csv(file_path)
    # Save the Combined & Satandardized files as a csv in the Data_Lake
    standardized_df.to_csv(rf"C:\biautomation\ETL\Transform\Data_Lake\{output_file_name}.csv")
    end_time = time.time()
    elapsed_time = end_time - start_time
    formatted_time = time.strftime("%H:%M:%S", time.gmtime(elapsed_time))
    print(f"Finished {_ReportTypesToCombine}...{formatted_time}")

def main():
    # A list of each report category (order by longest process time first to imporove performance)
    ReportTypesToCombine = ["SQL_Calls",
                            "SQL_Invoices",                            
                            "SQL_Customers",
                            "SQL_ScheduledJobs",
                            "SQL_InventoryLineItems",
                            "SQL_CompletedJobs",
                            "SQL_EstimatesCreatedOn",
                            "SQL_AllPayments_",
                            "SQL_AllPaymentsExcludeDeleted",
                            "SQL_AppliedPayments",
                            "SQL_EstimatesSoldOn",
                            "SQL_Notes",
                            "SQL_Memberships",
                            "SQL_ARTransactions_Dec24",
                            "SQL_ARTransactions_Jan25",
                            "BI_ARTransactions",
                            "SQL_MarketingCampaigns_Yesterday",
                            "SQL_MarketingCampaigns_Today",
                            "SQL_Rehash",
                            "SQL_OfficeAuditTrail", 
                            "SQL_CSR",
                            "SQL_Appointments"]

    # Multi-Process report combining
    print("Number of Reports to Combine : ", len(ReportTypesToCombine))
    print("Number of CPU's available : ", multiprocessing.cpu_count())
    pool_size = min(len(ReportTypesToCombine), multiprocessing.cpu_count())
    with multiprocessing.Pool(processes=pool_size) as pool:
        pool.map(Transform_Files, ReportTypesToCombine)

if __name__ == '__main__':
    main()