import shutil
import datetime
from pathlib import Path
import sys


def day_checker(path):
    p = Path(path)
    if not p.exists() or not p.is_file():
        return False
    mtime = datetime.datetime.fromtimestamp(p.stat().st_mtime)
    return mtime.date() == datetime.datetime.today().date()

def find_latest_file_in_dir(directory):
    p = Path(directory)
    if not p.exists() or not p.is_dir():
        return None
    files = [f for f in p.iterdir() if f.is_file()]
    if not files:
        return None
    return max(files, key=lambda f: f.stat().st_mtime)

def copy_file(source_path, destination_path, historical_path):
    src = Path(source_path)
    dest = Path(destination_path)
    hist = Path(historical_path)

    # Special-case: only when the destination filename is the EV file do we allow a directory source
    ev_target_name = "new ev inventory sales trend.xlsm"
    if dest.name.lower() == ev_target_name and src.is_dir():
        latest = find_latest_file_in_dir(src)
        if latest is None:
            print(f"No files found in EV directory: {src}", file=sys.stderr)
            return
        actual_source = latest
        print("Using latest EV file:", actual_source)
    else:
        if src.is_dir():
            print(f"Source is a directory but not the EV target: {src}", file=sys.stderr)
            return
        actual_source = src

    # Ensure destination and historical dirs exist
    dest.parent.mkdir(parents=True, exist_ok=True)
    hist.parent.mkdir(parents=True, exist_ok=True)

    # Create historical copy if the file was NOT modified today
    # We want to move the old file
    if not day_checker(actual_source):
        try:
            shutil.move(actual_source, hist)
        except Exception as e:
            print(f"Failed to create historical copy: {e}", file=sys.stderr)

    # Always copy latest to destination
    try:
        shutil.copy2(actual_source, dest)
    except Exception as e:
        print(f"Failed to copy to destination: {e}", file=sys.stderr)


if __name__ == "__main__":

    # Dynamic dates setup
    today = datetime.date.today()
    today = today.strftime("%m-%d-%Y")
    month_num = datetime.date.today().month
    month_num = str(month_num).zfill(2)

    print(today)
    print(month_num)

    # List of all file paths:
    files_to_copy = [
        # Allocation Tracker
        {
            "source": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Allocation_Tracker\Allocation_Tracker.xlsm",
            "destination": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Allocation_Tracker\Allocation_Tracker.xlsm",
            "historical_destination": rf"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Allocation_Tracker\Historical\Allocation_Tracker {today}.xlsm"
        },
        # Aston Martin
        {
            "source": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Aston_Martin\MY24 Aston Martin Sales.xlsm",
            "destination": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Aston_Martin_Monitoring\MY24 Aston Martin Sales.xlsm",
            "historical_destination": rf"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Aston_Martin_Monitoring\Historical\MY24 Aston Martin Sales {today}.xlsm"
        },
        # BP Tracker
        {
            "source": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Brand_President_Tracker\BP_Tracker.xlsm",
            "destination": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\BP_Tracker\BP_Tracker.xlsm",
            "historical_destination": rf"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\BP_Tracker\Historical\BP_Tracker {today}.xlsm"
        },
        # Discounted Inventory
        {
            "source": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Brand_President_Tracker\BP_Tracker.xlsm",
            "destination": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Discounted_Inventory\Discounted_Inventory_Tracking_HC.xlsm",
            "historical_destination": rf"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Discounted_Inventory\Historical\Discounted_Inventory_Tracking_HC {today}.xlsm"
        },
        # # Earnings (only done once a quarter...do not update regularly)
        # {
        #     "source": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Earnings\Report5.xlsx",
        #     "destination": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Earnings\Report5.xlsx",
        #     "historical_destination": rf"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Earnings\Historical\Report5 {today}.xlsx"
        # },
        # EV Sales & Inventory
        {
            "source": rf"W:\Corporate\Inventory\Reporting\EV Inventory Trend\NEW\{month_num}",
            "destination": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\EV_Sales_&_Inventory\New EV Inventory Sales Trend.xlsm",
            "historical_destination": rf"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\EV_Sales_&_Inventory\Historical\New EV Inventory Sales Trend {today}.xlsm"
        },
        # Low PVR
        {
            "source": r"W:\Corporate\Inventory\Reporting\Low PVR Deals\LOW PVR REPORT NEW & USED.xlsm",
            "destination": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Low_PVR\LOW PVR REPORT NEW & USED.xlsm",
            "historical_destination": rf"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Low_PVR\Historical\LOW PVR REPORT NEW & USED {today}.xlsm"
        }
    ]


    for file in files_to_copy:
        copy_file(file["source"], file["destination"], file["historical_destination"])