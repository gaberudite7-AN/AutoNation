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

def find_latest_file_in_dir(directory, exclude_patterns=None):
    p = Path(directory)
    if not p.exists() or not p.is_dir():
        return None
    files = [f for f in p.iterdir() if f.is_file()]
    if not files:
        return None

    # normalize exclude patterns to lowercase substrings
    exclude_patterns_low = [pat.lower() for pat in (exclude_patterns or [])]

    if exclude_patterns_low:
        files = [f for f in files if not any(pat in f.name.lower() for pat in exclude_patterns_low)]
        if not files:
            return None

def copy_file(source_path, destination_path, historical_path):
    src = Path(source_path)
    dest = Path(destination_path)
    hist = Path(historical_path)

    # List of lowercase substrings to match against dest.name (contains, case-insensitive)
    target_name_patterns = [
        "new ev inventory sales trend",
        "used margin vs actual 2022-2025 trend sales w price bucket w ic",
        # add more patterns here
    ]

    # List of substrings to ignore when picking the latest file from a source directory
    # update this list to avoid picking backups, temporary files, archived copies, etc.
    exclude_name_patterns = [
        "Used PVR Tracking",
    ]

    # If destination filename contains any of the target patterns and the source is a directory,
    # use the latest file from that directory
    dest_name_low = dest.name.lower().strip()
    matches_target = any(pat in dest_name_low for pat in target_name_patterns)

    if src.is_dir() and matches_target:
        latest = find_latest_file_in_dir(src,  exclude_patterns=exclude_name_patterns)
        if latest is None:
            print(f"No suitable files found in directory for target '{dest.name}': {src}", file=sys.stderr)
            return
        actual_source = latest
        print("Using latest file from directory for target:", actual_source)
    else:
        if src.is_dir():
            print(f"Source is a directory but destination '{dest.name}' does not match configured targets: {src}", file=sys.stderr)
            return
        actual_source = src

    # Ensure destination and historical dirs exist
    dest.parent.mkdir(parents=True, exist_ok=True)
    hist.parent.mkdir(parents=True, exist_ok=True)

    # Create historical copy if the file was NOT modified today
    # We want to move the old file
    if not day_checker(actual_source):
        try:
            shutil.copy2(actual_source, hist)
        except Exception as e:
            print(f"Failed to create historical copy: {e}", file=sys.stderr)

    # Always copy latest to destination
    try:
        shutil.copy2(actual_source, dest)
    except Exception as e:
        print(f"Failed to copy to destination: {e}", file=sys.stderr)


"""for finding latest file with include/exclude patterns"""
def find_latest_file_in_dir_include(directory, include_patterns=None, exclude_patterns=None):
    p = Path(directory)
    if not p.exists() or not p.is_dir():
        return None
    files = [f for f in p.iterdir() if f.is_file()]
    if not files:
        return None

    exclude_low = [pat.lower() for pat in (exclude_patterns or [])]
    if exclude_low:
        files = [f for f in files if not any(pat in f.name.lower() for pat in exclude_low)]
        if not files:
            return None

    include_low = [pat.lower() for pat in (include_patterns or [])]
    if include_low:
        files = [f for f in files if any(pat in f.name.lower() for pat in include_low)]
        if not files:
            return None

    return max(files, key=lambda f: f.stat().st_mtime)

def copy_file_with_hc(source_path, destination_path, historical_path, hc_substring="hc"):
    src = Path(source_path)
    dest = Path(destination_path)
    hist = Path(historical_path)

    # substrings to ignore when picking the latest file
    exclude_name_patterns = [
        "Used PVR Tracking",
        # add other excludes if needed
    ]

    # If source is a directory, pick the latest file that contains the hc_substring (case-insensitive)
    if src.is_dir():
        latest = find_latest_file_in_dir_include(src, include_patterns=[hc_substring], exclude_patterns=exclude_name_patterns)
        if latest is None:
            print(f"No suitable HC files found in directory: {src}", file=sys.stderr)
            return
        actual_source = latest
        print("Using latest HC file from directory:", actual_source)
    else:
        if not src.exists():
            print(f"Source does not exist: {src}", file=sys.stderr)
            return
        actual_source = src

    # Ensure destination and historical dirs exist
    dest.parent.mkdir(parents=True, exist_ok=True)
    hist.parent.mkdir(parents=True, exist_ok=True)

    # Create historical copy if the file was NOT modified today
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
            "source": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Discounted_Inventory_Tracking\Discounted_Inventory_Tracking_HC.xlsm",
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
        # Low PVR (New Side)
        {
            "source": r"W:\Corporate\Inventory\Reporting\Low PVR Deals\LOW PVR REPORT NEW & USED.xlsm",
            "destination": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Low_PVR\LOW PVR REPORT NEW & USED.xlsm",
            "historical_destination": rf"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Low_PVR\Historical\LOW PVR REPORT NEW & USED {today}.xlsm"
        }, 
        # Low PVR (Used Side)
        {
            "source": r"W:\Corporate\Inventory\Reporting\Low PVR Deals\LOW PVR REPORT NEW & USED.xlsm",
            "destination": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\Used\Low_PVR\LOW PVR REPORT NEW & USED.xlsm",
            "historical_destination": rf"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\Used\Low_PVR\Historical\LOW PVR REPORT NEW & USED {today}.xlsm"
        }, 
        # Industry
        {
            "source": r"W:\Corporate\Inventory\Reporting\JDPower Industry vs AN",
            "destination": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Industry\Industry Insights Summary File.xlsx",
            "historical_destination": rf"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\New\Industry\Historical\Industry Insights Summary File {today}.xlsx"
        },
        # Used Margin vs Actual
        {
            "source": rf"W:\Corporate\Inventory\Reporting\Used Margin vs Actual\2025\{month_num}",
            "destination": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\Used\Used_Margin_vs_Actual\Used Margin vs Actual 2022-2025 Trend Sales w Price Bucket w IC.xlsm",
            "historical_destination": rf"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\Used\Used_Margin_vs_Actual\Historical\Used Margin vs Actual 2022-2025 Trend Sales w Price Bucket w IC {today}.xlsm"
        },
    ]

    # List of all file paths:
    files_with_HC_to_copy = [
        # Used Margin vs Actual
        {
            "source": rf"W:\Corporate\Inventory\Reporting\Used Key Metrics - Trend Report\Used Buckets Report\2025\{month_num}",
            "destination": r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\Used\Used_Bucket\Used_Buckets_Report_HC.xlsx",
            "historical_destination": rf"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Sharepoint_File_Centralization\Used\Used_Bucket\Historical\Used_Buckets_Report_HC {today}.xlsx"
        },
    ]

    # Loop through all files and copy them
    for file in files_to_copy:
        copy_file(file["source"], file["destination"], file["historical_destination"])

    # Loop through all HC files ONLY and copy them
    for file in files_with_HC_to_copy:
        copy_file_with_hc(file["source"], file["destination"], file["historical_destination"], hc_substring="hc")