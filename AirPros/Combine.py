import os
import glob
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

# Mapping dictionaries
TENANT_MAP = {
    'AirPros':           "Air Pros, LLC",
    'SEFL':              "Air Pros, LLC",
    'Ocala':             "Air Pros, LLC",
    'Orlando':           "Air Pros, LLC",
    'Tampa':             "Air Pros, LLC",
    'West':              "Air Pros, LLC",
    'AirForce':          "Airforce Heating and Air",
    'CMHeating':         "CM Heating, Inc.",
    'Dallas':            "Dallas Plumbing & Air Pros LLC",
    'Dougs':             "Doug's Service Air Pros LLC",
    'DreamTeam':         "Dream Team Air Pros, LLC",
    'Hansen':            "Hansen Air Pros, LLC",
    'OneSource':         "One Source",
    'PersonalizedPower': "PPS",
    'AF':                "Airforce Heating and Air",
    'CM':                "CM Heating, Inc.",
    'DP':                "Dallas Plumbing & Air Pros LLC",
    'DPC':               "Dallas Plumbing & Air Pros LLC",
    'DT':                "Dream Team Air Pros, LLC",
    'OS':                "One Source",
    'PPS':               "PPS"
}

GROUP_MAP = {
    'AirPros':           "AirPros",
    'SEFL':              "AirPros",
    'Ocala':             "AirPros",
    'Orlando':           "AirPros",
    'Tampa':             "AirPros",
    'West':              "AirPros",
    'AirForce':          "AirForce",
    'CMHeating':         "CM Heating",
    'Dallas':            "Dallas Plumbing",
    'Dougs':             "Doug's",
    'DreamTeam':         "DreamTeam",
    'Hansen':            "Hansen",
    'OneSource':         "OneSource",
    'PersonalizedPower': "PPS",
    'AF':                "AirForce",
    'CM':                "CM Heating",
    'DP':                "Dallas Plumbing",
    'DPC':               "Dallas Plumbing",
    'DT':                "DreamTeam",
    'OS':                "OneSource",
    'PPS':               "PPS"
}

REGION_MAP = {
    'SEFL':              "South East Florida",
    'Ocala':             "Ocala",
    'Orlando':           "Orlando",
    'Tampa':             "Tampa",
    'West':              "West",
    'AirForce':          "Commercial AFH",
    'CMHeating':         "Everett",
    'Dallas':            "Dallas Plumbing",
    'Dougs':             "Doug's Hvac",
    'DreamTeam':         "DreamTeam",
    'Hansen':            "Hansen Hvac",
    'OneSource':         "OneSource Hvac",
    'PersonalizedPower': "PPS Hvac",
    'AF':                "Commercial AFH",
    'CM':                "Everett",
    'DP':                "Dallas Plumbing",
    'DPC':               "Dallas Plumbing",
    'DT':                "DreamTeam",
    'OS':                "OneSource Hvac",
    'PPS':               "PPS Electric"
}

def lookup_value_in_map(file_name: str, lookup_map: dict) -> str | None:
    for key, value in lookup_map.items():
        if key in file_name:
            return value
    return None

def read_file(file_path: str) -> pd.DataFrame | None:
    try:
        if file_path.endswith('.xlsx'):
            return pd.read_excel(file_path)
        elif file_path.endswith('.csv'):
            return pd.read_csv(file_path, encoding='latin1', dtype=str)
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
    return None

def process_file(file_path: str) -> pd.DataFrame | None:
    df = read_file(file_path)
    if df is None or df.empty:
        print(f"File {file_path} is empty or could not be read, skipping.")
        return None

    file_name = os.path.splitext(os.path.basename(file_path))[0]
    # Used for debugging - prints what file is being processed
    # print(f"Processing file: {file_path}")
    # na_cols = df.columns[df.isna().all()].tolist()
    # if na_cols:
    #     print(f"File {file_path} has all-NA columns: {na_cols}")

    if 'BI_ARTransactions' in file_name or 'SQL_ARTransactions' in file_name:
        df['group'] = lookup_value_in_map(file_name, GROUP_MAP)
        df['region'] = lookup_value_in_map(file_name, REGION_MAP)

    if 'pinned_notes' in df.columns:
        df['pinned_notes'] = df['pinned_notes'].astype(str).str.slice(0, 1000)

    df['tenant'] = lookup_value_in_map(file_name, TENANT_MAP)
    if df['tenant'] is None or pd.isna(df['tenant']).any():
        raise ValueError(f"Tenant could not be determined for file: {file_name}")

    # Add file_path to DataFrame for debugging
    df['file_path'] = file_path
    return df.iloc[:-1]

def combine_files(report_type: str, environment: str) -> pd.DataFrame | None:
    if environment == "DEV":
        source_dir = r'C:\biautomation\ETL\Extract\ServiceTitan'
    else:
        source_dir = r'C:\Users\BIAutomation\OneDrive - Air Pros USA\biautomation\ETL\Extract\ServiceTitan_Hold'

    patterns = [os.path.join(source_dir, f"{report_type}*.xlsx"),
                os.path.join(source_dir, f"{report_type}*.csv")]

    files = []
    for pattern in patterns:
        files.extend(glob.glob(pattern))

    if not files:
        print(f"No files found for report type '{report_type}'.")
        return None

    with ThreadPoolExecutor(max_workers=4) as executor:  # Increased max_workers
        dfs = list(executor.map(process_file, files))

    # Filter out invalid DataFrames (None, empty, or all-NA columns only)
    valid_dfs = [
        df for df in dfs
        if df is not None and not df.empty and df.dropna(axis=1, how='all').shape[1] > 0
    ]

    if not valid_dfs:
        raise ValueError(f"No valid DataFrames to concatenate for report type '{report_type}'.")

    # Get all unique columns across valid DataFrames
    all_columns = set().union(*[df.columns for df in valid_dfs])
    
    final_dfs = []
    for i, df in enumerate(valid_dfs):
        # Create a copy to avoid modifying the original
        df = df.copy()
        # Drop all-NA columns
        df = df.dropna(axis=1, how='all')
        # Add missing columns with pd.NA
        missing_cols = all_columns - set(df.columns)
        for col in missing_cols:
            df.loc[:, col] = pd.NA
        # Drop all-NA columns again after adding missing ones
        df = df.dropna(axis=1, how='all')
        # Only include DataFrame if it has non-NA columns
        if df.shape[1] > 0:
            # Debug: Check for all-NA columns
            na_cols = df.columns[df.isna().all()].tolist()
            if na_cols:
                print(f"DataFrame {i} for report '{report_type}' has all-NA columns: {na_cols}")
            final_dfs.append(df[sorted(all_columns & set(df.columns))])
        else:
            print(f"Skipping DataFrame {i} for report '{report_type}' with no non-NA columns.")

    if not final_dfs:
        raise ValueError(f"No DataFrames with non-NA columns to concatenate for report type '{report_type}'.")

    # Used for Debugging - Prints how many dataframes have NaN columns
    # print(f"Combining {len(final_dfs)} non-empty DataFrames for report '{report_type}'...")
    combined_df = pd.concat(final_dfs, ignore_index=True, sort=False)
    combined_df.index.names = ['key']

    return combined_df

if __name__ == "__main__":
    pass