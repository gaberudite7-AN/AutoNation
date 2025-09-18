# %%
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta


def Update_Historic_MS():
    
    today = datetime.today()
    Current_Month = today.month
    Last_Month = (today - relativedelta(months=1)).month
    Current_Year = today.year

    print(f"Updating with current month as {Current_Month}, Last month as {Last_Month}, and Current Year as {Current_Year}")

    #---------------------------------Load files to df--------------------------------------------------

    # File paths
    file_path = r"W:\Corporate\Inventory\Urban Science\For Append\AutoNation_SalesFile_NationalSales.txt"
    file_path2 = r"W:\Corporate\Inventory\Urban Science\For Append\AutoNation_SalesFile_NationalSales_Historical.txt"

    # Load files

    df_curr = pd.read_csv(file_path, delimiter=',', quotechar='"')

    df_hist = pd.read_csv(file_path2, delimiter=',', quotechar='"')

    print("Files loaded")

    #---------------------------------Transformation--------------------------------------------------
    print("Transformation ongoing")
    df_filtered = df_curr[(df_curr['SALE_MONTH'].isin([Last_Month, Current_Month])) & (df_curr['SALE_YEAR'] == Current_Year)]


    df_combined = pd.concat([df_hist, df_filtered], ignore_index=True)

    #print(df_combined['SALE_MONTH'].unique())

    #---------------------------------Save output file--------------------------------------------------
    print("Saving file. Please wait")
    output_path = fr"W:\Corporate\Inventory\Urban Science\For Append\AppendedHist_{Current_Year}{Current_Month}.txt"

    df_combined.to_csv(output_path, sep=',', index=False)

    print("File savedat: {output_path}")
    
    return

#run function
if __name__ == '__main__':
    
    Update_Historic_MS()