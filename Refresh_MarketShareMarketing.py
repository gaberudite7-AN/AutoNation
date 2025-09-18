# Imports
import xlwings as xw
import os
import time
import psutil

# Begin timer
start_time = time.time()

# Run with low priority ( will allow script to run in background and yield CPU to other apps)
try:
    p = psutil.Process(os.getpid())
    p.nice(psutil.IDLE_PRIORITY_CLASS) # Windows only
except Exception as e:
    print(f"Could not set low priority: {e}")

def Refresh_MarketShare():

    # Open file and process macro/Sql
    MarketShare_MarketingFile = r"C:\Users\BesadaG\OneDrive - AutoNation\Cutie Baez, Aili's files - Market Share\Market Share - For Marketing.xlsm"    
    # MarketShare_MarketingFile = r"C:\Users\cutiebaeza\OneDrive - AutoNation\Cutie Baez, Aili's files - Market Share\Market Share - For Marketing.xlsm"    
   
    app = xw.App(visible=True)
    app.display_alerts = True
    app.screen_updating = True # Optional: improve performance  
    MarketShare_Marketingwb = app.books.open(MarketShare_MarketingFile, update_links=False)
    # Run Macro
    Run_Macro = MarketShare_Marketingwb.macro("Refresh_MarketShare")
    Run_Macro()
    MarketShare_Marketingwb.save()
    time.sleep(60)

    # Save and close the excel document(s)    
    if MarketShare_Marketingwb:
        MarketShare_Marketingwb.close()
    if app:
        app.quit()

    # End Timer
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Updated Market share...completed in {elapsed_time:.2f} seconds")

    return

#run function
if __name__ == '__main__':
    
    Refresh_MarketShare()