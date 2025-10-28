# Imports
from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.events import EVENT_JOB_MAX_INSTANCES
import datetime as dt
import gc
import sys
import logging
import os
import importlib.util
import Update_UrbanScience
import Scripts.Python_Scripts.Archive.Allocation_Tracker as Allocation_Tracker
import Weekly_Report
import Prevent_Logoff
import Refresh_MarketShareMarketing
import time
import Discounted_Inventory_Tracking
import Brand_President
import multiprocessing
import Low_PVR
import MarketShare_Images_Scrape
import Update_UrbanScience
import Industry_Pull
from Allocation_Tracker import AllocationTracker



# Configure logging to suppress APScheduler's warning
logging.getLogger('apscheduler').setLevel(logging.ERROR)  # Suppress APScheduler warnings

# Set-up the Scheduler (Global scheduler declaration)
sched = BlockingScheduler()

def ScheduledEvents():
    try:
        now = dt.datetime.now()
        print(now.strftime('%Y-%m-%d %H:%M:%S'))
        # Run Transform, Load, and Garbage Collection between 3:00 AM and 7:00 PM
        if now.hour == 7 and now.minute == 0:
            # Update Urban Science Data
            Update_UrbanScience.Update_Historicals()
            time.sleep(2)
            Update_UrbanScience.Update_Daily_UrbanScience()
            time.sleep(2)
            
            """MONDAY TASKS"""
            if now.weekday() == 0: # Monday
                # @ 730AM on MONDAY, RUN DISCOUNTED INVENTORY TRACKING
                if now.hour == 7 and now.minute == 30:
                    # Update Discounted Inventory Tracking
                    Discounted_Inventory_Tracking.Discounted_Inventory_Tracking_Update()
                    # Create Email Template/Send Email
                    Discounted_Inventory_Tracking.Discounted_Inventory_Tracking_Email()                
                # @ 745AM on MONDAY, RUN ALLOCATION TRACKER
                if now.hour == 7 and now.minute == 45:
                    tracker = AllocationTracker(r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Allocation_Tracker')
                    tracker.run_allocation_tracker()
                # @ 8AM on MONDAY, RUN LOW PVR UPDATE
                if now.hour == 8 and now.minute == 0:
                    Low_PVR.Low_PVR_Update()
                if now.hour == 16 and now.minute == 30:
                        Brand_President.Update_DOC_AND_BUDGET_file()
                        Brand_President.Update_BPU_File()
            
            """WEDNESDAY TASKS"""
            if now.weekday() == 2: # Wednesday
                # @ 630AM on WEDNESDAY, SCRAPE LATEST MAKE FILE FROM URBAN SCIENCE
                if now.hour == 6 and now.minute == 30:
                    # Update Make File
                    Update_UrbanScience.Update_Make_File()
                    time.sleep(5)
                    # Create Email Template/Send Email
                    Industry_Pull.Update_Industry_UrbanScience()


            """FRIDAY TASKS"""            
            if now.weekday() == 4: # Friday
                if now.hour == 8 and now.minute == 30:
                    MarketShare_Images_Scrape.MarketShare_Images_Scrape_Update()
                if now.hour == 7 and now.minute == 15:
                    Weekly_Report.Process_Daily_Sales_File()    
                    Weekly_Report.Download_PWB()
                    Weekly_Report.Update_PWB_Data()
                    Weekly_Report.Weekly_Data_Update()
            print("Performing Garbage Collection...")
            gc.collect()
    except Exception as e:
        print(f"Error occurred: {e}")

# Event listener for skipped jobs
def on_job_skipped(event):
    print("Running...")

if __name__ == '__main__':
    now = dt.datetime.now()
    print("Starting AutoNation Scheduler")
    # Add event listener for skipped jobs
    sched.add_listener(on_job_skipped, EVENT_JOB_MAX_INSTANCES)
    # Schedule the job
    sched.add_job(ScheduledEvents, 'interval', minutes=1, max_instances=1)
    sched.start()