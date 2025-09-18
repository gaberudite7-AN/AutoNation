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
import Used_Car_Program
import Refresh_MarketShareMarketing
import time
import Discounted_Inventory_Tracking
import Brand_President

# Configure logging to suppress APScheduler's warning
logging.getLogger('apscheduler').setLevel(logging.ERROR)  # Suppress APScheduler warnings

# Set-up the Scheduler (Global scheduler declaration)
sched = BlockingScheduler()

def ScheduledEvents():
    try:
        now = dt.datetime.now()
        print(now.strftime('%Y-%m-%d %H:%M:%S'))
        # Run Transform, Load, and Garbage Collection between 3:00 AM and 7:00 PM
        if now.hour == 8 and now.minute == 30:
            # Run functions to update latest 
            print("Daily Sales script Ran in Scheduler")
            time.sleep(10)
            # Refresh MarketShare file
            Refresh_MarketShareMarketing.Refresh_MarketShare()
            # Update Daily file in Market Share
            #Web_Scraping.Web_Scraping_UrbanScienceHistorics.Move_Current_to_Historics()
            #Web_Scraping.Web_Scraping_UrbanScienceHistorics.Update_Daily_UrbanScience()
            time.sleep(5)
            if now.weekday() == 0: # Monday at 8:30am
                # Update Discounted Inventory Tracking
                Discounted_Inventory_Tracking.Discounted_Inventory_Tracking_Update()
                # Create Email Template/Send Email
                Discounted_Inventory_Tracking.Discounted_Inventory_Tracking_Email()                
                Allocation_Tracker.Allocation_Tracker_Update()
            if now.weekday() == 4: # Friday
                Weekly_Report.Weekly_Data_Update() 
            print("Performing Garbage Collection...")
            gc.collect()
        if now.hour == 16 and now.minute == 30:
           if now.weekday() == 0: # Monday at 4:30pm
                Brand_President.Update_DOC_AND_BUDGET_file()
                Brand_President.Update_BPU_File()
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