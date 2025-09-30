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
from Allocation_Tracker import AllocationTracker

keep_awake_process = None

def start_keep_awake():
    global keep_awake_process
    if keep_awake_process is None or not keep_awake_process.is_alive():
        keep_awake_process = multiprocessing.Process(target=Prevent_Logoff.keep_awake)
        keep_awake_process.start()
        print("Keep-awake process started.")

def stop_keep_awake():
    global keep_awake_process
    if keep_awake_process and keep_awake_process.is_alive():
        keep_awake_process.terminate()
        keep_awake_process.join()
        print("Keep-awake process stopped.")



# Configure logging to suppress APScheduler's warning
logging.getLogger('apscheduler').setLevel(logging.ERROR)  # Suppress APScheduler warnings

# Set-up the Scheduler (Global scheduler declaration)
sched = BlockingScheduler()

def ScheduledEvents():
    try:
        now = dt.datetime.now()
        print(now.strftime('%Y-%m-%d %H:%M:%S'))
        stop_keep_awake()
        print("Stopping keep awake process to allow script to run.")
        # Run Transform, Load, and Garbage Collection between 3:00 AM and 7:00 PM
        if now.hour == 7 and now.minute == 0:
            # Update Urban Science Data
            Update_UrbanScience.Update_Historicals()
            time.sleep(2)
            Update_UrbanScience.Update_Daily_UrbanScience()
            time.sleep(2)
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
            # @ 730AM on FRIDAY, RUN WEEKLY REPORT
            if now.hour == 7 and now.minute == 30:
                if now.weekday() == 4: # Friday
                    Weekly_Report.Weekly_Data_Update() 
            print("Performing Garbage Collection...")
            gc.collect()
        if now.hour == 16 and now.minute == 30:
           if now.weekday() == 0: # Monday at 4:30pm
                Brand_President.Update_DOC_AND_BUDGET_file()
                Brand_President.Update_BPU_File()
        # Restart the keep awake process after jobs finish
        start_keep_awake()
        print("Restarting keep awake process.")
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