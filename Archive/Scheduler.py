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
#import Python_Scripts.Allocation_Tracker

# Configure logging to suppress APScheduler's warning
logging.getLogger('apscheduler').setLevel(logging.ERROR)  # Suppress APScheduler warnings

# Set-up the Scheduler (Global scheduler declaration)
sched = BlockingScheduler()

def ScheduledEvents():
    try:
        now = dt.datetime.now()
        print(now.strftime('%Y-%m-%d %H:%M:%S'))
        # Run Transform, Load, and Garbage Collection between 3:00 AM and 7:00 PM
        if now.hour in range(3, 20) and now.minute == 1:
            # Run functions
            Update_UrbanScience.Update_Daily_UrbanScience()
            #Python_Scripts.Allocation_Tracker.Allocation_Tracker_Update()
            print("Performing Garbage Collection...")
            gc.collect()

    except Exception as e:
        print(f"Error occurred: {e}")

# Event listener for skipped jobs
def on_job_skipped(event):
    print("Running...")

if __name__ == '__main__':
    now = dt.datetime.now()
    # Add event listener for skipped jobs
    sched.add_listener(on_job_skipped, EVENT_JOB_MAX_INSTANCES)
    # Schedule the job
    sched.add_job(ScheduledEvents, 'interval', minutes=1, max_instances=1)
    sched.start()