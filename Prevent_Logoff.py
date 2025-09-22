# Create the keep_awake.py script file with the required content
import pyautogui
import time
import random

def keep_awake():
    print("Keep-awake script started. Press Ctrl+C to stop.")
    while True:
        # Simulate small random mouse movement
        x_move = random.randint(-100, 100)
        y_move = random.randint(-100, 100)
        pyautogui.moveRel(x_move, y_move, duration=5)
        pyautogui.moveRel(-x_move, -y_move, duration=5)

        # Simulate a harmless key press
        pyautogui.press('shift')

        print("Activity simulated. Waiting 5 seconds...")
        time.sleep(5)  # Wait for 5 seconds

# Run the function
if __name__ == "__main__":
    keep_awake()