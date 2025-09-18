import pyautogui
import time

def keep_awake():
    while True:
        pyautogui.moveRel(0, 50, duration=1)  # Move mouse slightly
        pyautogui.moveRel(0, -50, duration=1)
        time.sleep(2)  # Wait 5 minutes
        pyautogui.moveRel(0, 50, duration=1)  # Move mouse slightly
        pyautogui.moveRel(0, -50, duration=1)
        time.sleep(2)  # Wait 5 minutes
        pyautogui.moveRel(0, 50, duration=1)  # Move mouse slightly
        pyautogui.moveRel(0, -50, duration=1)
        time.sleep(2)  # Wait 5 minutes
        pyautogui.moveRel(0, 50, duration=1)  # Move mouse slightly
        pyautogui.moveRel(0, -50, duration=1)
        time.sleep(10)  # Wait 5 minutes
        print("Moved")

if __name__ == "__main__":
    keep_awake()