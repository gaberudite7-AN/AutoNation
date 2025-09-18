import subprocess
import traceback

try:
    subprocess.run([
        r"C:\Development\.venv\Scripts\python.exe",
        r"C:\Development\.venv\Scripts\Python_Scripts\Docker\UrbanScience_Update.py"
    ], check=True)
except Exception as e:
    with open(r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\error_log.txt", "a") as f:
        f.write(traceback.format_exc())