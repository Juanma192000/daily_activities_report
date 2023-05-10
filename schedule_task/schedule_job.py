import schedule
import time
import subprocess
import win32gui, win32con
from pathlib import Path

is_task_completed=False
def run_script():
    path_file=f"{Path(__file__).parent.parent}/src/main.py"
    path_file=path_file.replace("\\", "/")    
    subprocess.run(["python", path_file])
    global is_task_completed
    is_task_completed=True
# Define the time to run the script
scheduled_time = "13:30"  # Replace with the desired time in HH:MM format
schedule.every().day.at(scheduled_time).do(run_script)
# Keep the script running indefinitely
print("Starting shchedule execution...")
the_program_to_hide = win32gui.GetForegroundWindow()
win32gui.ShowWindow(the_program_to_hide , win32con.SW_HIDE)
while is_task_completed is not True:
    schedule.run_pending()
    time.sleep(1)
