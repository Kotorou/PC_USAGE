from win32gui import GetForegroundWindow
import psutil
import time
import win32process
from datetime import datetime
import requests
from tabulate import tabulate

process_time = {}
timestamp = {}
line_token = '***********************************'
line_url = 'https://notify-api.line.me/api/notify'
header = {'Content-type': 'application/x-www-form-urlencoded', 'Authorization': 'Bearer ' + line_token}

# Define the target notification time
target_time = "12:00:00"

while True:
    try:
        # Get the current time
        current_time = datetime.now().strftime("%H:%M:%S")

        # Get the process ID of the foreground window
        current_window_handle = GetForegroundWindow()
        pid = win32process.GetWindowThreadProcessId(current_window_handle)[1]
        current_app = psutil.Process(pid).name().replace(".exe", "")

        # Update timestamp for the current application
        if current_app:
            if current_app not in process_time:
                process_time[current_app] = 0
            if current_app not in timestamp:
                timestamp[current_app] = int(time.time())
            else:
                elapsed_time = int(time.time()) - timestamp[current_app]
                process_time[current_app] += elapsed_time
                timestamp[current_app] = int(time.time())

        # Check if the current time matches the target time
        if current_time == target_time:
            # Prepare data for the table
            table_data = []
            for app, elapsed_time in process_time.items():
                hours = elapsed_time // 3600
                minutes = (elapsed_time % 3600) // 60
                seconds = elapsed_time % 60
                time_str = f"{hours:02}:{minutes:02}:{seconds:02}"
                table_data.append([app, time_str])

            # Generate the table
            table = tabulate(table_data, headers=["Application", "Time Spent"], tablefmt="grid")

            # Send the notification
            requests.post(line_url, headers=header, data={'message': "\n" + table + "\n"})
        else:
            # Print the time spent on the current application
            hours = process_time.get(current_app, 0) // 3600
            minutes = (process_time.get(current_app, 0) % 3600) // 60
            seconds = process_time.get(current_app, 0) % 60
            print(f"{current_app}: {hours} hours, {minutes} minutes, {seconds} seconds")

        # Sleep for a while before checking again
        time.sleep(1)  # Sleep for 1 second

    except psutil.NoSuchProcess:
        pass

    except Exception as e:
        print(f"Error: {e}")
