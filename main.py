import configparser
import itertools
import logging
import os
import re
import sys
import time

import pandas as pd
import pystray
import win32con
import win32gui
from PIL import Image
from pystray import MenuItem
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

log_file = "sys_log.log"
logging.basicConfig(filename=log_file, level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Get the path to the INI file
ini_path = 'std-settings.ini'

# Create a ConfigParser instance
config = configparser.ConfigParser()

# Check if the INI file exists
if os.path.exists(ini_path):
    # Read the INI file
    config.read(ini_path)
else:
    # Create default folder names
    sources_folder_name = 'source'
    reports_folder_name = 'reports'
    completed_folder_name = 'completed'

    # Create default folder paths based on the current working directory
    WATCH_FOLDER = sources_folder_name
    REPORTS_FOLDER = reports_folder_name
    COMPLETED_FOLDER = completed_folder_name

    # Create the default directories if they don't exist
    os.makedirs(WATCH_FOLDER, exist_ok=True)
    os.makedirs(REPORTS_FOLDER, exist_ok=True)
    os.makedirs(COMPLETED_FOLDER, exist_ok=True)

    # Set the folder paths in the ConfigParser instance
    config['Folders'] = {
        'InputFolder': WATCH_FOLDER,
        'OutputFolder': REPORTS_FOLDER,
        'CompletedFolder': COMPLETED_FOLDER,
    }

    # Write the ConfigParser instance to the INI file
    with open(ini_path, 'w') as config_file:
        config.write(config_file)

# Read the folder paths from the INI file
SOURCE_FOLDER = config.get('Folders', 'InputFolder')
REPORTS_FOLDER = config.get('Folders', 'OutputFolder')
COMPLETED_FOLDER = config.get('Folders', 'CompletedFolder')

# Create the folders if they don't exist
os.makedirs(SOURCE_FOLDER, exist_ok=True)
os.makedirs(REPORTS_FOLDER, exist_ok=True)
os.makedirs(COMPLETED_FOLDER, exist_ok=True)

atm_section_pattern = re.compile(r"UPTIME TOTALS FOR ATM (\d+)")


def minimize_console():
    win = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(win, win32con.SW_HIDE)


def log_message(message):
    logging.info(message)


def parse_uptime_data(data):
    uptime_info = []
    lines = data.strip().split('\n')

    atm_id = None
    uptime_percent = None
    t_downtime_percent = None
    total_downtime = "00:00.00"
    downtime_reasons = []
    capturing_total_downtime = False

    for index, line in enumerate(lines):
        match = atm_section_pattern.search(line)
        if match:
            if atm_id is not None:
                if not downtime_reasons:
                    downtime_reasons.append(f"No Recorded Reason For ATM-{atm_id}")
                downtime_info = ', '.join(downtime_reasons)
                uptime_info.append([atm_id, uptime_percent, t_downtime_percent, total_downtime, downtime_info])

            atm_id = match.group(1)
            uptime_percent = None
            t_downtime_percent = None
            total_downtime = "00:00.00"
            downtime_reasons = []

        if "Uptime Adjustment" in line:
            capturing_total_downtime = True
            line_index = index
        elif capturing_total_downtime and index == line_index + 2:
            columns = line.split()
            total_downtime = columns[-2]
            t_downtime_percent = columns[-1]
            capturing_total_downtime = False

        if "Online" in line:
            columns = line.split()
            uptime_percent = columns[3]
        if "ACCUMULATED UPTIME TOTALS FOR *ALL* ATMS" in line:
            break

        if "No totals received from ATM" in line:
            downtime_reasons.append(f"No Totals Received For ATM-{atm_id}")
            uptime_percent = ""

        for reason in [
            "Closed from Sparrow", "Waiting For Comms", "Supervisor",
            "Diagnostics", "Re-entry", "Downloading HCF", "Downloading Other", "Hardware Fault",
            "Power Fail Recovery", "Uptime Adjustment"
        ]:
            if reason in line:
                columns = line.split()
                downtime_percent = columns[-1]
                downtime_reason = f'{reason} ({downtime_percent}%)' if columns[-2] != "0:00.00" else ""
                if downtime_reason:
                    downtime_reasons.append(downtime_reason)

    if atm_id is not None and downtime_reasons:
        downtime_info = ', '.join(downtime_reasons)
        uptime_info.append([atm_id, uptime_percent, t_downtime_percent, total_downtime, downtime_info])

    return uptime_info


def process_file(file_path):
    try:
        if file_path.endswith(".spa"):
            log_message(
                f"########################## FILE IN TASK: {os.path.basename(file_path)} >> STARTED << ##############################")
            log_message(
                f"Processing new source file.\nSOURCE FILE NAME  >> {os.path.basename(file_path)}\nSOURCES HOME >> {SOURCE_FOLDER}\n")

            with open(file_path, "r") as file:
                content = file.read()

            uptime_info = parse_uptime_data(content)

            if uptime_info:
                df = pd.DataFrame(uptime_info, columns=["ATM ID", "Uptime %", "Downtime %", "Total Downtime(hh:mm.ss)",
                                                        "Downtime Reasons"])
                excel_file = os.path.splitext(os.path.basename(file_path))[0] + ".xlsx"
                excel_path = os.path.join(REPORTS_FOLDER, excel_file)
                df.to_excel(excel_path, index=False)
                log_message(f"Report generated.\nREPORT FILE NAME  >> {excel_file}\nREPORTS HOME >> {REPORTS_FOLDER}\n")
                file.close()
                completed_path = os.path.join(COMPLETED_FOLDER, os.path.basename(file_path))

                if os.path.exists(completed_path):
                    os.remove(completed_path)

                os.rename(file_path, completed_path)
                log_message(
                    f"Archived processed file: \nARCHIVED FILE NAME  >> {os.path.basename(file_path)}\nARCHIVES HOME >> {COMPLETED_FOLDER}\n")
                log_message(
                    f"########################## FILE IN TASK: {os.path.basename(file_path)} >> COMPLETED << ##############################\n")

    except Exception as e:
        log_message(
            f"########################## FILE IN TASK: {os.path.basename(file_path)} >> STARTED << ##############################")
        log_message(
            f"\nProcessing new source file.\nSOURCE FILE NAME  >> {os.path.basename(file_path)}\nSOURCES HOME >> {SOURCE_FOLDER}\n")
        log_message(f"Error processing file: {os.path.basename(file_path)}")
        log_message(f"Error message: {str(e)}")
        log_message(
            f"########################## FILE IN TASK: {os.path.basename(file_path)} >> COMPLETED << ##############################\n")


existing_files = [os.path.join(SOURCE_FOLDER, filename) for filename in os.listdir(SOURCE_FOLDER)]

for existing_file in existing_files:
    process_file(existing_file)

for folder in [COMPLETED_FOLDER, REPORTS_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)


class MyHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.src_path.endswith(".spa"):
            process_file(event.src_path)


minimize_console()


def is_another_instance_running():
    # Create a lock file to ensure only one instance can run
    lock_file = "program.lock"
    if os.path.exists(lock_file):
        return True
    else:
        with open(lock_file, "w") as f:
            f.write(str(os.getpid()))
        return False


def menu_exit_callback(icon, item):
    log_message("ATM Reporter Closed.")
    icon.stop()
    os.remove("program.lock")
    os._exit(0)


def show_reports_folder():
    os.startfile(REPORTS_FOLDER)


def show_log_file():
    os.startfile(log_file)


def setup_tray_icon():
    try:
        icon_path = 'std_white.ico'
        image = Image.open(icon_path)
        title = "ATM Reporter"
        menu = (
            MenuItem('Reports', show_reports_folder),
            MenuItem('View Log', show_log_file),
            MenuItem('Exit', menu_exit_callback)
        )
        icon = pystray.Icon("Mike", image, "ATM Reporter", menu)
        return icon
    except Exception as e:
        log_message("Error setting up tray icon:")
        log_message(str(e))
        return None


def main():
    if is_another_instance_running():
        log_message("Another instance is already running. Exiting.")
        os._exit(0)
        return


if __name__ == "__main__":
    main()
    icon = setup_tray_icon()
    if icon is None:
        log_message("Exiting due to tray icon setup error.")
        os.remove("program.lock")
    else:
        observer = Observer()
        event_handler = MyHandler()
        observer.schedule(event_handler, path=SOURCE_FOLDER, recursive=False)
        log_message(f"ATM Reporter Started >> [SOURCES FOLDER >> {SOURCE_FOLDER}")
        observer.start()
        icon.run()
        try:
            while True:
                time.sleep(1)
                icon.update_menu()
        except KeyboardInterrupt:
            observer.stop()
            icon.stop()
            os.remove("program.lock")

        observer.join()
