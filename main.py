import logging
import os
import re
import time

import pandas as pd
import pystray
import win32con
import win32gui
from PIL import Image
from pystray import MenuItem
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sys_log.log")
logging.basicConfig(filename=log_file, level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

WATCH_FOLDER = os.path.dirname(os.path.abspath(__file__))
SOURCE_FOLDER = os.path.join(WATCH_FOLDER, "source")
COMPLETED_FOLDER = os.path.join(WATCH_FOLDER, "completed")
REPORTS_FOLDER = os.path.join(WATCH_FOLDER, "reports")
atm_section_pattern = re.compile(r"UPTIME TOTALS FOR ATM (\d+)")


def minimize_console():
    win = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(win, win32con.SW_HIDE)


def log_message(message):
    logging.info(message)


for folder in [SOURCE_FOLDER, COMPLETED_FOLDER, REPORTS_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder, mode=0o755)


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
            continue

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
                f"\nProcessing new source file.\nSOURCE FILE NAME  >> {os.path.basename(file_path)}\nSOURCES HOME >> {SOURCE_FOLDER}\n")

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
                    f"Archeved processed file: \nARCHIVED FILE NAME  >> {os.path.basename(file_path)}\nARCHIVES HOME >> {COMPLETED_FOLDER}\n")
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


# minimize_console()


def menu_exit_callback(icon, item):
    log_message("ATM Reporter Closed.")
    icon.stop()
    os._exit(0)


def show_reports_folder():
    os.startfile(REPORTS_FOLDER)


def setup_tray_icon():
    try:
        icon_path = 'std_white.ico'
        image = Image.open(icon_path)
        title = "ATM Reporter"
        menu = (
            MenuItem('Reports', show_reports_folder),
            MenuItem('Exit', menu_exit_callback)
        )
        icon = pystray.Icon("Mike", image, "ATM Reporter", menu)
        return icon
    except Exception as e:
        log_message("Error setting up tray icon:")
        log_message(str(e))
        return None


if __name__ == "__main__":
    icon = setup_tray_icon()
    if icon is None:
        log_message("Exiting due to tray icon setup error.")
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

        observer.join()
