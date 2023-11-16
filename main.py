import subprocess
import csv
import time
import serial
from datetime import datetime
import os

def get_current_date():
    return datetime.now().strftime('%Y-%m-%d')

# CSV file configuration
csv_file_path = f'C:\\USB-LOG\\usb_events_log{get_current_date()}.csv'
csv_header = ['Event', 'Timestamp', 'Device Caption', 'Device ID']

def log_usb_event(event, device_info):
    with open(csv_file_path, mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([event, time.strftime('%Y-%m-%d %H:%M:%S')] + device_info)

def log_usb_devices():
    try:
        # Run the wmic command to list USB devices
        result = subprocess.run(['wmic', 'path', 'win32_pnpentity', 'get', 'caption,deviceid'],
                                capture_output=True, text=True, check=True)

        # Extract device information from the output
        lines = result.stdout.strip().split('\n')
        devices_info = [line.split() for line in lines[1:] if line]

        # Check if CSV file exists
        if not os.path.exists(csv_file_path):
            # Set up CSV file with header
            with open(csv_file_path, mode='w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(csv_header)

        # Log USB events
        while True:
            new_result = subprocess.run(['wmic', 'path', 'win32_pnpentity', 'get', 'caption,deviceid'],
                                        capture_output=True, text=True, check=True)

            new_lines = new_result.stdout.strip().split('\n')
            new_devices_info = [line.split() for line in new_lines[1:] if line]

            # Compare current devices with previous devices to find changes
            for device_info in new_devices_info:
                if device_info not in devices_info:
                    log_usb_event('Plugged', device_info)
            for device_info in devices_info:
                if device_info not in new_devices_info:
                    log_usb_event('Unplugged', device_info)

            devices_info = new_devices_info
            time.sleep(1)  # Log events every second

    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    log_usb_devices()
