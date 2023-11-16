import subprocess
import csv
import time
import serial
import json
import pyautogui
from datetime import datetime
import os

def get_current_date():
    return datetime.now().strftime('%Y-%m-%d')

ser = serial.Serial('COM4', 9600)

def receive_and_decode():
    try:
        while True:
            # Read a line from the serial port
            data_bytes = ser.readline()
            try:
                # Attempt to decode the byte string using a more permissive approach
                data_str = data_bytes.decode('utf-8', errors='ignore')
                # or use errors='ignore' to skip invalid characters
                # Attempt to load the JSON data
                data_dict = json.loads(data_str)
                # Accessing individual elements (uncomment this section if needed)
                student_type = data_dict.get('type', '').split(" ")
                uid = data_dict.get('uid', '').split(" ")
                name = data_dict.get('name', '').split(" ")
                enrolment = data_dict.get('enrolment', '').split(" ")
                stream = data_dict.get('stream', '').split(" ")
                year = data_dict.get('year', '').split(" ")
                boarding = data_dict.get('boarding', '').split(" ")
                device_id = data_dict.get('device_id', '').split(" ")
                # Print or use the decoded values
                print("Student Type:", student_type[0])
                print("UID:", uid[0])
                print("Name:", name[0])
                print("Enrolment:", enrolment[0])
                print("Stream:", stream[0])
                print("Year:", year[0])
                print("Boarding:", boarding[0])
                if(student_type[0] == "student"):
                    pyautogui.press('enter')
                    pyautogui.press('enter')
                    pyautogui.typewrite('2003')
                # Return the decoded serial data
                value_send = [student_type[0], uid[0], name[0], enrolment[0], stream[0], year[0], boarding[0], device_id[0]]
                return value_send
            except json.JSONDecodeError as e:
                print("Error decoding JSON:", e)
    except serial.SerialException as se:
        print(f"Serial port error: {se}")
    finally:
        ser.close()

# CSV file configuration
csv_file_path = f'C:\\LOG\\usb_events_log{get_current_date()}.csv'
csv_header = ['Event', 'Timestamp', 'Device Caption', 'Type', 'UID', 'Name', 'Enrolment No', 'Stream', 'Year', 'Boarding', 'Device ID']

def log_usb_event(event, device_info, serial_data):
    with open(csv_file_path, mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([event, time.strftime('%Y-%m-%d %H:%M:%S')] + device_info + serial_data)

def log_usb_devices(serial_data):
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
                    log_usb_event('Plugged', device_info, serial_data)
            for device_info in devices_info:
                if device_info not in new_devices_info:
                    log_usb_event('Unplugged', device_info, serial_data)

            devices_info = new_devices_info
            time.sleep(1)  # Log events every second

    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    try:
        while True:
            serial_data = receive_and_decode()
            log_usb_devices(serial_data)
    except KeyboardInterrupt:
        # Handle keyboard interrupt (Ctrl+C) to gracefully exit the program
        ser.close()
        print("Program terminated by user.")
