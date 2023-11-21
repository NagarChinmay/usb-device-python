import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import subprocess
import csv
import time
import serial
import json
import pyautogui
from datetime import datetime
import os
import sys
import logging
# Disable PyAutoGUI fail-safe
pyautogui.FAILSAFE = False

class SerialService(win32serviceutil.ServiceFramework):
    _svc_name_ = '(PBL) Service USB To LOG WITH RFID Serial v1'
    _svc_display_name_ = '(PBL) Service USB To LOG with RFID Serial v1'
    _svc_description_ = 'Service for log all usb device with rfid'
    
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.is_alive = True
        self.ser = serial.Serial('COM4', 9600)  # Adjust the COM port and baud rate as needed
        self.csv_file_path = f'C:\\LOG\\usb_events_log_{self.get_current_date()}.csv'  # Include current date in the CSV file path
        self.csv_header = ['Event', 'Timestamp', 'Device Caption', 'Type', 'UID', 'Name', 'Enrolment No', 'Stream', 'Year', 'Boarding', 'Device ID']

        # Configure logging
        logging.basicConfig(filename='C:\\debug(pbl)\\your_log_file.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

    def get_current_date(self):
        return datetime.now().strftime('%Y-%m-%d')

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False
        self.ser.close()  # Close the serial port when stopping the service

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def main(self):
        try:
            while self.is_alive:
                serial_data = self.receive_and_decode()
                self.log_usb_devices(serial_data)
        except serial.SerialException as se:
            logging.error(f"Serial port error: {se}")
        finally:
            self.ser.close()

    def receive_and_decode(self):
        try:
            while True:
                # Read a line from the serial port
                data_bytes = self.ser.readline()
                try:
                    # Attempt to decode the byte string using a more permissive approach
                    data_str = data_bytes.decode('utf-8', errors='ignore')
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
                    logging.debug("decoding JSON: %s",value_send)
                    return value_send
                except json.JSONDecodeError as e:
                    logging.error("Error decoding JSON: %s", e)
        except (serial.SerialException, Exception) as se:  # Catch more general exception types
            logging.error("Serial port error or other exception: %s", se)
        finally:
            self.ser.close()

    def log_usb_event(self, event, device_info, serial_data):
        with open(self.csv_file_path, mode='a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow([event, time.strftime('%Y-%m-%d %H:%M:%S')] + device_info + serial_data)
            logging.debug("log to usb called")

    def log_usb_devices(self, serial_data):
        try:
            # Run the wmic command to list USB devices
            result = subprocess.run(['wmic', 'path', 'win32_pnpentity', 'get', 'caption,deviceid'],
                                    capture_output=True, text=True, check=True)

            # Extract device information from the output
            lines = result.stdout.strip().split('\n')
            devices_info = [line.split() for line in lines[1:] if line]

            # Check if CSV file exists
            if not os.path.exists(self.csv_file_path):
                # Set up CSV file with header
                with open(self.csv_file_path, mode='w', newline='') as file:
                    writer = csv.writer(file)
                    writer.writerow(self.csv_header)

            # Log USB events
            while True:
                new_result = subprocess.run(['wmic', 'path', 'win32_pnpentity', 'get', 'caption,deviceid'],
                                            capture_output=True, text=True, check=True)

                new_lines = new_result.stdout.strip().split('\n')
                new_devices_info = [line.split() for line in new_lines[1:] if line]

                # Compare current devices with previous devices to find changes
                for device_info in new_devices_info:
                    if device_info not in devices_info:
                        self.log_usb_event('Plugged', device_info, serial_data)
                for device_info in devices_info:
                    if device_info not in new_devices_info:
                        self.log_usb_event('Unplugged', device_info, serial_data)

                devices_info = new_devices_info
                time.sleep(1)  # Log events every second

        except subprocess.CalledProcessError as e:
            logging.error("Error: %s", e)

if __name__ == "__main__":
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(SerialService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(SerialService)
