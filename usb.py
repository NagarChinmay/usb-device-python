import serial
import json

# Replace 'COMx' with the appropriate serial port
ser = serial.Serial('COM4', 9600)

def receive_and_decode():
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
            # Print or use the decoded values
            print("Student Type:", student_type[0])
            print("UID:", uid[0])
            print("Name:", name[0])
            print("Enrolment:", enrolment[0])
            print("Stream:", stream[0])
            print("Year:", year[0])
            print("Boarding:", boarding[0])
        except json.JSONDecodeError as e:
            print("Error decoding JSON:", e)

if __name__ == "__main__":
    try:
        receive_and_decode()
    except KeyboardInterrupt:
        ser.close()
