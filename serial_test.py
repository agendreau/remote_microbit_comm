import serial
import requests
import time

port = "/dev/cu.usbmodem14102"
conn = serial.Serial(port, 115200)
count = 0;

while True:
    if(conn.inWaiting()>0):
        print("Message from Micro:bit")
        print(conn.readline().decode("utf-8").strip())
    #value = input("Write a short message:\n")
    #r = requests.get("http://flipacoinapi.com/json")
    #result = r.text
    count+=1
    #print(value)
    result_bytes = (str(count) + ",").encode()
    conn.write(result_bytes)
    time.sleep(0.01)