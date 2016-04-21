import datetime
import sys
import time
import xml.etree.ElementTree as etree
from requests import HTTPError
import serial


ser = serial.Serial(10, 38400, timeout=0.2,parity=serial.PARITY_NONE, rtscts=1)
s = ser.read(100)   # read up to one hundred bytes
                        # or as much is in the buffer
line = ser.readline()   # read a '\n' terminated line
          # write a string  ser.write("hello")
  
    
while True:
        s = ser.read(1000)   # read up to one hundred bytes
        if len(s)>0:
                print "received message" + str(s)
        else:
                continue
ser.close()
