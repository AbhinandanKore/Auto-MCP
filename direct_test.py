import pythoncom
import win32com.client
from win32com.client import gencache
import time

# Initialize COM
pythoncom.CoInitialize()

# Connect AutoCAD
acad = gencache.EnsureDispatch("AutoCAD.Application.26")

acad.Visible = True

print("Waiting for AutoCAD...")

# WAIT IMPORTANT
time.sleep(10)

# Get active document
doc = acad.ActiveDocument

print("CONNECTED SUCCESSFULLY")

# WAIT AGAIN
time.sleep(3)

# Send command
doc.SendCommand('_LINE\n')

print("LINE COMMAND SENT")