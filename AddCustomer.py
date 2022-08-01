import pyautogui
pyautogui.FAILSAFE=False
from win32api import GetSystemMetrics
import sys, os

sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'Common'))

import ExcelUtil

print("Executing Add Customer Test Case")

print("Create a Database Backup directory called SMP")

path="D:/SMP"
if not os.path.exists(path):
    os.makedirs(path)

#get the width and height of the monitor
# width= GetSystemMetrics(0)
# height=GetSystemMetrics(1)
#click on the botton left to launch the start menu
print("click on the botton left to launch the start menu")
pyautogui.click(x=26, y=1058,clicks=1,interval=2)
print("Select the SalesMate + Application")
pyautogui.typewrite("SalesMate +",0.5)
print("Press Enter Key to Launch SalesMate + Application and Wait for 10 Sec")
pyautogui.press("enter",1,5)
#Sometimes SalesMate + might take 10 seconds to load it . So 10 sec delay

#To close the setup wizard box
pyautogui.press("tab",1,2)
pyautogui.press("enter",1,2)

print("Press Enter Key again to close the Tip Of the Day Dialog")
pyautogui.press("enter",1,2)

print("Now alt and c shortcut key in  Salesmate to invoke the Customer Root menu")
pyautogui.press(['alt','c'],1,2)
print("Now press 'a' to add a customer")
pyautogui.press("a",1,2)

pyautogui.typewrite("31/7/22")
pyautogui.press("tab",1,2)
pyautogui.typewrite("10")
pyautogui.press("tab",1,2)
pyautogui.typewrite("Manisha")
pyautogui.press("tab",1,2)
pyautogui.typewrite("Patel")
pyautogui.press("tab",1,2)
pyautogui.typewrite("Ms")
pyautogui.press("tab",1,2)
pyautogui.typewrite("Delhi")
pyautogui.press("tab",1,2)
pyautogui.typewrite("234876")
pyautogui.press("tab",1,2)
pyautogui.typewrite("manisha@patel.com")
pyautogui.click(x=1074, y=865)
pyautogui.press("enter",1,2)
pyautogui.press("alt",1,2)
pyautogui.press("f",1,2)
pyautogui.press("x",1,2)
#pyautogui.press(['alt','f','x'],1,3)


print("Log the Results to the CSV File")
if os.path.exists(path+"/SalesMatePlus.mdb"):
    ExcelUtil.SaveTestResultToCSV("1","2","Pass",path+"/SalesMatePlus.mdb")
else:
    ExcelUtil.SaveTestResultToCSV("1","2","Fail","")

