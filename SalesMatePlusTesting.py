import pyautogui
pyautogui.FAILSAFE=False
from subprocess import call
import sys, os
import os.path
from win32api import GetSystemMetrics

sys.path.append(os.path.join(os.path.dirname(__file__), 'Modules\Common'))

import ExcelUtil

print("Running SalesMate + Ver 5.0 Test Cases ..")

my_path = os.path.abspath(os.path.dirname(__file__))

call(["python", os.path.join(my_path, "Modules\CustomerMenu.py")])

print("Dumping all the Results form CSV file to Original test case Document... ")

ExcelUtil.SaveTestResultToExcel("salesmatep_plus_test.csv")