# ******************************************************
#
# Southeastern Railway Museum, Project Task Scheduler
#
# Use Python 3.5+ or better
#
# Author/Owner: C. Hardt
#
# ******************************************************

import openpyxl
import sys

from ScheduleExcelDateStamp  import ExcelUpdateDateStamps
from ScheduleExcelDateFormat import ExcelDate

try:

    print("Schedule Task Date/Time Processor")
    print("Using... OpenXLpy Version: " + openpyxl.__version__)

    print ("Processing File... ", end = " ")

    ExcelFileName = sys.argv[1]
    if (sys.argv[1] == "" ):
        ExcelFileName = "DefaultTestFile.xlsx"
        print (" Default File: DefaultTestFile.xlsx")
    else:
        ExcelFileName = sys.argv[1]
        print (ExcelFileName)

    # Let�s Process the file 
    ProcessExcelDate(ExcelFileName)

except Exception as e:
    print("Exception - Main() " + str(e))





