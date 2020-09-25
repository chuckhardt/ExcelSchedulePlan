import ScheduleExcelRowUtilities as MyExcel
import os.path
import sys
import openpyxl

from shutil import copyfile
from ScheduleExcelSort import ScheduleExcelSortSheet
from ScheduleExcelSort import ScheduleExcelFormatDate
from ScheduleExcelUtilites import ExcelColumnUtil
from ScheduleExcelParseCmdLine import ExcelPasreCmdLine
from ScheduleExcelCalculateTaskStartAndEndDates import ScheduleExcelCalculateTaskStartAndEndDates
from os import path

ParsedListOfCmdArgs = ["Museum Project Schedule.xlsx", "Week 20.01A", "01/03/20"]
kExcelFileNameIndex = 0
kWorkBlkNameIndex = 1
kWorkBlkStartDateIndex = 2

# ***************************************
#
# Entry Point
#
# Everything is called from here
#
# ***************************************
try:

    print("Southeastern Railway Museum Project Scheduler - v1.02");

    # parse the command line, and convert to usable data
    ExcelPasreCmdLine (sys.argv, ParsedListOfCmdArgs)  
    ExcelFileName = ParsedListOfCmdArgs[kExcelFileNameIndex]

    # Let's make sure the file we passed, actually exists
    if (path.exists(ExcelFileName) != True):
        print ("ERROR: FileName " + str(ExcelFileName) + " NOT Found")
        exit(1)

    sWorkBlkName = ParsedListOfCmdArgs[kWorkBlkNameIndex]
    sWorkBlkStartDate = ParsedListOfCmdArgs[kWorkBlkStartDateIndex]

    # open the workbook and worksheet.  Use the first worksheet in the workbook
    print ("Processing File... " + ExcelFileName)

    WorkBookName = openpyxl.load_workbook(ExcelFileName)

    WorksheetName = WorkBookName.active
    WorksheetName = "Sheet1"

    #print("Workbook Name: " + str(WorkBookName) + " Worksheet: " + WorksheetName)

    # now we are going to sort the excel file by WorkBlk number, then assignee, and then finally by task number
    ScheduleExcelSortSheet (WorksheetName)
    #ScheduleExcelCalculateTaskStartAndEndDates(WorksheetName, sWorkBlkName, sWorkBlkStartDate)

    # WorkBookName.save(ExcelFileName)

except Exception as e:
        print("Exception in ExcelMain()" + str(e))

