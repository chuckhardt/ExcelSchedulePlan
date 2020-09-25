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
import datetime
import array as arr
import sys

from openpyxl.styles import NamedStyle
from ScheduleExcelUtilites import ExcelColumnUtil

# ***************************************
#
# IsDayAholiday()
#
# Convert the date used by Excel to a number we can work with
# determine if the day in question is a holiday,
#
# TRUE = Holiday
# FALSE = NOT a Holiday
#
# ***************************************
def IsDayAholiday(ExcelTaskStartDate):

    bReturnFlag = False

    try:

        USholidays = [datetime.datetime(2019, 12, 24, 00, 00),
                      datetime.datetime(2019, 12, 25, 00, 00),
                      datetime.datetime(2019, 12, 31, 00, 00),
                      datetime.datetime(2020, 1, 1, 00, 00),
                      datetime.datetime(2020, 1, 20, 00, 00),
                      datetime.datetime(2020, 5, 25, 00, 00),
                      datetime.datetime(2020, 7, 3, 00, 00),
                      datetime.datetime(2020, 5, 25, 00, 00),
                      datetime.datetime(2020, 7, 3, 00, 00),
                      datetime.datetime(2020, 9, 7, 00, 00),
                      datetime.datetime(2020, 11, 26, 00, 00),
                      datetime.datetime(2020, 11, 27, 00, 00),
                      datetime.datetime(2020, 12, 24, 00, 00),
                      datetime.datetime(2020, 12, 25, 00, 00),
                      datetime.datetime(2020, 12, 31, 00, 00)]

        for iIndex in range(0,len(USholidays),1):

            # if day is a holiday, return True
            if (ExcelTaskStartDate == USholidays[iIndex]):

                print ("Holiday Date Match: " + str(ExcelTaskStartDate))
                bReturnFlag = True
                break

            else:
                iIndex +=1

    except Exception as e:
        print("Exception - IsDayAHoliday() " + str(e))

    return (bReturnFlag)


# ***************************************
#
# DetermineNextWorkingDayAfterHoliday()
#
# Determine what the next working day is, after the holiday passed
#
# ***************************************
def DetermineNextWorkingDayAfterHoliday(ExcelTaskStartDate):

    try:

        USholidays = [datetime.datetime(2019, 12, 24, 00, 00),
                      datetime.datetime(2019, 12, 25, 00, 00),
                      datetime.datetime(2019, 12, 31, 00, 00),
                      datetime.datetime(2020, 1, 1, 00, 00),
                      datetime.datetime(2020, 1, 20, 00, 00),
                      datetime.datetime(2020, 5, 25, 00, 00),
                      datetime.datetime(2020, 7, 3, 00, 00),
                      datetime.datetime(2020, 5, 25, 00, 00),
                      datetime.datetime(2020, 7, 3, 00, 00),
                      datetime.datetime(2020, 9, 7, 00, 00),
                      datetime.datetime(2020, 11, 26, 00, 00),
                      datetime.datetime(2020, 11, 27, 00, 00),
                      datetime.datetime(2020, 12, 24, 00, 00),
                      datetime.datetime(2020, 12, 25, 00, 00),
                      datetime.datetime(2020, 12, 31, 00, 00)]

        for iIndex in range(0,len(USholidays),1):

            if (ExcelTaskStartDate == USholidays[iIndex]):

                print ("Holiday Date Match: " + str(ExcelTaskStartDate))
                ExcelTaskStartDate = ExcelTaskStartDate + datetime.timedelta(days=1)

                # check the next date and make sure it is not also a holiday
                ExcelTaskStartDate = DetermineNextWorkingDayAfterHoliday(ExcelTaskStartDate)
                break

            else:
                iIndex +=1

    except Exception as e:
        print("Exception - DetermineNextWorkingDayAfterHoliday() " + str(e))


    return (ExcelTaskStartDate)

# ***************************************
#
# CalculateNextWorkingDay()
#
# we are going to test to see if the day of the week passed to the function
# is a Saturday or Sunday, and it it is, we will increment the day to the next
# working day.  If Saturday is passed, the next working day will be Monday.
# Likewise Sunday would be converted to Monday
#
# ***************************************
def CalculateNextWorkingDay(ExcelTaskStartDate):

    kMonday = 1
    kFriday = 5
    bWeekDayFound = False

    try:

        while bWeekDayFound == False:

            # did we find a weekday (Monday thru Friday)
            iweekday = ExcelTaskStartDate.isoweekday()
            if ((iweekday >= kMonday) and (iweekday <= kFriday)):

                # Is the day a holiday by any chance?
                if (IsDayAholiday(ExcelTaskStartDate) == False):
                    bWeekDayFound = True

                # it is a holiday, so advance to the next day and look at the next day
                else:
                    ExcelTaskStartDate = ExcelTaskStartDate + datetime.timedelta(days=1)
            else:
                ExcelTaskStartDate = ExcelTaskStartDate + datetime.timedelta(days=1)

    except Exception as e:
        print("Exception - CalculateNextWorkingDay() " + str(e))


    return (ExcelTaskStartDate)

