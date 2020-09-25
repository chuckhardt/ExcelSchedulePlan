# ******************************************************
#
# Southeastern Railway Museum, Project Task Scheduler
#
# Use Python 3.5+ or better
#
# Author/Owner: C. Hardt
#
# ******************************************************

import math
import datetime

# ***************************************
#
# ConvertExcelDate()
#
# Convert the date and time into a printable formate
#
# ***************************************
def ConvertExcelDate(xldatetime):

    try:

        sTempDate = xldatetime.strftime("%m/%d/%Y")
        ExcelNumberDate = xldatetime.strptime(sTempDate, '%m/%d/%Y').date()

    except Exception as e:
        print("Exception in ConvertExcelDate()" + str(e))

    return ExcelNumberDate
