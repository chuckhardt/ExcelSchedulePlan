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


# ***************************************
#
# ExcelPasreCmdLine()
#
# Copies the arguments passed to main, and
# into the specified array
#
# ***************************************
def ExcelPasreCmdLine (CmdArgsPassedToMain, ParsedListofCmdArgs):

    try:

        iRawCmdLineIndex = 1
        iParsedCmdLineIndex = 0
        while (iRawCmdLineIndex < len(CmdArgsPassedToMain)):

            ParsedListofCmdArgs[iParsedCmdLineIndex] = CmdArgsPassedToMain[iRawCmdLineIndex]
            print ("CmdLine Arg " + str(iRawCmdLineIndex) + ": " + CmdArgsPassedToMain[iRawCmdLineIndex])
            iRawCmdLineIndex += 1
            iParsedCmdLineIndex += 1

    except Exception as e:
        print("Exception in ExcelPasreCmdLine()" + str(e))


    return (iRawCmdLineIndex)
