import math
import os
from datetime import datetime
import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter
import helper

REPORT_NAME = 'Timesheet_Report.xlsx'
REPORT_HEADER_COLOR = 'FFFFD8B1'    # in HEX8 format
NAME_REPORT_HEADER = 'Name'
DAY_REPORT_HEADER = 'Day'
HANA_HOURS_REPORT_HEADER = 'S4Hana Hours'
CLIENT_HOURS_REPORT_HEADER = 'Client Hours'
LEAVE_REPORT_HEADER = 'Leave'       # based on yellow color in client timesheet

DEFAULT_SHEET_NAME_OF_NEW_WORKBOOK = 'Sheet'
OPENPYXL_ENGINE = 'openpyxl'    # Note: Don't change its value. It is a standard value.
S4HANA_TIMESHEET_DUMP_FILE = 'hana.xlsx'
LEAVE_CELL_COLOR = '#FFFF00'
YES = 'Yes'
NO = 'No'
EXCEL_FILE_NAME_LENGTH_LIMIT = 31
EXTRA_CELL_WIDTH = 4

CTS_COLUMN_NAME_INDEX = 0
CTS_COLUMN_COUNTRY_INDEX = 3
CTS_FIRST_DATE_COLUMN_INDEX = 4  # index starts from 0
CTS_NAME_PREFIX_TO_REMOVE_LENGTH = 18
CTS_PROCESS_STOP = 'India'
CTS_ROW_PROCESS_STOP_COLUMN = 'Total Billable Hours'
CTS_FOLDER = 'Client'     # contains all timesheets of different squads
CTS_ACTUAL_START_ROW = 12  # row starts from 1. It is count of rows not index

STS_COLUMN_NAME_INDEX = 1
STS_FIRST_DATE_COLUMN_INDEX = 4  # index starts from 0
STS_ACTUAL_START_ROW = 2 # row starts from 1. It is count of rows not index

def parseClientTimesheets(directory, hanaResult):
    print('INFO: Start parse timesheets from "%s" folder' % directory)
    wb = openpyxl.Workbook()

    fileCount = helper.getFileCount(directory)
    print('%s files found' % fileCount)

    failedFileCount = 0
    skippedFileCount = 0

    for fileName in os.listdir(directory):
        if not helper.isFileHasSpecifiedExtension(fileName, ['.xlsx', '.xls']):
            skippedFileCount += 1
            print('WARN: "%s" is not an excel file' % fileName)
            continue

        try:
            filePath = os.path.join(directory, fileName)
            dataFrame = pd.read_excel(filePath, skiprows = CTS_ACTUAL_START_ROW - 1, engine = OPENPYXL_ENGINE) # columns of interest starts from 12th row

            clientDayHoursByName = parseSingleClientTimesheet(dataFrame)
            print('INFO: "%s" parsed successfully' % fileName)

            leavesByName  = getLeavesByName(filePath)
            comparisonResult = compareHanaWithClientDetails(hanaResult, clientDayHoursByName, leavesByName)
            
            filenameWithoutExtension = helper.getFileNameWithoutExtension(fileName)
            sheetName = filenameWithoutExtension[CTS_NAME_PREFIX_TO_REMOVE_LENGTH:]
            sheetName = sheetName[0:EXCEL_FILE_NAME_LENGTH_LIMIT]

            wb = generateReport(wb, comparisonResult, sheetName)
            print('INFO: Report for "%s" generated successfully' % fileName)
        except Exception as e:
            failedFileCount += 1
            print('ERROR: "%s" occured when processing %s' %(e, fileName))

    reportName = '_'.join([helper.getCurrentMonthName(), REPORT_NAME])
    wb.save(reportName)
    print('INFO: "%s" generated successfully' % reportName)
    print('INFO: Report for Client timesheets: %s SUCCESS, %s FAILED %s SKIPPED' % (fileCount - failedFileCount - skippedFileCount, failedFileCount, skippedFileCount))

def parseHana(dataFrame):
    columns = dataFrame.columns
    dayHoursByName = {}

    for r in range(len(dataFrame.index)):
        name = dataFrame.loc[r, columns[STS_COLUMN_NAME_INDEX]]

        for c in range(STS_FIRST_DATE_COLUMN_INDEX, len(columns)):
            currentDateColumn = columns[c]

            dayHour = {}
            day = int(currentDateColumn[:2])   # first 2 chars represent day. Column Format : dd.mm.yyyy
            hours = dataFrame.loc[r, currentDateColumn]
            dayHour[day] = 0 if math.isnan(hours) else hours
            dayHoursByName.setdefault(name, []).append(dayHour)
    return dayHoursByName
    
def parseSingleClientTimesheet(dataFrame):
    columns = dataFrame.columns
    totalRows = len(dataFrame.index)
    dayHoursByName = {}

    for r in range(totalRows):
        name = dataFrame.loc[r, columns[CTS_COLUMN_NAME_INDEX]]
        country = dataFrame.loc[r, columns[CTS_COLUMN_COUNTRY_INDEX]]
        
        if(country != CTS_PROCESS_STOP):     # if its a valid row, can also check person number...
            break

        for c in range(CTS_FIRST_DATE_COLUMN_INDEX, len(columns)):
            if (columns[c] == CTS_ROW_PROCESS_STOP_COLUMN):
                break
            else:                
                fridayOfCurrentWeek = helper.getSpecifDayOfCurrentWeek(4)   # 4 - index of Friday
                today = datetime.today()
                lastDayOfCurrentMonth = helper.getLastDayOfMonth(today.year, today.month)
                checkDay = lastDayOfCurrentMonth if today.day > fridayOfCurrentWeek else fridayOfCurrentWeek

                currentDateColumn = columns[c]
                day = int(currentDateColumn)
                if day > checkDay:    # process data only till checkDay
                    break
                
                dayHour = {}
                hours = dataFrame.loc[r,currentDateColumn]
                dayHour[day] = 0 if math.isnan(hours) else hours
                dayHoursByName.setdefault(name, []).append(dayHour)
    return dayHoursByName

def compareHanaWithClientDetails(hanaResult, clientResult, leavesByName):
    comparisonResult = []
    for person in hanaResult:
        clientResultForPerson = clientResult.get(person)
        if clientResultForPerson is None:
            continue
        for dh in hanaResult[person]:
            [(day, hour)] = dh.items()      # single the object has only one key-value pair
            for dh2 in clientResultForPerson:
                [(day2, hour2)] = dh2.items()
                if day == day2:
                    if (hour != hour2) or (hour == 0 and hour2 == 0):       # either mismatch hours, or hours are empty in both timesheets
                        hasPersonAnyLeaveInTheMonth = leavesByName.get(person)     # First check if name exists in the leave result
                        isPersonOnLeave = YES if hasPersonAnyLeaveInTheMonth and int(day) in leavesByName[person] else NO

                        verifyData = {}
                        verifyData[NAME_REPORT_HEADER] = person
                        verifyData[DAY_REPORT_HEADER] = day
                        verifyData[HANA_HOURS_REPORT_HEADER] = hour
                        verifyData[CLIENT_HOURS_REPORT_HEADER] = hour2
                        verifyData[LEAVE_REPORT_HEADER] = isPersonOnLeave
                        comparisonResult.append(verifyData)
                    break
    return comparisonResult

def getLeavesByName(clientTimesheetPath):
    wb = openpyxl.load_workbook(clientTimesheetPath, data_only = True)
    sheet = wb.active
    stop = False
    headerRow = ''
    
    leavesByName = {}
    for row, row_cells in enumerate(sheet.iter_rows()):
        cellNum = 0   
                                 
        if row < CTS_ACTUAL_START_ROW - 1:    # skip unnecessary rows. -1 becoz row index starts from 0                                                                          
            continue
        elif row == CTS_ACTUAL_START_ROW - 1: # contains column names for data of interest. -1 becoz row index starts from 0
            headerRow = row_cells
            continue

        for cell in row_cells:
            if cellNum == 0 and cell.value is None: # if the row is not resource row, it means all resource's timings have been processed
                stop = True
                break

            color = helper.getCellColor(cell)
            if color == LEAVE_CELL_COLOR:
                name = row_cells[CTS_COLUMN_NAME_INDEX].value
                day = headerRow[cellNum].value
                leavesByName.setdefault(name, []).append(day)
                # print(cell.value, color, row_cells[0].value, headerRow[cellNum].value)
            cellNum += 1

        if stop:
            break
    return leavesByName            

def generateReport(wb, reportResult, sheetName) :
    helper.removeDefaultSheet(wb)
    sheet = helper.createSheet(wb, sheetName)
    headerColumns = [NAME_REPORT_HEADER, DAY_REPORT_HEADER,HANA_HOURS_REPORT_HEADER,
                        CLIENT_HOURS_REPORT_HEADER, LEAVE_REPORT_HEADER]

    helper.renderDataInSheet(sheet, reportResult, headerColumns, REPORT_HEADER_COLOR)
    helper.adjustCellWidthToContent(sheet, EXTRA_CELL_WIDTH) # add EXTRA_CELL_WIDTH becoz otherwise the width does not fit the content correctly
    helper.hideGridLines(sheet)
    helper.applyFilter(sheet)
    return wb

if __name__ == '__main__':
    try:
        print('INFO: "%s" read start' % S4HANA_TIMESHEET_DUMP_FILE )
        dataFrame = pd.read_excel(S4HANA_TIMESHEET_DUMP_FILE, skiprows = STS_ACTUAL_START_ROW - 1, engine = OPENPYXL_ENGINE)   # columns of interest starts from 2nd row
        print('INFO: "%s" read successfully' % S4HANA_TIMESHEET_DUMP_FILE )

        print('INFO: "%s" parse start' % S4HANA_TIMESHEET_DUMP_FILE )
        hanaResult = parseHana(dataFrame)
        print('INFO: "%s" parsed successfully' % S4HANA_TIMESHEET_DUMP_FILE )

        parseClientTimesheets(CTS_FOLDER, hanaResult)

    except Exception as e:
        print('ERROR: %s' % e)
    