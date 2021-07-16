import win32com.client, os, re
from shutil import copy, move
from datetime import datetime, timedelta


# Function: find file- x + any chars
def FileFinder(fileList, strBegin):
    regex = strBegin + '.+'
    pattern = re.compile(regex)
    for f in fileList:
        if re.findall(pattern, f):  # empty seq = False
            found = re.findall(pattern, f)
    return found


# Function: String date (YYYY/MM/DD) to Datetime stamp
def strDatetoDateStamp(strDate):
    datestamp = datetime.strptime(strDate, '%Y.%m.%d')
    return datestamp


# Function: Datetime stamp to string date (YYYY/MM/DD)
def dateStampToStrDate(dateStamp):
    strDate = datetime.strftime(dateStamp, '%Y.%m.%d')
    return strDate


path = '\\\\SERVER\\DIRECTORY\\PATH\\'
fileList = os.listdir(path)  # list of desktop items
fn = FileFinder(fileList, 'Office Time Sheet - Week Ending ')[0]  # regex search for time sheet file

# Open, Rename, Print date.
for x in range(1, 7):
    d = datetime.today() - timedelta(days=x)  # Today's date backwards until sunday

    if d.strftime('%A') == 'Sunday':

        # Open time sheet
        ExObj = win32com.client.Dispatch("Excel.Application")
        ExObj.Visible = 1
        wb = ExObj.Workbooks.Open(path + fn)
        ws = wb.Sheets(1)

        # Rename date
        ws.Cells(4, 1).Value = datetime.strftime(d, 'Week Ending: %d-%m-%Y')

        # Print
        ws.PrintOut()
        wb.Close(True)
        ExObj.Application.Quit()
        break

# 1. Capture the date of the time sheet;  2. Add on 7 days to get the next Monday;  3. Implement new date into a new filename.
strDate = fn[-14:-4]
dateStamp = strDatetoDateStamp(strDate)
newDateStamp = dateStamp + timedelta(days=7)
strDate = dateStampToStrDate(newDateStamp)
newFn = fn[0:-14] + strDate + '.xls'

# Move completed time sheet to 'Time Sheets' folder in 'Work for Staff'
move(path + fn, 'DRIVE_LETTER:\\DIRECTORY_B\\Time Sheets\\')

# Create a copy of the blank time sheet, rename it to 'newFn', and move it to the Desktop.
copy('DRIVE_LETTER:\\DIRECTORY_B\\Office Time Sheet - Week Ending BLANK.xls', path + newFn)
