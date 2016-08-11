'########################################################
'#####      Track-IT Reports                        #####
'#####      Automated Daily Reports from Track-IT   #####
'#####      Author: Spencer Alessi                  #####
'#####      Modified: 8/10/2016                     #####  
'########################################################

Dim sStartDate, sEndDate, startDate, endDate, IS_MONDAY, today

Do
endDate = Date
'is it monday?
IS_MONDAY = WeekdayName(2)
today = WeekDayName(WeekDay(Now()))

intCompare = StrComp(today, IS_MONDAY, vbTextCompare)

'If monday: use friday's date for startdate, otherwise use normal startdate
If intCompare = 0 Then
    startDate = DateAdd("d",-3,endDate)
Else
	startDate = DateAdd("d",-1,endDate)
End If

sStartDate = Year(startDate) & "-" & Month(startDate) & "-" & Day(startDate)  
sEndDate = Year(endDate) & "-" & Month(endDate) & "-" & Day(endDate)


Loop While Len(Trim(sStartDate)) = 0
'Wscript.Echo "Starting Date: " & sStartDate
'Wscript.Echo "Ending Date: " & sEndDate

Dim username, password, server, inputfile, outputfile, filetype, rangevalue, daterange, yearFolder, theMonth, monthFolder
yearFolder = Year(sEndDate)
'Wscript.Echo "Year: " & yearFolder
theMonthNumber = Month(sEndDate)
theMonthName = MonthName(theMonthNumber)
monthFolder = theMonthNumber & "_" & theMonthName
'Wscript.Echo "Month: " & monthFolder

Dim oFso, oFolder, sDirectory
Set oFso = CreateObject("Scripting.FileSystemObject")
sDirectory = "[path to where you want to save your reports]" & yearFolder & "\" & monthFolder
'msgbox sDirectory
If (oFso.FolderExists(sDirectory)) Then
	'msgbox "Folder exists already - skipping create new folder"
Else
	'msgbox "Folder does not exist - creating new folder now"
  Set oFolder = oFSO.CreateFolder(sDirectory)
End If

username = "-U " & "[TrackIT Database User]"
'clear text password is bad practice, create a read only account
password = "-P " & "[TrackIT Database password]"
server = "-S " & "Trackit "
inputfile = "-F " & "daily_report.rpt "
outputfile = "-O " & CHR(34) & sDirectory & "\" & "trackit-report-" & sEndDate & ".xls" & CHR(34)
filetype = " -E " & "xls "
rangevalue = "-a " & CHR(34) & "Date_ Entered_Range:" & "(" & sStartDate & "," & sEndDate & ")" & CHR(34)
parameters = username&password&server&inputfile&outputfile&filetype&rangevalue

Dim oShell
Set oShell = Wscript.CreateObject ("Wscript.Shell")

oShell.Run "crexport.exe" & " " & parameters, 1, true

Set oShell = Nothing




 