# project-exception-import
Import Microsoft Project workday exceptions from Excel sheets. Requires both Microsoft Excel & Project installed.

Exceptions are days which have shift schedules different from default work calendar. 

## Usage
1. Create an Excel file, including these headers:
    - `Name`: name of the exception
    - `Workday`: incidates whether the exception is an overtime (1) or a leave (0)
    - `Start`: starting date of the exception
    - `Finish`: ending date of the exception
    
    If there are any overtime exceptions, a shift schedule of them should also be defined:
    - `Shift Start`: starting time of the overtime shift
    - `Shift Finish`: ending time of the overtime shift
    MS Project supports only 5 work periods per day. The script will take the first 5 periods and ignore the rest.
    
2. Open powershell, and call the script as below. Omit `-ClearExceptions` if you want to keep existing exceptions from the project file.

``.\exception_import.ps1 -ProjectFile .\myproject.mpp -CalendarSheetFile .\mysheet.xlsx -ClearExceptions``

