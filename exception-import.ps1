param (
    [Parameter(Mandatory = $true)]  [String] $ProjectFile,
    [Parameter(Mandatory = $true)]  [String] $CalendarSheetFile,
    [Parameter(Mandatory = $false)] [Switch] $ClearExceptions
)

Set-StrictMode -Version Latest

if(-not (Test-Path $ProjectFile)) {
    throw [System.IO.FileNotFoundException] "$ProjectFile does not exist" 
}
if(-not (Test-Path $CalendarSheetFile)) {
    throw [System.IO.FileNotFoundException] "$CalendarSheetFile does not exist" 
}
$ProjectFilePath = (Get-Item $ProjectFile).FullName
$CalendarSheetFilePath = (Get-Item $CalendarSheetFile).FullName

# open worksheet
Write-Output "Opening $CalendarSheetFile"
$ExcelApp = New-Object -ComObject 'Excel.Application'
$Workbook = $ExcelApp.Workbooks.Open($CalendarSheetFilePath, 0, $true) #read only

# open project
Write-Output "Opening $ProjectFile"
$ProjectApp = New-Object -ComObject 'MSProject.Application'
$null = $ProjectApp.FileOpenEx($ProjectFilePath, $false)

$Project = $null
$Exceptions = $null
foreach($proj in $ProjectApp.Projects) {
    if($proj.FullName -eq $ProjectFilePath) {
        $Exceptions = $proj.Calendar.Exceptions
        $Project = $proj
        break
    }
}
if($null -eq $Project) {
    throw "$ProjectFileName is not found as FullName in all opened projects"
}

# analyze table headers
$Worksheet = $Workbook.ActiveSheet
$TotalRows = $Worksheet.UsedRange.Rows.Count
$TotalCols = $Worksheet.UsedRange.Columns.Count

$ExNameCol = -1
$ExWorkdayCol = -1
$ExFinishDateCol = -1
$ExStartDateCol = -1

$ShiftStartCol = -1
$ShiftFinishCol = -1

for($c = 1; $c -le $TotalCols; $c += 1) {
    $val = $Worksheet.Cells(1, $c).Text
    if($null -eq $val) {
        continue
    }
    $val = $val.toString().trim().toUpper()
    switch($val) {
        'NAME'         {$ExNameCol       = $c; break}
        'WORKDAY'      {$ExWorkdayCol    = $c; break}
        'START'        {$ExStartDateCol  = $c; break}
        'FINISH'       {$ExFinishDateCol = $c; break}
        'SHIFT START'  {$ShiftStartCol   = $c; break}
        'SHIFT FINISH' {$ShiftFinishCol  = $c; break}
    }
}

# get shift info
$WorkdayShifts = [System.Collections.ArrayList]@() #array
$cnt = 0
for($r = 2; $r -le $TotalRows; $r += 1){
    if($cnt -gt 5) {
        # only 5 shifts
        break
    }
    else {
        $st = $Worksheet.Cells($r, $ShiftStartCol).Text
        $fn = $Worksheet.Cells($r, $ShiftFinishCol).Text
        if(('' -ne $st) -and ('' -ne $fn)) {
            $null = $workdayShifts.Add(@{st = $st; fn = $fn})
            $cnt += 1
        }
    }
}
Write-Output "Using shifts as workday exceptions:"
$cnt = 0
foreach($sh in $WorkdayShifts) {
    $cnt += 1
    Write-Output ("{0}: {1} - {2}" -f $cnt, $sh.st, $sh.fn)
}
# loop and add exceptions
if($ClearExceptions) {
    Write-Output "Clearing all previous exceptions"
    foreach ($ex in $Exceptions) {
        $ex.Delete()
    }
}
for($r = 2; $r -le $TotalRows; $r += 1){
    $name = $Worksheet.Cells($r, $ExNameCol).Text
    $isWorkday = $Worksheet.Cells($r, $ExWorkdayCol).Text
    $start = $Worksheet.Cells($r, $ExStartDateCol).Text
    $finish = $Worksheet.Cells($r, $ExFinishDateCol).Text
    Write-Output "Adding exception: $name, Workday = $isWorkday, Start = $start, Finish = $finish"
    $ex = $null
    $ex = $Exceptions.Add(
        1,  #pjDaily
        $start,
        $finish,
        $null, #$duration,
        $name
    )
    if($null -eq $ex) {
        continue
    }
    if([bool]$isWorkday) {
        $cnt = 0
        :shiftloop foreach($sh in $WorkdayShifts) {
            $cnt += 1
            switch($cnt) {
                1 {$ex.Shift1.Start = $sh.st; $ex.Shift1.Finish = $sh.fn; break}
                2 {$ex.Shift2.Start = $sh.st; $ex.Shift2.Finish = $sh.fn; break}
                3 {$ex.Shift3.Start = $sh.st; $ex.Shift3.Finish = $sh.fn; break}
                4 {$ex.Shift4.Start = $sh.st; $ex.Shift4.Finish = $sh.fn; break}
                5 {$ex.Shift5.Start = $sh.st; $ex.Shift5.Finish = $sh.fn; break}
                default {break shiftloop}
            }
        }
    }
}

Write-Output "Saving project: $ProjectFile"
$null = $ProjectApp.FileExit(1) #pjSave
$ExcelApp.Quit()
