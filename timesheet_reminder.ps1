# Function to display a file open dialog
Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Excel Woprkbook (*.xlsx)| *.xlsx"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}


# Create an Excel COM object
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

# Open spreadsheet
$InputFile = Get-FileName("$env:userprofile\Downloads")
if ($InputFile -eq "") {
    exit
}
if (!(Test-Path $InputFile -PathType Leaf)) {
    exit
}
$Excel.Workbooks.Open($InputFile)
$Sheet = $Excel.worksheets.item(1)

# Contains all the pending timesheets per manager
# (key=supervisor, value=array of timesheets)
$managerTimesheets = @{}

# Contains all the pending timesheets per employees
# (key=employee, value=array of timesheets)
$employeeTimesheets = @{}

# Browse each line of the spreadsheet starting at row 2
$row = 2
While ($true) {

    # Create a timesheet off the current row
    $timeSheet = [PSCustomObject]@{
        DateRange          = $Sheet.Cells.Item($row, 1).Text
        EmployeeName       = $Sheet.Cells.Item($row, 2).Text
        EmployeeSupervisor = $Sheet.Cells.Item($row, 5).Text
        Status             = $Sheet.Cells.Item($row, 6).Text
        ActionRequired     = $Sheet.Cells.Item($row, 7).Text
    }

    # Stop when the employee name cell is empty
    if ($timeSheet.EmployeeName -eq "") {
        break
    }

    # When the timesheet is not submitted
    if ($timeSheet.Status -ne "Submitted") {

        # Add it to employeeTimesheets
        if (!($employeeTimesheets.ContainsKey($timeSheet.EmployeeName))) {
            $employeeTimesheets[$timeSheet.EmployeeName] = @()
        }
        $employeeTimesheets[$timeSheet.EmployeeName] += $timeSheet

        # Add it to managerTimesheets
        if (!($managerTimesheets.ContainsKey($timeSheet.EmployeeSupervisor))) {
            $managerTimesheets[$timeSheet.EmployeeSupervisor] = @()
        }
        $managerTimesheets[$timeSheet.EmployeeSupervisor] += $timeSheet

    } elseif ($timeSheet.ActionRequired -ne "Have Approved") {

        # Add it to managerTimesheets
        if (!($managerTimesheets.ContainsKey($timeSheet.EmployeeSupervisor))) {
            $managerTimesheets[$timeSheet.EmployeeSupervisor] = @()
        }
        $managerTimesheets[$timeSheet.EmployeeSupervisor] += $timeSheet

    }

    #Write-Output  $timeSheet
    $row++
}


$Outlook = New-Object -ComObject Outlook.Application

# Prepare email for each employee
$employeeTimesheets.GetEnumerator() | ForEach-Object {

    Write-Output "Employee: " $_.Key

    $Mail = $Outlook.CreateItem(0)
    # Set recipient to employee's short name
    $_.Key -match '(.+), (\w+)'
    $Mail.To = $Matches.0
    $Mail.Subject = "REMIDER: Please submit your timesheets"
    $Mail.Body = "Please submit the following timesheets:"

    Foreach ($t in $_.Value) {
        $Mail.Body += "  o " + $t.DateRange + " (" + $t.Status + ")"
    }

    $Mail.Display()
    #$Mail.Send()
    # Write-Output $Mail
}


# Prepare email for each manager
$managerTimesheets.GetEnumerator() | ForEach-Object {

    Write-Output "Supervisor: " $_.Key

    $Mail = $Outlook.CreateItem(0)
    # Set recipient to supervisor's short name
    $_.Key -match '(.+), (\w+)'
    $Mail.To = $Matches.0
    $Mail.Subject = "REMIDER: Please have a look at the missing timesheets"
    $Mail.Body = "Please have a look at the following employee's timesheets:"

    Foreach ($t in $_.Value) {
        $Mail.Body += "  o " + $t.DateRange + " - " + $t.EmployeeName + " (" + $t.Status + ")"
    }

    $Mail.Display()
    #$Mail.Send()
    #Write-Output $Mail
}


$Excel.Quit()
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Sheet)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
[GC]::Collect()