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

# Open spreadsheet
$InputFile = Get-FileName("$env:userprofile\Downloads")
if ($InputFile -eq "") {
    exit
}
if (!(Test-Path $InputFile -PathType Leaf)) {
    exit
}

# Create an Excel COM object
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true
$Excel.DisplayAlerts = $false
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


if (Get-Process Outlook) {
    $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} else {
    $Outlook = New-Object -ComObject Outlook.Application
}


# Function to display an outlook email prefilled
Function Show-Email($to, $subject, $body) {
    $Mail = $Outlook.CreateItem(0)
    $Mail.To      = '' + $to
    $Mail.Subject = '' + $subject
    $Mail.Body    = '' + $body
    $Mail.Display()
}


# Function to ask for email action
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Function Show-EmailDialog($to, $subject, $body) {

    $dialogY = 600
    $buttonsY = $dialogY - 75

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Timesheet Email Reminder'
    $form.Size = New-Object System.Drawing.Size(475, $dialogY)
    $form.StartPosition = 'CenterScreen'

    $SendButton = New-Object System.Windows.Forms.Button
    $SendButton.Location = New-Object System.Drawing.Point(10, $buttonsY)
    $SendButton.Size = New-Object System.Drawing.Size(75, 23)
    $SendButton.Text = 'Send'
    $SendButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $SendButton
    $form.Controls.Add($SendButton)

    $SendAllButton = New-Object System.Windows.Forms.Button
    $SendAllButton.Location = New-Object System.Drawing.Point(100, $buttonsY)
    $SendAllButton.Size = New-Object System.Drawing.Size(75, 23)
    $SendAllButton.Text = 'Send All'
    $SendAllButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $form.Controls.Add($SendAllButton)

    $EditButton = New-Object System.Windows.Forms.Button
    $EditButton.Location = New-Object System.Drawing.Point(190, $buttonsY)
    $EditButton.Size = New-Object System.Drawing.Size(75, 23)
    $EditButton.Text = 'Edit'
    $EditButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    $form.Controls.Add($EditButton)

    $SkipButton = New-Object System.Windows.Forms.Button
    $SkipButton.Location = New-Object System.Drawing.Point(280, $buttonsY)
    $SkipButton.Size = New-Object System.Drawing.Size(75, 23)
    $SkipButton.Text = 'Skip'
    $SkipButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $form.Controls.Add($SkipButton)

    $AbortButton = New-Object System.Windows.Forms.Button
    $AbortButton.Location = New-Object System.Drawing.Point(370, $buttonsY)
    $AbortButton.Size = New-Object System.Drawing.Size(75, 23)
    $AbortButton.Text = 'Abort'
    $AbortButton.DialogResult = [System.Windows.Forms.DialogResult]::Abort
    $form.CancelButton = $AbortButton
    $form.Controls.Add($AbortButton)

    $labelTo = New-Object System.Windows.Forms.Label
    $labelTo.Location = New-Object System.Drawing.Point(10, 20)
    $labelTo.Size = New-Object System.Drawing.Size(400, 20)
    $labelTo.Text = '[To] ' + $to
    $form.Controls.Add($labelTo)

    $labelSubject = New-Object System.Windows.Forms.Label
    $labelSubject.Location = New-Object System.Drawing.Point(10, 40)
    $labelSubject.Size = New-Object System.Drawing.Size(400, 20)
    $labelSubject.Text = '[Subject] ' + $subject
    $form.Controls.Add($labelSubject)

    $labelBody = New-Object System.Windows.Forms.Label
    $labelBody.Location = New-Object System.Drawing.Point(10, 80)
    $labelBody.Size = New-Object System.Drawing.Size(400, 550)
    $labelBody.Text = $body
    $form.Controls.Add($labelBody)

    $form.Topmost = $true

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Send
    } elseif ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Send All
    } elseif ($result -eq [System.Windows.Forms.DialogResult]::Retry) {
        # Edit
        Show-Email "$to" "$subject" "$body"
    } elseif ($result -eq [System.Windows.Forms.DialogResult]::No) {
        # Skip
    } elseif ($result -eq [System.Windows.Forms.DialogResult]::Abort) {
        # Abort
        $Excel.Quit()
        exit
    }

    $result
}


# Prepare email for each employee
$employeeTimesheets.GetEnumerator() | ForEach-Object {

    Write-Output "Employee: " $_.Key

    # Set recipient to employee's short name
    $_.Key -match '(.+), (\w+)'
    $to = $Matches.0
    $subject = "REMIDER: Please submit your timesheets"
    $body = "Please submit the following timesheets:`n"

    Foreach ($t in $_.Value) {
        $body += "`n    " + $t.DateRange + " (" + $t.Status + ")"
    }

    $result = Show-EmailDialog "$to" "$subject" "$body"
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Send All
    }
}


# Prepare email for each manager
$managerTimesheets.GetEnumerator() | ForEach-Object {

    Write-Output "Supervisor: " $_.Key

    # Set recipient to supervisor's short name
    $_.Key -match '(.+), (\w+)'
    $to = $Matches.0
    $subject = "REMIDER: Please have a look at the missing timesheets"
    $body = "Please have a look at the following employee's timesheets:`n"

    Foreach ($t in $_.Value) {
        $body += "`n    " + $t.DateRange + " - " + $t.EmployeeName + " (" + $t.Status + ")"
    }

    $result = Show-EmailDialog "$to" "$subject" "$body"
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Send All
    }
}


$Excel.Quit()
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Sheet)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
[GC]::Collect()