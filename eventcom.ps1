# eventcom.ps1
# 
# Read an Excel workbook and create Outlook appointments for rows with a Date.
# Default behavior: detect header row columns (Date, Subject, Body, Duration, Location, AllDay).
# Creates appointments and sets a reminder `ReminderDays` before start (default 30 days).
# 
# Usage examples:
#   powershell -ExecutionPolicy Bypass -File .\eventcom.ps1 -ExcelPath 'Appointments VBA - Public.xlsm' -DryRun
#   powershell -ExecutionPolicy Bypass -File .\eventcom.ps1 -ExcelPath '.\MyAppointments.xlsx' -ReminderDays 30
# Parameters:
#   -ExcelPath: path to workbook (required)
#   -SheetName: optional sheet name or index
#   -ReminderDays: days before start for reminder (default 30)
#   -DefaultDurationMinutes: fallback duration (default 60)
#   -DryRun: do not save appointments; just print actions
param(
	[Parameter(Mandatory=$true)] [string]$ExcelPath,
	[Parameter(Mandatory=$false)] [string]$SheetName = '',
	[Parameter(Mandatory=$false)] [int]$ReminderDays = 30,
	[Parameter(Mandatory=$false)] [int]$DefaultDurationMinutes = 60,
	[switch]$DryRun
)

function Release-ComObject($obj) {
	if ($null -ne $obj) {
		try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {}
	}
}

if (-not (Test-Path $ExcelPath)) {
	Write-Error "Excel file not found: $ExcelPath"
	exit 1
}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
	$wb = $excel.Workbooks.Open((Resolve-Path $ExcelPath).Path)
	if ($SheetName -ne '') {
		try { $ws = $wb.Worksheets.Item($SheetName) } catch { $ws = $wb.Worksheets.Item(1) }
	} else { $ws = $wb.Worksheets.Item(1) }

	$used = $ws.UsedRange
	$rows = $used.Rows.Count
	$cols = $used.Columns.Count
	# Read header row and detect columns
	$headerMap = @{}
	for ($c = 1; $c -le $cols; $c++) {
	$h = ($ws.Cells.Item(1, $c).Text -as [string]).Trim()
	$l = $h.ToLower()
	if ($l -match '^(date|start)') { $headerMap['Date'] = $c }
	elseif ($l -match '^(subject|title)') { $headerMap['Subject'] = $c }
	elseif ($l -match '^(body|description)') { $headerMap['Body'] = $c }
	elseif ($l -match '^duration') { $headerMap['Duration'] = $c }
		elseif ($l -match '^location') { $headerMap['Location'] = $c }
		elseif ($l -match '^(allday|all ?day)') { $headerMap['AllDay'] = $c }
	}
	if (-not $headerMap.ContainsKey('Date')) { $headerMap['Date'] = 1 }
	if (-not $headerMap.ContainsKey('Subject')) { $headerMap['Subject'] = 2 }

	$outlook = New-Object -ComObject Outlook.Application
	$created = 0
	for ($r = 2; $r -le $rows; $r++) {
	$rawDate = $ws.Cells.Item($r, $headerMap['Date']).Value2
	if ($null -eq $rawDate -or [string]::IsNullOrWhiteSpace([string]$rawDate)) { continue }

		# Convert Excel serial date (numeric) or text to DateTime
		try {
			if ($rawDate -is [double] -or $rawDate -is [int]) { $start = [DateTime]::FromOADate([double]$rawDate) }
			else { $start = [DateTime]::Parse([string]$rawDate) }
		} catch {
			Write-Warning "Row ${r}: couldn't parse date value '$rawDate' - skipping"
			continue
		}

		$subject = ''
	if ($headerMap.ContainsKey('Subject')) { $subject = ($ws.Cells.Item($r, $headerMap['Subject']).Text -as [string]).Trim() }
	if ([string]::IsNullOrWhiteSpace($subject)) { $subject = 'Appointment' }

	$body = ''
	if ($headerMap.ContainsKey('Body')) { $body = ($ws.Cells.Item($r, $headerMap['Body']).Text -as [string]).Trim() }

		$duration = $DefaultDurationMinutes
		if ($headerMap.ContainsKey('Duration')) {
			$dval = $ws.Cells.Item($r, $headerMap['Duration']).Value2
			if ($dval -ne $null -and [int]::TryParse([string]$dval,[ref]$null)) { $duration = [int]$dval }
		}

		$location = ''
	if ($headerMap.ContainsKey('Location')) { $location = ($ws.Cells.Item($r, $headerMap['Location']).Text -as [string]).Trim() }

	$allDay = $false
		if ($headerMap.ContainsKey('AllDay')) {
			$aval = ($ws.Cells.Item($r, $headerMap['AllDay']).Value2)
			if ($aval -ne $null -and $aval -ne '') {
				$allDay = ($aval -eq 1 -or ($aval -as [string]).ToLower() -in @('true','yes','y'))
			}
		}

		$reminderMinutes = [int]($ReminderDays * 24 * 60)

		if ($DryRun) {
			Write-Output "[DryRun] Row $r -> Subject: '$subject' Start: $start Duration: $duration min Reminder: $ReminderDays days Location: '$location' AllDay: $allDay"
		} else {
			try {
				$appt = $outlook.CreateItem(1)
				$appt.Start = $start
				$appt.Duration = $duration
				$appt.Subject = $subject
				$appt.Body = $body
				if ($location) { $appt.Location = $location }
				if ($allDay) { $appt.AllDayEvent = $true }
				$appt.ReminderSet = $true
				$appt.ReminderMinutesBeforeStart = $reminderMinutes
				$appt.Save()
				$created++
				Write-Output "Created: '$subject' on $start"
			} catch {
				Write-Warning "Row ${r}: failed to create appointment: $_"
			} finally {
				if ($null -ne $appt) { Release-ComObject $appt }
			}
		}
	}

	if (-not $DryRun) { Write-Output "Total appointments created: $created" }

} finally {
	if ($null -ne $wb) { $wb.Close($false) }
	if ($null -ne $excel) { $excel.Quit() }
	Release-ComObject $ws
	Release-ComObject $wb
	Release-ComObject $excel
	if ($null -ne $outlook) { Release-ComObject $outlook }
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
}
