Option Explicit

Sub AddAppointmentsToOutlookCalendar()
    ' Constants for Outlook
    Const olFolderCalendar As Integer = 9
    Const olAppointmentItem As Integer = 1
    
    Dim olApp As Object ' Outlook.Application
    Dim olNamespace As Object ' Outlook.Namespace
    Dim olFolder As Object ' Outlook.Folder
    Dim olApt As Object ' Outlook.AppointmentItem
    
    ' Excel variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim successCount As Long
    
    ' Date variables
    Dim startDate As Date
    Dim startTime As Date
    Dim endDate As Date
    Dim endTime As Date
    
    On Error GoTo ErrorHandler
    
    ' Create Outlook application object
    Set olApp = CreateObject("Outlook.Application")
    
    ' Get Outlook default namespace
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' Get default calendar folder
    Set olFolder = olNamespace.GetDefaultFolder(olFolderCalendar)
    
    ' Set the workbook and worksheet
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Sheet1")
    
    ' Find the last non-empty row in column A (Appointment_Name)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Check if there are appointments data
    If lastRow < 2 Then
        MsgBox "No appointments found in the dataset.", vbInformation
        GoTo Cleanup
    End If
    
    ' Set the range based on the longest row that isn't null
    Set rng = ws.Range("A2:H" & lastRow)
    
    successCount = 0
    
    ' Loop through each appointment in the range
    For Each cell In rng.Rows
        ' Validate essential data
        If IsDate(cell.Range("B1").Value) And IsDate(cell.Range("D1").Value) Then
            ' Create a new appointment item
            Set olApt = olFolder.Items.Add(olAppointmentItem)
            
            With olApt
                .Subject = cell.Range("A1").Value
                
                startDate = cell.Range("B1").Value
                startTime = cell.Range("C1").Value
                .Start = startDate + startTime
                
                endDate = cell.Range("D1").Value
                endTime = cell.Range("E1").Value
                .End = endDate + endTime
                
                .Location = cell.Range("F1").Value
                .Body = cell.Range("G1").Value
                .ReminderSet = True
                .ReminderMinutesBeforeStart = 43200 ' 30 days in minutes
                
                .Save
            End With
            
            successCount = successCount + 1
            Set olApt = Nothing
        End If
    Next cell
    
    MsgBox successCount & " appointments added to Outlook calendar successfully!", vbInformation

Cleanup:
    ' Release object references
    On Error Resume Next
    Set olApt = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
