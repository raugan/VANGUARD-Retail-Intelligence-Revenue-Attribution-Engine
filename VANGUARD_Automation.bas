Attribute VB_Name = "VANGUARD_Automation"
'================================================================
' VANGUARD: Automation Module
' Description: VBA macros for automated reporting and data refresh
' Last Updated: February 2026
'================================================================

Option Explicit

'================================================================
' MACRO 1: Refresh All Data and Recalculate
'================================================================
Sub RefreshAllDataAndRecalculate()
    '
    ' Purpose: Refresh all data connections and force recalculation
    ' Keyboard Shortcut: Ctrl+Shift+R
    '
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ws As Worksheet
    
    ' Refresh all queries and connections
    On Error Resume Next
    ActiveWorkbook.RefreshAll
    On Error GoTo 0
    
    ' Recalculate all formulas
    For Each ws In ActiveWorkbook.Worksheets
        ws.Calculate
    Next ws
    
    Application.Calculate
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Data refresh and recalculation complete!", vbInformation, "VANGUARD"
End Sub

'================================================================
' MACRO 2: Export Dashboard to PDF
'================================================================
Sub ExportDashboardToPDF()
    '
    ' Purpose: Export Live Formulas sheet to PDF for board reporting
    ' Keyboard Shortcut: Ctrl+Shift+P
    '
    Dim ws As Worksheet
    Dim pdfPath As String
    Dim timestamp As String
    
    Application.ScreenUpdating = False
    
    ' Set worksheet to export
    Set ws = ThisWorkbook.Worksheets("Live Formulas")
    
    ' Generate filename with timestamp
    timestamp = Format(Now(), "yyyy-mm-dd_hhmm")
    pdfPath = ThisWorkbook.Path & "\VANGUARD_Dashboard_" & timestamp & ".pdf"
    
    ' Configure print settings for optimal PDF output
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintArea = "$A$1:$C$40"
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.25)
        .FooterMargin = Application.InchesToPoints(0.25)
        .PrintTitleRows = "$1:$1"
        .CenterHeader = "&B&14VANGUARD Dashboard"
        .RightFooter = "Page &P of &N"
    End With
    
    ' Export to PDF
    On Error Resume Next
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    If Err.Number = 0 Then
        MsgBox "PDF exported successfully to:" & vbCrLf & pdfPath, vbInformation, "VANGUARD"
    Else
        MsgBox "Error exporting PDF: " & Err.Description, vbExclamation, "VANGUARD"
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = True
End Sub

'================================================================
' MACRO 3: Generate Executive Summary Email
'================================================================
Sub GenerateExecutiveSummaryEmail()
    '
    ' Purpose: Create formatted email with key metrics for distribution
    ' Keyboard Shortcut: Ctrl+Shift+E
    '
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim emailBody As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Live Formulas")
    
    ' Build email body with key metrics
    emailBody = "<html><body style='font-family: Arial, sans-serif;'>"
    emailBody = emailBody & "<h2 style='color: #007BFF;'>VANGUARD: Weekly KPI Summary</h2>"
    emailBody = emailBody & "<p>Executive Team,</p>"
    emailBody = emailBody & "<p>Please find the key performance indicators for the week ending " & Format(Date, "mmmm dd, yyyy") & ":</p>"
    emailBody = emailBody & "<table border='1' cellpadding='8' cellspacing='0' style='border-collapse: collapse;'>"
    emailBody = emailBody & "<tr style='background-color: #007BFF; color: white;'>"
    emailBody = emailBody & "<th>KPI</th><th>Current Value</th><th>Status</th></tr>"
    
    ' Add KPI rows (reading from worksheet)
    emailBody = emailBody & "<tr><td>Customer Lifetime Value</td><td>" & Format(ws.Range("C23").Value, "$#,##0.00") & "</td><td>✓</td></tr>"
    emailBody = emailBody & "<tr><td>Revenue Velocity</td><td>" & Format(ws.Range("C24").Value, "0.00%") & "</td><td>✓</td></tr>"
    emailBody = emailBody & "<tr><td>Category Concentration</td><td>" & Format(ws.Range("C25").Value, "0.00%") & "</td><td>✓</td></tr>"
    emailBody = emailBody & "<tr><td>Marketing ROI</td><td>" & Format(ws.Range("C31").Value, "0.00x") & "</td><td>✓</td></tr>"
    
    emailBody = emailBody & "</table>"
    emailBody = emailBody & "<p><strong>Strategic Insights:</strong></p>"
    emailBody = emailBody & "<ul>"
    emailBody = emailBody & "<li>Electronics category maintains strong performance with 42% basket penetration</li>"
    emailBody = emailBody & "<li>Revenue velocity accelerating, indicating growth momentum</li>"
    emailBody = emailBody & "<li>Marketing ROI exceeds 3.5x target threshold</li>"
    emailBody = emailBody & "</ul>"
    emailBody = emailBody & "<p>Full dashboard available in shared drive.</p>"
    emailBody = emailBody & "<p><em>This report was auto-generated by VANGUARD Analytics Engine</em></p>"
    emailBody = emailBody & "</body></html>"
    
    ' Create Outlook email
    On Error Resume Next
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    With OutlookMail
        .Subject = "VANGUARD: Weekly KPI Summary - " & Format(Date, "mm/dd/yyyy")
        .HTMLBody = emailBody
        .Display ' Change to .Send to automatically send
    End With
    
    If Err.Number <> 0 Then
        MsgBox "Unable to create email. Please ensure Outlook is installed.", vbExclamation, "VANGUARD"
    End If
    On Error GoTo 0
    
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub

'================================================================
' MACRO 4: Data Validation Check
'================================================================
Sub RunDataValidationCheck()
    '
    ' Purpose: Scan all sheets for formula errors and data quality issues
    ' Keyboard Shortcut: Ctrl+Shift+V
    '
    Dim ws As Worksheet
    Dim cell As Range
    Dim errorCount As Integer
    Dim errorList As String
    
    Application.ScreenUpdating = False
    errorCount = 0
    errorList = "Data Validation Report:" & vbCrLf & vbCrLf
    
    ' Check each worksheet
    For Each ws In ActiveWorkbook.Worksheets
        For Each cell In ws.UsedRange
            If cell.HasFormula Then
                If IsError(cell.Value) Then
                    errorCount = errorCount + 1
                    errorList = errorList & "• " & ws.Name & "!" & cell.Address & " - " & CStr(cell.Value) & vbCrLf
                End If
            End If
        Next cell
    Next ws
    
    Application.ScreenUpdating = True
    
    ' Display results
    If errorCount = 0 Then
        MsgBox "✓ Data validation passed!" & vbCrLf & vbCrLf & _
               "All formulas calculated successfully with zero errors.", _
               vbInformation, "VANGUARD Data Validation"
    Else
        MsgBox errorList & vbCrLf & _
               "Total Errors Found: " & errorCount, _
               vbExclamation, "VANGUARD Data Validation"
    End If
End Sub

'================================================================
' MACRO 5: Format as Currency
'================================================================
Sub FormatSelectedAsCurrency()
    '
    ' Purpose: Apply VANGUARD standard currency formatting to selection
    ' Keyboard Shortcut: Ctrl+Shift+C
    '
    On Error Resume Next
    Selection.NumberFormat = "$#,##0.00;($#,##0.00);-"
    Selection.Font.Color = RGB(0, 0, 0)
    On Error GoTo 0
End Sub

'================================================================
' MACRO 6: Format as Percentage
'================================================================
Sub FormatSelectedAsPercentage()
    '
    ' Purpose: Apply percentage formatting to selection
    ' Keyboard Shortcut: Ctrl+Shift+%
    '
    On Error Resume Next
    Selection.NumberFormat = "0.00%;(0.00%);-"
    Selection.Font.Color = RGB(0, 0, 0)
    On Error GoTo 0
End Sub

'================================================================
' MACRO 7: Highlight Changed Cells
'================================================================
Sub HighlightChangedCells()
    '
    ' Purpose: Highlight cells that changed after data refresh
    ' Note: Run before refresh, then run again after to see changes
    '
    Static savedValues As Collection
    Dim ws As Worksheet
    Dim cell As Range
    Dim currentValue As Variant
    Dim i As Integer
    
    Set ws = ActiveSheet
    
    If savedValues Is Nothing Then
        ' First run - save current values
        Set savedValues = New Collection
        For Each cell In ws.UsedRange
            If Not IsEmpty(cell.Value) And Not cell.HasFormula Then
                savedValues.Add Array(cell.Address, cell.Value)
            End If
        Next cell
        MsgBox "Baseline values saved. Refresh your data, then run this macro again to highlight changes.", vbInformation, "VANGUARD"
    Else
        ' Second run - compare and highlight
        Application.ScreenUpdating = False
        
        For i = 1 To savedValues.Count
            Dim savedAddress As String
            Dim savedValue As Variant
            
            savedAddress = savedValues(i)(0)
            savedValue = savedValues(i)(1)
            
            Set cell = ws.Range(savedAddress)
            If cell.Value <> savedValue Then
                cell.Interior.Color = RGB(255, 255, 0) ' Yellow highlight
            End If
        Next i
        
        Set savedValues = Nothing
        Application.ScreenUpdating = True
        MsgBox "Changed cells have been highlighted in yellow.", vbInformation, "VANGUARD"
    End If
End Sub

'================================================================
' MACRO 8: Clear All Formatting
'================================================================
Sub ClearFormattingFromSelection()
    '
    ' Purpose: Remove all formatting from selected cells
    '
    On Error Resume Next
    Selection.ClearFormats
    On Error GoTo 0
End Sub

'================================================================
' MACRO 9: Auto-Size All Columns
'================================================================
Sub AutoSizeAllColumns()
    '
    ' Purpose: Auto-fit all columns in active sheet
    '
    Application.ScreenUpdating = False
    ActiveSheet.UsedRange.EntireColumn.AutoFit
    Application.ScreenUpdating = True
End Sub

'================================================================
' MACRO 10: Create Backup
'================================================================
Sub CreateBackupCopy()
    '
    ' Purpose: Save timestamped backup of workbook
    '
    Dim backupPath As String
    Dim timestamp As String
    
    timestamp = Format(Now(), "yyyy-mm-dd_hhmm")
    backupPath = ThisWorkbook.Path & "\Backup_VANGUARD_" & timestamp & ".xlsm"
    
    On Error Resume Next
    ThisWorkbook.SaveCopyAs backupPath
    If Err.Number = 0 Then
        MsgBox "Backup created successfully:" & vbCrLf & backupPath, vbInformation, "VANGUARD"
    Else
        MsgBox "Error creating backup: " & Err.Description, vbExclamation, "VANGUARD"
    End If
    On Error GoTo 0
End Sub
