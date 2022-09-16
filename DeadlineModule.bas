Attribute VB_Name = "DeadlineModule"
Private Sub check_deadlines()
On Error Resume Next 'Avoiding the code stopping in case a cell doesn't match the desired date format.

'Declaring variables
    Dim ws As Worksheet
    Dim xLastRow As Integer
    Dim iDays As Integer
    Dim dte As Date
    Dim closeDeadlines As String
    Dim Chapter As String
    Dim FirstCell As String
    Dim index(2, 100) As String
    Dim i As Integer
    Dim cdTracker As Integer
    Dim closeDLs(10) As String
    
    FirstCell = "D2"
    dtToday = Date
    cdTracker = 0

'Running a loop for each worksheet in the workbook.
    For Each ws In ThisWorkbook.Sheets
        ws.Activate
        xLastRow = Cells(Rows.Count, "D").End(xlUp).row 'Noting the last row with data in the worksheet to identify the relevant range.
        LastCell = "D" & xLastRow
        i = 2 'Tracking the row number
        
    'Running a loop for each cell in the worksheet.
        For Each c In Range(FirstCell, LastCell)
            dte = c
            iDays = dte - Date 'Noting the number of days between deadline and current date.
            If iDays < 7 And Range("E" & i) <> "Yes" Then 'Checking if days to deadline is less than 7 and if the relevant material is already read.
                Book = Range("A" & i)
                Chapter = Range("B" & i)
                cdTracker = cdTracker + 1 'Tracking the number of items meeting the criteria.
                closeDLs(cdTracker) = Book & " " & Chapter & " should be read within " & iDays & " days."
            End If
            i = i + 1
        Next c
    Next ws
    
    'Removing the items in the array without data to provide cleaner message box.
    Dim newCloseDLs()
    For x = 1 To 20
        If closeDLs(x) <> "" Then
            displayText = displayText & vbCrLf & closeDLs(x)
        End If
    Next x
    
    MsgBox displayText
End Sub


