Attribute VB_Name = "AlanBerrie"

' Alan Berrie NCR Export date format

Private Sub NCR_dateformat()

    Dim MyRange As Range
    Dim thisCell As Range
    Dim isCellEmpty As Boolean
    Dim missingRows As Integer, missingRowsToBreak As Integer
    Dim prevRow As Integer, currRow As Integer, scrollRow As Integer
    
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    'Remember time when macro starts
    StartTime = Timer
    
    Set MyRange = Selection
    missingRowsToBreak = 20
    missingRows = 0
    prevRow = 1
    currRow = 1
    scrollRow = 1
    
    Application.ScreenUpdating = False
    
    For Each thisCell In MyRange.Cells
        currRow = thisCell.row
        isCellEmpty = IsEmpty(thisCell)
        If isCellEmpty = True Then
            missingRows = missingRows + 1
            If missingRows > missingRowsToBreak Then Exit For
        Else
            If currRow - prevRow >= scrollRow Then
                Application.ScreenUpdating = True
                missingRows = 0
                prevRow = currRow
                ActiveWindow.scrollRow = thisCell.row
                Application.ScreenUpdating = False
            End If
            thisCell.NumberFormat = "dd-mmm-yyyy;@"
            thisCell.Select
            Application.SendKeys "{F2}", True
            Application.SendKeys "{ENTER}", True
            DoEvents
        End If
    Next
    
    Application.ScreenUpdating = True
    
    'Determine how many seconds code took to run
    SecondsElapsed = Round(Timer - StartTime, 2)
    
    'Notify user in seconds
    MsgBox "Completed in " & SecondsElapsed & " seconds", vbInformation

End Sub

