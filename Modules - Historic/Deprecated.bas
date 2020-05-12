Attribute VB_Name = "Deprecated"
' ----- DEPRECATED SUBS -----


' CTRL+SHIFT + T
Private Sub trim_values()

    Dim MyRange As Range
    Dim thisCell As Range
    Dim isCellEmpty As Boolean
    Dim missingRowsToBreak As Integer
    Dim missingRows As Integer
    Dim trimmedCell As String

    Set MyRange = Selection
    trimmedCell = ""
    missingRowsToBreak = 1000
    missingRows = 0
    
    For Each thisCell In MyRange.Cells
        isCellEmpty = IsEmpty(thisCell)
        If isCellEmpty = True Then
            missingRows = missingRows + 1
            If missingRows > missingRowsToBreak Then Exit For
        Else
            missingRows = 0
            trimmedCell = "TRIM(SUBSTITUTE(""" & thisCell.Value & """,CHAR(160),CHAR(32)))"
            thisCell.Formula = "= " & trimmedCell
        End If
    Next
    
    Application.CutCopyMode = False
    MyRange.Cells(1, 1).Select
    
End Sub

