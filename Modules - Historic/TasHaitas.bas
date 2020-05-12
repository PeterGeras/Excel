Attribute VB_Name = "TasHaitas"

' ------------------------------ '
' --- MACRO REQUESTED BY TAS --- '
' ------------------------------ '


' --- ASSIGN GLOBAL CONSTANTS --- '

Private Const startRow = 6              'Row in outputSheet to drag down formulas
Private Const startCol = "A"            'Left-most column in outputSheet of formulas
Private Const endCol = "U"              'Right-most column in outputSheet of formulas
Private Const defaultMissingRowsToBreak = 10    'Will check in dataSheet for this many empty rows at end
                                                'before deciding that it is the end of the sheet
Private Const dataSheetStartString = "OBJECTNUMBER"   'Table string start in dataSheet
Private Const diffStartData = 6        'Row number start of data in dataSheet
Private Const outputSheetStartString = "Line #"   'Table string start in outputSheet

'Sheet parameters
Private outputSheet As Worksheet
Private dataSheet As Worksheet
Private outputSheetRowCount As Long

'Assign
Private Sub AssignGlobals()
    Set outputSheet = Sheets("Validation Sheet") 'Sheets(1) refers to first sheet
    Set dataSheet = Sheets("OperationsBody")   'Sheets("name") refers to name of sheet
    outputSheetRowCount = CountRows(outputSheet)
End Sub


' ---------- MAIN ---------- '
Private Sub Validation_Sheet_Shape()

    Dim wb As Workbook
    Dim lastRow As String
    
    Set wb = ActiveWorkbook
    wb.Activate
    
    AssignGlobals
    
    Application.Calculation = xlCalculationManual
    
    'Clear filter
    On Error Resume Next
    outputSheet.ShowAllData
    
    lastRow = "" & endCol & outputSheetRowCount
    
    outputSheet.Range(lastRow).Select
    If ContinueMessage = False Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ShapeTables
    
    outputSheet.Calculate
    Application.ScreenUpdating = True
    
    outputSheet.Range("A1").Select
    
    Application.Calculation = xlCalculationAutomatic
    
End Sub


Private Sub ShapeTables()

    Dim dataStartPos As Long, outputStartPos As Long
    Dim dataSheetRows As Long, outputSheetRows As Long
    Dim numRowsToAdd As Long
    
    dataStartPos = FindTableStart(dataSheet, dataSheetStartString)
    outputStartPos = FindTableStart(outputSheet, outputSheetStartString)
    
    dataSheetRows = CountRows(dataSheet, dataStartPos)
    outputSheetRows = CountRows(outputSheet, outputStartPos)
    
    'msg = msgBoxDot("dataStartPos", "" & dataStartPos)
    'msg = msg & msgBoxDot("outputStartPos", "" & outputStartPos)
    'msg = msg & msgBoxDot("dataSheetRows", "" & dataSheetRows)
    'msg = msg & msgBoxDot("outputSheetRows", "" & outputSheetRows)
    'MsgBox (msg)
    
    numRowsToAdd = dataSheetRows - dataStartPos
    
    DeleteRows outputStartPos + 1, outputSheetRows
    AddRows outputStartPos + 1, outputStartPos + numRowsToAdd
    FormulaDragger startCol, outputStartPos, endCol, outputStartPos + numRowsToAdd
    
End Sub

Function CountRows(sh As Worksheet, Optional startRow As Long = 1) As Long

    Dim currRow As Long
    Dim isCellEmpty As Boolean
    Dim missingRows As Integer

    isCellEmpty = True
    
    currRow = startRow
    missingRows = 0
    
    While missingRows < defaultMissingRowsToBreak
        isCellEmpty = IsEmpty(sh.Cells(currRow, 1))
        currRow = currRow + 1
        If isCellEmpty = True Then
            missingRows = missingRows + 1
        Else
            missingRows = 0
        End If
    Wend
    
    CountRows = currRow - defaultMissingRowsToBreak - 1
    
End Function

Function FindTableStart(sh As Worksheet, findStr As String, Optional currRow As Long = 1) As Long

    Dim missingRows As Long
    
    missingRows = 0
    
    If missingRowsToBreak < defaultMissingRowsToBreak Then
        missingRowsToBreak = defaultMissingRowsToBreak
    End If
    
    While missingRows < missingRowsToBreak
        If sh.Cells(currRow, 1).Value = findStr Then
            FindTableStart = currRow + 1
            Exit Function
        ElseIf IsEmpty(sh.Cells(currRow, 1)) = True Then
            missingRows = missingRows + 1
        Else
            missingRows = 0
        End If
        currRow = currRow + 1
    Wend
    
    FindTableStart = currRow - defaultMissingRowsToBreak - 1
    
End Function

Private Sub DeleteRows(xRow As Long, yRow As Long)

    If yRow < xRow Then Exit Sub

    Dim sh As Worksheet
    Dim rowDeleteString As String
    
    Set sh = outputSheet
    rowDeleteString = "" & xRow & ":" & yRow
    
    sh.Rows(rowDeleteString).Delete Shift:=xlUp

End Sub

Private Sub AddRows(xRow As Long, yRow As Long)
    
    If yRow < xRow Then Exit Sub
    
    Dim sh As Worksheet
    Dim rowAddString As String
    
    Set sh = outputSheet
    rowAddString = "" & xRow & ":" & yRow
    
    sh.Rows(rowAddString).Insert Shift:=xlDown

End Sub


Private Sub FormulaDragger(xCol As String, xRow As Long, yCol As String, yRow As Long)
    
    If yRow <= xRow Then Exit Sub
    
    Dim sh As Worksheet
    Dim initialRange As String
    Dim startRange As String
    Dim endRange As String
    
    Set sh = outputSheet
    initialRange = xCol & xRow & ":" & yCol  'A1:F
    startRange = initialRange & xRow   'A1:F1
    endRange = initialRange & yRow   'A1:F9
    
    sh.Range(startRange).AutoFill Destination:=Range(endRange), Type:=xlFillDefault
    
End Sub

Function ContinueMessage() As Boolean

    Dim strMsg As String
    
    strMsg = "" & msgBoxDot("Validation worksheet name", outputSheet.name)
    strMsg = strMsg & msgBoxDot("Last row in sheet", "" & outputSheetRowCount)
    
    ContinueMessage = True
    
    If MsgBox(strMsg & vbNewLine & "Are these details correct?", vbYesNo) = vbNo Then ContinueMessage = False
    
End Function

Private Sub WarningMessage(name As String, Optional row As Long = 0)
    
    Dim strMsg As String
    
    strMsg = "" & "WARNING!!!" & vbNewLine
    strMsg = strMsg & msgBoxDot("Expected worksheet missing", name)
    
    If row > 0 Then
        strMsg = strMsg & msgBoxDot("Expected order number", "A" & row)
    End If
    
    MsgBox strMsg

End Sub

Function msgBoxDot(msgA As String, msgB As String) As String

    msgBoxDot = msgA & " " & Chr(149) & " " & msgB & vbNewLine
    
End Function


