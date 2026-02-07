Public strFileName As String
Public currentWB As Workbook
Public dataWB As Workbook
Public strCopyRange As String

Sub GetData()
    Dim strWhereToCopy As String, strStartCellColName As String
    Dim strListSheet As String
    
    strListSheet = "List"
    
    On Error GoTo ErrH
    Sheets(strListSheet).Select
    Range("B2").Select
    
    'this is the main loop, we will open the files one by one and copy their data into the masterdata sheet
    Set currentWB = ActiveWorkbook
    Do While ActiveCell.Value <> ""
        
        strFileName = ActiveCell.Offset(0, 1) & ActiveCell.Value
        strCopyRange = ActiveCell.Offset(0, 2) & ":" & ActiveCell.Offset(0, 3)
        strWhereToCopy = ActiveCell.Offset(0, 4).Value
        strStartCellColName = Mid(ActiveCell.Offset(0, 5), 2, 1)
        
        Application.Workbooks.Open strFileName, UpdateLinks:=False, ReadOnly:=True
        Set dataWB = ActiveWorkbook
        
        Range(strCopyRange).Select
        Selection.Copy
        
        currentWB.Activate
        Sheets(strWhereToCopy).Select
        lastRow = LastRowInOneColumn(strStartCellColName)
        Cells(lastRow + 1, 1).Select
        
        Selection.PasteSpecial xlPasteValues, xlPasteSpecialOperationNone
        Application.CutCopyMode = False
        dataWB.Close False
        Sheets(strListSheet).Select
        ActiveCell.Offset(1, 0).Select
    Loop
    Exit Sub
    
ErrH:
    MsgBox "It seems some file was missing. The data copy operation is not complete."
    Exit Sub
End Sub

Public Function LastRowInOneColumn(col)
  
    Dim lastRow As Long
    With ActiveSheet
    lastRow = .Cells(.Rows.Count, col).End(xlUp).Row
    End With
    LastRowInOneColumn = lastRow
End Function


