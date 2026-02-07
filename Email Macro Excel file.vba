Sub Send_Files()
'File to e-mail out the corrosponding sheets for Suppliers were needed
    Dim OutApp As Object
    Dim OutMail As Object
    Dim sh As Worksheet
    Dim CurrentRow As Long
    
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    Sheets("sheet1").Select
    Set sh = Sheets("sheet1")
    Set OutApp = CreateObject("Outlook.Application")
    
    For Each cell In sh.Columns("B").Cells.SpecialCells(xlCellTypeConstants)
    'For Each cell In sh.Columns("B").Cells.SpecialCells(xlCellTypeFormulas)
        'Enter the path/file names in the C column in each row
         CurrentRow = cell.Row
        If Cells(CurrentRow, 4) Like "?*@?*.?*" And _
        Cells(CurrentRow, 3) <> "" And Not Dir(Cells(CurrentRow, 3) & "") = "" Then 'Assumes Filepath is column C
            Set OutMail = OutApp.CreateItem(0)
		On Error Resume Next
 With OutMail
 
 .to = Cws.Cells(Rnum, 1).Value
 .Subject = "subject line here"
 .HTMLBody = StrBody & RangetoHTML(rng) & sBody 'changed
 .DeferredDeliveryTime = Delay
 .Send 'use Display or Send
 End With
 On Error GoTo 0

                'Assumes Name is column A
                .Attachments.Add Cells(CurrentRow, 3).Value
                .Send  'Or use .Display
            End With
            
            Cells(CurrentRow, 5) = Now
            Cells(CurrentRow, 6) = Environ("Username")
            Set OutMail = Nothing
        End If
    Next cell
    Set OutApp = Nothing
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    Sheets("Summary Sheet").Select
    MsgBox "Completed"
End Sub


.SentOnBehalfOfName = " shared email box name here"
