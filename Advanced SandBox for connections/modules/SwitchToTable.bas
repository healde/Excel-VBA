Attribute VB_Name = "SwitchToTable"
Sub CreateListObject()
' In this sub : thisWB.getExtSource, thisWB.getTargetSheet, thisWB.getTable, getPosArray function
ThisWorkbook.ResetSettings
 
 
Dim wb As Workbook
Set wb = ActiveWorkbook
'Set wb = Workbooks(ThisWorkbook.getExtSource)

Dim ws As Worksheet
Set ws = ActiveSheet
'Set ws = wb.Worksheets(ThisWorkbook.getTargetSheet)


' Set a default table if the sheet is Empty
If IsEmpty(ws.UsedRange) Then
    
    'mySheets = Array("sheetOne", "sheetTwo", "sheetThree")
    'myArray = Array("index", "label", "info", "table")
    myArray = ThisWorkbook.getTable(posArray)
    
    posArray = getPosArray(ThisWorkbook.getSheets, ThisWorkbook.getTargetSheet)
    
    If posArray > -1 And UBound(myArray) > -1 Then
    
    For i = 0 To UBound(myArray, 1)
    
        ws.Cells(1, i + 1) = myArray(i)
        Next
    End If
    
End If



Dim rng As Range
On Error GoTo ErreurRange
Set rng = Range("[" & wb.Name & "]" & ws.Name & "!" & ws.Name & "TabRef")

On Error GoTo 0
' Application.Goto Reference:="" & ws.Name & "TabRef"

    Dim tbl As ListObject
    
On Error Resume Next ' another error 1004 on trying to set a data table over an existing one
    Set tbl = _
                ws.ListObjects.Add( _
                SourceType:=xlSrcRange, _
                Source:=rng, _
                XlListObjectHasHeaders:=xlYes)
                 
        tbl.Name = "MyDataTable"
        
Exit Sub ' Exit to avoid the following error handler.

ErreurRange:
 Select Case Err.Number
 Case 1004
    With ws.UsedRange.Resize(RowSize:=ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1)
                
        .value = .value
        .Name = "" & ws.Name & "TabRef"
        End With
        
    Case Else
       MsgBox "error number : " & Err.Number
       Exit Sub
    End Select
    
Err.Clear
Resume

 
End Sub


Sub DeleteListObject()

Dim wb As Workbook
Set wb = ActiveWorkbook

Dim ws As Worksheet
Set ws = ActiveSheet

Dim back As Range
Set back = Selection
            
Dim tbl As ListObject
Dim nam As Name

With ws
errorFlag = False
Application.ScreenUpdating = False
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row + 1
    
        ws.UsedRange.Copy
        ws.Cells(lastRow + 1, 1).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, SkipBlanks:=True
        Application.CutCopyMode = False
    
    For Each tbl In .ListObjects
        tbl.Delete
    Next
    
        back.Select ' come back to selection must be done before deleting rows in case the range is concerned
        Rows("1:" & lastRow).Delete
    
    For Each nam In wb.Names
            
            If nam Like "*[#]*" Then
                ' MsgBox "error, skipping " & nam
                
            Else
                On Error GoTo ErreurRef
                If nam.RefersToRange.Parent.Name = .Name Then nam.Delete
                errorFlag = False
                
            End If
        Next nam
        On Error GoTo 0
       
Application.ScreenUpdating = True
End With
Exit Sub

ErreurRef:

    If errorFlag = True Then
    
        Resume
    Else
        MsgBox "error number " & Err.Number & ", try to delete " & nam
        On Error Resume Next
        nam.Delete
        
        errorFlag = True
    End If
    
Err.Clear
Resume

End Sub

