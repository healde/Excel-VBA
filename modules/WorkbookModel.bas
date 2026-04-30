Attribute VB_Name = "WorkbookModel"
Sub ExportAsDistant()
' In this sub : thisWB.getExtSource, thisWB.getSheets, thisWB.getTable :: thisWB.getLimitArray function
ThisWorkbook.ResetSettings


'    chemin = ThisWorkbook.Path & "\distant3"                           ' Example without global variable : toggle comment
    chemin = ThisWorkbook.getExtSource
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

        Dim ws As Worksheet
        Dim rng As Range
        Dim str As Variant
        
'        extSheets = Array("sheetOne", "sheetTwo", "sheetThree")        ' Example without global variable : toggle comment
        extSheets = ThisWorkbook.getSheets
        
        ' ------------------------------------------------
        
            Dim extColumn As Variant
            
'            ReDim extColumn(0 To 2, 0 To 3)                            ' Example without global variable : toggle comment
            
'                underArrA = Array("index", "label", "info", "refer")
'                For i = 0 To UBound(underArrA)
'                    extColumn(0, i) = underArrA(i)
'                Next
'
'                underArrB = Array("index", "label", "info", "local")
'                For i = 0 To UBound(underArrB)
'                    extColumn(1, i) = underArrB(i)
'                Next
'
'                underArrC = Array("index", "label", "info", "scope")
'                For i = 0 To UBound(underArrC)
'                    extColumn(2, i) = underArrC(i)
'                Next

            Dim colHeader As Variant
            ReDim extColumn(ThisWorkbook.getLimitArray(1), ThisWorkbook.getLimitArray(2))
            
            ' After setting the table frame, Get each table to an array and transfer it into a new line of this multidimensional array
            For j = 0 To UBound(extColumn, 1)

                    i = 0
                If UBound(ThisWorkbook.getTable(j)) > -1 Then
                For Each colHeader In ThisWorkbook.getTable(j)

                    extColumn(j, i) = colHeader
                    i = i + 1
                Next
                End If
                
                Next
                
        ' ------------------------------------------------
        
            For j = 0 To UBound(extColumn, 1)
            
                str = extSheets(j)
                
                    result = ""
                    On Error Resume Next
                    result = (Worksheets(str).Name = str)
                
                    If result = "" Then
                        Set ws = Worksheets.Add
                        ws.Name = str
                        
                    Else
                        Set ws = Worksheets(str)
                        ws.Visible = True
                    End If
                    
                    
                ws.Cells.ClearContents
                
                For i = 0 To UBound(extColumn, 2)
                    ws.Cells(1, i + 1) = extColumn(j, i)
                Next
            Next
            
        ' ------------------------------------------------
        
        
    Worksheets(extSheets).Copy
    
    With ActiveWorkbook
    
        For Each ws In .Worksheets
        
            With ws.UsedRange.Resize(RowSize:=1)
            
                .value = .value
                .Name = ws.Name & "Range"
                End With
        Next
    
         .SaveAs Filename:=chemin, FileFormat:=xlOpenXMLStrictWorkbook
         .Close SaveChanges:=False
    End With
    
    
        ' ------------------------------------------------
    
        For Each str In extSheets
            
            Set ws = Worksheets(str)
                
            ' Worksheets(str).Visible = False
        Next
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True


End Sub


