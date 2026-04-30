Attribute VB_Name = "ManageQueries"
Sub DoQueries()
' In this sub : thisWB.getExtSource, thisWB.getTargetSheet, thisWB.getTable, getPosArray function >> thisWB.getQueryName
ThisWorkbook.ResetSettings


' Get position of the defaultSheet in the sheets variable array
posArray = getPosArray(ThisWorkbook.getSheets, ThisWorkbook.getTargetSheet)


' Do mFormula if we get table from this index
If posArray > -1 And UBound(ThisWorkbook.getTable(posArray)) > -1 Then

Dim mFormula As String
mFormula = _
"let SourceRef = Excel.Workbook(File.Contents(""" & ThisWorkbook.getExtSource & """), null, true), " & _
    "DataRef = SourceRef{[Name=""" & ThisWorkbook.getTargetSheet & """]}[Data], " & _
    "#""Promoted headers"" = Table.PromoteHeaders(DataRef, [PromoteAllScalars=true]), " & _
    "#""Type modified"" = Table.TransformColumnTypes(#""Promoted headers"",{"
    
i = 0
s = UBound(ThisWorkbook.getTable(posArray))
For Each colHeader In ThisWorkbook.getTable(posArray)

    mFormula = mFormula & "{""" & colHeader & """, type text}"
    If i = s Then Exit For
    
    mFormula = mFormula & ", "
    i = i + 1
    Next

mFormula = mFormula & "}) in #""Type modified"""
    
    
    
With ActiveWorkbook

initialQueries = .Queries.Count

    If initialQueries > 0 Then
    
        Dim quy As WorkbookQuery
        Dim quyNames() As Variant
        Dim quyLimits() As Integer
        
        Count = 0
        For Each quy In .Queries
        
            Count = Count + 1
            ReDim Preserve quyNames(Count)
            quyNames(Count) = quy.Name
        Next
        
        quyLimits = getIndexLimits(quyNames)
        
        dialogAnswer = MsgBox("Do you want remove all the queries", vbYesNo + vbDefaultButton2, "Remove all")
        If dialogAnswer = vbYes Then
            
            While .Queries.Count > 0
            
                On Error Resume Next
                .Queries.Item(1).Delete
            Wend
            
        Else
            If QueryExists(ThisWorkbook.getQueryName, .Queries) Then
            
                dialogAnswer = MsgBox("Do you want replace the Query which get default index", vbYesNo + vbDefaultButton2, "Change default index")
                If dialogAnswer = vbYes Then
                
                    On Error Resume Next
                    .Queries.Item(ThisWorkbook.getQueryName).Delete
                    On Error Resume Next
                    query1 = ActiveWorkbook.Queries.Add(ThisWorkbook.getQueryName, mFormula)
                    
                Else
                If initialQueries > 1 Then
                
                dialogAnswer = MsgBox("Do you want remove only Query which get default index", vbYesNo + vbDefaultButton2, "Remove default index")
                If dialogAnswer = vbYes Then
                
                    On Error Resume Next
                    .Queries.Item(ThisWorkbook.getQueryName).Delete
                    
                End If
                End If
                
                End If
            Else
                
                dialogAnswer = MsgBox("Do you want add the default index Query", vbYesNo + vbDefaultButton2, "Change default index")
                If dialogAnswer = vbYes Then
                    
                    On Error Resume Next
                    query1 = ActiveWorkbook.Queries.Add(ThisWorkbook.getQueryName, mFormula)
                    
                End If
            End If
            
            dialogAnswer = MsgBox("Do you want add a Query with higher index", vbYesNo, "Change highest index")
            If dialogAnswer = vbYes Then
                
                    indexGot$ = getOnlyIndex(ThisWorkbook.getQueryName)
                    If Not indexGot = "" Then
                    
                        indexInt% = CInt(indexGot)
                        If indexInt > quyLimits(2) And QueryExists(ThisWorkbook.getQueryName, .Queries) Then
                        
                            newQuery$ = Replace(ThisWorkbook.getQueryName, indexGot, indexInt + 1)
                        Else
                            newQuery$ = Replace(ThisWorkbook.getQueryName, indexGot, quyLimits(2) + 1)
                        End If
                    Else
                            newQuery$ = ThisWorkbook.getQueryName & quyLimits(2) + 1
                    End If
                
                    On Error Resume Next
                    query2 = ActiveWorkbook.Queries.Add(newQuery, mFormula)
                
            End If
        End If
        
    End If
    If .Queries.Count = 0 Then
        
        If initialQueries = 0 Then
            dialogAnswer = MsgBox("No query : do you want add one ?", vbYesNo, "Add first query")
        Else
            dialogAnswer = MsgBox("Do you want add the default index query again ?", vbYesNo, "Add first query")
        End If
        If dialogAnswer = vbYes Then
            
            On Error Resume Next
            query1 = ActiveWorkbook.Queries.Add(ThisWorkbook.getQueryName, mFormula)
        End If
        
    End If
    
End With
End If

End Sub
