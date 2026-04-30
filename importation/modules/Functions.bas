Attribute VB_Name = "Functions"
' Two next functions are used as error fixers in importation sub routines


Function errorCommand(state As Boolean) As Boolean
errorCommand = True

    ' ".Connections.add" triggers error 5
    If Err.Number = 5 Then
        If ThisWorkbook.getCommandType <> 3 Then
            MsgBox "Error while xlCmdType = " & ThisWorkbook.getCommandType & Chr(10) & _
                    "Trying again with setting xlCmdType to 3"
            ThisWorkbook.setCommandType = 3
        End If
    Else
       MsgBox "Add method : Error number " & Err.Number & Chr(10) & Err.Description
    End If
    
Err.Clear
End Function


Function errorName(ByVal toIndent As String) As String

' ".ListObject.DisplayName" triggers error 1004 while ".WorkbookConnection" triggers error 5
If Err.Number = 5 Or Err.Number = 1004 Then
    
        If toIndent = "" Then
            errorName = "_1"
            
    ElseIf toIndent = "_15" Then
    
            MsgBox "Name affectation : Error number " & Err.Number & Chr(10) & _
                    "Tried over 15 indentations : maximum reached or problem elsewhere"
            errorName = ""
            
        Else
            errorName = "_" & getOnlyIndex(toIndent) + 1
        End If
        
    Else
            MsgBox "Name affectation : Error number " & Err.Number & Chr(10) & Err.Description
            errorName = ""
    End If

Err.Clear
End Function




' All next Functions are currently used in ManageQueries()


' shared between ManageQueries(), DeleteListObject(), and ThisWB.setColumnValu
Function getPosArray(ByRef thisArray As Variant, ByVal thisString As String) As Integer
    
    getPosArray = -1
    
    If Not IsEmpty(thisArray) Then
    If UBound(thisArray) > -1 And thisString <> "" Then

        For i = 0 To UBound(thisArray)
        
        If StrComp(thisString, thisArray(i), vbTextCompare) = 0 Then
        
            getPosArray = i
            Exit For
            End If
        Next
    
    End If
    End If

End Function



Function getIndexLimits(elements As Variant) As Integer()

Dim element As Variant
Dim result(0 To 2) As Integer

quantity = 0    ' set number of index found
lowest = 32767  ' start with maximum integer value
biggest = 0     ' start with minimum index value
                              
                              
For Each element In elements
    
    indexGot$ = getOnlyIndex(element)
    
    If Not indexGot = "" Then
    
        indexInt% = CInt(indexGot)
        
        quantity = quantity + 1
        If indexInt < lowest Then lowest = indexInt
        If indexInt > biggest Then biggest = indexInt
    End If
    
Next

    If biggest = 0 Then lowest = 0
    result(1) = lowest
    result(2) = biggest
    result(0) = quantity
    
    getIndexLimits = result
                
End Function


' shared between ManageQueries(), and errorName() Function
Function getOnlyIndex(ByVal originame As String) As String

    matchIn = 0   ' reset the flag in
    matchOut = 0  ' reset the flag out

    For i = 1 To Len(originame)
        If Mid(originame, i) Like "#*" Then
        
            matchIn = 1
            Exit For
        End If
    Next
    
        If i < Len(originame) Then

    For j = i + 1 To Len(originame)
        If Mid(originame, j) Like "[!0-9]*" Then
        
            matchOut = 1
            Exit For
        End If
    Next
            j = j - i + 1
        Else
            j = 1
        End If
    
    If matchOut = 1 Then j = j - 1
    If matchIn = 1 Then
    
        getOnlyIndex = Mid(originame, i, j)
        
    Else
        getOnlyIndex = ""
        
    End If

End Function



Function QueryExists(qName As String, Queries As Object) As Boolean

    Dim q As Object
    Err.Clear
    
    On Error Resume Next
    Set q = Queries.Item(qName)
    QueryExists = (Err.Number = 0)
    
    Err.Clear
    
End Function
