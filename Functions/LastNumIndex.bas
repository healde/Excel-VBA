Sub testgetIndexLimits()

    Dim myResult() As Integer
    Dim myString() As Variant
    myString = Array("espece654", "terrain45aaa", "reintroduction4896", "observation456", "essai sans digit")
    
    Count = 5
    Dim nam As Worksheet
    
    For Each ws In ActiveWorkbook.Sheets

        Count = Count + 1
        ReDim Preserve myString(Count)
        myString(Count) = ws.Name
    Next
    
    myResult = getIndexLimits(myString)
    
    MsgBox " " & myResult(0) & " elements [ " & myResult(1) & " ; " & myResult(2) & " ]"
    
    indexGot$ = getOnlyIndex(myString(0))
    
    If Not indexGot = "" Then
        indexInt% = CInt(indexGot)
        MsgBox myString(0) & " -> " & Replace(myString(0), indexGot, indexGot + 1, Count:=1)
    Else
        MsgBox "Get """ & indexGot & " "" from """ & myString(0) & """ > Vartype: " & VarType(indexGot)
    End If
    
End Sub

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
