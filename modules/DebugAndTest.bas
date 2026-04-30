Attribute VB_Name = "DebugAndTest"
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



Sub myWeirdestDebug()

Dim lio As ListObject
Dim conn As WorkbookConnection
Dim connA As WorkbookConnection
Dim connB As WorkbookConnection

Count = 0

On Error Resume Next
For Each lio In ActiveSheet.ListObjects

    Count = Count + 1
    
    sot = "lio_SourceType"
    objA = "ConnA is Empty"
    objB = "ConnB is Empty"
    
    Debug.Print Chr(10)
    
For Each conn In ActiveWorkbook.Connections
'    ' CORRECT STATEMENT BELOW
'    If lio.QueryTable.WorkbookConnection.Name = conn.Name Then Debug.Print " v" & Chr(10) & Count & ".TRUE", lio.SourceType & "QT filtered", lio.QueryTable.WorkbookConnection.Name & " = " & conn.Name
'    If lio.TableObject.WorkbookConnection.Name = conn.Name Then Debug.Print " v" & Chr(10) & Count & ".TRUE", lio.SourceType & "TO filtered", lio.TableObject.WorkbookConnection.Name & " = " & conn.Name
'
'    ' Next lines are ERRORS of logical relations
'    If lio.QueryTable.WorkbookConnection.Name = conn.Name Then Debug.Print Count & ".FALSE", lio.SourceType & "TO", lio.TableObject.WorkbookConnection.Name & " = " & conn.Name
'    If lio.TableObject.WorkbookConnection.Name = conn.Name Then Debug.Print Count & ".FALSE", lio.SourceType & "QT", lio.QueryTable.WorkbookConnection.Name & " = " & conn.Name
'
'    ' These lines for encapsing the two lines which are displayed in the same loop
'    If lio.QueryTable.WorkbookConnection.Name = conn.Name Then Debug.Print Chr(CInt(Not (lio.QueryTable.WorkbookConnection.Name = conn.Name)) * 10) & "^"
'    If lio.TableObject.WorkbookConnection.Name = conn.Name Then Debug.Print Chr(CInt(Not (lio.TableObject.WorkbookConnection.Name = conn.Name)) * 10) & "^"
Next


' 3 <-> connA  from .QueryTable
If lio.SourceType = 3 Then sot = "QT"
Set connA = ActiveWorkbook.Connections(lio.QueryTable.WorkbookConnection.Name)

If IsEmpty(lio.TableObject.WorkbookConnection.Name) = False Then objA = ".FALSE"
If IsEmpty(lio.QueryTable.WorkbookConnection.Name) = False Then objA = "." & (IsEmpty(lio.QueryTable.WorkbookConnection.Name) = False)
'If IsEmpty(lio.TableObject.WorkbookConnection.Name) = False Then objA = ".?? ConnA"


' 4 <-> connB  from .TableObject
If lio.SourceType = 4 Then sot = "TO"
Set connB = ActiveWorkbook.Connections(lio.TableObject.WorkbookConnection.Name)

If IsEmpty(lio.TableObject.WorkbookConnection.Name) = False Then objB = ".FALSE"
If IsEmpty(lio.TableObject.WorkbookConnection.Name) = False Then objB = "." & (IsEmpty(lio.TableObject.WorkbookConnection.Name) = False)
'If IsEmpty(lio.TableObject.WorkbookConnection.Name) = False Then objB = ".?? connB"


Debug.Print Chr(10)
Debug.Print Count & objA, lio.SourceType & sot & " result", connA.Name, , "As WorkbookConnection"
Debug.Print , "'-> from", lio.QueryTable.WorkbookConnection.Name, , "As Name Property"
Debug.Print
Debug.Print Count & objB, lio.SourceType & sot & " result", connB.Name, , "As WorkbookConnection"
Debug.Print , "'-> from", lio.TableObject.WorkbookConnection.Name, , "As Name Property"
Debug.Print Chr(10)

Next
End Sub

