Attribute VB_Name = "ManageByCleaning"
Sub cleanListObjects()

Dim aws As Worksheet
Dim tbl As ListObject

For Each aws In ActiveWorkbook.Worksheets
For Each tbl In aws.ListObjects
    
    tbl.Delete
    
Next
Next

End Sub


Sub cleanConnections()

    ' Set the name of fantom connection for storing Data Model
    fantomConnection = "ThisWorkbookDataModel"

    Dim wbc As WorkbookConnection
    
    For Each wbc In ActiveWorkbook.Connections
        If wbc.Name <> fantomConnection Then initialConnections = initialConnections + 1
    Next
    
    
    If initialConnections > 0 Then
    
    Dim stepByStep As Boolean
        
        If initialConnections = 1 Then
            
            If ActiveWorkbook.Connections(1).Name = fantomConnection Then
            
                connectionToDelete$ = ActiveWorkbook.Connections(2)
            Else
                connectionToDelete$ = ActiveWorkbook.Connections(1)
            End If
            
            If MsgBox("1 connection deletable : " & Chr(10) & _
                      "Do you want to delete : " & connectionToDelete & " ?", _
                        vbYesNo, "Connection cleaning") = vbYes Then
                        
                ActiveWorkbook.Connections(connectionToDelete).Delete
            Else
                stepByStep = True
            End If
                
        Else
        ' Way to delete
        stepByStep = False
    
            If MsgBox(initialConnections & " connections deletable : " & Chr(10) & _
                      "Do you want confirm for Each one before deleting ?", vbYesNo, "Deleting mode") = vbYes Then
                stepByStep = True
            End If
        
            For Each wbc In ActiveWorkbook.Connections
                
                On Error Resume Next
                If Not wbc Is Nothing Then
                
                ' Skip fantom connection treatment
                If wbc.Name <> fantomConnection Then
                        
                    ' Step by step deleting
                    If stepByStep = True Then
                        
                        dialogAnswer = MsgBox("Do you want to delete : " & wbc.Name & " ?", _
                                                vbYesNoCancel, "Connection cleaning")
                                                
                        If dialogAnswer = vbYes Then wbc.Delete
                        
                        If dialogAnswer = vbCancel Then Exit For
    
                    ' One shot deleting
                    Else
                        wbc.Delete
                    End If
            
            End If
            End If
            Next
                
        End If
    
    Else
    
        If ActiveWorkbook.Connections.Count = 1 Then
            If ActiveWorkbook.Connections(1).Name = fantomConnection Then _
                        MsgBox "No connection deletable" & Chr(10) & _
                                "The connection " & fantomConnection & " is persisting but should be empty"
        Else
                        MsgBox "No connections deletable"
        End If
        
    End If
    
    
    If initialConnections > 0 Then
    
        ' If there were some connections and all have been deleted
        If ActiveWorkbook.Connections.Count = 0 Then
        
                        MsgBox "Every connection cleared"
        Else
        
            ' If only the fantom connection remains
            If ActiveWorkbook.Connections.Count = 1 And _
                ActiveWorkbook.Connections(1).Name = fantomConnection _
            Then
                
                        MsgBox "Every connection cleared" & Chr(10) & _
                                "The connection " & fantomConnection & " is persisting but should be empty"
                        
            ' count again with skiping the fantom Connection
            Else
            
                For Each wbc In ActiveWorkbook.Connections
                    If wbc.Name <> fantomConnection Then finalConnections = finalConnections + 1
                Next
            
                ' If any other connections persist after deleting in one shot, then avoid to quote the fantom connection
                If stepByStep = False Then
            
                    If ActiveWorkbook.Connections(1).Name = fantomConnection Then
                    
                        MsgBox "1/" & finalConnections & ": The connection " & _
                                ActiveWorkbook.Connections(2) & " is persisting"
                    Else
                        MsgBox "1/" & finalConnections & ": The connection " & _
                                ActiveWorkbook.Connections(1) & " is persisting"
                    End If
                    
                ' If any other connections persist after deleting step by step
                Else
                        MsgBox finalConnections & " connections have been kept out of " & initialConnections
                End If
                    
            End If
        End If
     
    End If
        
End Sub

