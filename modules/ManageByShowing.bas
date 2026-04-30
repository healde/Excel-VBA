Attribute VB_Name = "ManageByShowing"
Sub showListobjects()

    If ActiveSheet.ListObjects.Count > 0 Then
        dialogAnswer = MsgBox("ListObject 1 : " & ActiveSheet.ListObjects(1) & Chr(10) & "Do you want search for other ?", _
                        vbYesNo, "First ListObject in this sheet")
    Else
        dialogAnswer = vbNo
    End If
    
    If dialogAnswer = vbYes Then
    
        Dim aws As Worksheet
        Dim tbl As ListObject
        
        For Each aws In ActiveWorkbook.Worksheets
            If dialogAnswer = vbNo Then Exit For
        
        For Each tbl In aws.ListObjects
        
            If dialogAnswer = vbYes Then
            
                dialogAnswer = MsgBox("ListObject found in " & aws.Name & " : " & tbl.Name & Chr(10) & "Display next ?", _
                    vbYesNo, "ListObjects in this workbook")
            Else
                If tbl.Name <> aws.ListObjects(1) Then Exit For
            End If
            
        Next
        Next
    End If

End Sub

Sub showConnections()
    
    dialogQuestion = "Do you want display names ?"
    dialogButton = vbYesNo
    
    Dim wcs As Connections
    Set wcs = ActiveWorkbook.Connections
    If wcs.Count = 0 Then
    
        dialogQuestion = ""
        dialogButton = vbOKOnly
    End If
    
    dialogAnswer = MsgBox(wcs.Count & " connections found" & Chr(10) & dialogQuestion, _
                    dialogButton, "Number of connections")
    
    If dialogAnswer = vbYes Then
    
    dialogQuestion = "Do you want display next ?"
    Dim wbc As WorkbookConnection
    
    i = 1
    For Each wbc In wcs
    
        If i = wcs.Count Then
            dialogQuestion = ""
            dialogButton = vbOKOnly
        Else
            i = i + 1
        End If
    
        dialogAnswer = MsgBox("Connection found : " & wbc.Name & Chr(10) & dialogQuestion, _
                        dialogButton, "Connections")
        If dialogAnswer = vbNo Then Exit For
        
    Next
    End If

End Sub

