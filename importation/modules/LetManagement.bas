Attribute VB_Name = "LetManagement"
Sub AutoManagement()
' In these sub : thisWb.{get|set}{Delete|Show}{Listobject|Connection}Boolean

ThisWorkbook.ResetSettings

    ' ActiveWorkbook.Worksheets.Add
    Application.CutCopyMode = False
    
    If ThisWorkbook.getShowConnectionBoolean Then showConnections
    If ThisWorkbook.getCleanConnectionBoolean Then cleanConnections

    If ThisWorkbook.getShowListobjectBoolean Then showListobjects
    If ThisWorkbook.getCleanListobjectBoolean Then cleanListObjects
    
End Sub


Sub setManagement()
    
    If Not ThisWorkbook.getShowConnectionBoolean Then
        DefaultValue = vbDefaultButton2
    Else
        DefaultValue = 0
    End If
    dialogValue = MsgBox("Do you want set the macro ""show connections"" on ?", vbYesNo + DefaultValue, "showConnections boolean")
    ThisWorkbook.setShowConnectionBoolean = (dialogValue = vbYes)
    
    If Not ThisWorkbook.getCleanConnectionBoolean Then
        DefaultValue = vbDefaultButton2
    Else
        DefaultValue = 0
    End If
    dialogValue = MsgBox("Do you want set the macro ""clean connections"" on ?", vbYesNo + DefaultValue, "cleanConnections boolean")
    ThisWorkbook.setCleanConnectionBoolean = (dialogValue = vbYes)

    If Not ThisWorkbook.getShowListobjectBoolean Then
        DefaultValue = vbDefaultButton2
    Else
        DefaultValue = 0
    End If
    dialogValue = MsgBox("Do you want set the macro ""show Listobjects"" on ?", vbYesNo + DefaultValue, "showListobjects boolean")
    ThisWorkbook.setShowListobjectBoolean = (dialogValue = vbYes)
    
    If Not ThisWorkbook.getCleanListobjectBoolean Then
        DefaultValue = vbDefaultButton2
    Else
        DefaultValue = 0
    End If
    dialogValue = MsgBox("Do you want set the macro ""clean Listobjects"" on ?", vbYesNo + DefaultValue, "cleanListobjects boolean")
    ThisWorkbook.setCleanListobjectBoolean = (dialogValue = vbYes)
    
End Sub


Sub getManagement()

    If Not ThisWorkbook.getShowConnectionBoolean Then
        DefaultValue = vbDefaultButton2
    Else
        DefaultValue = 0
    End If
    dialogValue = MsgBox("Do you want to display connections ? (" & ActiveWorkbook.Connections.Count & ")", _
                        vbYesNoCancel + DefaultValue + vbQuestion, "showConnections macro")
    If dialogValue = vbCancel Then Exit Sub
    If dialogValue = vbYes Then showConnections
    
    If Not ThisWorkbook.getCleanConnectionBoolean Then
        DefaultValue = vbDefaultButton2
    Else
        DefaultValue = 0
    End If
    dialogValue = MsgBox("Do you want to clean connections ?", vbYesNoCancel + DefaultValue + vbQuestion, "cleanConnections macro")
    If dialogValue = vbCancel Then Exit Sub
    If dialogValue = vbYes Then cleanConnections
    
    If Not ThisWorkbook.getShowListobjectBoolean Then
        DefaultValue = vbDefaultButton2
    Else
        DefaultValue = 0
    End If
    dialogValue = MsgBox("Do you want to display Listobjects ? (" & ActiveSheet.ListObjects.Count & ")", _
                        vbYesNoCancel + DefaultValue + vbQuestion, "showListobjects macro")
    If dialogValue = vbCancel Then Exit Sub
    If dialogValue = vbYes Then showListobjects
    
    If Not ThisWorkbook.getCleanListobjectBoolean Then
        DefaultValue = vbDefaultButton2
    Else
        DefaultValue = 0
    End If
    dialogValue = MsgBox("Do you want to clean Listobjects ?", vbYesNoCancel + DefaultValue + vbQuestion, "cleanListobjects macro")
    If dialogValue = vbCancel Then Exit Sub
    If dialogValue = vbYes Then cleanListObjects
    
End Sub







