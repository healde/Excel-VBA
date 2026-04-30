Attribute VB_Name = "ModifySettings"
Sub ChangeSettings()
' This sub can be called for setting parameters changing from a process to another

    ThisWorkbook.ResetSettings
   
    ThisWorkbook.setTargetValue = True
    changeSettingExtSource
    If ThisWorkbook.getTargetValue = False Then Exit Sub
    changeSettingProvider
    If ThisWorkbook.getTargetValue = False Then Exit Sub
    changeSettingCommandType
    If ThisWorkbook.getTargetValue = False Then Exit Sub
    changeSettingExtSheet
    If ThisWorkbook.getTargetValue = False Then Exit Sub
    changeSettingExtTable
    If ThisWorkbook.getTargetValue = False Then Exit Sub
    changeSettingQueryName
    If ThisWorkbook.getTargetValue = False Then Exit Sub
    changeSettingConnectionName
   
End Sub

Sub changeSettingExtSource()

    targetValue = Application.InputBox("Enter full path and name for distant workbook", "Parameter : source", ThisWorkbook.getExtSource)
    If VarType(targetValue) = 11 Then
        If targetValue = False Then ThisWorkbook.setTargetValue = False
    Else
        ThisWorkbook.setExtSource = targetValue
    End If
End Sub

Sub changeSettingProvider()

    If Not ThisWorkbook.getIdProvider Then
        DefaultValue = vbDefaultButton2
    Else
        DefaultValue = 0
    End If
    dialogValue = MsgBox("Provider ""yes"" : Use of queries with Mashup" & Chr(10) & _
                         "Provider ""no""  : Direct access to table", vbYesNoCancel + DefaultValue, _
                         "OLEDB Provider ...")

    If dialogValue = vbCancel Then
        ThisWorkbook.setTargetValue = False
    Else
        ThisWorkbook.setProvidingSource = (dialogValue = vbYes)
    End If
End Sub

Sub changeSettingCommandType()

    targetValue = Application.InputBox("{ xlCmdSql : 2 | xlCmdTable : 3 | xlCmdTableCollection : 6 }", "Parameter : way for collecting datas ", ThisWorkbook.getCommandType)
    If VarType(targetValue) = 11 Then
        If targetValue = False Then ThisWorkbook.setTargetValue = False
    Else
        ThisWorkbook.setCommandType = CInt(targetValue)
    End If
End Sub

Sub changeSettingExtSheet()

    targetValue = Application.InputBox("Enter the name of an accessible sheet", "Parameter : sheet", ThisWorkbook.getTargetSheet)
    If VarType(targetValue) = 11 Then
        If targetValue = False Then ThisWorkbook.setTargetValue = False
    Else
        ThisWorkbook.setTargetSheet = targetValue
    End If
End Sub

Sub changeSettingExtTable()

    targetValue = Application.InputBox("Enter the name of a valid distant table", "Parameter : distant table", ThisWorkbook.getTargetTable)
    If VarType(targetValue) = 11 Then
        If targetValue = False Then ThisWorkbook.setTargetValue = False
    Else
        ThisWorkbook.setTargetTable = targetValue
    End If
End Sub

Sub changeSettingQueryName()

    targetValue = Application.InputBox("Enter the first query name", "Parameter : default query name", ThisWorkbook.getQueryName)
    If VarType(targetValue) = 11 Then
        If targetValue = False Then ThisWorkbook.setTargetValue = False
    Else
        ThisWorkbook.setQueryName = targetValue
    End If
End Sub

Sub changeSettingConnectionName()

    targetValue = Application.InputBox("Enter the first connection name", "Parameter : default connection name", ThisWorkbook.getConnectionName)
    If VarType(targetValue) = 11 Then
        If targetValue = False Then ThisWorkbook.setTargetValue = False
    Else
        ThisWorkbook.setConnectionName = targetValue
    End If
End Sub






Sub changeTable()
ThisWorkbook.ResetSettings

    ' get the default worksheet name, and set to first sheet name value in case of false or empty value
    targetValue = ThisWorkbook.getTargetSheet
    If targetValue = "" Or targetValue = False Then ThisWorkbook.setTargetSheet = ThisWorkbook.getSheets(0)
    
targetFlag% = vbRetry
While targetFlag = vbRetry
    
    ' anytime the auxiliary value of default worksheet will be empty or false, set it to the first sheet name
    If targetValue = "" Or targetValue = False Then targetValue = ThisWorkbook.getTargetSheet
        
        ' input default worksheet and check if it does exist in the recorded worksheets list
        targetValue = Application.InputBox("Enter the name of an accessible sheet", "Parameter : sheet", targetValue)
        If VarType(targetValue) = 11 Then If targetValue = False Then Exit Sub
        
            targetSheet = getPosArray(ThisWorkbook.getSheets, targetValue)
            If targetSheet = -1 Then
                
                targetFlag = MsgBox("This distant sheet seems not recorded", vbRetryCancel, "Parameter : sheet")
            Else
                ThisWorkbook.setTargetSheet = targetValue
                targetFlag = vbOK
            End If
            
Wend
If targetFlag = vbCancel Then Exit Sub
    
    
targetFlag = vbRetry
While targetFlag = vbRetry
' after choosing the default sheet, keep repeating column treatment operations
    
    While targetFlag = vbRetry
    
    ' input index for column, and check if it either numerical or alpha-numerical
    targetIndex = Application.InputBox("Enter an alphabetic or numerical index to get column name" & Chr(10) & "Or press cancel", "Parameter : column index", 0)
    If VarType(targetIndex) = 11 Then If targetIndex = False Then Exit Sub
    
        If targetIndex Like "[A-z]*" And Not targetIndex Like "[A-z]*[!A-z]*" Then
            
                targetIndex = ActiveSheet.Columns(targetIndex).Column - 1
                targetFlag = vbOK
                
        Else
        If targetIndex Like "[0-9]*" And Not targetIndex Like "[0-9]*[!0-9]*" Then
        
                targetIndex = CInt(targetIndex)
                targetFlag = vbOK
        
        Else
                targetFlag = MsgBox("Column format is incorrect" & Chr(10) & "Do you want to retry ?", vbRetryCancel, "Parameter : column index")
        End If
        End If
        
        ' before leaving the loop if the index get the right format, check to it is in the range
        If targetFlag = vbOK Then
            
            If targetIndex >= 0 And targetIndex <= ThisWorkbook.getLimitArray(2) Then
            
                ThisWorkbook.setDefaultColumn = targetIndex
            Else
                targetFlag = MsgBox("Out of range, maximum : " & ThisWorkbook.getLimitArray(2) & Chr(10) & "Do you want to retry ?", _
                            vbRetryCancel, "Parameter : column")
            End If
        End If
        
    Wend
    If targetFlag = vbCancel Then Exit Sub
    
        ' get column value targeted in the recorded tables
        thisTable = ThisWorkbook.getTable(targetSheet)
        targetValue = thisTable(ThisWorkbook.getDefaultColumn)
    
    targetFlag = vbRetry
    While targetFlag = vbRetry
    
    ' prompt for modifying or just applying the column name got
    targetValue = Application.InputBox("Enter a name for this column", "Parameter : Column value", targetValue)
    If VarType(targetValue) = 11 Then If targetValue = False Then Exit Sub
    
        If targetValue Like "[ A-z]*" And Not targetValue Like "[ A-z]*[! A-z]*" Then
    
            ThisWorkbook.setColumnValue = targetValue
            targetFlag = vbOK
        Else
            targetFlag = MsgBox("Name column format is incorrect" & Chr(10) & "Do you want to retry ?", vbRetryCancel, "Parameter : column value")
        End If
    
    Wend
    targetFlag = vbRetry
Wend

End Sub
