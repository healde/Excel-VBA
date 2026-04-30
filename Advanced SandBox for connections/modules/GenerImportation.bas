Attribute VB_Name = "GenerImportation"
Sub ImportFromConnection()
' In this sub : thisWB.getConnectionName >> thisWB.getQueryName
ThisWorkbook.ResetSettings
Dim errorFlag As Boolean

AutoManagement

' Get infos from connection
CommandDesc = "Table"

    infoHig = ActiveWorkbook.Connections(ThisWorkbook.getConnectionName).OLEDBConnection.Connection
    If InStr(1, infoHig, "Mashup.OleDb", vbTextCompare) > 0 Then CommandDesc = "Query"

CommandValue = ActiveWorkbook.Connections(ThisWorkbook.getConnectionName).OLEDBConnection.CommandText

    posCharLeft = InStr(CommandValue, "[") + 1
    lenghtString = InStr(CommandValue, "]") - posCharLeft
    If lenghtString > 0 Then CommandValue = Mid(CommandValue, posCharLeft, lenghtString)

    CommandValue = Replace(CommandValue, """", "")

' Start importation
Dim lio As TableObject

On Error GoTo errorCommand
Set lio = ActiveSheet.ListObjects.Add(SourceType:=xlSrcModel, _
        Source:=ActiveWorkbook.Connections(ThisWorkbook.getConnectionName), _
        Destination:=ActiveCell).TableObject
        
On Error GoTo 0
        
    With lio
    
        ' Optional
        .RowNumbers = False
        .PreserveFormatting = True
        .AdjustColumnWidth = True
        .RefreshStyle = xlInsertDeleteCells


        indent = ""
        On Error GoTo errorName
        'This parameter cannot start by number
        .ListObject.DisplayName = "TableDisplayed_" & ThisWorkbook.getConnectionName & indent
        indent = ""
        .WorkbookConnection = "ConnectionSource_" & CommandValue & "_" & _
                                CommandDesc & "Target" & indent
                                
    On Error GoTo errorRefresh
    .Refresh
    End With
    

ActiveCell.Offset(2, 0).Select
Exit Sub
errorRefresh:

    If errorFlag = True Then Exit Sub
    MsgBox "Error Number " & Err.Number & " while targeting " & ThisWorkbook.getCommandValue
    errorFlag = True
    
Resume
errorCommand:

    If errorFlag = True Then Exit Sub
    errorFlag = errorCommand(errorFlag)
    
Resume
errorName:

    indent = errorName(indent)
    If indent = "" Then Exit Sub

' default listObject name : Table_ExternalData_1
' default connection name returned : ModelConnection_ExternalData_1 + ThisWorkbookDataModel
' Opened from dataModel : "ModelConnection_" & WorkbookQuery.Name
Resume
End Sub


Sub ImportFromExternal()
' In this sub : thisWB.getQueryName
ThisWorkbook.ResetSettings
Dim errorFlag As Boolean

AutoManagement

Dim lio As QueryTable

On Error GoTo errorCommand
Set lio = ActiveSheet.ListObjects.Add(SourceType:=xlSrcExternal, _
        Source:=ThisWorkbook.getProvidingSource, _
        Destination:=ActiveCell).QueryTable
        
On Error GoTo 0
        
    With lio
        
        ' Required
        .CommandType = ThisWorkbook.getCommandType
        .CommandText = ThisWorkbook.getCommandText

        ' Optional
        .RowNumbers = False
        .PreserveFormatting = True
        .AdjustColumnWidth = True
'        .PreserveColumnInfo = True
'        .FillAdjacentFormulas = False
        .RefreshStyle = xlInsertDeleteCells
'        .RefreshOnFileOpen = False
'        .RefreshPeriod = 0
'        .SavePassword = False
'        .SaveData = True
'        .BackgroundQuery = True


        indent = ""
        On Error GoTo errorName
        'auto-indentation for this parameter but cannot start by number
        .ListObject.DisplayName = "TableGot_" & ThisWorkbook.getCommandValue & "_" & indent
        indent = ""
        .WorkbookConnection = "ExternalSource_" & ThisWorkbook.getCommandValue & "_" & _
                                ThisWorkbook.getCommandDesc & "Target" & indent
        
    
    On Error GoTo errorRefresh
    .Refresh 'BackgroundQuery:=False
    End With


ActiveCell.Offset(2, 0).Select
Exit Sub
errorRefresh:

    If errorFlag = True Then Exit Sub
    MsgBox "Error Number " & Err.Number & " while targeting " & ThisWorkbook.getCommandValue
    errorFlag = True
    
Resume
errorCommand:

    If errorFlag = True Then Exit Sub
    errorFlag = errorCommand(errorFlag)
    
Resume
errorName:

    indent = errorName(indent)
    If indent = "" Then Exit Sub

' default listObject name : Table_ExternalData_1
' default connection name returned : Connection
Resume
End Sub
