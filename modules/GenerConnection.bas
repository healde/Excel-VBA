Attribute VB_Name = "GenerConnection"
Sub SafeAceImportADO()

ThisWorkbook.ResetSettings

    Dim cn As Object ' ADODB.Connection
    Dim rs As Object ' ADODB.Recordset
    Dim sql As String
    Dim target As Range

    ' Your SQL
    
    ThisWorkbook.setProvidingSource = False
    ThisWorkbook.setCommandType = 2
    sql = ThisWorkbook.getCommandText

    ' Target range
    Set target = ActiveSheet.Range("A1")

    ' Create ADO connection
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & ThisWorkbook.getExtSource & ";" & _
            "Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"

    ' Execute query
    Set rs = cn.Execute(sql)

    ' Dump data to sheet WITHOUT creating any Excel connection
    target.CopyFromRecordset rs



    ' Cleanup
'    rs.Close
'    cn.Close
'    Set rs = Nothing
'    Set cn = Nothing

End Sub

Sub AddConnectionOnly()
' In this sub : thisWB.getQueryName >> thisWB.getConnectionName
ThisWorkbook.ResetSettings
Dim errorFlag As Boolean


    ' Modèle: ActiveWorkbook.Connections.Add _
            Name:=, _
            Description:=, _
            ConnectionString:=, _
            CommandText:=, _
            lCmdtype:=
            
    Dim wbc As WorkbookConnection
    
    On Error GoTo errorCommand
    Set wbc = ActiveWorkbook.Connections.Add( _
            Name:=ThisWorkbook.getConnectionName, _
            Description:="Connection to the « " & ThisWorkbook.getCommandValue & " » " & ThisWorkbook.getCommandDesc, _
            ConnectionString:=ThisWorkbook.getProvidingSource, _
            CommandText:=ThisWorkbook.getCommandText, _
            lCmdtype:=ThisWorkbook.getCommandType)
            
    On Error GoTo 0
            
    ' It is then possible to add into the Data Model
    If MsgBox("Do you want add the connection as model ?", vbYesNo, "Load into Data Model") = vbYes Then
    
        Dim wbcc As WorkbookConnection
        Set wbcc = ActiveWorkbook.Model.AddConnection(wbc)
        
            wbcName = wbc.Name
            wbc.Delete
        wbcc.Name = wbcName

    End If
    
'' AutoManagement

Exit Sub
errorCommand:

    If errorFlag = True Then Exit Sub
    errorFlag = errorCommand(errorFlag)
    
Resume
End Sub

Sub AddConnectionWithDataModel()
' In this sub : thisWB.getQueryName >> thisWB.getConnectionName
ThisWorkbook.ResetSettings
Dim errorFlag As Boolean


    ' Modèle: ActiveWorkbook.Connections.Add2 _
            Name:=, _
            Description:=, _
            ConnectionString:=, _
            CommandText:=, _
            lCmdtype:=, _
            CreateModelConnection:=, _
            ImportRelationShips:=
            
    Dim wbc As WorkbookConnection
    
    On Error GoTo errorCommand ' lCmdtype:=xlCmdTableCollection seems only possible in the method .add2
    Set wbc = ActiveWorkbook.Connections.Add2( _
            Name:=ThisWorkbook.getConnectionName, _
            Description:="Connection to the « " & ThisWorkbook.getCommandValue & " » " & ThisWorkbook.getCommandDesc, _
            ConnectionString:=ThisWorkbook.getProvidingSource, _
            CommandText:=ThisWorkbook.getCommandText, _
            lCmdtype:=ThisWorkbook.getCommandType, _
            CreateModelConnection:=True, _
            ImportRelationShips:=False)
            
    On Error GoTo 0
            
'' AutoManagement

Exit Sub
errorCommand:

    If errorFlag = True Then Exit Sub
    errorFlag = errorCommand(errorFlag)
    
Resume
End Sub
