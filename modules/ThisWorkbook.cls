VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private needToReset As Boolean

Private externalWorkbook As String
Private externalWorksheets As Variant
Private externalTables As Variant

Private defaultColumn As Integer
Private defaultSheet As String
Private defaultQueryName As String
Private defaultTable As String

Private CommandType As Integer
Private CommandText As String
Private CommandValue As String
Private CommandDesc As String
Private idProvider As Boolean
Private defaultProvider As String

Private defaultConnectionName As String

Private executeShowListobject As Boolean
Private executeShowConnection As Boolean
Private executeCleanListobject As Boolean
Private executeCleanConnection As Boolean

Private myValue As String


Private Sub Workbook_Open()

    ResetSettings

End Sub

Sub ResetSettings()

' In the two next sub, you can define all global variable which are necessary to pilot the script

' In others sub, the global variables used are indicate in header comment by "thisWB.getTheVariable"
' Every time an unhandled runtime error occurs, the global variables get empty so this sub should be called before each process using some of these
' https://stackoverflow.com/questions/7041138/what-is-the-lifetime-of-a-global-variable-in-excel-vba
    
    ' init and target
    If ThisWorkbook.getExtSource = "" Then ThisWorkbook.setExtSource = "C:\Users\33745\Desktop\distant3.xlsx"
    If IsEmpty(ThisWorkbook.getSheets) Then ThisWorkbook.setSheets = Array("sheetOne", "sheetTwo", "sheetThree")
    If UBound(ThisWorkbook.getTable(0)) < 0 Then resetTables

    ' target
    If ThisWorkbook.getTargetSheet = "" Then ThisWorkbook.setTargetSheet = "sheetOne"
    If ThisWorkbook.getTargetTable = "" Then ThisWorkbook.setTargetTable = ThisWorkbook.getTargetSheet & "Range"
    If ThisWorkbook.getQueryName = "" Then ThisWorkbook.setQueryName = "myQuery1"
    If ThisWorkbook.getConnectionName = "" Then ThisWorkbook.setConnectionName = "myConnection1"
    
    
    'ThisWorkbook.noNeedToReset = False ' Debug
    If isNeedToReset Then
    
        ' method of getting datas
        ThisWorkbook.setProvidingSource = True
        ThisWorkbook.setCommandType = 3
        
        ' config management
        ThisWorkbook.setShowConnectionBoolean = False
        ThisWorkbook.setCleanConnectionBoolean = False
        ThisWorkbook.setShowListobjectBoolean = False
        ThisWorkbook.setCleanListobjectBoolean = False
    
    ' end
    ThisWorkbook.noNeedToReset = True
    Range("A1").Select
    
    End If
End Sub

Sub resetTables()
    ReDim externalTables(0 To 2, 0 To 3)
            
    underArrA = Array("index", "label", "info", "refer")
    For i = 0 To UBound(underArrA)
        externalTables(0, i) = underArrA(i)
    Next

    underArrB = Array("index", "label", "info", "local")
    For i = 0 To UBound(underArrB)
        externalTables(1, i) = underArrB(i)
    Next

    underArrC = Array("index", "label", "info", "scope")
    For i = 0 To UBound(underArrC)
        externalTables(2, i) = underArrC(i)
    Next
End Sub

' shared between ExportAsDistant(), and ChangeTable()

Function getLimitArray(Optional ByVal ctc As Integer) As Integer

    If Not ctc > 0 Then ctc = 1
    
    If Not IsEmpty(externalTables) Then If UBound(externalTables) > -1 Then getLimitArray = UBound(externalTables, ctc)

End Function




Public Property Get getTargetValue() As String

    'invoked when getTargetValue is on the right-hand side expression of an assignment,
    'e.g. var = ThisWorkbook.getTargetValue
     
     getTargetValue = myValue
     
End Property

Public Property Let setTargetValue(ByVal value As String)

    'invoked when setTargetValue is on the left-hand side of an assignment,
    'e.g. ThisWorkbook.setTargetValue = value; the 'value' parameter is the result of the right-hand side expression
    
    myValue = value
    
End Property






Public Property Get getExtSource() As String
    getExtSource = externalWorkbook
End Property

Public Property Let setExtSource(ByVal value As String)
    externalWorkbook = value
End Property


Public Property Get getTargetSheet() As String
     getTargetSheet = defaultSheet
End Property

Public Property Let setTargetSheet(ByVal value As String)
    defaultSheet = value
End Property


Public Property Get getSheets() As Variant
     getSheets = externalWorksheets
End Property

Public Property Let setSheets(ByRef thisArray As Variant)
    externalWorksheets = thisArray
End Property



' Get all columns from a table into an array
Public Property Get getTable(ByVal ctab As Integer) As Variant
    
    ' To declare as dimension is needed to Redim in the loop
    Dim thisArray() As Variant
    ' To affect an empty array is needed to get an empty array in case of failure
    thisArray = Array()
    
    If Not IsEmpty(externalTables) Then
    If ctab >= 0 And ctab <= UBound(externalTables, 1) Then
    
        For i = 0 To UBound(externalTables, 2)
    
            ReDim Preserve thisArray(i)
            thisArray(i) = externalTables(ctab, i)
        Next
    
    End If
    End If
    
    getTable = thisArray
    ' use "UBound(getTable(x)) > -1" to test an empty array in result
    
End Property


Public Property Get getDefaultColumn()
    getDefaultColumn = defaultColumn
End Property

Public Property Let setDefaultColumn(ByVal ccol As Integer)
    defaultColumn = ccol
End Property


Public Property Let setColumnValue(ByVal value As String)
    
If Not IsEmpty(externalWorksheets) Then
If defaultSheet <> "" And UBound(externalTables) > -1 Then

    posArray = getPosArray(externalWorksheets, defaultSheet)
         
        If posArray > -1 And defaultColumn > -1 Then
           If ccol > UBound(externalTables, 2) Then ReDim Preserve externalTables(0 To UBound(externalTables, 1), 0 To ccol)
           
           externalTables(posArray, defaultColumn) = value
        End If

End If
End If

End Property




Public Property Get getQueryName() As String
     getQueryName = defaultQueryName
End Property

Public Property Let setQueryName(ByVal value As String)
    defaultQueryName = value
    If idProvider = True Then CommandValue = value
End Property

Public Property Get getTargetTable() As String
     getTargetTable = defaultTable
End Property

Public Property Let setTargetTable(ByVal value As String)
    defaultTable = value
    If idProvider = False Then CommandValue = value
End Property



Public Property Get getCommandText() As String
    getCommandText = CommandText
End Property

Public Property Let setCommandText(ByVal value As String)
    If CommandType = 2 Then
        CommandText = "SELECT * FROM [" & value & "]"
    Else
        CommandText = value
    End If
End Property

Public Property Get getCommandType() As String
    getCommandType = CommandType
End Property

Public Property Let setCommandType(ByVal value As Integer)
    Select Case value
    Case 2
        CommandType = 2 ' xlCmdSql
    Case 6
        CommandType = 6 ' xlCmdTableCollection
    Case Else
        CommandType = 3 ' xlCmdTable
    End Select
    
    setCommandText = CommandValue
End Property


Public Property Get getIdProvider() As Boolean
    getIdProvider = idProvider
End Property

Public Property Get getCommandDesc() As String
     getCommandDesc = CommandDesc
End Property

Public Property Get getCommandValue() As String
     getCommandValue = CommandValue
End Property

Public Property Get getProvidingSource() As String
    getProvidingSource = defaultProvider
End Property

Public Property Let setProvidingSource(ByVal value As Boolean)
    idProvider = value
    
    ' Target : Query > dataSheet
    If value = True Then
        CommandValue = defaultQueryName
        CommandDesc = "Query"
        
        defaultProvider = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" _
                            & defaultQueryName ' & _
                            "commandDesc=" & CommandDesc & ";commandValue=" & CommandValue
    
    ' Target : dataName, dataRange
    Else
        CommandValue = defaultTable
        CommandDesc = "Table"
        
        defaultProvider = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" _
                            & externalWorkbook & _
                            ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";" ' & _
                            "commandDesc=" & CommandDesc & ";commandValue=" & CommandValue
    End If
    
        setCommandText = CommandValue
End Property



Public Property Get getConnectionName() As String
     getConnectionName = defaultConnectionName
End Property

Public Property Let setConnectionName(ByVal value As String)
    defaultConnectionName = value
End Property





Public Property Get isNeedToReset() As Boolean
    isNeedToReset = Not needToReset
End Property

Public Property Let noNeedToReset(ByVal value As Boolean)
    needToReset = value
End Property


Public Property Get getShowListobjectBoolean() As Boolean
    getShowListobjectBoolean = executeShowListobject
End Property

Public Property Let setShowListobjectBoolean(ByVal value As Boolean)
    executeShowListobject = value
End Property


Public Property Get getCleanListobjectBoolean() As Boolean
    getCleanListobjectBoolean = executeCleanListobject
End Property

Public Property Let setCleanListobjectBoolean(ByVal value As Boolean)
    executeCleanListobject = value
End Property


Public Property Get getShowConnectionBoolean() As Boolean
    getShowConnectionBoolean = executeShowConnection
End Property

Public Property Let setShowConnectionBoolean(ByVal value As Boolean)
    executeShowConnection = value
End Property


Public Property Get getCleanConnectionBoolean() As Boolean
    getCleanConnectionBoolean = executeCleanConnection
End Property

Public Property Let setCleanConnectionBoolean(ByVal value As Boolean)
    executeCleanConnection = value
End Property
