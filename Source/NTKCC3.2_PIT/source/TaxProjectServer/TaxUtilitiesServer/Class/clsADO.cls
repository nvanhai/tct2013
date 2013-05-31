VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum Provider
    [MSDAORA.1] = 0
End Enum

Private Const strOriginConnectionString = "Provider=[+PROVIDER+];Password=[+PASSWORD+];User ID=[+USER ID+];Data Source=[+DATA SOURCE+];Persist Security Info=True"


Private mProvider       As Provider
Private strDataSource   As String
Private strDBUserName   As String
Private strDBPassword   As String
Private IsConnected     As Boolean
Private IsConnected_qhs     As Boolean

Private conExcute       As New ADODB.Connection
Private conExcute_qhs       As New ADODB.Connection
Private xmlPackages     As New MSXML.DOMDocument
Private xmlPackage      As MSXML.IXMLDOMNode

Private serverDB   As String
Private userDB   As String
Private passDB   As String
Private vatPath   As String
Private QHSPath   As String

Public Property Get Connection() As ADODB.Connection
    Set Connection = conExcute
End Property

Public Property Get Provider() As Provider
    Provider = mProvider
End Property

Public Property Let Provider(ByVal pProvider As Provider)
    mProvider = pProvider
End Property

Public Property Get DataSource() As String
    DataSource = strDataSource
End Property

Public Property Let DataSource(ByVal pDataSource As String)
    strDataSource = pDataSource
End Property

Public Property Get DBUserName() As String
    DBUserName = strDBUserName
End Property

Public Property Let DBUserName(ByVal pDBUserName As String)
    strDBUserName = pDBUserName
End Property

Public Property Get DBPassword() As String
    DBPassword = strDBPassword
End Property

Public Property Let DBPassword(ByVal pDBPassword As String)
    strDBPassword = pDBPassword
End Property

Public Sub CreateConnectionString(ByVal strDir As String)
'    Dim strConnectionString As String
'
'    If pProvider <> -1 Then mProvider = pProvider
'    If pDataSource <> vbNullString Then strDataSource = pDataSource
'    If pDBUserName <> vbNullString Then strDBUserName = pDBUserName
'    If pDBUserPassword <> vbNullString Then strDBPassword = pDBUserPassword
'
'    strConnectionString = strOriginConnectionString
'    Select Case mProvider
'        Case [MSDAORA.1]
'            strConnectionString = Replace(strConnectionString, "[+PROVIDER+]", "MSDAORA.1", , , vbTextCompare)
'    End Select
'    strConnectionString = Replace(strConnectionString, "[+PASSWORD+]", strDBPassword, , , vbTextCompare)
'    strConnectionString = Replace(strConnectionString, "[+USER ID+]", strDBUserName, , , vbTextCompare)
'    strConnectionString = Replace(strConnectionString, "[+DATA SOURCE+]", strDataSource, , , vbTextCompare)
    Dim strConnection As String
    Dim strFiletype As String
'On Error GoTo OpenCnnErr
     
    If UCase(Mid(strDir, Len(strDir) - 3, 4)) = ".DBC" Then
       strFiletype = "DBC"
    Else
       strFiletype = "DBF"
    End If
  
    strConnection = "Provider=MSDASQL;Driver={Microsoft Visual FoxPro Driver};" & _
    "SourceType=" & strFiletype & ";" & _
    "SourceDB=" & strDir & ";" & _
    "Exclusive=Yes;OLE DB Services = -4"
    conExcute.ConnectionString = strConnection
    conExcute.CursorLocation = adUseClient
End Sub
'dhdang sua insert QLT
Public Sub CreateConnectionStringSQL(ByVal strDir As String)

    Dim strConnection As String
    GetConfigFile strDir
    'strConnection = "Provider=SQLOLEDB.1;Password=" & strPass & ";Persist Security Info=True;User ID=sa;Initial Catalog=QHSCC;Data Source=NAMHL\SQLEXPRESS"
    strConnection = "Provider=SQLNCLI.1; Data Source=" + serverDB + ";Initial Catalog=QHSCC;User Id=" + userDB + ";Password=" + passDB + ";"
    conExcute.ConnectionString = strConnection
    conExcute.CursorLocation = adUseClient
End Sub
Public Sub CreateConnectionStringCheckSQL(ByVal strDir As String)

    Dim strConnection As String
    GetConfigFile strDir
    'strConnection = "Provider=SQLOLEDB.1;Password=" & strPass & ";Persist Security Info=True;User ID=sa;Initial Catalog=QHSCC;Data Source=NAMHL\SQLEXPRESS"
    strConnection = "Provider=SQLNCLI.1; Data Source=" + serverDB + ";Initial Catalog=QHSCC;User Id=" + userDB + ";Password=" + passDB + ";"
    conExcute_qhs.ConnectionString = strConnection
    conExcute_qhs.CursorLocation = adUseClient
End Sub

Public Sub Connect()
On Error GoTo Err
    If IsConnected = False Then
        conExcute.Open
        IsConnected = True
        Exit Sub
    End If
Err:
    IsConnected = False
End Sub
Public Sub Connect_qhs()
On Error GoTo Err
    If IsConnected_qhs = False Then
        conExcute_qhs.Open
        IsConnected_qhs = True
        Exit Sub
    End If
Err:
    IsConnected_qhs = False
End Sub

Public Sub DisConnect()
    If IsConnected = True Then
        conExcute.Close
        IsConnected = False
    End If
End Sub
Public Sub DisConnect_qhs()
    If IsConnected_qhs = True Then
        conExcute_qhs.Close
        IsConnected_qhs = False
    End If
End Sub
Public Property Get Connected() As Boolean
    Connected = IsConnected
End Property
Public Property Get Connected_qhs() As Boolean
    Connected_qhs = IsConnected_qhs
End Property
Public Function ExecuteDLL(pSQL As String) As Boolean
    On Error GoTo Err
    conExcute.Execute pSQL
    ExecuteDLL = True
    Exit Function
Err:
    ExecuteDLL = False
    SaveErrorLog "Thuc hien cau lenh SQL", "ExecullteDLL", Err.Number, Err.Description
End Function
Public Sub SaveErrorLog(pFormName As String, pFunctionName As String, pErrorNumber As Long, pErrorDesc As String)
    Dim msg As String
    Dim FileNum As Long
    Dim path As String
    path = App.path & "\LogFile.txt"
    msg = Now & " " & pFormName & " " & pFunctionName & vbCrLf
    msg = msg & str(pErrorNumber) & " : " & pErrorDesc
    FileNum = FreeFile
    Open path For Append As FileNum
        Print #FileNum, msg
        Print #FileNum, "------------------------------------------------------------"
    Close #FileNum
End Sub
Public Function Execute(pSQL As String) As Recordset
    Dim recExcute As New ADODB.Recordset

    recExcute.Open pSQL, conExcute, adOpenDynamic, adLockReadOnly
    If recExcute.Fields.Count <> 0 Then
        If Not (recExcute.EOF = True And recExcute.BOF = True) Then Set Execute = recExcute
    End If

    Set recExcute = Nothing
        
End Function

Public Function Execute_Qhs(pSQL As String) As Recordset
    Dim recExcute As New ADODB.Recordset

    recExcute.Open pSQL, conExcute_qhs, adOpenDynamic, adLockReadOnly
    If recExcute.Fields.Count <> 0 Then
        If Not (recExcute.EOF = True And recExcute.BOF = True) Then Set Execute_Qhs = recExcute
    End If

    Set recExcute = Nothing
        
End Function
Public Property Let Package(ByVal pPackageID As String)
    Set xmlPackage = xmlPackages.nodeFromID(pPackageID)
End Property

Public Sub ParameterByIndex(ByVal pParaIndex As Long, ByVal pValue As Variant)
    If xmlPackage Is Nothing Then Exit Sub
    xmlPackage.childNodes(pParaIndex).Attributes.getNamedItem("Value").nodeValue = CStr(pValue)
End Sub

Public Sub ParameterByName(ByVal pParaName As String, ByVal pValue As Variant)
    Dim xmlPara As MSXML.IXMLDOMNode
    
    If xmlPackage Is Nothing Then Exit Sub
    For Each xmlPara In xmlPackage.childNodes
        If UCase(xmlPara.Attributes.getNamedItem("Name").nodeValue) = UCase(pParaName) Then
            xmlPara.Attributes.getNamedItem("Value").nodeValue = CStr(pValue)
            Exit For
        End If
    Next
    Set xmlPara = Nothing
End Sub

Private Function CreatePara(xmlPara As MSXML.IXMLDOMNode) As String
    Select Case xmlPara.Attributes.getNamedItem("DataType").nodeValue
        Case "String"
            CreatePara = "'" & xmlPara.Attributes.getNamedItem("Value").nodeValue & "'"
        Case "Numeric"
            CreatePara = xmlPara.Attributes.getNamedItem("Value").nodeValue
    End Select
End Function

Private Function CreateParaString(xmlPackage As MSXML.IXMLDOMNode) As String
    Dim xmlPara As MSXML.IXMLDOMNode
    Dim strTemp As String
    
    CreateParaString = xmlPackage.Attributes.getNamedItem("PackageName").nodeValue & "."
    CreateParaString = CreateParaString & xmlPackage.Attributes.getNamedItem("ProcedureName").nodeValue
    
    For Each xmlPara In xmlPackage.childNodes
        strTemp = strTemp & CreatePara(xmlPara) & ","
    Next
    strTemp = Left(strTemp, Len(strTemp) - 1)
    strTemp = "(" & strTemp & ")"
    CreateParaString = CreateParaString & strTemp
    CreateParaString = "SELECT " & CreateParaString & " a FROM DUAL"
End Function

Public Function ExcutePackage() As Recordset
    Dim recExcute As New ADODB.Recordset
    
    If xmlPackage Is Nothing Then Exit Function
    recExcute.Open CreateParaString(xmlPackage), conExcute, adOpenDynamic, adLockReadOnly
    If Not (recExcute.EOF = True And recExcute.BOF = True) Then Set ExcutePackage = recExcute

    Set recExcute = Nothing
End Function

Private Sub Class_Initialize()
    xmlPackages.resolveExternals = True
    xmlPackages.validateOnParse = True
    xmlPackages.async = False
End Sub

Private Sub Class_Terminate()
    DisConnect
    Set conExcute = Nothing
    Set xmlPackage = Nothing
    Set xmlPackages = Nothing
End Sub

Public Sub BeginTrans()
    conExcute.BeginTrans
End Sub

Public Sub RollbackTrans()
    conExcute.RollbackTrans
End Sub

Public Sub CommitTrans()
    conExcute.CommitTrans
End Sub
Public Sub GetConfigFile(ByVal strDir As String)
    ' Get a free file number
    On Error GoTo KT
    Dim nFileNum As String
    Dim sNextLine As String
    Dim clsMgs As New clsMessageBox
    nFileNum = FreeFile
    Dim objFSO As New FileSystemObject
        If objFSO.FileExists(strDir & "\QHSCC.config") Then
             ' Open a text file for input. inputbox returns the path to read the file
            Open strDir & "\QHSCC.config" For Input As nFileNum
            Line Input #nFileNum, sNextLine
            serverDB = GiaiMa(sNextLine)
            Line Input #nFileNum, sNextLine
            userDB = GiaiMa(sNextLine)
            Line Input #nFileNum, sNextLine
            passDB = GiaiMa(sNextLine)
            Line Input #nFileNum, sNextLine
            vatPath = sNextLine
        Else
            serverDB = ""
            userDB = ""
            passDB = ""
        End If
  
    ' Close the file
KT:
    Close nFileNum

End Sub
Public Function GiaiMa(ByVal pString As String) As String
On Error Resume Next
    Dim FirstString
    Dim SecondString
    Dim i As Integer
    Dim j As Integer
    Dim LengOfString As Integer
    Dim mCount As Integer
    Dim mKey As Integer
    mKey = Val(Right(pString, Right(pString, 1) + 1)) / 10
    FirstString = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "a", "b", "c", _
        "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "x", "y", "z", _
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "X", "Y", "Z")
    SecondString = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    For i = 0 To 59
        j = i + mKey
        If j > 59 Then j = (j Mod 59) - 1
            SecondString(i) = FirstString(j)
    Next i
    LengOfString = Len(pString)
    For mCount = 1 To LengOfString - Len(str(mKey))
        For i = 0 To 59
            If Mid(pString, mCount, 1) = SecondString(i) Then
                Mid(pString, mCount) = FirstString(i)
                Exit For
            End If
        Next i
    Next mCount
GiaiMa = Left(pString, Len(pString) - Len(str(mKey)))
End Function
