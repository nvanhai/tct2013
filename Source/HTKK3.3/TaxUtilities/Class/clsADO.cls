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
'******************************************************
' Company           : CMC Software
' Center Name       : FIS (Financial Insurance Solution)
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Class Modules
' Form, Module
'   or Class name   : clsADO
' Descriptions      : Class to create connection to VATCC and QLT
' Start date        : 10/08/2007 (dd/mm/yyyy)
' Finish date       :
' Coder             : hlnam
' Integrate         :
' Project manager   :
' Last modify       :
' Reason of modify  :
'******************************************************

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

Private conExcute       As New ADODB.Connection
Private xmlPackages     As New MSXML.DOMDocument
Private xmlPackage      As MSXML.IXMLDOMNode

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

Public Sub CreateConnectionString(Optional pProvider As Provider = -1, Optional pDataSource As String = vbNullString, Optional pDBUserName As String = vbNullString, Optional pDBUserPassword As String = vbNullString)
    Dim strConnectionString As String
    
    If pProvider <> -1 Then mProvider = pProvider
    If pDataSource <> vbNullString Then strDataSource = pDataSource
    If pDBUserName <> vbNullString Then strDBUserName = pDBUserName
    If pDBUserPassword <> vbNullString Then strDBPassword = pDBUserPassword
    
    strConnectionString = strOriginConnectionString
    Select Case mProvider
        Case [MSDAORA.1]
            strConnectionString = Replace(strConnectionString, "[+PROVIDER+]", "MSDAORA.1", , , vbTextCompare)
    End Select
    strConnectionString = Replace(strConnectionString, "[+PASSWORD+]", strDBPassword, , , vbTextCompare)
    strConnectionString = Replace(strConnectionString, "[+USER ID+]", strDBUserName, , , vbTextCompare)
    strConnectionString = Replace(strConnectionString, "[+DATA SOURCE+]", strDataSource, , , vbTextCompare)
    conExcute.ConnectionString = strConnectionString
End Sub

Public Sub Connect()
    If IsConnected = False Then
        conExcute.Open
        IsConnected = True
    End If
End Sub

Public Sub DisConnect()
    If IsConnected = True Then
        conExcute.Close
        IsConnected = False
    End If
End Sub

Public Property Get Connected() As Boolean
    Connected = IsConnected
End Property

Public Function Execute(pSQL As String) As Recordset
    Dim recExcute As New ADODB.Recordset

    recExcute.Open pSQL, conExcute, adOpenDynamic, adLockReadOnly
    If recExcute.Fields.Count <> 0 Then
        If Not (recExcute.EOF = True And recExcute.BOF = True) Then Set Execute = recExcute
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
