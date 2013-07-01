VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVariables"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************
' Company           : CMC Software
' Center Name       : FIS (Financial Insurance Solution)
' Project           : Du an ho tro ke khai thue version 1.3.0
' Package           : Class Modules
' Form, Module
'   or Class name   : clsVariables
' Descriptions      : Class to define variables
' Start date        : 10/08/2007 (dd/mm/yyyy)
' Finish date       :
' Coder             : hlnam
' Integrate         :
' Project manager   :
' Last modify       :
' Reason of modify  :
'******************************************************


Option Explicit

Public Property Get DateKHBS() As String
    DateKHBS = strDateKHBS
End Property

Public Property Let DateKHBS(ByVal pDate As String)
    strDateKHBS = pDate
End Property


Public Property Get Day() As String
    Day = strDay
End Property

Public Property Let Day(ByVal pDay As String)
    strDay = pDay
End Property

Public Property Get Year() As String
    Year = strFinYear
End Property

Public Property Let Year(ByVal pFinYear As String)
    strFinYear = pFinYear
End Property

Public Property Get LastDay() As String
    LastDay = strLastDayOfQuarter
End Property

Public Property Let LastDay(ByVal strDay As String)
    strLastDayOfQuarter = strDay
End Property

Public Property Get FirstDay() As String
    FirstDay = strFirstDayOfQuarter
End Property

Public Property Let FirstDay(ByVal strDay As String)
    strFirstDayOfQuarter = strDay
End Property

Public Property Get Month() As String
    Month = strMonth
End Property

Public Property Let Month(ByVal pMonth As String)
    strMonth = pMonth
End Property

Public Property Get ThreeMonths() As String
    ThreeMonths = str3Months
End Property

Public Property Let ThreeMonths(ByVal pThreeMonth As String)
    str3Months = pThreeMonth
End Property

Public Property Get DataFolder() As String
    DataFolder = strDataFolder
End Property

Public Property Let DataFolder(ByVal strFolder As String)
    strDataFolder = strFolder
End Property

Public Property Get NodeMenu() As MSXML.IXMLDOMElement
    Set NodeMenu = xmlNodeMenu
End Property

Public Property Let NodeMenu(ByVal pNodeMenu As MSXML.IXMLDOMElement)
    Set xmlNodeMenu = pNodeMenu
End Property

Public Sub xmlDataReDim(pIndex As Long, Optional pPreserve As Boolean = False)
    If pPreserve = False Then
        ReDim xmlData(pIndex)
    Else
        ReDim Preserve xmlData(pIndex)
    End If
End Sub

Public Property Get xmlDataCount() As Long
    xmlDataCount = UBound(xmlData)
End Property

Public Property Get Data(pIndex As Long) As MSXML.DOMDocument
    Set Data = xmlData(pIndex)
End Property

Public Property Let Data(pIndex As Long, ByVal pData As MSXML.DOMDocument)
    Set xmlData(pIndex) = pData
End Property

Public Property Get DataKHBS() As MSXML.DOMDocument
    Set DataKHBS = xmlDataKHBS
End Property

Public Property Let DataKHBS(ByVal pDataKHBS As MSXML.DOMDocument)
    Set xmlDataKHBS = pDataKHBS
End Property


Public Property Get NodeValidity() As MSXML.IXMLDOMNode
    Set NodeValidity = xmlNodeValidity
End Property

Public Property Let NodeValidity(pNodeValidity As MSXML.IXMLDOMNode)
    Set xmlNodeValidity = pNodeValidity
End Property

Public Property Get NodeMessage() As MSXML.IXMLDOMNodeList
    Set NodeMessage = xmlNodeListMessage
End Property

Public Property Let NodeMessage(pNodeMessage As MSXML.IXMLDOMNodeList)
    Set xmlNodeListMessage = pNodeMessage
End Property

Public Property Get NodeCaption() As MSXML.IXMLDOMNode
    Set NodeCaption = xmlNodeCaption
End Property

Public Property Let NodeCaption(pNodeCaption As MSXML.IXMLDOMNode)
    Set xmlNodeCaption = pNodeCaption
End Property

Public Sub AdjustDataReDim(pIndex As Long, Optional pPreserve As Boolean = False)
    If pPreserve = False Then
        ReDim bolAdjustData(pIndex)
    Else
        ReDim Preserve bolAdjustData(pIndex)
    End If
End Sub

Public Property Get AdjustDataCount() As Long
    AdjustDataCount = UBound(bolAdjustData)
End Property

Public Property Get AdjustData(pIndex As Long) As Boolean
    AdjustData = bolAdjustData(pIndex)
End Property

Public Property Let AdjustData(pIndex As Long, ByVal pData As Boolean)
    Dim intCtrl As Integer
    bolAdjustData(pIndex) = pData
    If pData Then
        blnDataChanged = True
    End If
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = blnDataChanged
End Property

Public Property Let DataChanged(ByVal pData As Boolean)
    blnDataChanged = pData
End Property

' phuc vu BC 26
Public Property Get CheckNewDataBC26() As Boolean
    CheckNewDataBC26 = isNewDataBC26
End Property

Public Property Let CheckNewDataBC26(ByVal pData As Boolean)
    isNewDataBC26 = pData
End Property

' Phuc vu check GH
Public Property Get CheckToKhaiGH() As Boolean
    CheckToKhaiGH = isCheckGH
End Property

Public Property Let CheckToKhaiGH(ByVal pData As Boolean)
    isCheckGH = pData
End Property

