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
Option Explicit

Public Property Get Year() As String
    Year = strFinYear
End Property

Public Property Let Year(ByVal pFinYear As String)
    strFinYear = pFinYear
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
    bolAdjustData(pIndex) = pData
End Property

Public Property Get NodeListError() As MSXML.IXMLDOMNodeList
    Set NodeListError = xmlNodeListError
End Property

Public Property Let NodeError(pNodeListError As MSXML.IXMLDOMNodeList)
    Set xmlNodeListError = pNodeListError
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

Public Property Get FinanceStartDate() As String
    FinanceStartDate = strFinanceStartDate
End Property

Public Property Let FinanceStartDate(ByVal pFinDate As String)
    strFinanceStartDate = pFinDate
End Property

'PIT
Public Property Get isCheckPIT() As Boolean
    isCheckPIT = isUsePIT
End Property
'version MV
Public Property Get verMaVach() As String
    verMaVach = strMaVach
End Property
'--Check MST de ghi du lieu sang QLT hoac QCT (ghi sang QCT cap nhat gia tri Y)--
Public Property Get isCheckQCT() As Boolean
    isCheckQCT = isQCT
End Property
Public Property Let isCheckQCT(ByVal checkQCT As Boolean)
    isQCT = checkQCT
End Property

Public Property Let isCheckPIT(ByVal checkPIT As Boolean)
    isUsePIT = checkPIT
End Property
Public Property Let verMaVach(ByVal strVer As String)
    strMaVach = strVer
End Property
