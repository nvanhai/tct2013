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

'Dinh nghia type luu du lieu bang HDR
Public Type TNCN_HDR
    tin As Variant
    ten_dtnt As Variant
    DIA_CHI As Variant
    loai_tkhai As Variant
    ngay_nop As Variant
    'kylb_tu_ngay As Variant
    'kylb_den_ngay As Variant
    'KYKK_TU_NGAY As Variant
    'KYKK_DEN_NGAY As Variant
    ngay_cap_nhat As Variant
    nguoi_cn As Variant
    co_loi_dda As Variant
    so_hieu_tep As Variant
    so_tt_tk As Variant
    da_nhan As Variant
    ghi_chu_loi As Variant
    khoa_so As Variant
    phong_xly As Variant
    kkbs As Variant
    tthtk As Variant
    KYLBO As Variant
    KYKKHAI As Variant
    ma_cqt As Variant
    kieu_ky As Variant
    id As Variant
    thueondinh As Variant
 
    ma_dl_thue As Variant
    so_hd_dl As Variant
    ngay_hd_dl As Variant
    lan_bs As Variant
End Type
'end
'Dinh nghia type luu du lieu bang DTL
Public Type TNCN_DTL
    MATKHAI As Variant
    madtnt As Variant
    KYLBO As Variant
    KYKKHAI As Variant
    tthtk As Variant
    NGNOP As Variant
    cttn As Variant
    giatri As Variant
    lan_quet As Variant
    ky_hieu As Variant
    ma_cqt As Variant
    danhan As Variant
    id As Variant
    hdr_id As Variant
End Type
'end
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
Public Property Get DataChanged() As Boolean
    DataChanged = blnDataChanged
End Property

'PIT
Public Property Get isCheckPIT() As Boolean
    isCheckPIT = isUsePIT
End Property

Public Property Let isCheckPIT(ByVal checkPIT As Boolean)
    isUsePIT = checkPIT
End Property

