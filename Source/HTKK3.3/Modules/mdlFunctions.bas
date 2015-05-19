Attribute VB_Name = "mdlFunctions"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As Any) As Long
Public Declare Function SetLocaleInfo _
   Lib "kernel32.dll" _
     Alias "SetLocaleInfoA" _
       (ByVal Locale As Long, _
        ByVal LCType As Long, _
        ByVal lpLCData As String) As Boolean
        
Public Declare Function GetSystemDefaultLCID _
   Lib "kernel32.dll" () As Long
   
Public Type activeForm
    id As String
    showed As Boolean
End Type

Public Type Quy
    q As Integer
    Y As Integer
    dNgayDauQuy As Date
    dNgayCuoiQuy As Date
End Type

Global zoomindex As Integer

'Public APP_VERSION As String
'Public Const APP_VERSION = "3.0.0"
'Demo
Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_STHOUSAND = &HF
'Ket xuat XML
Public Const maDVu = "HTKK"
Public Const tenDVu = "HTKK"
Public Const pbanDVu = "9.9.9"
Public Const ttinNhaCCapDVu = ""
Public Const pbanTKhaiXML = "2.0.8"
'End XML

'Trien khai GD1
Public Const TK_GD1 = True
'End TKGD1

Public Const APP_VERSION = "9.9.9"
Public Const KIEU_KY_THANG = "M"
Public Const KIEU_KY_QUY = "Q"
Public Const KIEU_KY_NAM = "Y"
Public Const KIEU_KY_NGAY_NAM = "D_Y"
Public Const KIEU_KY_NGAY_THANG = "D_M"
Public Const KIEU_KY_NGAY_PS = "D"
Public Const KIEU_KY_THANG_NAM = "M_Y"
Public Const KIEU_KY_TU_NGAY_DEN_NGAY = "D_D"


Public Const DDMMYYY = "DD/MM/YYYY"
Public Const DDMM = "DD/MM"
Public Const MMYYYY = "MM/YYYY"
Public Const YYYY = "YYYY"

Public Const SS_SORT_ORDER_ASCENDING = 1

Public Const SS_BORDER_TYPE_NONE = 0
Public Const SS_BORDER_TYPE_LEFT = 1
Public Const SS_BORDER_TYPE_RIGHT = 2
Public Const SS_BORDER_TYPE_TOP = 4
Public Const SS_BORDER_TYPE_BOTTOM = 8
Public Const SS_BORDER_TYPE_OUTLINE = 16

Public Const SS_BORDER_STYLE_DEFAULT = 0
Public Const SS_BORDER_STYLE_SOLID = 1
Public Const SS_BORDER_STYLE_FINE_DOT = 13

Public Const SS_BDM_CURRENT_ROW = 4


Public Const strIdKHBS_TT156 = "~02~04~71~72~11~12~73~03~70~80~81~82~06~77~05~86~87~88~90~94~96~97~98~99~93~92~89~85~55~56~84~"


Const mYear_____ = "T_2"
Const mMonth____ = "T_3"
Const mThreeMonths = "T_4"
Const mTuNgay = "T_6"
Const mDenNgay = "T_7"

Public strKieuKy As String
Public strNgayTaiChinh As String
Public iNgayTaiChinh As Integer
Public iThangTaiChinh As Integer

Public xmlNodeListMenu As MSXML.IXMLDOMNodeList             ' xml node list for menu
Public xmlHeaderData As New MSXML.DOMDocument               ' xml document for header data
Public strPrinterName As String
Public mCurrentSheet As Integer                       ' save value of current sheet
Public strTaxIdString As String
Public arrActiveForm() As activeForm
Public hasActiveForm As Boolean
Public hasDefaultForm As Boolean
Public strDataBarcode() As String                           ' String input to barcode
Public arrErrCells As Scripting.Dictionary
Public intPrintingSession As Integer
Public intDataSession As Integer
Public strHiddenFormName As String                          ' Save name of hidden form
Public strInterfaceUnloadEventName As String                          ' Name of unload event
Public strKHBS As String                          ' Save name KHBS
Public strSolanBS As String                          ' Save name KHBS
Public strSolanKK As String

Public strSoLanXuatBan As String

Public strfileFont As String
Public themDuLieu As Boolean
Public themXoaDuLieu As Boolean

Public flgloadToKhai As Boolean
Public listInBoSung2A() As String
Public listInBoSung5A() As String
Public listInBoSung5B() As String

Public listInBoSung6B() As String

Public flgPrintBoSung As Boolean

Public isNewdata As Boolean ' Phuc vu BC26

Public isNewdataBS As Boolean ' Phuc vu BS

Public strDataFileBS As String ' Lay ten file de phuc vu to khai bo sung

Public strLoaiTKThang_PS As String

Public strLoaiTKQT As String

Public strQuy As String

Public hanNopTk As String

Public ngayLapTkBs As String

Public strLoaiNNKD As String

Public strDauTho As String
Public strCondensate As String
Public strKhiThienNhien As String


Public pbanTKhaiXML_TK As String
Public pbanTKhaiXML_TK_KHBS As String
Public maTKhaiXML As String
Public tenTKhaiXML As String
Public moTaBMauXML As String
 
' bien xu ly luu datafile cho to khai TAIN_DK
' DT dau tho
' CD condensate
' KTN khi thien nhien
Public strLoaiTkDk As String

 
Public strBarcodeInPDF As String    'Chua chuoi ma vach duoc in ra file PDF cuoi cung (Them vao) dung cho iHTKK


''' GetAttribute description
''' Get an attribute value of xmlNode
''' Parameter1 xmlNodeCell      : xmlNode the node need get attribute value
''' Parameter2 pAttributeName   : attribute name
''' Output                      : attribute value
Public Function GetAttribute(xmlNodeCell As MSXML.IXMLDOMNode, pAttributeName As String) As String
    On Error Resume Next
    GetAttribute = xmlNodeCell.Attributes.getNamedItem(pAttributeName).nodeValue
End Function

''' SetAttribute description
''' Set an attribute value to xmlNode
''' Parameter1 xmlNodeCell      : xmlNode the node need set attribute value
''' Parameter2 pAttributeName   : attribute name
''' Parameter3 pAttributeName   : attribute value
Public Sub SetAttribute(xmlNodeCell As MSXML.IXMLDOMNode, pAttributeName As String, pValue As String)
    On Error Resume Next
    xmlNodeCell.Attributes.getNamedItem(pAttributeName).nodeValue = pValue
End Sub

''' SetActiveCell description
''' Set Col & Row properties for grid
''' Parameter1 pGrid        : the fpSpread which set Col & Row properties
''' Parameter2 pCellString  : CellID value
Private Sub SetActiveCell(pGrid As fpSpread, pCellString As String)
    On Error GoTo ErrorHandle
    
    Dim lAnchor As Integer
    
    lAnchor = InStr(1, pCellString, "_")
    pGrid.Col = pGrid.ColLetterToNumber(Left(pCellString, lAnchor - 1))
    pGrid.Row = Val(Right(pCellString, Len(pCellString) - lAnchor))
    
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "SetActiveCell", Err.Number, Err.Description
End Sub

''' LoadHeaderData description
''' Set Header data for last sheet in Excel book
''' Parameter1 pGrid    : the fpSpread which set Header data
''' Parameter2 pSheet   : header data sheet index
Private Sub LoadHeaderData(pGrid As fpSpread, pSheet As Integer)
    On Error GoTo ErrorHandle
    
    Dim strDataFileName  As String
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
        
    strDataFileName = TAX_Utilities_v2.DataFolder & "Header_01.xml"
    xmlHeaderData.Load strDataFileName
    With pGrid
        .sheet = pSheet
        
        Set xmlNodeListCell = xmlHeaderData.getElementsByTagName("Cell")

        For Each xmlNodeCell In xmlNodeListCell
            ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID2"), lCol, lRow
            .SetText lCol, lRow, GetAttribute(xmlNodeCell, "Value")
        Next

        SetActiveCell pGrid, mYear_____
        .Text = TAX_Utilities_v2.Year
        SetActiveCell pGrid, mMonth____
        .Text = TAX_Utilities_v2.month
        SetActiveCell pGrid, mThreeMonths
        .Text = TAX_Utilities_v2.ThreeMonths
        SetActiveCell pGrid, mTuNgay
        .Text = TAX_Utilities_v2.FirstDay
        SetActiveCell pGrid, mDenNgay
        .Text = TAX_Utilities_v2.LastDay
        SetActiveCell pGrid, "T_1"
        .Text = TAX_Utilities_v2.Day
    End With
    
    Set xmlNodeCell = Nothing
    Set xmlNodeListCell = Nothing
    
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "LoadHeaderData", Err.Number, Err.Description
End Sub

Public Function GetValidityNode() As MSXML.IXMLDOMNode
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListValidity As MSXML.IXMLDOMNodeList
    Dim xmlNodeValidity As MSXML.IXMLDOMNode
    
    Dim ValidityDate As Date, StartDate As Date, MaxDate As Date
        
    If strLoaiTKThang_PS = "TK_LANPS" Then
        ValidityDate = DateSerial(TAX_Utilities_v2.Year, TAX_Utilities_v2.month, TAX_Utilities_v2.Day)
    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then
        ValidityDate = GetNgayCuoiThang(CInt(TAX_Utilities_v2.Year), CInt(TAX_Utilities_v2.month))
    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then
        Select Case TAX_Utilities_v2.ThreeMonths
            Case "1", "2", "3", "4"
                ValidityDate = GetNgayCuoiQuy(CInt(TAX_Utilities_v2.ThreeMonths), _
                            CInt(TAX_Utilities_v2.Year), iNgayTaiChinh, iThangTaiChinh)
'            Case "2"
'                ValidityDate = format("30/06/" & TAX_Utilities_v2.Year, "dd/mm/yyyy")
'            Case "3"
'                ValidityDate = format("30/09/" & TAX_Utilities_v2.Year, "dd/mm/yyyy")
'            Case "4"
'                ValidityDate = format("31/12/" & TAX_Utilities_v2.Year, "dd/mm/yyyy")
        End Select
    '*******************************************
    ' ThanhDX modified
    ' Date: 04/04/06
    ' ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" Then
    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "1" Then
    '*******************************************
        ValidityDate = NgayCuoiNamTaiChinh(CInt(TAX_Utilities_v2.Year), iThangTaiChinh, iNgayTaiChinh)
    Else
        ValidityDate = Date
    End If
    
    Set xmlNodeListValidity = TAX_Utilities_v2.NodeMenu.selectNodes("Validity")
    For Each xmlNodeValidity In xmlNodeListValidity
        StartDate = format(GetAttribute(xmlNodeValidity, "StartDate"), "dd/mm/yyyy")
        If ValidityDate >= StartDate Then
            If StartDate > MaxDate Then
                MaxDate = StartDate
                Set GetValidityNode = xmlNodeValidity
            End If
        End If
    Next
    Exit Function
ErrorHandle:
    SaveErrorLog "mdlFunctions", "GetValidityNode", Err.Number, Err.Description
End Function

''' LoadTemplate description
''' Load a Excel template to grid, the name and the path of MS Excel get from TAX_Utilities_v2.NodeMenu (attribute "InterfaceTemplate")
''' Hide last sheet of Excel book, the last sheet containt result of business rule and the header informations
''' Parameter1 pGrid    : the fpSpread which set the template to
''' Modify by ThanhDX
''' Modify date: 08/11/2005
Public Sub LoadTemplate(pGrid As fpSpread, Optional IsInterface As Boolean = True)
    On Error GoTo ErrorHandle
    
    Dim lFileName As String
    Dim lSheetName() As String
    Dim lSheetCount As Integer
    Dim lWorkBookHandle As Integer
    Dim i As Integer
    Dim xmlNodeSheet As MSXML.IXMLDOMNode
    Dim lSheetExist As Boolean
        
    On Error GoTo ErrorHandle
    
    If TAX_Utilities_v2.NodeMenu Is Nothing Then Exit Sub
    If TAX_Utilities_v2.NodeValidity Is Nothing Then _
        TAX_Utilities_v2.NodeValidity = GetValidityNode
    
    If IsInterface = True Then
        lFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity, "InterfaceTemplate"))
    Else
        lFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity, "ReportTemplate"))
    End If
'*************************************************
'ThanhDX added
'Date: 28/02/06
    With pGrid
        '****************************
        'ThanhDX modified
        'Date: 12/05/06
'        If .IsExcelFile(lFileName) <> 1 Then GoTo ErrorHandle
'        .ImportExcelBook lFileName, vbNullString
        .LoadFromFile lFileName
        .sheet = .SheetCount
        LoadHeaderData pGrid, .sheet
        .SheetVisible = False
        .sheet = .SheetCount - 1
        .SheetVisible = False
    End With

    'Set Cursor type and turn off beep sound
    pGrid.CursorType = CursorTypeLockedCell
    pGrid.CursorStyle = CursorStyleArrow
    pGrid.NoBeep = True
    
    '***************************************
    
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "LoadTemplate", Err.Number, Err.Description
End Sub

''' ParserCellID description
''' Parser CellID string to column and row value
''' Parameter1 pGrid    : the current fpSpread grid
''' Parameter2 pCellID  : the CellID value of the xmlNode need parser
''' Parameter2 pCol     : Column value of cell (Output value)
''' Parameter2 pRow     : Row number of cell (Output value)
Public Sub ParserCellID(pGrid As fpSpread, pCellID As String, pCol As Long, pRow As Long)
    On Error GoTo ErrorHandle
    
    Dim lPos As Long
    
    lPos = InStr(1, pCellID, "_", vbTextCompare)
    
    If lPos > 0 Then
        pCol = pGrid.ColLetterToNumber(Left(pCellID, lPos - 1))
        pRow = Val(Right(pCellID, Len(pCellID) - lPos))
    'ThanhDX added
    Else
        pCol = 0
        pRow = 0
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "ParserCellID", Err.Number, Err.Description
End Sub

''' SetupData description
''' Load data from Data Files, fill data to correct cell
''' Parameter1 pGrid    : the fpSpread grid which want fill data
'Public Sub SetupData(pGrid As fpSpread)
'    On Error GoTo ErrorHandle
'
'    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
'    Dim xmlNodeCell As MSXML.IXMLDOMNode
'    Dim lSheet As Long, lCol As Long, lRow As Long
'    Dim strDataFileName As String
'    Dim strOriginDataFileName As String
'
'    TAX_Utilities_v2.xmlDataReDim (TAX_Utilities_v2.NodeValidity.childNodes.length - 1)
'
'    With pGrid
'        '.EventEnabled(EventAllEvents) = False
'        For lSheet = 0 To TAX_Utilities_v2.xmlDataCount
'            'If GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "Active") <> "0" Then
'                .Sheet = lSheet + 1
'                TAX_Utilities_v2.Data(lSheet) = New MSXML.DOMDocument
'                TAX_Utilities_v2.Data(lSheet).resolveExternals = True
'                TAX_Utilities_v2.Data(lSheet).validateOnParse = True
'                TAX_Utilities_v2.Data(lSheet).async = False
'                strOriginDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "TemplateFolder")) & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
'                If GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "0" Then
'                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
'                Else
'                    If Val(TAX_Utilities_v2.month) <> 0 Then
'                        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'                    Else
'                        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
'                    End If
'                End If
'                TAX_Utilities_v2.Data(lSheet).Load strDataFileName
'                If TAX_Utilities_v2.Data(lSheet).parseError.reason <> vbNullString Then
'                    If InStr(1, TAX_Utilities_v2.Data(lSheet).parseError.reason, "The system cannot locate the object specified.") <> 0 Then
'                        TAX_Utilities_v2.Data(lSheet).Load strOriginDataFileName
'                        If TAX_Utilities_v2.Data(lSheet).parseError.reason <> vbNullString Then
'                            MsgBox TAX_Utilities_v2.Data(lSheet).parseError.reason
'                        End If
'                    Else
'                        MsgBox TAX_Utilities_v2.Data(lSheet).parseError.reason
'                    End If
'                End If
'
'                ' If load original data -> not fill
'                Set xmlNodeListCell = TAX_Utilities_v2.Data(lSheet).getElementsByTagName("Cell")
'
'                For Each xmlNodeCell In xmlNodeListCell
'                    ParserCellID pGrid, vCellID, lCol, lRow
'                    If GetAttribute(xmlNodeCell, "FirstCell") = "1" Then
'                        .MaxRows = .MaxRows + 1
'                        .InsertRows lRow, 1
'                        .CopyRowRange lRow - 1, lRow - 1, lRow
'                    End If
'                    .col = lCol
'                    .Row = lRow
'                    Select Case .CellType
'                        Case CellTypeCheckBox
'                            ' Check box
'                            If UCase(vValue) = UCase("x") Then
'                                .Text = "1"
'                            Else
'                                .Text = "0"
'                            End If
'                        Case CellTypeComboBox
'                            .SetText lCol, lRow, vValue
'                        Case CellTypeDate
'                            .CellType = CellTypeEdit
'                            .SetText lCol, lRow, vValue
'                            .CellType = CellTypeDate
''*******************************
''ThanhDX added
''Date: 09/01/2006
'                        Case CellTypePic
'                            .SetText lCol, lRow, vValue
''*******************************
'                        Case Else
'                            .Value = vValue
'                    End Select
'                Next
'
'                Set xmlNodeCell = Nothing
'                Set xmlNodeListCell = Nothing
'            'End If
'        Next
'        '.EventEnabled(EventAllEvents) = True
'    End With
'
'    Exit Sub
'ErrorHandle:
'    SaveErrorLog "mdlFunctions", "SetupData", Err.Number, Err.Description
'End Sub
Public Sub SetupData(pGrid As fpSpread)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim lSheet As Long, lCol As Long, lRow As Long
    Dim lRows As Long
    Dim blnNewData As Boolean, blnHasSetActiveCell As Boolean
    Dim blnExistData As Boolean
    Dim strDataFileName As String
    Dim strDataFileNameBS As String
    Dim strOriginDataFileName As String
    Dim varTemp As Variant
    
    Dim fso As New FileSystemObject
    Dim totalCell, countCell As Long
    
    ' test to khai 01/TBVMT
    Dim LocaleDecimal As String
    Dim LocaleComma As String
    ' end
    
    
    TAX_Utilities_v2.xmlDataReDim (TAX_Utilities_v2.NodeValidity.childNodes.length - 1)
    
    Set arrErrCells = New Scripting.Dictionary
    blnExistData = True
    With pGrid
        '.EventEnabled(EventAllEvents) = False
        For lSheet = 0 To TAX_Utilities_v2.xmlDataCount
            'If GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "Active") <> "0" Then
                blnNewData = False
                .sheet = lSheet + 1
                
                mCurrentSheet = lSheet + 1
                
                blnHasSetActiveCell = False
                
                TAX_Utilities_v2.Data(lSheet) = New MSXML.DOMDocument
                TAX_Utilities_v2.Data(lSheet).resolveExternals = True
                TAX_Utilities_v2.Data(lSheet).validateOnParse = True
                TAX_Utilities_v2.Data(lSheet).async = False
                strOriginDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "TemplateFolder")) & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                If strKHBS = "" Or strKHBS = "TKCT" Then
                   If GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "0" Then
                        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                    Else
                        If GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "0" Then
                            ' to khai GTGT co to khai thang va quy
                            If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "04" _
                            Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "36" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "25" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "96" _
                            Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "94" Then
                               If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Then
                                    If strQuy = "TK_THANG" Then
                                        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                    ElseIf strQuy = "TK_QUY" Then
                                        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_Q0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                    ElseIf strQuy = "TK_LANPS" Then
                                        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                    End If
                               Else
                                    If strQuy = "TK_THANG" Then
                                        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                    ElseIf strQuy = "TK_QUY" Then
                                        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_Q0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                    End If
                               End If
                            ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "98" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "92" Then
'                                If strLoaiTKThang_PS = "TK_THANG" Then
'                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'                                ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
'                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'                                End If
                                 If strQuy = "TK_THANG" Then
                                     strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                 ElseIf strQuy = "TK_LANPS" Then
                                     strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                 ElseIf strQuy = "TK_LANXB" Then
                                     strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                 End If
                            Else
                                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                            End If
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then
                            If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "74" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "75" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "23" Then
                                ' To khai 08/TNCN co to khai tu thang va to khai quy
                                If strQuy = "TK_TU_THANG" Then
                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                Else
                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                End If
                            ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "73" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "56" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "55" Then
                                ' To khai 02/TNDN
                                If strLoaiTKThang_PS = "TK_LANPS" Then
                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                ElseIf strLoaiTKThang_PS = "TK_NAM" Then
                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Year & ".xml"
                                Else
                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                End If
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "68" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "14" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "13" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "18" Then
                                ' BC 26
                                If strQuy = "TK_THANG" Then
                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_T" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                ElseIf strQuy = "TK_QUY" Then
                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                End If
                            Else
                                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                            End If
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "0" Then
                                'Data file contain Day from and to.
                            If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "80" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "82" Then
                                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                            Else
                                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                            End If
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then
                                'Data file contain Day from and to.
                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                    & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                        Else
                                If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "93" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "89" Then
                                    'Data file not contain Day from and to.
                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                    & strLoaiTkDk & "_" & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "87" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "97" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "77" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "88" Then
                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "76" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "59" _
                                Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "43" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "41" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "17" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "26" Then
                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "95" Then
                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_L" & strSolanKK & "_" & TAX_Utilities_v2.Year & ".xml"
                                Else
                                    'Data file not contain Day from and to.
                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                    & TAX_Utilities_v2.Year & ".xml"
                                End If
                        '*********************************
                        End If
                    End If
                strDataFileNameBS = ""
                Else
                    If GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "0" Then
                        strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                    Else
                        If GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "0" Then
                            ' to khai thang quy GTGT
                            If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "04" _
                            Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "36" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "25" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "96" _
                            Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "94" Then
                                If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Then
                                    If strQuy = "TK_THANG" Then
                                        strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                    ElseIf strQuy = "TK_QUY" Then
                                        strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_Q0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                    ElseIf strQuy = "TK_LANPS" Then
                                        strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                    End If
                                Else
                                    If strQuy = "TK_THANG" Then
                                        strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                    ElseIf strQuy = "TK_QUY" Then
                                        strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_Q0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                    End If
                                End If
                            ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "98" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "92" Then
'                                If strLoaiTKThang_PS = "TK_THANG" Then
'                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'                                ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
'                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'                                End If
                                 If strQuy = "TK_THANG" Then
                                     strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                 ElseIf strQuy = "TK_LANPS" Then
                                     strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                 ElseIf strQuy = "TK_LANXB" Then
                                     strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                 End If
                            Else
                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                            End If
                            ' set ten file de lay du lieu phuc vu to khai BS
                            If lSheet = 0 Then
                                If Trim(strSolanBS) = "1" And Trim(TAX_Utilities_v2.month) <> "" Then
                                    If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "04" _
                                    Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "36" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "25" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "96" _
                                    Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "94" Then
                                        If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Then
                                            If strQuy = "TK_THANG" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            ElseIf strQuy = "TK_QUY" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_Q0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                            ElseIf strQuy = "TK_LANPS" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            End If
                                        Else
                                            If strQuy = "TK_THANG" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            ElseIf strQuy = "TK_QUY" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_Q0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                            End If
                                        End If
                                    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "98" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "92" Then
'                                        If strLoaiTKThang_PS = "TK_THANG" Then
'                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'                                        ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
'                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'                                        End If
                                        If strQuy = "TK_THANG" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        ElseIf strQuy = "TK_LANPS" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        ElseIf strQuy = "TK_LANXB" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "70" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "06" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "90" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "05" Then
                                        If strLoaiTKThang_PS = "TK_THANG" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    Else
                                        strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                    End If
                                ElseIf Trim(strSolanBS) = "1" And Trim(TAX_Utilities_v2.ThreeMonths) <> "" Then
                                    strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                Else
                                     If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "04" _
                                    Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "96" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "94" Then
                                        If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Then
                                            If strQuy = "TK_THANG" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            ElseIf strQuy = "TK_QUY" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_Q0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                            ElseIf strQuy = "TK_LANPS" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            End If
                                        Else
                                            If strQuy = "TK_THANG" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            ElseIf strQuy = "TK_QUY" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_Q0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                            End If
                                        End If
                                    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "98" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "92" Then
'                                        If strLoaiTKThang_PS = "TK_THANG" Then
'                                            strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'                                        ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
'                                            strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'                                        End If
                                        If strQuy = "TK_THANG" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        ElseIf strQuy = "TK_LANPS" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        ElseIf strQuy = "TK_LANXB" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "93" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "89" Then
                                        strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "70" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "06" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "90" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "05" Then
                                        If strLoaiTKThang_PS = "TK_THANG" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                        
                                    Else
                                    ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                        strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                    End If
                                End If
                            End If
                            ' end
                            
                            ' Doi voi to khai thang TNCN
                            If (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_11" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_15" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_1" _
                            Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "05" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "06" _
                            Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "86" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "89" _
                            Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "85" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "90" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "96" _
                            Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "98" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "92" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "94" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "93") And mCurrentSheet = 1 Then
                                ' Kiem tra xem da ton tai lan bo sung nay chua?
                                If Not fso.FileExists(strDataFileName) Then
                                    ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                    If Trim(strSolanBS) = "1" Then
                                        If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "04" _
                                        Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "36" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "25" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "96" _
                                        Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "94" Then
                                            If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Then
                                                If strQuy = "TK_THANG" Then
                                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                                ElseIf strQuy = "TK_QUY" Then
                                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_Q0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                                ElseIf strQuy = "TK_LANPS" Then
                                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                                End If
                                            Else
                                                If strQuy = "TK_THANG" Then
                                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                                ElseIf strQuy = "TK_QUY" Then
                                                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_Q0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                                End If
                                            End If
                                        ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "98" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "92" Then
'                                            If strLoaiTKThang_PS = "TK_THANG" Then
'                                                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'                                            ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
'                                                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'                                            End If
                                             If strQuy = "TK_THANG" Then
                                                 strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                             ElseIf strQuy = "TK_LANPS" Then
                                                 strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                             ElseIf strQuy = "TK_LANXB" Then
                                                 strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                             End If
                                        Else
                                            strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    Else
                                        ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                        If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "04" _
                                        Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "36" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "25" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "96" _
                                        Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "94" Then
                                            If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Then
                                                If strQuy = "TK_THANG" Then
                                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                                ElseIf strQuy = "TK_QUY" Then
                                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_Q0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                                ElseIf strQuy = "TK_LANPS" Then
                                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                                End If
                                            Else
                                                If strQuy = "TK_THANG" Then
                                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                                ElseIf strQuy = "TK_QUY" Then
                                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_Q0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                                End If
                                            End If
                                        ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "98" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "92" Then
'                                            If strLoaiTKThang_PS = "TK_THANG" Then
'                                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'                                            ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
'                                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'                                            End If
                                             If strQuy = "TK_THANG" Then
                                                 strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                             ElseIf strQuy = "TK_LANPS" Then
                                                 strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                             ElseIf strQuy = "TK_LANXB" Then
                                                 strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & strLoaiTkDk & "_L" & strSoLanXuatBan & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                             End If
                                        Else
                                            strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    End If
                                End If
                            End If
                        ' set to khai TTDb
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" Then
                            If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "70" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "72" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "06" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "90" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "81" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "05" Then
                                If strLoaiTKThang_PS = "TK_THANG" Then
                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                End If
                            Else
                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                            End If
                            ' set ten file de lay du lieu phuc vu to khai BS
                            If lSheet = 0 Then
                                If Trim(strSolanBS) = "1" And Trim(TAX_Utilities_v2.month) <> "" Then
                                    If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "70" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "72" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "06" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "90" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "81" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "05" Then
                                        If strLoaiTKThang_PS = "TK_THANG" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    Else
                                        strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                    End If
                                ElseIf Trim(strSolanBS) = "1" And Trim(TAX_Utilities_v2.ThreeMonths) <> "" Then
                                    strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                Else
                                    If Val(strSolanBS) > 1 And Trim(TAX_Utilities_v2.month) <> "" Then
                                        If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "70" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "72" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "06" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "90" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "81" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "05" Then
                                            If strLoaiTKThang_PS = "TK_THANG" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            End If
                                        Else
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    ElseIf Val(strSolanBS) > 1 And Trim(TAX_Utilities_v2.ThreeMonths) <> "" Then
                                        strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                    Else
                                        ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                        strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                    End If
                                End If
                            End If
                            ' end
                            
                            ' Doi voi to khai thang TNCN
                            If (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_11" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_15" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_1" _
                            Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "05" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "06" _
                            Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "81" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "70" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "85" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "90") And mCurrentSheet = 1 Then
                                ' Kiem tra xem da ton tai lan bo sung nay chua?
                                If Not fso.FileExists(strDataFileName) Then
                                    ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                    If Trim(strSolanBS) = "1" Then
                                        If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "70" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "72" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "06" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "90" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "81" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "05" Then
                                            If strLoaiTKThang_PS = "TK_THANG" Then
                                                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
                                                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            End If
                                        Else
                                            strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    Else
                                        If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "70" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "72" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "06" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "90" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "81" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "05" Then
                                            If strLoaiTKThang_PS = "TK_THANG" Then
                                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            ElseIf strLoaiTKThang_PS = "TK_LANPS" Then
                                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            End If
                                        Else
                                        ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                            strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    End If
                                End If
                            End If
                            
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then
                            ' Neu to khai 08_TNCN se xu ly rieng
                            If (GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "74" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "75" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "23") And strQuy = "TK_TU_THANG" Then
                                ' set ten file de lay du lieu phuc vu to khai BS
                                If lSheet = 0 Then
    '                                If Trim(strSolanBS) = "1" And Trim(TAX_Utilities_v2.month) <> "" Then
    '                                    strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
    '                                ElseIf Trim(strSolanBS) = "1" And Trim(TAX_Utilities_v2.ThreeMonths) <> "" Then
    '                                    strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
    '                                Else
    '                                    ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
    '                                    strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
    '                                End If
                                    If Trim(strSolanBS) = "1" Then
                                        strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                    Else
                                        ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                        strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                    End If
                                End If
                                ' end
                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                ' Doi voi to khai quy TNCN
                                If ((TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_11") Or (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_15")) And mCurrentSheet = 1 Then
                                    ' Kiem tra xem da ton tai lan bo sung nay chua?
                                    If Not fso.FileExists(strDataFileName) Then
                                        ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                        If Trim(strSolanBS) = "1" Then
                                            strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                        Else
                                            ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                            strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                        End If
                                    End If
                                End If
                            ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "73" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "56" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "55" Then
                                 If strLoaiTKThang_PS = "TK_LANPS" Then
                                     ' set ten file de lay du lieu phuc vu to khai BS
                                    If lSheet = 0 Then
        '                                If Trim(strSolanBS) = "1" And Trim(TAX_Utilities_v2.month) <> "" Then
        '                                    strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
        '                                ElseIf Trim(strSolanBS) = "1" And Trim(TAX_Utilities_v2.ThreeMonths) <> "" Then
        '                                    strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
        '                                Else
        '                                    ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
        '                                    strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
        '                                End If
                                        If Trim(strSolanBS) = "1" Then
                                            'strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        Else
                                            ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                            'strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    End If
                                    ' end
        
                                
                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                    ' Doi voi to khai quy TNCN
                                    If (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_11" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_15" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "11" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "12" _
                                    Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "73" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "56" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "55") And mCurrentSheet = 1 Then
                                        ' Kiem tra xem da ton tai lan bo sung nay chua?
                                        If Not fso.FileExists(strDataFileName) Then
                                            ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                            If Trim(strSolanBS) = "1" Then
                                                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            Else
                                                ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                                            End If
                                        End If
                                    End If
                                 ElseIf strLoaiTKThang_PS = "TK_NAM" Then
                                     ' set ten file de lay du lieu phuc vu to khai BS
                                    If lSheet = 0 Then
                                        If Trim(strSolanBS) = "1" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.Year & ".xml"
                                        Else
                                            ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    End If
                                    ' end
                                        
                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Year & ".xml"
                                    ' Doi voi to khai quy TNCN
                                    If (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_11" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_15" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "11" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "12" _
                                    Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "73" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "56" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "55") And mCurrentSheet = 1 Then
                                        ' Kiem tra xem da ton tai lan bo sung nay chua?
                                        If Not fso.FileExists(strDataFileName) Then
                                            ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                            If Trim(strSolanBS) = "1" Then
                                                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Year & ".xml"
                                            Else
                                                ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.Year & ".xml"
                                            End If
                                        End If
                                    End If
                                 Else
                                     ' set ten file de lay du lieu phuc vu to khai BS
                                    If lSheet = 0 Then
        '                                If Trim(strSolanBS) = "1" And Trim(TAX_Utilities_v2.month) <> "" Then
        '                                    strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
        '                                ElseIf Trim(strSolanBS) = "1" And Trim(TAX_Utilities_v2.ThreeMonths) <> "" Then
        '                                    strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
        '                                Else
        '                                    ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
        '                                    strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
        '                                End If
                                        If Trim(strSolanBS) = "1" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                        Else
                                            ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    End If
                                    ' end
        
                                
                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                    ' Doi voi to khai quy TNCN
                                    If (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_11" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_15" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "11" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "12" _
                                    Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "73" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "56") And mCurrentSheet = 1 Then
                                        ' Kiem tra xem da ton tai lan bo sung nay chua?
                                        If Not fso.FileExists(strDataFileName) Then
                                            ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                            If Trim(strSolanBS) = "1" Then
                                                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                            Else
                                                ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                ' set ten file de lay du lieu phuc vu to khai BS
                                If lSheet = 0 Then
    '                                If Trim(strSolanBS) = "1" And Trim(TAX_Utilities_v2.month) <> "" Then
    '                                    strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
    '                                ElseIf Trim(strSolanBS) = "1" And Trim(TAX_Utilities_v2.ThreeMonths) <> "" Then
    '                                    strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
    '                                Else
    '                                    ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
    '                                    strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
    '                                End If
                                    If Trim(strSolanBS) = "1" Then
                                        strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                    Else
                                        ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                        strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                    End If
                                End If
                                ' end
    
                            
                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                ' Doi voi to khai quy TNCN
                                If (TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_11" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_15" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "11" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "12" _
                                Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "73" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "56" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "99") And mCurrentSheet = 1 Then
                                    ' Kiem tra xem da ton tai lan bo sung nay chua?
                                    If Not fso.FileExists(strDataFileName) Then
                                        ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                        If Trim(strSolanBS) = "1" Then
                                            strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                        Else
                                            ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                            strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "0" Then
                                'Data file contain Day from and to.
                                If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "80" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "82" Then
                                     strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                     & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                     If lSheet = 0 Then
                                             ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                             If Trim(strSolanBS) = "1" Then
                                                 strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                                 & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                             Else
                                                 ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                                 strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                                 & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                             End If
                                     End If
                                         
                                    If Not fso.FileExists(strDataFileName) And mCurrentSheet = 1 Then
                                        ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                        If Trim(strSolanBS) = "1" Then
                                            strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                            & Replace$(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace$(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                        Else
                                            ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                            strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                            & Replace$(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace$(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                        End If
                                    End If
                                ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "03" Then
                                        'Data file not contain Day from and to.
                                        strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                        & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                        If lSheet = 0 Then
                                            ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                            If Trim(strSolanBS) = "1" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                                & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                            Else
                                                ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                                & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                            End If
                                        End If
                                        
                                        If Not fso.FileExists(strDataFileName) And mCurrentSheet = 1 Then
                                            ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                            If Trim(strSolanBS) = "1" Then
                                                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                                & TAX_Utilities_v2.Year & "_" & Replace$(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace$(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                            Else
                                                ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                                & TAX_Utilities_v2.Year & "_" & Replace$(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace$(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                            End If
                                        End If
                                Else
                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                    & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                End If
                        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then
                                'Data file contain Day from and to.
                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                        
                        Else
                                If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "93" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "89" Then
                                    'Data file not contain Day from and to.
                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                    & strLoaiTkDk & "_" & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                    If lSheet = 0 Then
                                        ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                        If Trim(strSolanBS) = "1" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                            & strLoaiTkDk & "_" & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                        Else
                                            ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                            & strLoaiTkDk & "_" & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                        End If
                                    End If
                                    
                                    If Not fso.FileExists(strDataFileName) And mCurrentSheet = 1 Then
                                        ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                        If Trim(strSolanBS) = "1" Then
                                            strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                            & strLoaiTkDk & "_" & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                        Else
                                            ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                            strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                            & strLoaiTkDk & "_" & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                        End If
                                    End If
                                 ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "87" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "97" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "77" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "88" Then
                                    'Data file not contain Day from and to.
                                        strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                        & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                        If lSheet = 0 Then
                                            ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                            If Trim(strSolanBS) = "1" Then
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                                & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                            Else
                                                ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                                strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                                & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                            End If
                                        End If
                                        
                                        If Not fso.FileExists(strDataFileName) And mCurrentSheet = 1 Then
                                            ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                            If Trim(strSolanBS) = "1" Then
                                                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                                & TAX_Utilities_v2.Year & "_" & Replace$(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace$(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                            Else
                                                ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                                strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                                & TAX_Utilities_v2.Year & "_" & Replace$(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace$(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                                            End If
                                        End If
                                Else
                                
                                    'Data file not contain Day from and to.
                                    strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & strSolanBS & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                    & TAX_Utilities_v2.Year & ".xml"
                                    If lSheet = 0 Then
                                        ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                        If Trim(strSolanBS) = "1" Then
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                            & TAX_Utilities_v2.Year & ".xml"
                                        Else
                                            ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                            strDataFileBS = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                            & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    End If
                                    
                                    If Not fso.FileExists(strDataFileName) And mCurrentSheet = 1 Then
                                        ' Neu chua ton tai lan bo sung nay va lan bo sung la 1 thi se lay to khai chinh thuc de cap nhat du lieu
                                        If Trim(strSolanBS) = "1" Then
                                            strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                            & TAX_Utilities_v2.Year & ".xml"
                                        Else
                                            ' Neu bo sung tu lan thu 2 tro di thi se lay lan gan voi lan bo sung nhat
                                            strDataFileName = TAX_Utilities_v2.DataFolder & "bs" & Val(strSolanBS) - 1 & "_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                                            & TAX_Utilities_v2.Year & ".xml"
                                        End If
                                    End If
                                End If
                        '*********************************
                        End If
                    End If
                strDataFileNameBS = strDataFileName
                End If
                isNewdataBS = True
                
                'kiem tra ton tai TK BS
                If strKHBS = "TKBS" And (GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "01" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "02" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "04" _
                Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "11" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "12" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "06" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "05" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "77" _
                Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "71" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "72" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "03" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "80" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "81" _
                Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "82" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "86" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "87" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "89" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "73" _
                Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "56" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "55" _
                Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "83" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "85" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "70" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "90" _
                Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "96" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "98" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "94" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "99" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "92" _
                Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "97" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "93" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "88") And fso.FileExists(strDataFileNameBS) Then
                     isNewdataBS = False
                End If
                
                
                If blnExistData Then
                    TAX_Utilities_v2.Data(lSheet).Load strDataFileName
                    ' Phuc vu BC 26
                    isNewdata = False
                Else
                    TAX_Utilities_v2.Data(lSheet).Load strOriginDataFileName
                    'New data
                    blnNewData = True
                    ' Phuc vu BC 26
                    isNewdata = True
                End If
               ' TAX_Utilities_v2.Data(lSheet).Load "D:\HTKK\HTKK1.3\HTKK140\HTKK\InterfaceTemplates\xml\01_GTGT.xml"
                If TAX_Utilities_v2.Data(lSheet).parseError.reason <> vbNullString Then
                    If InStr(1, TAX_Utilities_v2.Data(lSheet).parseError.errorCode, "2146697210") <> 0 Then
                        If lSheet = 0 Then
                            'To khai khong ton tai
                            blnExistData = False
                        End If
                        
                        'New data
                        blnNewData = True
                        
                        TAX_Utilities_v2.Data(lSheet).Load strOriginDataFileName
                        If TAX_Utilities_v2.Data(lSheet).parseError.reason <> vbNullString Then
                            MsgBox TAX_Utilities_v2.Data(lSheet).parseError.reason
                        End If
                    Else
                        MsgBox TAX_Utilities_v2.Data(lSheet).parseError.reason
                    End If
                Else
                    If blnExistData And GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "ID") <> "KHBS" Then
                        SetAttribute TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "Active", "1"
                    End If
                End If
                
                
                Dim vCellID As String
                Dim vValue As String
                ' If load original data -> not fill
                Set xmlNodeListCell = TAX_Utilities_v2.Data(lSheet).getElementsByTagName("Cell")
                ' Tinh tong so cell do trong datafile
                totalCell = xmlNodeListCell.length
                ' Dat chi so ban dau cua cell la 1
                countCell = 1
                .EventEnabled(EventChange) = False
                For Each xmlNodeCell In xmlNodeListCell
                    '18/11/2011 dntai
                    ' Trong truong ho la to khai thang/quy TNCN va Tong so cell < tong so cell - 7 (Cac cell tu ngay ky den ... So lan bo sung ko duoc cap nhat,thong tin nhan vien dai ly thue) thi thoat khoi vong for
                    If ((TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_11" Or TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_15") And (countCell = totalCell - 1) And mCurrentSheet = 1) Then Exit For
                    ' Ket thuc Trong truong ho la to khai thang/quy TNCN
                    
                    vCellID = GetAttribute(xmlNodeCell, "CellID")
                    vValue = GetAttribute(xmlNodeCell, "Value")
                    ParserCellID pGrid, vCellID, lCol, lRow
                    If GetAttribute(xmlNodeCell, "FirstCell") = "1" Then
                        lRows = GetDynRowCount(pGrid, xmlNodeCell.parentNode)
                        InsertRow pGrid, lRow, lRows, True
                    End If
                    
                    'Xu ly mst phan thong tin header
                    If GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "0" Then
                        If lCol = 3 And (lRow = 34 Or lRow = 39) Then
                            If Len(vValue) = 13 Then
                                vValue = Left$(vValue, 10) & "-" & Right$(vValue, 3)
                            End If
                        End If
                    End If
                    'Next
                    '.EventEnabled(EventChange) = True
                
                'For Each xmlNodeCell In xmlNodeListCell
                    
                
                  '  ParserCellID pGrid, vCellID, lCol, lRow
                    
                    
                    .Col = lCol
                    .Row = lRow
                    If Not .Lock And Not blnHasSetActiveCell Then
                        .SetActiveCell lCol, lRow
                        blnHasSetActiveCell = True
                    End If
                    Select Case .CellType
                        Case CellTypeCheckBox
                            ' Check box
                            If UCase(vValue) = UCase("x") Then
                                .Text = "1"
                            Else
                                .Text = "0"
                                If vValue <> "" And vValue <> "0" Then
                                    'Set note
                                    arrErrCells.Add .sheet & "_" & vCellID, .BackColor
                                    .CellNote = GetAttribute(GetMessageCellById("0080"), "Msg")
                                    .BackColor = 12713215 ' vbRed
                                End If
                                
                                SetAttribute xmlNodeCell, "Value", ""
                            End If
                        Case CellTypeComboBox ', CellTypeEdit, CellTypePic
                            If blnNewData And .Text <> vValue Then
                                SetAttribute xmlNodeCell, "Value", .Text
                            Else
                                .Text = vValue
                                .Col = lCol
                                .Row = lRow
                                If vValue <> .Text Then
                                    SetAttribute xmlNodeCell, "Value", .Text
                                    '.Text = vValue
                                    'Set note
                                    arrErrCells.Add .sheet & "_" & vCellID, .BackColor
                                    .CellNote = GetAttribute(GetMessageCellById("0079"), "Msg")
                                    .BackColor = 12713215 ' vbRed
                                End If
                            End If
'                        Case CellTypeDate
'                            .CellType = CellTypeEdit
'                            .SetText lCol, lRow, vValue
'                            .CellType = CellTypeDate
'*******************************
'ThanhDX added
'Date: 09/01/2006
                        Case CellTypePic
                            If blnNewData And .Text <> vValue Then
                                SetAttribute xmlNodeCell, "Value", .Text
                            Else
                                .Text = vValue
                                .Col = lCol
                                .Row = lRow
                                If vValue <> .Text Then
                                    SetAttribute xmlNodeCell, "Value", .Text
                                    '.Text = vValue
                                    'Set note
                                    arrErrCells.Add .sheet & "_" & vCellID, .BackColor
                                    .CellNote = GetAttribute(GetMessageCellById("0080"), "Msg")
                                    .BackColor = 12713215 ' vbRed
                                End If
                            End If
'*******************************
                        Case CellTypeNumber
                            If Not .Lock Or (.Lock And .Formula = "") Then
                                If blnNewData And .value <> vValue Then
                                    SetAttribute xmlNodeCell, "Value", .value
                                Else
                                    'Format numeric
                                    If Not IsNumeric(vValue) Then
                                        arrErrCells.Add .sheet & "_" & vCellID, .BackColor
                                        .CellNote = GetAttribute(GetMessageCellById("0080"), "Msg")
                                        .BackColor = 12713215 ' vbRed
                                    End If
                                    
                                    SetAttribute xmlNodeCell, "Value", IIf(Not IsNumeric(vValue), "0", vValue)
                                    
                                    'Neu gia tri nam ngoai pham vi
                                    'If Not .Lock Then
                                        If Val(vValue) > .TypeNumberMax Or Val(vValue) < .TypeNumberMin Then
                                            SetAttribute xmlNodeCell, "Value", "0"
                                            'Set note
                                            arrErrCells.Add .sheet & "_" & vCellID, .BackColor
                                            .CellNote = GetAttribute(GetMessageCellById("0077"), "Msg") & "[" & .TypeNumberMin & ";" & .TypeNumberMax & "]"
                                            .BackColor = 12713215 ' vbRed
                                        End If
                                    'End If
                                    ' Xu ly rieng truong hop to khai 01/TBVMT
                                    If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "90" Then
                                            LocaleDecimal = Mid$(CStr(11 / 10), 2, 1)
                                            If InStr(1, vValue, ",") > 0 Then
                                                vValue = Replace$(vValue, ",", LocaleDecimal)
                                            ElseIf InStr(1, vValue, ".") > 0 Then
                                                vValue = Replace$(vValue, ",", LocaleDecimal)
                                            End If
                                            .value = vValue
                                    Else
                                        .value = vValue
                                    End If
                                End If
                            End If
                        Case CellTypeEdit
                            If blnNewData And .Text <> vValue Then
                                SetAttribute xmlNodeCell, "Value", .Text
                            Else
                                .Text = vValue
                                .Col = lCol
                                .Row = lRow
                                If vValue <> .Text Then
                                    SetAttribute xmlNodeCell, "Value", .Text
                                    '.Text = vValue
                                    'Set note
                                    arrErrCells.Add .sheet & "_" & vCellID, .BackColor
                                    .CellNote = GetAttribute(GetMessageCellById("0078"), "Msg") & .TypeMaxEditLen
                                    .BackColor = 12713215 ' vbRed
                                End If
                            End If
'*******************************
                        Case CellTypePercent
                            If Not .Lock Or (.Lock And .Formula = "") Then
                                If blnNewData And .value <> vValue Then
                                    SetAttribute xmlNodeCell, "Value", .value
                                Else
                                .Text = vValue
                                .Col = lCol
                                .Row = lRow
                                If vValue <> .Text Then
                                    SetAttribute xmlNodeCell, "Value", .Text
                                    '.Value = vValue
                                    'Set note
'                                    arrErrCells.Add .sheet & "_" & vCellID, .BackColor
'                                    .CellNote = GetAttribute(GetMessageCellById("0080"), "Msg")
'                                    .BackColor = 12713215 ' vbRed
                                End If
                                End If
                            End If

                        Case Else
                            If blnNewData And .value <> vValue Then
                                SetAttribute xmlNodeCell, "Value", .value
                            Else
                                .value = vValue
                                .Col = lCol
                                .Row = lRow
                                If vValue <> .value Then
                                    SetAttribute xmlNodeCell, "Value", .value
                                    '.Value = vValue
                                    'Set note
                                    arrErrCells.Add .sheet & "_" & vCellID, .BackColor
                                    .CellNote = GetAttribute(GetMessageCellById("0080"), "Msg")
                                    .BackColor = 12713215 ' vbRed
                                End If
                            End If
                    End Select
                    countCell = countCell + 1
                Next
        
                Set xmlNodeCell = Nothing
                Set xmlNodeListCell = Nothing
            'End If
        Next
        .EventEnabled(EventAllEvents) = True
    End With
    
    Exit Sub
ErrorHandle:
    Debug.Print "Sheet " & lSheet & " - Row: " & lRow
    SaveErrorLog "mdlFunctions", "SetupData", Err.Number, Err.Description
End Sub



''' GetCellSpan description
''' Get cell span of current cell
''' Parameter1 pGrid    : the current fpSpread grid (input value)
''' Parameter2 pCol     : the current column (input/ output value)
''' Parameter3 pRow     : the current row (input/ output value)
''' Parameter4 pNumsRow : number of row with span (output value)
''' Parameter5 pNumsCol : number of column with span (output value)
Public Sub GetCellSpan(pGrid As fpSpread, pCol As Long, pRow As Long, Optional pNumsRow As Variant, Optional pNumsCol As Variant)
    On Error GoTo ErrorHandle
    
    Dim lRowAnchor As Variant, lColAnchor As Variant
    
    pGrid.GetCellSpan pCol, pRow, lColAnchor, lRowAnchor, pNumsCol, pNumsRow
    If lRowAnchor <> -1 And lColAnchor <> -1 Then
        pRow = Val(lRowAnchor)
        pCol = Val(lColAnchor)
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "GetCellSpan", Err.Number, Err.Description
End Sub

''' GetCellID description
''' Get CellID of current cell
''' Parameter1 pGrid    : the current fpSpread grid (input value)
''' Parameter2 pCol     : the current column (input value)
''' Parameter3 pRow     : the current row (input value)
Public Function GetCellID(pGrid As fpSpread, ByVal pCol As Long, ByVal pRow As Long) As String
    GetCellID = pGrid.ColNumberToLetter(pCol) & "_" & CStr(pRow)
End Function

'*******************************************************
'Description: SetupReportData function load data to frmReportData
'             This function refer to SetupData() function
'Author:ThanhDX
'Date:12/10/2005
'Input:
'   fpsGrid: Contain data to print
'   IsInterface: true if form load data is interface form, false if it is report form,
'       default value is true (interface form).
'OutPut:
'Return:
'*******************************************************
'Public Sub SetupReportData(fpsGrid As fpSpread, Optional IsInterface As Boolean = True)
'    On Error GoTo ErrorHandle
'
'    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
'    Dim xmlNodeCell As MSXML.IXMLDOMNode
'    Dim lSheet As Long, lCol As Long, lRow As Long, intRowHeight As Integer
'    Dim strDataFileName As String
'    Dim strOriginDataFileName As String
'
''    If IsInterface Then TAX_Utilities_v2.xmlDataReDim (TAX_Utilities_v2.NodeValidity.childNodes.length - 1)
'
'    With fpsGrid
'    For lSheet = 0 To TAX_Utilities_v2.xmlDataCount
'        If GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "Active") <> "0" Then
'            .Sheet = lSheet + 1
'    '        If IsInterface Then
'    '            Set TAX_Utilities_v2.Data(lSheet) = New MSXML.DOMDocument
'    '            TAX_Utilities_v2.Data(lSheet).resolveExternals = True
'    '            TAX_Utilities_v2.Data(lSheet).validateOnParse = True
'    '            TAX_Utilities_v2.Data(lSheet).async = False
'    '            strOriginDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
'    '            If GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "0" Then
'    '                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
'    '            Else
'    '                strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
'    '            End If
'    '            TAX_Utilities_v2.Data(lSheet).Load strDataFileName
'    '            If TAX_Utilities_v2.Data(lSheet).parseError.reason <> vbNullString Then
'    '                If InStr(1, TAX_Utilities_v2.Data(lSheet).parseError.reason, "The system cannot locate the object specified.") <> 0 Then
'    '                    TAX_Utilities_v2.Data(lSheet).Load strOriginDataFileName
'    '                    If TAX_Utilities_v2.Data(lSheet).parseError.reason <> vbNullString Then
'    '                        MsgBox TAX_Utilities_v2.Data(lSheet).parseError.reason
'    '                    End If
'    '                Else
'    '                    MsgBox TAX_Utilities_v2.Data(lSheet).parseError.reason
'    '                End If
'    '            End If
'    '        End If
'
'            Set xmlNodeListCell = TAX_Utilities_v2.Data(lSheet).getElementsByTagName("Cell")
'
'            For Each xmlNodeCell In xmlNodeListCell
'    '            If IsInterface Then
'    '                ParserCellID fpsGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
'    '            Else
'                    ParserCellID fpsGrid, GetAttribute(xmlNodeCell, "CellID2"), lCol, lRow
'    '            End If
'                If GetAttribute(xmlNodeCell, "FirstCell") = "1" Then
'                    .MaxRows = .MaxRows + 1
'                    .InsertRows lRow, 1
'                    .CopyRowRange lRow - 1, lRow - 1, lRow
'                    ResetRow fpsGrid, lRow
'                End If
'                If Not IsNullNumber(GetAttribute(xmlNodeCell, "Value")) And GetAttribute(xmlNodeCell, "Value") <> "" And lRow <> 0 And lCol <> 0 Then
'                    .col = lCol
'                    .Row = lRow
'                    .Value = GetAttribute(xmlNodeCell, "Value")
'                    If .RowHeight(lRow) < .MaxTextRowHeight(lRow) Then _
'                        .RowHeight(lRow) = .MaxTextRowHeight(lRow)
'                Else
'                    .SetText lCol, lRow, ""
'                End If
'
'            Next
'
'            Set xmlNodeCell = Nothing
'            Set xmlNodeListCell = Nothing
'        End If
'    Next
'
'    End With
'
'    Exit Sub
'ErrorHandle:
'    SaveErrorLog "mdlFunctions", "SetupReportData", Err.Number, Err.Description
'End Sub

Public Sub SetupReportData(fpsGrid As fpSpread, Optional IsInterface As Boolean = True)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim lSheet As Long, lCol As Long, lRow As Long, intRowHeight As Integer
    Dim strDataFileName As String
    Dim strOriginDataFileName As String
    Dim lRow2s As Long
    Dim iCountPageBreak As Byte
    Dim varTemp As Variant
    With fpsGrid
    iCountPageBreak = 0
    For lSheet = 0 To TAX_Utilities_v2.xmlDataCount
        If GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "Active") <> "0" Then
            .sheet = lSheet + 1
            If lSheet = 0 And strKHBS = "frmKHBS_BS" Then
               PrintLabelKHBS GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID"), fpsGrid, 1
            End If
        
            
            Set xmlNodeListCell = TAX_Utilities_v2.Data(lSheet).getElementsByTagName("Cell")
    
            For Each xmlNodeCell In xmlNodeListCell
                ParserCellID fpsGrid, GetAttribute(xmlNodeCell, "CellID2"), lCol, lRow
                If GetAttribute(xmlNodeCell, "FirstCell") = "1" Then
                    GetDynRowCount fpsGrid, xmlNodeCell.parentNode, lRow2s
                    InsertRow fpsGrid, lRow, lRow2s, True
                    ResetRow xmlNodeCell, fpsGrid, lRow, lRow2s
                End If
                .Col = lCol
                .Row = lRow
                
                If GetAttribute(xmlNodeCell, "PageBreak") = "1" Then
                    If Not xmlNodeCell.parentNode.nextSibling Is Nothing Then
                        .RowPageBreak = True
                    Else
                        ' Xu ly rieng cho to quyet toan 09/TNCN (05TNCN->09TNCN)
                        If TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "45" Then
                            iCountPageBreak = iCountPageBreak + 1 ' To quyet toan 09 co 2 trang in tren cung mot sheet
                            If iCountPageBreak < 2 Then ' Trong truong hop 1 trang dau thi ngat thanh tung trang
                                .RowPageBreak = True
                            Else ' Den trang thu 2 thi khong ngat nua
                                .RowPageBreak = False
                            End If
                        ' Ket thuc
                        Else ' Cac truong hop to khai khac xu ly nhu binh thuong
                            .RowPageBreak = False
                        End If
                        
                    End If
                End If
                'dhdang sua loi các mâu an chi mau in de trang khi nhap "0000000"
                'ngay 14-05-2011
                Dim IsNullNumber_ As Variant
                
                If TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ParentID").nodeValue = "101_7" Then
                       IsNullNumber_ = IsNullNumber_ac(GetAttribute(xmlNodeCell, "Value"))
                Else
                       IsNullNumber_ = IsNullNumber(GetAttribute(xmlNodeCell, "Value"))
                End If
                If Not IsNullNumber_ And GetAttribute(xmlNodeCell, "Value") <> "" And lRow <> 0 And lCol <> 0 Then
                'htphuong edit 19/05/2008
                'cut "000" sau dau ","
                    If .CellType = CellTypeNumber Then

                                varTemp = GetAttribute(xmlNodeCell, "Value")
                                If IsNumeric(CStr(varTemp)) Then
                                    If Right(GetAttribute(xmlNodeCell, "Value"), Len(GetAttribute(xmlNodeCell, "Value")) - InStr(1, GetAttribute(xmlNodeCell, "Value"), ".")) = "0000" Then
                                       .TypeNumberDecPlaces = 0
                                    ElseIf Right(GetAttribute(xmlNodeCell, "Value"), Len(GetAttribute(xmlNodeCell, "Value")) - InStr(1, GetAttribute(xmlNodeCell, "Value"), ".")) = "000" Then
                                       .TypeNumberDecPlaces = 0
                                    ElseIf Right(GetAttribute(xmlNodeCell, "Value"), Len(GetAttribute(xmlNodeCell, "Value")) - InStr(1, GetAttribute(xmlNodeCell, "Value"), ".")) = "00" Then
                                       .TypeNumberDecPlaces = 0
                                    ElseIf InStr(1, GetAttribute(xmlNodeCell, "Value"), ".") > 0 Then
                                        If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "17" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "42" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "43" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "26" Then
                                            .TypeNumberDecPlaces = 0
                                        ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "94" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "98" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "99" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "92" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "24" _
                                        Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "96" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "19" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "93" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "89" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "76" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "03" Then
                                            .TypeNumberDecPlaces = Len(GetAttribute(xmlNodeCell, "Value")) - InStr(1, GetAttribute(xmlNodeCell, "Value"), ".")
                                        Else
                                            .TypeNumberDecPlaces = 3 'Len(GetAttribute(xmlNodeCell, "Value")) - InStr(1, GetAttribute(xmlNodeCell, "Value"), ".")
                                        End If
                                    
                                    ElseIf Right(GetAttribute(xmlNodeCell, "Value"), Len(GetAttribute(xmlNodeCell, "Value")) - InStr(1, GetAttribute(xmlNodeCell, "Value"), ",")) = "0000" Then
                                       .TypeNumberDecPlaces = 0
                                    ElseIf Right(GetAttribute(xmlNodeCell, "Value"), Len(GetAttribute(xmlNodeCell, "Value")) - InStr(1, GetAttribute(xmlNodeCell, "Value"), ",")) = "000" Then
                                       .TypeNumberDecPlaces = 0
                                    ElseIf Right(GetAttribute(xmlNodeCell, "Value"), Len(GetAttribute(xmlNodeCell, "Value")) - InStr(1, GetAttribute(xmlNodeCell, "Value"), ",")) = "00" Then
                                       .TypeNumberDecPlaces = 0
                                    ElseIf InStr(1, GetAttribute(xmlNodeCell, "Value"), ",") > 0 Then
                                            .TypeNumberDecPlaces = 3 'Len(GetAttribute(xmlNodeCell, "Value")) - InStr(1, GetAttribute(xmlNodeCell, "Value"), ".")
                                    End If
                                End If
                    End If
'end edit
'                    Debug.Print xmlNodeCell.xml
                    .value = GetAttribute(xmlNodeCell, "Value")
                Else
                    .SetText lCol, lRow, ""
                  'Kiem tra cac ID nao cua AC thi moi sua nhe
                       'dntai 10/05/2011
                'sua loi mat so khong trong cot <tuso - denso> khi in doi voi nhung to an chi
                 '   .SetText lCol, lRow, GetAttribute(xmlNodeCell, "Value")
                End If
                
                If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "12" And .Col = .ColLetterToNumber("CD") And (.Row = "54" Or .Row = "56" Or .Row = "57" Or .Row = "58") Then
                   If .Text <> "" Then
                        If Len(.Text) <= 2 Then
                            .Text = Left(.Text & ".000", 6)
                        ElseIf Len(.Text) > 2 Then
                            .Text = Left$(.Text & "000", 6)
                        End If
                        
                        If Right(.Text, Len(.Text) - InStr(1, .Text, ".")) = "000" Then
                              .CellType = CellTypeEdit
                              .TypeHAlign = TypeHAlignRight
                             ' .Text = Left(.Text, Len(.Text) - 4) & "%"
                              .Text = Left(.Text, Len(.Text) - 4)
                          Else
                              .CellType = CellTypeEdit
                              .TypeHAlign = TypeHAlignRight
                              '.Text = Left(.Text, Len(.Text) - 4) & "," & Right(.Text, 3) & "%"
                              .Text = Left(.Text, Len(.Text) - 4) & "," & Right(.Text, 3)
                          End If
                    End If
                End If
                
                 If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "12" And .Col = .ColLetterToNumber("CD") And .Row = "59" And .Text <> "" Then
                    If GetAttribute(TAX_Utilities_v2.Data(0).nodeFromID("H_47"), "Value") = "x" Or GetAttribute(TAX_Utilities_v2.Data(0).nodeFromID("H_47"), "Value") = "1" Then
                        .CellType = CellTypeEdit
                        .TypeHAlign = TypeHAlignRight
                        '.Text = .Text & "%"
                    Else
                        ' FormatTextPercent fpsGrid, 1, .Col, .Row
                        If Len(.Text) <= 2 Then
                            .Text = Left(.Text & ".000", 6)
                        ElseIf Len(.Text) > 2 Then
                            .Text = Left$(.Text & "000", 6)
                        End If
                        If Right(.Text, Len(.Text) - InStr(1, .Text, ".")) = "000" Then
                            .CellType = CellTypeEdit
                            .TypeHAlign = TypeHAlignRight
                            '.Text = Left(.Text, Len(.Text) - 4) & "%"
                            .Text = Left(.Text, Len(.Text) - 4)
                        Else
                            .CellType = CellTypeEdit
                            .TypeHAlign = TypeHAlignRight
                            '.Text = Left(.Text, Len(.Text) - InStr(1, .Text, ".")) & "," & Right(.Text, 3)
                            .Text = Mid(.Text, 1, InStr(1, .Text, ".") - 1) & "," & Mid(.Text, InStr(1, .Text, ".") + 1, 3)
                        End If
                    End If
                End If
                
                
                
                
'                If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "12" And .Col = .ColLetterToNumber("CD") And (.Row = "27" Or .Row = "25") Then
'                      If Right(.Text, Len(.Text) - InStr(1, .Text, ".")) = "000" Then
'                          .CellType = CellTypePercent
'                          .TypePercentDecPlaces = 0
'                      Else
'                          .CellType = CellTypePercent
'                          .TypePercentDecPlaces = 3
'                          .TypePercentDecimal = TypePercentNegStyle1
'                      End If
'
'                End If
'
'                 If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "12" And .Col = .ColLetterToNumber("CD") And .Row = "28" Then
'                    If GetAttribute(TAX_Utilities_v2.Data(0).nodeFromID("G_16"), "Value") = "x" Then
'                        .CellType = CellTypeEdit
'                        .TypeHAlign = TypeHAlignRight
'                    Else
'                       ' FormatTextPercent fpsGrid, 1, .Col, .Row
'                            If Right(.Text, Len(.Text) - InStr(1, .Text, ".")) = "000" Then
'                                .CellType = CellTypePercent
'                                .TypePercentDecPlaces = 0
'                            Else
'                                .CellType = CellTypePercent
'                                .TypePercentDecPlaces = 3
'                           '     .TypePercentDecimal = ","
'                            End If
'                    End If
'                End If

                
                If lRow <> 0 And lCol <> 0 Then
                    If .RowHeight(lRow) < .MaxTextRowHeight(lRow) Then
                        .RowHeight(lRow) = .MaxTextRowHeight(lRow) - 1
                    End If
                End If
            Next
    
            Set xmlNodeCell = Nothing
            Set xmlNodeListCell = Nothing
        End If
    Next

    End With
    
    Exit Sub
ErrorHandle:
    'MsgBox lSheet & " - " & lCol & " - " & lRow
    SaveErrorLog "mdlFunctions", "SetupReportData", Err.Number, Err.Description
End Sub
'*******************************************************
'Description: CutStringByNumByte function divide a
'   string into some strings by limited bytes
'
'Author:ThanhDX
'Date:17/10/2005
'Input: strData: The input string is divided
'       numByte: number of limited byte value,
'                this value must be an even number
'OutPut:
'Return:Variant array which are the normal strings.
'*******************************************************
Public Function CutStringByNumByte(ByVal strData As String, _
                          ByVal numByte As Integer) As Variant
On Error GoTo ErrHandle
    
    Dim tmpArray() As String
    Dim num As Integer
    Dim i As Integer
        
    num = Int(LenB(strData) / numByte) + 1
    
    ReDim tmpArray(num)
    
    For i = 1 To num
        tmpArray(i) = CStr(MidB(strData, 1, numByte))
        strData = CStr(MidB(strData, numByte + 1))
    Next
    CutStringByNumByte = tmpArray()

Exit Function
ErrHandle:
    SaveErrorLog "mdlFunction", "CutStringByNumByte", Err.Number, Err.Description
End Function


'*******************************************************
'Description: CutStringByNumByte function divide a
'   string into some strings by limited bytes
'
'Author:Namhl
'Date:15/07/2009
'Input: strData: The input string is divided
'       numByte: number of limited byte value,
'                this value must be an even number
'OutPut:
'Return:Variant array which are the normal strings.
'*******************************************************
Public Function CutStringByNumChar(ByVal strData As String, _
                          ByVal numChar As Integer) As Variant
On Error GoTo ErrHandle
    
    Dim tmpArray() As String
    Dim num As Integer
    Dim i As Integer
        
    num = Int(Len(strData) / numChar) + 1
    
    ReDim tmpArray(num)
    
    For i = 1 To num
        tmpArray(i) = CStr(Mid(strData, 1, numChar))
        strData = CStr(Mid(strData, numChar + 1))
    Next
    CutStringByNumChar = tmpArray()

Exit Function
ErrHandle:
    SaveErrorLog "mdlFunction", "CutStringByNumChar", Err.Number, Err.Description
End Function

'*******************************************************
'Description: ValidNumber function check a strings if it is a valid number
'Author:TuanLM
'Date:17/10/2005
'Paramter: s: the string to check
'          max: the max value of number
'          min: the min value of number
'Return:True if it is a valid number, false if it is not a valid number
'*******************************************************

Public Function ValidNumber(ByVal s As String, Optional max As Integer, Optional min As Integer = 0) As Boolean
   Dim i As Long
   Dim sNumber As String
   Dim bReturn As Boolean

   bReturn = True
   If IsNumeric(s) Then
        If CInt(s) > max Or CInt(s) <= min Then
            bReturn = False
        End If
   Else
        bReturn = False
   End If
   
   ValidNumber = bReturn
   
End Function

'*******************************************************
'Description: ValidFormatDate function check a strings if it is a valid number
'Author:TuanLM
'Date:17/10/2005
'Paramter: s: the string to check
'          max: the max value of number
'          min: the min value of number
'Return:True if it is a valid number, false if it is not a valid number
'*******************************************************

Public Function ValidFormatDate(txtDate As TextBox, format As String) As Boolean

    Select Case format
        Case "M"
            If Not ValidNumber(txtDate.Text, 12) Then
                DisplayMessage "0018", msOKOnly, miInformation
                txtDate.SetFocus
                Exit Function
            ElseIf Len(txtDate.Text) = 1 Then
                txtDate.Text = "0" & txtDate.Text
            End If
        Case "Q"
            If Not ValidNumber(txtDate.Text, 4) Then
                DisplayMessage "0018", msOKOnly, miInformation
                txtDate.SetFocus
                Exit Function
            End If
        Case "Y"
            If Not IsNumeric(txtDate.Text) Then
                DisplayMessage "0018", msOKOnly, miInformation
                txtDate.SetFocus
                Exit Function
            ElseIf Len(txtDate.Text) = 3 Then
'                If CInt(txtDate.Text) >= 100 Then
'                    txtDate.Text = "1" & txtDate.Text
'                Else
'                    txtDate.Text = "2" & txtDate.Text
'                End If
                txtDate.Text = "2" & txtDate.Text
            ElseIf Len(txtDate.Text) = 2 Then
'                If CInt(txtDate.Text) >= 80 Then
'                    txtDate.Text = "19" & txtDate.Text
'                Else
'                    txtDate.Text = "20" & txtDate.Text
'                End If
                txtDate.Text = "20" & txtDate.Text
            ElseIf Len(txtDate.Text) = 1 Then
                txtDate.Text = "200" & txtDate.Text
            End If
            
            If Val(txtDate.Text) < 2000 Then
                DisplayMessage "0043", msOKOnly, miInformation
                txtDate.SetFocus
                Exit Function
            End If
        Case Else
        
    End Select
    ValidFormatDate = True
End Function

'*******************************************************
'Description: SaveErrorLog sub write errors to log file
'Author:TuanLM
'Date:17/10/2005
'Paramter: pFormName: form has error
'          pFunctionName: function has error
'          pErrorNumber: error number
'          pErrorDesc: description of error
'Return:
'*******************************************************

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

Public Function getNode(pID As String) As MSXML.IXMLDOMNode
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim lID As String
    
    For Each xmlNode In xmlNodeListMenu
        lID = xmlNode.Attributes.getNamedItem("ID").nodeValue
        If lID = pID Then
            Exit For
        End If
    Next
    
    Set getNode = xmlNode
    Set xmlNode = Nothing
End Function

Public Function getFormIndex(pID As String) As Integer
    Dim i As Long
    
    For i = 1 To UBound(arrActiveForm)
        If arrActiveForm(i).id = pID Then
            getFormIndex = i
            Exit For
        End If
    Next
End Function

Public Function IsActiveForm() As Boolean
    Dim i As Long
    
    For i = 0 To UBound(arrActiveForm)
        If arrActiveForm(i).showed = True Then
            IsActiveForm = True
            Exit For
        End If
    Next
End Function

Private Function CreateCell(xmlNodeCell As MSXML.IXMLDOMNode) As String
    Dim strReturn As String
    
    On Error GoTo ErrorHandle
    
    'CreateCell = GetAttribute(xmlNodeCell, "Value") & "~"
'*******************************************
'ThanhDX added
'Date: 13/01/2006
    'Repalce character control by character code 20
    strReturn = Replace(GetAttribute(xmlNodeCell, "Value"), "~", Chr$(20)) ' & "~"
    'Replace special characters of xml structure
    '' "&" character
    strReturn = Replace(strReturn, "&", "&amp;")
    '' "'" character
    strReturn = Replace(strReturn, "'", "&apos;")
    '' """ character
    strReturn = Replace(strReturn, """", "&quot;")
    '' ">" character
    strReturn = Replace(strReturn, ">", "&gt;")
    '' "<" character
    strReturn = Replace(strReturn, "<", "&lt;")
    
    CreateCell = Replace(strReturn, "#", "1" & Chr$(20) & Chr$(20) & "1")
    
'*******************************************
    Exit Function
ErrorHandle:
    SaveErrorLog "mdlFunction", "CreateCell", Err.Number, Err.Description
End Function

Private Function CreateCells(xmlNodeCells As MSXML.IXMLDOMNode) As String
    On Error GoTo ErrorHandle
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    For Each xmlNodeCell In xmlNodeCells.childNodes
        If Not GetAttribute(xmlNodeCell, "Encode") = "0" Then
                CreateCells = CreateCells & CreateCell(xmlNodeCell) & "~"
        End If
    Next
    
    Exit Function
ErrorHandle:
    SaveErrorLog "mdlFunction", "CreateCells", Err.Number, Err.Description
End Function

Private Function CreateSection(xmlNodeSection As MSXML.IXMLDOMNode) As String
    On Error GoTo ErrorHandle
    Dim xmlNodeCells As MSXML.IXMLDOMNode
            
    For Each xmlNodeCells In xmlNodeSection.childNodes
        CreateSection = CreateSection & CreateCells(xmlNodeCells)
    Next
    If CreateSection <> "" Then
        'The section is encoded
        CreateSection = Left(CreateSection, Len(CreateSection) - 1)
        CreateSection = "<S>" & CreateSection & "</S>"
    End If
    
    Exit Function
ErrorHandle:
    SaveErrorLog "mdlFunction", "CreateSection", Err.Number, Err.Description
End Function

Private Function CreateSections(xmlNodeSections As MSXML.IXMLDOMNode, pSheet As String) As String
    On Error GoTo ErrorHandle
    Dim xmlNodeSection As MSXML.IXMLDOMNode
    
    For Each xmlNodeSection In xmlNodeSections.childNodes
        CreateSections = CreateSections & CreateSection(xmlNodeSection)
    Next
    CreateSections = "<S" & pSheet & ">" & CreateSections & "</S" & pSheet & ">"
    
    Exit Function
ErrorHandle:
    SaveErrorLog "mdlFunction", "CreateSections", Err.Number, Err.Description
End Function

Public Sub CreateExcelBook()
    On Error GoTo ErrorHandle
    Dim i As Long
    Dim xmlNodeSections As MSXML.IXMLDOMNode
    Dim strTemp As String
    
    For i = 0 To TAX_Utilities_v2.xmlDataCount
        If GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(i), "Active") <> "0" Then
            Set xmlNodeSections = TAX_Utilities_v2.Data(i).getElementsByTagName("Sections")(0)
            ReDim Preserve strDataBarcode(i)
            'strTemp = CreateSections(xmlHeaderData.getElementsByTagName("Sections")(0), "H")
            If i = 0 Then
                strDataBarcode(i) = format(iNgayTaiChinh, "0#") & "/" & format(iThangTaiChinh, "0#") _
                & GetAttribute(TAX_Utilities_v2.NodeValidity, "StartDate") _
                & IIf(GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1", TAX_Utilities_v2.FirstDay & TAX_Utilities_v2.LastDay, "") _
                & CreateSections(xmlNodeSections, GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(i), "ID"))
            Else
                strDataBarcode(i) = CreateSections(xmlNodeSections, GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(i), "ID"))
            End If
            'strDataBarcode(i) = "<T" & Right(GetAttribute(TAX_Utilities_v2.NodeValidity, "Class"), 2) & ">" & strTemp & "</T" & Right(GetAttribute(TAX_Utilities_v2.NodeValidity, "Class"), 2) & ">"
        End If
    Next
    
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunction", "CreateExcelBook", Err.Number, Err.Description
End Sub

'****************************************************
'Description: ResetRow procedure reset all of data in row
'Author:ThanhDX.
'Modify by:
'Date:14/11/2005
'Input: lRow: Row is reset
'Output:
'Return:

'****************************************************

Private Sub ResetRow(ByVal xmlCellNode As MSXML.IXMLDOMNode, fpsGrid As fpSpread, ByVal lRow As Long, ByVal lRows As Long)
    Dim lRowCtrl As Long, lColCtrl As Long
    Dim xmlCellsNode As MSXML.IXMLDOMNode
    Dim xmlTempCellNode As MSXML.IXMLDOMNode
    Dim lngCol As Long, lngRow As Long
    
    Set xmlCellsNode = xmlCellNode.parentNode
    For Each xmlTempCellNode In xmlCellsNode.childNodes
        ParserCellID fpsGrid, GetAttribute(xmlTempCellNode, "CellID2"), lngCol, lngRow
        With fpsGrid
            .Col = lngCol
            .Row = lngRow
            If .CellType = CellTypeNumber Then
                .value = 0
            Else
                .Text = vbNullString
            End If
        End With
    Next
    With fpsGrid
        For lRowCtrl = lRow To lRow + lRows
            '.Row = lRowCtrl
            If .RowHeight(lRowCtrl) > 13 Then
                .RowHeight(lRowCtrl) = 14
            End If
        Next lRowCtrl
    End With
End Sub
'
''****************************************************
''Description:SetFormCaption procedure set caption for form
''Author:TuanLM
''Modify by:
''Date:11/11/2005
''Input: frmForm: form need set caption
''       bkGround: picture back ground
''       lblCaption: lable caption
''Output:
''Return:
'
''****************************************************
'Public Sub SetFormCaption(frmForm As Form, bkGround As Image, lblCaption As MSForms.Label)
'
'    'set style for background of caption
'    bkGround.Picture = LoadPicture(GetRootDirectory &  "Pictures\caption.bmp")
'    bkGround.Move 0, 0, frmForm.Width, 320
'    bkGround.Stretch = True
'
'    'set style for lable of caption
'    lblCaption.Top = 30
'    lblCaption.Left = 50
'    lblCaption.Width = bkGround.Width
'    lblCaption.Height = bkGround.Height
'    lblCaption.BackStyle = fmBackStyleTransparent
'    lblCaption.TextAlign = fmTextAlignLeft
'End Sub

'****************************************************
'Description:IsNullNumber function check a number whether is null (0 value)
'Author:ThanhDX
'Modify by:
'Date:18/11/2005
'Input:
'       strValue: the input number has type string
'Output:
'Return:True if the number is null (0 value)

'****************************************************
Public Function IsNullNumber(ByVal strValue As String) As Boolean
    strValue = Replace$(strValue, "0", "")
    strValue = Replace$(strValue, ".", "")
    If Trim(strValue) = "" Then IsNullNumber = True
End Function
Public Function IsNullNumber_ac(ByVal strValue As String) As Boolean
    If strValue <> "0000000" Then
        strValue = Replace$(strValue, "0", "")
        strValue = Replace$(strValue, ".", "")
    End If
    If Trim(strValue) = "" Then IsNullNumber_ac = True
End Function

'****************************************************
'Description:IsNullNumber function check a number whether is null (0 value)
'Author:ThanhDX
'Modify by:
'Date:18/11/2005
'Input:
'       strValue: the input number has type string
'Output:
'Return:True if the number is null (0 value)

'****************************************************
Public Function GetDaysOfMonth(intMonth As Integer, intYear As Integer) As Integer
    On Error GoTo ErrHandle
    Select Case intMonth
        Case 1, 3, 5, 7, 8, 10, 12
            GetDaysOfMonth = 31
        Case 2
            If intYear / 4 = intYear \ 4 And intYear \ 100 <> intYear / 100 Then
                GetDaysOfMonth = 29
            Else
                GetDaysOfMonth = 28
            End If
        Case 4, 6, 9, 11
            GetDaysOfMonth = 30
    End Select
    Exit Function
ErrHandle:
    SaveErrorLog "mdlFunctions", "GetDaysOfMonth", Err.Number, Err.Description
End Function

'****************************************************
'Description:GetLastMonthOfPeriod function check a number whether is null (0 value)
'Author:ThanhDX
'Modify by:
'Date:18/11/2005
'Input:
'       strValue: the input number has type string
'Output:
'Return:True if the number is null (0 value)

'****************************************************
Public Function GetFirstMonthOfPeriod(strPeriod As String) As Integer
    On Error GoTo ErrHandle
    Select Case strPeriod
        Case "1"
            GetFirstMonthOfPeriod = 3
        Case "2"
            GetFirstMonthOfPeriod = 6
        Case "3"
            GetFirstMonthOfPeriod = 9
        Case "4"
            GetFirstMonthOfPeriod = 12
    End Select
    Exit Function
ErrHandle:
    SaveErrorLog "mdlFunctions", "GetLastMonthOfPeriod", Err.Number, Err.Description
End Function

'******************************
'Description: CheckTaxCode function check whether
'             tax code is valid
'Author: ThanhDX
'Date:29/12/2005
'Input:
'******************************

Public Function CheckTaxCode(ms1 As String, ms2 As String, ms3 As String, _
    ms4 As String, ms5 As String, ms6 As String, ms7 As String, _
    ms8 As String, ms9 As String, ms10 As String) As Boolean
    Dim a As Long
    
    On Error GoTo ErrorHandle
    
    a = 31 * Val(ms1) + 29 * Val(ms2) + 23 * Val(ms3) + 19 * Val(ms4) + 17 * Val(ms5) + 13 * Val(ms6) + 7 * Val(ms7) + 5 * Val(ms8) + 3 * Val(ms9)
    If ms10 <> 10 - (a Mod 11) Then
        CheckTaxCode = False
    Else
        CheckTaxCode = True
    End If
    
    Exit Function
ErrorHandle:
    
    SaveErrorLog "mdlFunctions", "CheckTaxCode", Err.Number, Err.Description
End Function

'******************************
'Description: IsValidTaxId function check whether
'             tax id string is valid
'Author: ThanhDX
'Date:29/12/2005
'Input:
'******************************
Public Function IsValidTaxId(strTaxID As String) As Boolean
    If Not IsNumeric(strTaxID) Then _
        Exit Function
    If Len(strTaxID) <> 10 And Len(strTaxID) <> 13 Then _
        Exit Function
    If Not CheckTaxCode(Mid$(strTaxID, 1, 1), Mid$(strTaxID, 2, 1), _
        Mid$(strTaxID, 3, 1), Mid$(strTaxID, 4, 1), Mid$(strTaxID, 5, 1), _
        Mid$(strTaxID, 6, 1), Mid$(strTaxID, 7, 1), Mid$(strTaxID, 8, 1), _
        Mid$(strTaxID, 9, 1), Mid$(strTaxID, 10, 1)) Then
            Exit Function
    End If
    IsValidTaxId = True
    
End Function

'*********************************************
'Description:
'Author: ThanhDX
'Date:
'Input:
'*********************************************

Public Function CheckPeriod(ByVal strMonth As String, ByVal strYear As String) As Boolean
    On Error GoTo ErrHandle
    
    If GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then ' strKieuKy = KIEU_KY_QUY
        If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "68" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "14" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "13" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "18" Then
            If strQuy = "TK_THANG" Then
                If CInt(strYear) > CInt(Year(Date)) Then
                    DisplayMessage "0044", msOKOnly, miInformation
                    Exit Function
                ElseIf CInt(strYear) = CInt(Year(Date)) Then
                    If CInt(strMonth) > CInt(month(Date)) Then
                        DisplayMessage "0044", msOKOnly, miInformation
                        Exit Function
                    End If
                End If
            Else
                If CInt(strYear) > CInt(Year(Date)) Then
                    DisplayMessage "0188", msOKOnly, miInformation
                    Exit Function
                ElseIf CInt(strYear) = CInt(Year(Date)) Then
                    If GetNgayDauQuy(CInt(strMonth), CInt(strYear), iNgayTaiChinh, iThangTaiChinh) > Date Then
                        DisplayMessage "0188", msOKOnly, miInformation
                        Exit Function
                    End If
                End If
            End If
        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "65" Then
            If CInt(strYear) > CInt(Year(Date)) Then
                DisplayMessage "0188", msOKOnly, miInformation
                Exit Function
            ElseIf CInt(strYear) = CInt(Year(Date)) Then
                If strQuy = "TK_QUY" Then
                    If CInt(strYear) > CInt(Year(Date)) Then
                        DisplayMessage "0188", msOKOnly, miInformation
                        Exit Function
                    ElseIf CInt(strYear) = CInt(Year(Date)) Then
                        If GetNgayDauQuy(CInt(strMonth), CInt(strYear), iNgayTaiChinh, iThangTaiChinh) > Date Then
                            DisplayMessage "0188", msOKOnly, miInformation
                            Exit Function
                        End If
                    End If
                Else
                    If CInt(strMonth) = 1 And 1 > CInt(month(Date)) Then
                        DisplayMessage "0188", msOKOnly, miInformation
                        Exit Function
                    ElseIf CInt(strMonth) = 2 And 7 > CInt(month(Date)) Then
                        DisplayMessage "0188", msOKOnly, miInformation
                        Exit Function
                    End If
                End If
            End If
        Else
            'dhdang sua bo chan doi voi to 02 quy
            'If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "37" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "16" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "51" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") <> "38" Then
            ' To khai TNCN se khong tinh theo nam tai chinh
            If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "51" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "16" Then
                If GetNgayDauQuy(CInt(strMonth), CInt(strYear), 1, 1) > Date Then
                    DisplayMessage "0045", msOKOnly, miInformation
                    Exit Function
                End If
            ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "73" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "56" Then
            Else
                If GetNgayDauQuy(CInt(strMonth), CInt(strYear), iNgayTaiChinh, iThangTaiChinh) > Date Then
                    DisplayMessage "0045", msOKOnly, miInformation
                    Exit Function
                End If
            End If
        End If
    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then 'strKieuKy = KIEU_KY_THANG
        If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "04" _
            Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "36" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "25" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "96" _
            Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "94" Then
            If strQuy = "TK_THANG" Then
                If CInt(strYear) > CInt(Year(Date)) Then
                    DisplayMessage "0044", msOKOnly, miInformation
                    Exit Function
                ElseIf CInt(strYear) = CInt(Year(Date)) Then
                    If CInt(strMonth) > CInt(month(Date)) Then
                        DisplayMessage "0044", msOKOnly, miInformation
                        Exit Function
                    End If
                End If
            ElseIf strQuy = "TK_QUY" Then
                If GetNgayDauQuy(CInt(strMonth), CInt(strYear), iNgayTaiChinh, iThangTaiChinh) > Date Then
                    DisplayMessage "0045", msOKOnly, miInformation
                    Exit Function
                End If
            End If
        Else
            If CInt(strYear) > CInt(Year(Date)) Then
                DisplayMessage "0044", msOKOnly, miInformation
                Exit Function
            ElseIf CInt(strYear) = CInt(Year(Date)) Then
                If CInt(strMonth) > CInt(month(Date)) Then
                    DisplayMessage "0044", msOKOnly, miInformation
                    Exit Function
                End If
            End If
        End If
    '************************************
    ' ThanhDX modified
    ' Date: 04/04/06
    ' ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" Then 'strKieuKy = KIEU_KY_NAM
    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "1" Then 'strKieuKy = KIEU_KY_NAM
    '************************************
        If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "24" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "95" Then
        Else
            If CInt(strYear) > CInt(Year(Date)) Then
                DisplayMessage "0063", msOKOnly, miInformation
            Exit Function
        End If
            
        End If
'    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "1/2" And GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "68" Then  'Bao cao an chi
'        If CInt(strYear) > CInt(Year(Date)) Then
'            DisplayMessage "0188", msOKOnly, miInformation
'            Exit Function
'        End If
    End If
    
    CheckPeriod = True
    Exit Function
ErrHandle:
    SaveErrorLog "mdlFunctions", "CheckPeriod", Err.Number, Err.Description
End Function

'****************************************************
'Description: GetDynRowCount function get count of interface rows in
'             one Cells node.
'Author: ThanhDX
'Date:14/12/2006
'Input:
'       pGrid: fpSpread
'       xmlNodeCells: Cells node in dynamic section
'       lReportRows: Count of report rows in Cells node
'       lMinRow: Min row in Cells node
'       lMaxRow: Max row in Cells node
'****************************************************
Public Function GetDynRowCount(pGrid As fpSpread, xmlNodeCells As MSXML.IXMLDOMNode, Optional ByRef lReportRows As Long, Optional ByRef lMinRow As Long, Optional lMaxRow As Long)
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim lRow As Long, lCol As Long
    Dim lRow2 As Long, lCol2 As Long
    Dim lMaxRow2 As Long, lMinRow2 As Long
    
    lMinRow = 100000
    lMaxRow = 0
    lMinRow2 = 100000
    lMaxRow2 = 0
    
    If Not xmlNodeCells Is Nothing Then
        For Each xmlNodeCell In xmlNodeCells.childNodes
            'Get CellID
            ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
            
            'Get CellID2
            ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID2"), lCol2, lRow2
            
            'Get max row
            If lRow > lMaxRow Then
                lMaxRow = lRow
            End If
            
            'Get min row
            If lRow < lMinRow Then
                lMinRow = lRow
            End If
            
            'Get max row
            If lRow2 > lMaxRow2 Then
                lMaxRow2 = lRow2
            End If
            
            'Get min row
            If lRow2 < lMinRow2 Then
                lMinRow2 = lRow2
            End If
        Next
        
        GetDynRowCount = lMaxRow - lMinRow + 1
        lReportRows = lMaxRow2 - lMinRow2 + 1
    End If
End Function

Public Sub InsertRow(fpSpread1 As fpSpread, ByVal pRow As Long, lRows As Long, Optional blnFillingData As Boolean = False)
    On Error GoTo ErrorHandle
    
    Dim i As Long, lBgColor As Long
    Dim lRowCtrl As Long, lColCtrl As Long
    'Dim mCurrentSheet As Long
    
    With fpSpread1
        '.Visible = False
        .ReDraw = False
        '.Sheet = mCurrentSheet
        .MaxRows = .MaxRows + lRows
        .InsertRows pRow, lRows
        For lRowCtrl = 1 To lRows
        
            .CopyRowRange pRow - lRowCtrl, pRow - lRowCtrl, pRow + lRows - lRowCtrl
            .Row = pRow - lRowCtrl
            '.RowHeight(pRow - lRowCtrl) = 14
            If Not blnFillingData Then
                For i = 1 To fpSpread1.MaxCols
                    '***************************
                    'ThanhDX added
                    'Date: 26/12/2005
                    .Col = i
                    lBgColor = .BackColor
                    .Row = pRow + lRows - lRowCtrl
                    If Not .Lock Then
                        'Set BgColor to inserted cell
                        If lBgColor <> &HC0C0FF And lBgColor <> 12713215 Then 'vbRed
                            .BackColor = lBgColor
                        Else
                            .BackColor = vbWhite
                        End If
                    '***************************
                    ' ThanhDX added
                    ' Date: 29/04/06
                    Else
                        If Not TAX_Utilities_v2.Data(mCurrentSheet - 1).nodeFromID( _
                           GetCellID(fpSpread1, i, pRow - lRowCtrl)) Is Nothing Then
                            If .BackColor = &HC0C0FF Or .BackColor = 12713215 Then
                                .BackColor = vbWhite
                            End If
                        End If
                    '***************************
                    End If
                    '***************************
                    ' Reset empty value for new row on grid
                    If .Lock = False Then
                        Select Case .CellType
                            Case CellTypeNumber
                                .SetText i, .Row, 0
                            Case Else
                                .SetText i, .Row, vbNullString
                        End Select
                        .CellNote = vbNullString
                    '***************************
                    ' ThanhDX added
                    ' Date: 08/04/06
                    Else
                        If Not TAX_Utilities_v2.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, i, pRow - lRowCtrl)) Is Nothing Then
                            Select Case .CellType
                                Case CellTypeNumber
                                    .SetText i, .Row, 0
                                Case Else
                                    .SetText i, .Row, vbNullString
                            End Select
                            .CellNote = vbNullString
                        End If
                    '***************************
                    End If
                Next i
            End If
        Next lRowCtrl
        '.Visible = True
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "mdlFunctions", "InsertRow", Err.Number, Err.Description
End Sub

Public Sub IncreaseRowInDOM(fpSpread1 As fpSpread, xmlDomData As MSXML.DOMDocument, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, i As Long
        
    If xmlDomData Is Nothing Then Exit Sub
    Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
    
    For i = xmlNodeListCell.length - 1 To 0 Step -1
        ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
        If lRow >= pRow Then
            ' Increase value of row attribute + 1 (CellID)
            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fpSpread1, lCol, lRow + lRows)
            
            ' Increase value of row attribute + 1 (CellID2)
            ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fpSpread1, lCol, lRow + lRow2s)
        End If
    Next
        
    Set xmlNodeListCell = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "mdlFunctions", "IncreaseRowInDOM", Err.Number, Err.Description
End Sub

Public Function GetKieuKy() As String
    Dim month As String
    Dim threemonth As String
    Dim strDay As String
    Dim strYear As String ' Phuc vu an chi
    Dim i As Integer
    
    i = getFormIndex(TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ID").nodeValue)
    arrActiveForm(i).showed = True
    
    month = TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("Month").nodeValue
    threemonth = TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("ThreeMonth").nodeValue
    strDay = TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("Day").nodeValue
' phuc vu an chi
    strYear = TAX_Utilities_v2.NodeMenu.Attributes.getNamedItem("Year").nodeValue
    If strYear = "1/2" Then
        GetKieuKy = "H_Y"
        ' end
    ElseIf month = "1" And strDay = "1" Then
        GetKieuKy = KIEU_KY_NGAY_THANG
    ElseIf month = "1" Then
        GetKieuKy = KIEU_KY_THANG
    ElseIf threemonth = "1" Then
        GetKieuKy = KIEU_KY_QUY
    ElseIf strDay = "1" Then
        GetKieuKy = KIEU_KY_NGAY_NAM
    Else
        GetKieuKy = KIEU_KY_NAM
    End If
End Function

Public Function GetNgayBatDauNamTaiChinh() As String
    Dim xmlDomHeader As New MSXML.DOMDocument
    
    xmlDomHeader.Load GetAbsolutePath("..\DataFiles\" & strTaxIdString & "\Header_01.xml")
    GetNgayBatDauNamTaiChinh = GetAttribute(xmlDomHeader.getElementsByTagName("Cell")(23), "Value")
    
    Set xmlDomHeader = Nothing
End Function
Public Function KiemTraNgayTaiChinh(strDate As String, Optional blnShowMessage As Boolean = True) As Boolean
    Dim arrDateUnit() As String
    Dim d As Integer
    Dim m As Integer
    Dim i As Integer
    
    'KiemTraNgayTaiChinh = True
    KiemTraNgayTaiChinh = False
    If Len(strDate) > 0 Then
        arrDateUnit = Split(strDate, "/")
        For i = 0 To UBound(arrDateUnit)
            arrDateUnit(i) = Trim(arrDateUnit(i))
        Next
        d = Val(arrDateUnit(0))
        m = Val(arrDateUnit(1))
        If (d = 1 And (m = 1 Or m = 4 Or m = 7 Or m = 10)) Then
            KiemTraNgayTaiChinh = True
        ElseIf blnShowMessage Then
            DisplayMessage "0064", msOKOnly, miCriticalError
        End If
    ElseIf blnShowMessage Then
        DisplayMessage "0061", msOKOnly, miCriticalError
    End If
End Function
Public Function GetNgayTaiChinh(strDate As String) As Integer
    Dim arrDateUnit() As String
    Dim i As Integer
    
    GetNgayTaiChinh = -1
    If Len(strDate) > 0 Then
        arrDateUnit = Split(strDate, "/")
        arrDateUnit(0) = Trim(arrDateUnit(0))
        GetNgayTaiChinh = Val(arrDateUnit(0))
    End If
End Function
Public Function GetThangTaiChinh(strDate As String) As Integer
    Dim arrDateUnit() As String
    Dim i As Integer
    
    GetThangTaiChinh = -1
    If Len(strDate) > 0 Then
        arrDateUnit = Split(strDate, "/")
        arrDateUnit(1) = Trim(arrDateUnit(1))
        GetThangTaiChinh = Val(arrDateUnit(1))
    End If
End Function

Public Function GetQuyHienTai(dNgayTaiChinh As Integer, dThangTaiChinh As Integer) As Quy
    Dim dNgayBatDau As Date
    Dim dNgayDauNam As Date
    Dim iInterval As Integer
    Dim dNgayHienTai As Date
    
    dNgayDauNam = DateSerial(Year(Now), 1, 1)
    dNgayBatDau = DateSerial(Year(Now), dThangTaiChinh, dNgayTaiChinh)
    iInterval = DateDiff("D", dNgayDauNam, dNgayBatDau)
    dNgayHienTai = Now - iInterval
    
    GetQuyHienTai.q = DatePart("Q", dNgayHienTai)
    GetQuyHienTai.Y = Year(dNgayHienTai)
    GetQuyHienTai.dNgayDauQuy = GetNgayDauQuy(GetQuyHienTai.q, GetQuyHienTai.Y, dNgayTaiChinh, dThangTaiChinh)
    GetQuyHienTai.dNgayCuoiQuy = GetNgayCuoiQuy(GetQuyHienTai.q, GetQuyHienTai.Y, dNgayTaiChinh, dThangTaiChinh)
End Function
'dhdang sua ham lay ky hien tai phuc vu an chi
Public Function GetKyHienTai(dNgayTaiChinh As Integer, dThangTaiChinh As Integer) As Quy
    Dim dNgayBatDau As Date
    Dim dNgayDauNam As Date
    Dim iInterval As Integer
    Dim dNgayHienTai As Date

    dNgayDauNam = DateSerial(Year(Now), 1, 1)
    dNgayBatDau = DateSerial(Year(Now), dThangTaiChinh, dNgayTaiChinh)
    iInterval = DateDiff("D", dNgayDauNam, dNgayBatDau)
    dNgayHienTai = Now - iInterval

    GetKyHienTai.q = DatePart("Q", dNgayHienTai)
    If GetKyHienTai.q = 1 Or GetKyHienTai.q = 2 Then
        GetKyHienTai.q = 1
        GetKyHienTai.Y = Year(dNgayHienTai)
        GetKyHienTai.dNgayDauQuy = DateSerial(GetKyHienTai.Y, 1, 1)
        GetKyHienTai.dNgayCuoiQuy = DateSerial(GetKyHienTai.Y, 6, 31)
    Else
        GetKyHienTai.q = 2
        GetKyHienTai.Y = Year(dNgayHienTai)
        GetKyHienTai.dNgayDauQuy = DateSerial(GetKyHienTai.Y, 7, 1)
        GetKyHienTai.dNgayCuoiQuy = DateSerial(GetKyHienTai.Y, 12, 31)
    End If
End Function
' end
Public Function GetNamHienTai(dNgayTaiChinh As Integer, dThangTaiChinh As Integer) As Long
    Dim dNgayBatDau As Date
    Dim iInterval As Integer
    Dim dNgayHienTai As Date
        
    dNgayBatDau = DateSerial(Year(Now), dThangTaiChinh, dNgayTaiChinh)
    iInterval = DateDiff("D", Date, dNgayBatDau)
    
    If iInterval <= 0 Then
        GetNamHienTai = Year(Date)
    Else
        GetNamHienTai = Year(Date) - 1
    End If
    
End Function

Public Function GetNgayDauQuy(q As Integer, Y As Integer, dNgayTaiChinh As Integer, dThangTaiChinh As Integer) As Date
    Dim mTaiChinh As Integer
    Dim yTaiChinh As Integer
    
    mTaiChinh = (q - 1) * 3 + dThangTaiChinh
    yTaiChinh = Y
    If mTaiChinh > 12 Then
        mTaiChinh = mTaiChinh - 12
        yTaiChinh = Y + 1
    End If
    GetNgayDauQuy = DateSerial(yTaiChinh, mTaiChinh, dNgayTaiChinh)
End Function

Public Function GetNgayCuoiQuy(q As Integer, Y As Integer, dNgayTaiChinh As Integer, dThangTaiChinh As Integer) As Date
    Dim mTaiChinhDau As Integer
    Dim mTaiChinhCuoi As Integer
    Dim yTaiChinhDau As Integer
    Dim yTaiChinhCuoi As Integer
    Dim iInterval As Integer
    
    mTaiChinhDau = (q - 1) * 3 + dThangTaiChinh + 2 'Thang cuoi quy
    If dNgayTaiChinh = 1 Then
        mTaiChinhCuoi = mTaiChinhDau + 1 'Thang dau quy sau
        yTaiChinhDau = Y
        yTaiChinhCuoi = Y
        If mTaiChinhDau > 12 Then
            mTaiChinhDau = mTaiChinhDau - 12
            yTaiChinhDau = Y + 1
        End If
        If mTaiChinhCuoi > 12 Then
            mTaiChinhCuoi = mTaiChinhCuoi - 12
            yTaiChinhCuoi = Y + 1
        End If
        
        'Limitation of year
        If yTaiChinhCuoi >= 10000 Then
            yTaiChinhCuoi = 9999
        End If
        
        iInterval = DateDiff("D", DateSerial(yTaiChinhDau, mTaiChinhDau, 1), DateSerial(yTaiChinhCuoi, mTaiChinhCuoi, 1)) - 1
        GetNgayCuoiQuy = DateSerial(yTaiChinhDau, mTaiChinhDau, 1) + iInterval
    Else
        GetNgayCuoiQuy = DateSerial(yTaiChinhDau, mTaiChinhDau, 1)
    End If
End Function

Public Function GetNgayCuoiThang(intYear As Integer, intMonth As Integer) As Date
    Dim ValidityDate As Date
    
    Select Case intMonth
        Case 1
            ValidityDate = format("31/" & format(intMonth, "00") & "/" & format(intYear, "0000"), "dd/mm/yyyy")
        Case 2
             If CInt(format(intYear, "0000")) / 4 = CInt(format(intYear, "0000")) \ 4 And CInt(format(intYear, "0000")) \ 100 <> CInt(format(intYear, "0000")) / 100 Then
                ValidityDate = format("29/" & format(intMonth, "00") & "/" & format(intYear, "0000"), "dd/mm/yyyy")
            Else
                ValidityDate = format("28/" & format(intMonth, "00") & "/" & format(intYear, "0000"), "dd/mm/yyyy")
            End If
        Case 3
            ValidityDate = format("31/" & format(intMonth, "00") & "/" & format(intYear, "0000"), "dd/mm/yyyy")
        Case 4
            ValidityDate = format("30/" & format(intMonth, "00") & "/" & format(intYear, "0000"), "dd/mm/yyyy")
        Case 5
            ValidityDate = format("31/" & format(intMonth, "00") & "/" & format(intYear, "0000"), "dd/mm/yyyy")
        Case 6
            ValidityDate = format("30/" & format(intMonth, "00") & "/" & format(intYear, "0000"), "dd/mm/yyyy")
        Case 7
            ValidityDate = format("31/" & format(intMonth, "00") & "/" & format(intYear, "0000"), "dd/mm/yyyy")
        Case 8
            ValidityDate = format("31/" & format(intMonth, "00") & "/" & format(intYear, "0000"), "dd/mm/yyyy")
        Case 9
            ValidityDate = format("30/" & format(intMonth, "00") & "/" & format(intYear, "0000"), "dd/mm/yyyy")
        Case 10
            ValidityDate = format("31/" & format(intMonth, "00") & "/" & format(intYear, "0000"), "dd/mm/yyyy")
        Case 11
            ValidityDate = format("30/" & format(intMonth, "00") & "/" & format(intYear, "0000"), "dd/mm/yyyy")
        Case 12
            ValidityDate = format("31/" & format(intMonth, "00") & "/" & format(intYear, "0000"), "dd/mm/yyyy")
    End Select
    
    GetNgayCuoiThang = ValidityDate
End Function

Public Function s2d(d As String) As Date
   Dim strFormat() As String
    strFormat = Split(d, "/")
    s2d = DateSerial(strFormat(2), strFormat(1), strFormat(0))
    
End Function

Private Function numberb2d(fd As String, td As String) As Integer
    numberb2d = DateDiff("d", s2d(fd), s2d(td))
    If numberb2d <= 0 Then numberb2d = 0
End Function

' ham tinh so ngay chenh lech
Public Function getSoNgay(fd As String, td As String) As Long
    getSoNgay = DateDiff("d", s2d(fd), s2d(td))
    If getSoNgay <= 0 Then getSoNgay = 0
End Function

Public Function NgayCuoiNamTaiChinh(Y As Integer, dThangTaiChinh As Integer, dNgayTaiChinh As Integer) As Date
    Dim dNgayTC As Date
    
    dNgayTC = DateSerial(Y, dThangTaiChinh, dNgayTaiChinh)
    NgayCuoiNamTaiChinh = DateAdd("M", 12, dNgayTC)
    NgayCuoiNamTaiChinh = DateAdd("d", -1, NgayCuoiNamTaiChinh)
End Function

Public Sub SetSheetVisible(fpSpread1 As fpSpread)
    Dim xmlSheetNode As MSXML.IXMLDOMNode
    Dim intCtrl As Integer
    
    With fpSpread1
        For intCtrl = 1 To .SheetCount
            .sheet = intCtrl
            For Each xmlSheetNode In TAX_Utilities_v2.NodeValidity.childNodes
                If .SheetName = GetAttribute(xmlSheetNode, "Caption") Then
                    If GetAttribute(xmlSheetNode, "Active") = "0" Then
                        .SheetVisible = False
                    End If
                    Exit For
                End If
            Next
        Next intCtrl
    End With
End Sub

Public Function GetMessageCellById(ByVal strId As String) As MSXML.IXMLDOMNode
    Dim xmlInforNode As MSXML.IXMLDOMNode
    
    For Each xmlInforNode In TAX_Utilities_v2.NodeMessage
        If GetAttribute(xmlInforNode, "ID") = strId Then
            Set GetMessageCellById = xmlInforNode
            Exit Function
        End If
    Next
End Function

Public Function LoadSessionValueFromFile(ByVal strFileName As String) As Boolean
    Dim lFileNum As Long
    Dim intCtrl As Integer
    Dim strData As String
    Dim arrStrData() As String
    Dim fso As New FileSystemObject
    Dim strPeriod As String
    
    lFileNum = FreeFile
    
    On Error GoTo ErrInvalidFileHandle
    
    '**********************************
    ' ThanhDX added
    ' Date: 20/05/06
    ' Check exist of Session file
    
    'If Dir(strFileName) = "" Then
        ' Create file
        'Open strFileName For Binary Access Write As #lFileNum
        'PutString lFileNum, ""
        'Close #lFileNum
    'End If
    
    Open strFileName For Binary Access Read As #lFileNum
    strData = DeCompress(GetString(lFileNum))
    Close #lFileNum
    
    arrStrData = Split(strData, ":")
    
    If GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then
        strPeriod = format(TAX_Utilities_v2.month, "00") & "/" & format(TAX_Utilities_v2.Year, "0000")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then
        strPeriod = format(TAX_Utilities_v2.ThreeMonths, "00") & "/" & format(TAX_Utilities_v2.Year, "0000")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "1" Then
        strPeriod = format(TAX_Utilities_v2.Year, "0000")
    Else
        strPeriod = ""
    End If
    
    For intCtrl = 0 To UBound(arrStrData)
        If intCtrl Mod 4 = 0 And intCtrl < UBound(arrStrData) Then
            If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = arrStrData(intCtrl) And strPeriod = arrStrData(intCtrl + 1) Then
                intDataSession = CInt(arrStrData(intCtrl + 2))
                intPrintingSession = CInt(arrStrData(intCtrl + 3))
                Exit For
            End If
        End If
    Next intCtrl
    
    If intCtrl = UBound(arrStrData) + 1 Then
        intDataSession = 0
        intPrintingSession = 0
        Open strFileName For Binary Access Write As #lFileNum
        PutString lFileNum, Compress(strData & IIf(strData <> "", ":", "") & GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") & ":" & strPeriod & ":0:0")
        Close #lFileNum
    End If
    
    LoadSessionValueFromFile = True
    'Reset
    Exit Function
ErrInvalidFileHandle:
    'Invalid data
    SaveErrorLog "mdlFunctions", "LoadSessionValueFromFile", Err.Number, Err.Description
    DisplayMessage "", msOKOnly, miCriticalError, , mrOK
    
End Function

Private Function GetString(ByVal Filenumber As Integer) As String
    'Dim StrLengthlong As Long
    Dim StrLength As Long
    
    Get #Filenumber, , StrLength
    'StrLength = StrLengthInt

    GetString = String$(StrLength, " ")
    Get #Filenumber, , GetString
End Function

Private Sub PutString(ByVal Filenumber As Integer, Strng As String)
    Put #Filenumber, , CLng(Len(Strng))
    Put #Filenumber, , Strng
End Sub

Public Function SaveSessionValueToFile(ByVal strFileName As String, Optional ByVal blnPrintingSession As Boolean = True, Optional ByVal blnDataSession As Boolean = True) As Boolean
    Dim lFileNum As Long
    Dim intCtrl As Integer
    Dim strData As String
    Dim arrStrData() As String
'    Dim fso As New FileSystemObject

'    If Not fso.FileExists(strFileName) Then
'        'File not exist
'        'DisplayMessage "", msOKOnly, miCriticalError, , mrOK
'        Exit Function
'    End If
    
    Dim strPeriod As String
    
    lFileNum = FreeFile
    
    If GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then
        strPeriod = format(TAX_Utilities_v2.month, "00") & "/" & format(TAX_Utilities_v2.Year, "0000")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then
        strPeriod = format(TAX_Utilities_v2.ThreeMonths, "00") & "/" & format(TAX_Utilities_v2.Year, "0000")
    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "1" Then
        strPeriod = format(TAX_Utilities_v2.Year, "0000")
    Else
        strPeriod = format(TAX_Utilities_v2.Year, "")
    End If
    
    On Error GoTo ErrInvalidFileHandle
    Open strFileName For Binary Access Read As #lFileNum
    strData = GetString(lFileNum) 'DeCompress(GetString(lFileNum))
    If strData <> vbNullString Then
        strData = DeCompress(strData)
    End If
    
    Close #lFileNum
    
    arrStrData = Split(strData, ":")
    
    On Error GoTo ErrHandle
    For intCtrl = 0 To UBound(arrStrData)
        If intCtrl Mod 4 = 0 Then
            If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = arrStrData(intCtrl) And strPeriod = arrStrData(intCtrl + 1) Then
                If blnDataSession Then arrStrData(intCtrl + 2) = CStr(intDataSession)
                If blnPrintingSession Then arrStrData(intCtrl + 3) = CStr(intPrintingSession)
                Exit For
            End If
        End If
    Next intCtrl
    
    On Error GoTo ErrInvalidFileHandle
    Open strFileName For Binary Access Write As #lFileNum
    PutString lFileNum, Compress(Join(arrStrData, ":"))
    Close #lFileNum
    
    SaveSessionValueToFile = True
    
    Exit Function
ErrHandle:
    SaveErrorLog "mdlFunctions", "SaveSessionValueToFile", Err.Number, Err.Description
    Exit Function
ErrInvalidFileHandle:
    'Invalid data
    'DisplayMessage "", msOKOnly, miCriticalError, , mrOK
End Function

Public Sub PrinterKillDoc()
    Printer.KillDoc
    Printer.PaperSize = vbPRPSA4
End Sub

Public Sub PrinterEndDoc()
    Printer.EndDoc
    Printer.PaperSize = vbPRPSA4
End Sub
'
'Public Function LoadXML(ByRef xmlDom As MSXML.DOMDocument, ByVal strFileName As String) As Boolean
'    Dim lngFileNum As Long
'    Dim lngLength As Long
'    Dim strValue As String
'
'    If Dir(strFileName) = vbNullString Then Exit Function
'
'    lngFileNum = FreeFile
'    Open strFileName For Binary Access Read As #lngFileNum
'    Get #lngFileNum, , lngLength
'    strValue = String$(lngLength, " ")
'    Get #lngFileNum, , strValue
'    Close #lngFileNum
'
'    strValue = TAX_Utilities_v2.Convert(TAX_Utilities_v2.DeCompress(strValue), TCVN, UNICODE)
'
'    LoadXML = xmlDom.LoadXML(strValue)
'End Function

'Public Function SaveXML(ByVal xmlDom As MSXML.DOMDocument, ByVal strFileName As String) As Boolean
'    Dim lngFileNum As Long
'    Dim lngLength As Long
'    Dim strValue As String
'
'    If Dir(strFileName) <> vbNullString Then
'        Kill strFileName
'        'Exit Function
'    End If
'    strValue = xmlDom.xml
'    strValue = TAX_Utilities_v2.Compress(TAX_Utilities_v2.Convert(strValue, UNICODE, TCVN))
'
'    lngFileNum = FreeFile
'    Open strFileName For Binary Access Write As #lngFileNum
'
'    Put #lngFileNum, , CLng(Len(strValue))
'    Put #lngFileNum, , strValue
'    Close #lngFileNum
'
'    SaveXML = True
'    Exit Function
'ErrHandle:
'    SaveErrorLog "mdlFunctions", "SaveXML", Err.Number, Err.Description
'End Function

Public Sub SetupDataKHBS(pGrid As fpSpread)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim lSheet As Long
    Dim blnNewData As Boolean, blnHasSetActiveCell As Boolean
    Dim blnExistData As Boolean
    Dim strKHBSDataFileName As String
    Dim strDataFileName As String
    Dim strOriginDataFileName As String
    Dim varTemp As Variant
    Dim blnResetdata As Boolean
    Dim strLastFileNam As String
    Dim strDataLastKHBS As String
    blnResetdata = False

    
    blnExistData = False
                If GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_v2.NodeMenu, "Year") = "0" Then
                    strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                    strKHBSDataFileName = TAX_Utilities_v2.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                Else
                    If GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "0" Then
                        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                        strKHBSDataFileName = TAX_Utilities_v2.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & "_" & TAX_Utilities_v2.DateKHBS & ".xml"
                    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then
                        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & ".xml"
                        strKHBSDataFileName = TAX_Utilities_v2.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & "_" & TAX_Utilities_v2.DateKHBS & ".xml"
                    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "0" Then
                            'Data file contain Day from and to.
                            strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                            & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & ".xml"
                             strKHBSDataFileName = TAX_Utilities_v2.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                            & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & "_" & TAX_Utilities_v2.DateKHBS & ".xml"
                    ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then
                            'Data file contain Day from and to.
                            strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                            & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & ".xml"
                            strKHBSDataFileName = TAX_Utilities_v2.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                            & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & "_" & TAX_Utilities_v2.DateKHBS & ".xml"
                    Else
                            'Data file not contain Day from and to.
                            strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                            & TAX_Utilities_v2.Year & ".xml"
                            strKHBSDataFileName = TAX_Utilities_v2.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & "_" _
                            & TAX_Utilities_v2.Year & "_" & TAX_Utilities_v2.DateKHBS & ".xml"
                    '*********************************
                    End If
                End If
    
    Set arrErrCells = New Scripting.Dictionary
    TAX_Utilities_v2.xmlDataReDim (TAX_Utilities_v2.NodeValidity.childNodes.length - 1)
    TAX_Utilities_v2.Data(0) = New MSXML.DOMDocument
    TAX_Utilities_v2.Data(0).resolveExternals = True
    TAX_Utilities_v2.Data(0).validateOnParse = True
    TAX_Utilities_v2.Data(0).async = False
    
    TAX_Utilities_v2.Data(0).Load strKHBSDataFileName
    
    strDataLastKHBS = Replace(GetLastfileName, "KHBS_", "KHBS1_")
    
    
    TAX_Utilities_v2.DataKHBS = New MSXML.DOMDocument
    TAX_Utilities_v2.DataKHBS.resolveExternals = True
    TAX_Utilities_v2.DataKHBS.validateOnParse = True
    TAX_Utilities_v2.DataKHBS.async = False
    TAX_Utilities_v2.DataKHBS.Load strDataLastKHBS
    
    
    
    If TAX_Utilities_v2.Data(lSheet).parseError.reason <> vbNullString Then
        If InStr(1, TAX_Utilities_v2.Data(lSheet).parseError.errorCode, "2146697210") <> 0 Then
            
            strLastFileNam = GetLastfileName
            strDataLastKHBS = Replace(strLastFileNam, "KHBS_", "KHBS1_")
            TAX_Utilities_v2.Data(0).Load strLastFileNam
            
                If TAX_Utilities_v2.Data(0).parseError.reason <> vbNullString Then
                    If InStr(1, TAX_Utilities_v2.Data(0).parseError.errorCode, "2146697210") <> 0 Then
                                strOriginDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "TemplateFolder")) & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & ".xml"
                                
                                If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "05" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "06" Or _
                                    GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "08" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "09" Then
                                     TAX_Utilities_v2.Data(0).Load strDataFileName
                                Else
                                    TAX_Utilities_v2.Data(0).Load strOriginDataFileName
                                End If
                                
                                TAX_Utilities_v2.DataKHBS.Load strDataFileName
                    Else
                        MsgBox TAX_Utilities_v2.Data(0).parseError.reason
                    End If
                Else
                    If strKHBS = "frmKHBS_BS" Then
                        TAX_Utilities_v2.Data(0).getElementsByTagName("Sections")(0).removeChild TAX_Utilities_v2.Data(0).getElementsByTagName("Section")(0)
                        TAX_Utilities_v2.Data(0).getElementsByTagName("Sections")(0).removeChild TAX_Utilities_v2.Data(0).getElementsByTagName("Section")(0)
                        TAX_Utilities_v2.Data(0).getElementsByTagName("Sections")(0).removeChild TAX_Utilities_v2.Data(0).getElementsByTagName("Section")(0)
                        TAX_Utilities_v2.DataKHBS.Load strDataLastKHBS
                     End If
                End If
            
        Else
            MsgBox TAX_Utilities_v2.Data(0).parseError.reason
        End If
    Else
        blnExistData = True
'        If GetAttribute(TAX_Utilities_v2.Data(0).childNodes(2).firstChild, "loaiKHBS") <> strHiddenFormName Then
'         Dim lResult As VbMsgBoxResult
'            lResult = DisplayMessage("0119", msYesNo, miQuestion, , mrNo)
'                If lResult = mrYes Then
'                   strLastFileNam = GetLastfileName
'                   strDataLastKHBS = Replace(strLastFileNam, "KHBS_", "KHBS1_")
'                    TAX_Utilities_v2.Data(0).Load strDataLastKHBS
'                    If TAX_Utilities_v2.Data(0).parseError.reason <> vbNullString Then
'                        If InStr(1, TAX_Utilities_v2.Data(0).parseError.reason, "The system cannot locate the object specified.") <> 0 Then
'                                    TAX_Utilities_v2.Data(0).Load strDataFileName
'                                    strDataFileName = Replace(strDataFileName, "KHBS_", "KHBS1_")
'                                    TAX_Utilities_v2.DataKHBS.Load strDataFileName
'                        Else
'                            MsgBox TAX_Utilities_v2.Data(0).parseError.reason
'                        End If
'                    Else
''                    TAX_Utilities_v2.Data(0).getElementsByTagName("Sections")(0).removeChild TAX_Utilities_v2.Data(0).getElementsByTagName("Section")(0)
''                    TAX_Utilities_v2.Data(0).getElementsByTagName("Sections")(0).removeChild TAX_Utilities_v2.Data(0).getElementsByTagName("Section")(0)
''                    TAX_Utilities_v2.Data(0).getElementsByTagName("Sections")(0).removeChild TAX_Utilities_v2.Data(0).getElementsByTagName("Section")(0)
'
'                    TAX_Utilities_v2.DataKHBS.Load strDataLastKHBS
'                    End If
'                 blnExistData = False
'                Else
'                strHiddenFormName = GetAttribute(TAX_Utilities_v2.Data(0).childNodes(2).firstChild, "loaiKHBS")
'                End If
'        End If
'        If GetAttribute(TAX_Utilities_v2.Data(0).childNodes(2).firstChild, "loaiKHBS") = "frmKHBS_BS" Then
'           blnResetdata = True
'        End If
    End If
    
    If blnExistData = True Then
        Dim xmlNodeSession As MSXML.IXMLDOMNode
        Dim xmlNodeListSession As MSXML.IXMLDOMNodeList
        Dim xmlNode As MSXML.IXMLDOMNode
        Dim i As Integer
        TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1) = New MSXML.DOMDocument
        TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).resolveExternals = True
        TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).validateOnParse = True
        TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).async = False
        strOriginDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "TemplateFolder")) & "KHBS.xml"
        TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).Load strOriginDataFileName
        TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Sections")(0).replaceChild TAX_Utilities_v2.Data(0).getElementsByTagName("Section")(2), TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Section")(2)
        TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Sections")(0).replaceChild TAX_Utilities_v2.Data(0).getElementsByTagName("Section")(1), TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Section")(1)
        TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Sections")(0).replaceChild TAX_Utilities_v2.Data(0).getElementsByTagName("Section")(0), TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Section")(0)
        Set xmlNodeListCell = TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
        FillData pGrid, xmlNodeListCell, 1, False
        Set xmlNodeListCell = TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")
        FillData pGrid, xmlNodeListCell, pGrid.SheetCount - 1, False
        SetAttribute TAX_Utilities_v2.NodeValidity.childNodes(TAX_Utilities_v2.NodeValidity.childNodes.length - 1), "Active", "1"
        For lSheet = 1 To TAX_Utilities_v2.xmlDataCount - 1
                strOriginDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "TemplateFolder")) & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                TAX_Utilities_v2.Data(lSheet) = New MSXML.DOMDocument
                TAX_Utilities_v2.Data(lSheet).resolveExternals = True
                TAX_Utilities_v2.Data(lSheet).validateOnParse = True
                TAX_Utilities_v2.Data(lSheet).async = False
                TAX_Utilities_v2.Data(lSheet).Load strOriginDataFileName
                SetAttribute TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "Active", "0"
        Next
    Else
        For lSheet = 0 To TAX_Utilities_v2.xmlDataCount
           If lSheet = 0 Then
            Set xmlNodeListCell = TAX_Utilities_v2.Data(lSheet).getElementsByTagName("Cell")
                FillData pGrid, xmlNodeListCell, 1, False
                SetAttribute TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "Active", "1"
           Else
                strOriginDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "TemplateFolder")) & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                TAX_Utilities_v2.Data(lSheet) = New MSXML.DOMDocument
                TAX_Utilities_v2.Data(lSheet).resolveExternals = True
                TAX_Utilities_v2.Data(lSheet).validateOnParse = True
                TAX_Utilities_v2.Data(lSheet).async = False
                TAX_Utilities_v2.Data(lSheet).Load strOriginDataFileName
                SetAttribute TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "Active", "0"
           End If
           If lSheet = TAX_Utilities_v2.xmlDataCount Then
                SetAttribute TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "Active", "1"
                Dim songaynopcham As Long
                Dim hannop As String
                Dim ngayKHBS  As String
                 Dim dNgayCuoiKy As Date
                
                If TAX_Utilities_v2.month <> "" Then
                    If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "01" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "04" _
                            Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "96" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "94" Then
                        If strQuy = "TK_THANG" Then
                            If TAX_Utilities_v2.month = 12 Then
                                hannop = "20/" & "01" & "/" & TAX_Utilities_v2.Year + 1
                            Else
                                hannop = "20/" & Right("00" & TAX_Utilities_v2.month + 1, 2) & "/" & TAX_Utilities_v2.Year
                            End If
                        ElseIf strQuy = "TK_QUY" Then
                            If TAX_Utilities_v2.ThreeMonths = "04" Then
                               hannop = "31/" & "01" & "/" & TAX_Utilities_v2.Year + 1
                            ElseIf TAX_Utilities_v2.ThreeMonths = "03" Then
                                hannop = "31/" & "10" & "/" & TAX_Utilities_v2.Year
                            ElseIf TAX_Utilities_v2.ThreeMonths = "02" Then
                                hannop = "31/" & "07" & "/" & TAX_Utilities_v2.Year
                            ElseIf TAX_Utilities_v2.ThreeMonths = "01" Then
                                hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                            End If
                        End If
                    Else
                        If TAX_Utilities_v2.month = 12 Then
                            hannop = "20/" & "01" & "/" & TAX_Utilities_v2.Year + 1
                        Else
                            hannop = "20/" & Right("00" & TAX_Utilities_v2.month + 1, 2) & "/" & TAX_Utilities_v2.Year
                        End If
                    End If
                ElseIf TAX_Utilities_v2.ThreeMonths <> "" Then
                    If TAX_Utilities_v2.ThreeMonths = "04" Then
                       hannop = "31/" & "01" & "/" & TAX_Utilities_v2.Year + 1
                    ElseIf TAX_Utilities_v2.ThreeMonths = "03" Then
                        hannop = "31/" & "10" & "/" & TAX_Utilities_v2.Year
                    ElseIf TAX_Utilities_v2.ThreeMonths = "02" Then
                        hannop = "31/" & "07" & "/" & TAX_Utilities_v2.Year
                    ElseIf TAX_Utilities_v2.ThreeMonths = "01" Then
                        hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                    End If
'                    dNgayCuoiKy = DateAdd("D", 30, GetNgayCuoiQuy(TAX_Utilities_v2.ThreeMonths, CInt(TAX_Utilities_v2.Year), iNgayTaiChinh, iThangTaiChinh))
'                    hannop = format(dNgayCuoiKy, "dd/mm/yyyy")
                Else
                    dNgayCuoiKy = DateAdd("D", 90, NgayCuoiNamTaiChinh(TAX_Utilities_v2.Year, iThangTaiChinh, iNgayTaiChinh))
                    hannop = format(dNgayCuoiKy, "dd/mm/yyyy")
                End If
        
        'Neu vao ngay thu 7 thi cong them 2 ngay,  ngay CN thi cong them mot ngay
                If Weekday(CDate(hannop)) = 7 Then
                    hannop = DateAdd("D", 2, CDate(hannop))
                    hannop = format(hannop, "dd/mm/yyyy")
                ElseIf Weekday(CDate(hannop)) = 1 Then
                    hannop = DateAdd("D", 1, CDate(hannop))
                    hannop = format(hannop, "dd/mm/yyyy")
                End If
                
                
                ngayKHBS = Mid(TAX_Utilities_v2.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_v2.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_v2.DateKHBS, 5, 4)
                songaynopcham = numberb2d(hannop, ngayKHBS)
                SetAttribute TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).nodeFromID("B_24"), "Value", hannop
                SetAttribute TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).nodeFromID("BE_17"), "Value", CStr(songaynopcham)
                SetAttribute TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).nodeFromID("BG_23"), "Value", CStr(format(Date, "dd/mm/yyyy"))
                With pGrid
                    .sheet = .SheetCount - 1
                    .SetText .ColLetterToNumber("B"), 24, hannop
                    .SetText .ColLetterToNumber("BE"), 17, songaynopcham
                    .SetText .ColLetterToNumber("BG"), 23, CStr(format(Date, "dd/mm/yyyy"))
                    .Col = .ColLetterToNumber("BG")
                    .Row = 22
                     SetAttribute TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).nodeFromID("BG_22"), "Value", .Text
                    
                End With
                
            End If
        Next
        
        Set xmlNodeSession = Nothing
        Set xmlNodeListCell = Nothing
        Set xmlNode = Nothing
    End If
    
        
    If blnExistData = False Then
       ResetKHBSData pGrid, False
    End If
        
        
        
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "SetupData", Err.Number, Err.Description
End Sub
Public Sub SetupDataKHBS_TT28(pGrid As fpSpread)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim lSheet As Long
    Dim blnNewData As Boolean, blnHasSetActiveCell As Boolean
    Dim blnExistData As Boolean
    Dim strKHBSDataFileName As String
    Dim strDataFileName As String
    Dim strOriginDataFileName As String
    Dim varTemp As Variant
    Dim blnResetdata As Boolean
    Dim strLastFileNam As String
    Dim strDataLastKHBS As String
    
    Dim strarrdate() As String ' su dung cho to khai 02/NTNN va 04/NTNN
                'SetAttribute TAX_Utilities_v2.NodeValidity.childNodes(lSheet), "Active", "1"
                Dim songaynopcham As Long
                Dim hannop As String
                Dim ngayKHBS  As String
                 Dim dNgayCuoiKy As Date
                
                If TAX_Utilities_v2.month <> "" Then
                    ' To khai 01/GTGT gia han thang 4,5,6 nam 2012 -> tinh lai han nop
                    If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "01" Then
                         If strQuy = "TK_THANG" Then
                            If (TAX_Utilities_v2.month = 4 Or TAX_Utilities_v2.month = 5 Or TAX_Utilities_v2.month = 6) And TAX_Utilities_v2.Year = 2012 And TAX_Utilities_v2.CheckToKhaiGH = True Then
                                If TAX_Utilities_v2.month = 4 Then
                                    hannop = "20/" & "11" & "/" & TAX_Utilities_v2.Year
                                ElseIf TAX_Utilities_v2.month = 5 Then
                                    hannop = "20/" & "12" & "/" & TAX_Utilities_v2.Year
                                ElseIf TAX_Utilities_v2.month = 6 Then
                                    hannop = "21/" & "01" & "/" & TAX_Utilities_v2.Year + 1
                                End If
                            Else
                                ' cac ky ke khai khac van tinh han nop binh thuong
                                If TAX_Utilities_v2.month = 12 Then
                                    hannop = "20/" & "01" & "/" & TAX_Utilities_v2.Year + 1
'                                ElseIf TAX_Utilities_v2.month = 4 Then
'                                    hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                                Else
                                    hannop = "20/" & Right("00" & TAX_Utilities_v2.month + 1, 2) & "/" & TAX_Utilities_v2.Year
                                End If
                            End If
                        ElseIf strQuy = "TK_QUY" Then
                            If Val(TAX_Utilities_v2.ThreeMonths) = 4 Then
                               hannop = "31/" & "01" & "/" & TAX_Utilities_v2.Year + 1
                            ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 3 Then
                                hannop = "31/" & "10" & "/" & TAX_Utilities_v2.Year
                            ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 2 Then
                                hannop = "31/" & "07" & "/" & TAX_Utilities_v2.Year
                            ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 1 Then
                                hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                            End If
                            hannop = format(hannop, "dd/mm/yyyy")
                        End If
                    ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "04" _
                            Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "96" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "94" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "99" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "98" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "92" Then
                        If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "99" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "98" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "92" Then
                           If strQuy = "TK_THANG" Then
                                 ' cac to khai thang khac van tinh binh thuong
                                If TAX_Utilities_v2.month = 12 Then
                                    hannop = "20/" & "01" & "/" & TAX_Utilities_v2.Year + 1
    '                            ElseIf TAX_Utilities_v2.month = 4 Then
    '                                hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                                Else
                                    hannop = "20/" & Right("00" & TAX_Utilities_v2.month + 1, 2) & "/" & TAX_Utilities_v2.Year
                                End If
                            ElseIf strQuy = "TK_QUY" Then
                                If Val(TAX_Utilities_v2.ThreeMonths) = 4 Then
                                   hannop = "31/" & "01" & "/" & TAX_Utilities_v2.Year + 1
                                ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 3 Then
                                    hannop = "31/" & "10" & "/" & TAX_Utilities_v2.Year
                                ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 2 Then
                                    hannop = "31/" & "07" & "/" & TAX_Utilities_v2.Year
                                ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 1 Then
                                    hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                                End If
                                hannop = format(hannop, "dd/mm/yyyy")
                            ElseIf strQuy = "TK_LANPS" Then
                                hannop = format(DateAdd("D", 10, DateSerial(CInt(TAX_Utilities_v2.Year), CInt(TAX_Utilities_v2.month), CInt(TAX_Utilities_v2.Day))), "dd/mm/yyyy")
                            ElseIf strQuy = "TK_LANXB" Then
                                hannop = format(DateAdd("D", 35, DateSerial(CInt(TAX_Utilities_v2.Year), CInt(TAX_Utilities_v2.month), CInt(TAX_Utilities_v2.Day))), "dd/mm/yyyy")
                            End If

                        Else
                            
                            If strQuy = "TK_THANG" Then
                                 ' cac to khai thang khac van tinh binh thuong
                                If TAX_Utilities_v2.month = 12 Then
                                    hannop = "20/" & "01" & "/" & TAX_Utilities_v2.Year + 1
    '                            ElseIf TAX_Utilities_v2.month = 4 Then
    '                                hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                                Else
                                    hannop = "20/" & Right("00" & TAX_Utilities_v2.month + 1, 2) & "/" & TAX_Utilities_v2.Year
                                End If
                            ElseIf strQuy = "TK_QUY" Then
                                If Val(TAX_Utilities_v2.ThreeMonths) = 4 Then
                                   hannop = "31/" & "01" & "/" & TAX_Utilities_v2.Year + 1
                                ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 3 Then
                                    hannop = "31/" & "10" & "/" & TAX_Utilities_v2.Year
                                ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 2 Then
                                    hannop = "31/" & "07" & "/" & TAX_Utilities_v2.Year
                                ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 1 Then
                                    hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                                End If
                                hannop = format(hannop, "dd/mm/yyyy")
                            End If
                        End If
                    Else
                        If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "72" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "73" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "56" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "70" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "81" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "06" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "05" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "90" Then
                            If strLoaiTKThang_PS = "TK_LANPS" Then
                                hannop = DateAdd("D", 10, DateSerial(CInt(TAX_Utilities_v2.Year), CInt(TAX_Utilities_v2.month), CInt(TAX_Utilities_v2.Day)))
                            Else
                                ' cac to khai thang khac van tinh binh thuong
                                If TAX_Utilities_v2.month = 12 Then
                                    hannop = "20/" & "01" & "/" & TAX_Utilities_v2.Year + 1
        '                        ElseIf TAX_Utilities_v2.month = 4 Then
        '                            hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                                Else
                                    hannop = "20/" & Right("00" & TAX_Utilities_v2.month + 1, 2) & "/" & TAX_Utilities_v2.Year
                                End If
                            End If
                        Else
                            ' cac to khai thang khac van tinh binh thuong
                            If TAX_Utilities_v2.month = 12 Then
                                hannop = "20/" & "01" & "/" & TAX_Utilities_v2.Year + 1
    '                        ElseIf TAX_Utilities_v2.month = 4 Then
    '                            hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                            Else
                                hannop = "20/" & Right("00" & TAX_Utilities_v2.month + 1, 2) & "/" & TAX_Utilities_v2.Year
                            End If
                        End If
                   End If
                ElseIf TAX_Utilities_v2.ThreeMonths <> "" Then
                    If Val(TAX_Utilities_v2.ThreeMonths) = 4 Then
                       hannop = "31/" & "01" & "/" & TAX_Utilities_v2.Year + 1
                    ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 3 Then
                        hannop = "31/" & "10" & "/" & TAX_Utilities_v2.Year
                    ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 2 Then
                        hannop = "31/" & "07" & "/" & TAX_Utilities_v2.Year
                    ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 1 Then
                        hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                    End If
'                    dNgayCuoiKy = DateAdd("D", 30, GetNgayCuoiQuy(TAX_Utilities_v2.ThreeMonths, CInt(TAX_Utilities_v2.Year), iNgayTaiChinh, iThangTaiChinh))
'                    hannop = format(dNgayCuoiKy, "dd/mm/yyyy")
                    hannop = format(hannop, "dd/mm/yyyy")
                Else
                    If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "80" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "82" Then
                        strarrdate = Split(TAX_Utilities_v2.LastDay, "/")
                        dNgayCuoiKy = DateAdd("D", 45, DateSerial(CInt(strarrdate(2)), CInt(strarrdate(1)), CInt(strarrdate(0))))
                        hannop = format(dNgayCuoiKy, "dd/mm/yyyy")
                    Else
                        dNgayCuoiKy = DateAdd("D", 90, NgayCuoiNamTaiChinh(TAX_Utilities_v2.Year, iThangTaiChinh, iNgayTaiChinh))
                        hannop = format(dNgayCuoiKy, "dd/mm/yyyy")
                    End If
                End If
        
        'Neu vao ngay thu 7 thi cong them 2 ngay,  ngay CN thi cong them mot ngay
                If Weekday(CDate(hannop)) = 7 Then
                    hannop = DateAdd("D", 2, CDate(hannop))
                    hannop = format(hannop, "dd/mm/yyyy")
                ElseIf Weekday(CDate(hannop)) = 1 Then
                    hannop = DateAdd("D", 1, CDate(hannop))
                    hannop = format(hannop, "dd/mm/yyyy")
                End If
                
                ' chuyen ve dinh dang string dd/mm/yyyy
                hannop = Day(hannop) & "/" & month(hannop) & "/" & Year(hannop)
                
                ngayKHBS = Mid(TAX_Utilities_v2.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_v2.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_v2.DateKHBS, 5, 4)
                songaynopcham = numberb2d(hannop, ngayKHBS)
'                SetAttribute TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).nodeFromID("B_24"), "Value", hannop
'                SetAttribute TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).nodeFromID("BE_17"), "Value", CStr(songaynopcham)
'                SetAttribute TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).nodeFromID("BG_23"), "Value", CStr(format(Date, "dd/mm/yyyy"))
                With pGrid
                    .sheet = .SheetCount - 1
                    .SetText .ColLetterToNumber("E"), 24, hannop
                    .SetText .ColLetterToNumber("BG"), 5, ngayKHBS
                    .SetText .ColLetterToNumber("BD"), 5, songaynopcham
                
                    
                    'dhdang sua load tk BS da có du lieu se ko tinh lai theo cong thuc nua
                    Dim lCol_temp As Long
                    Dim lRow_temp As Long
                    Dim xmlNodeCell_temp As MSXML.IXMLDOMNode
                    
                    Dim strIdTkhaiTT156 As String
                    Dim strIdTkCheck As String
                    strIdTkhaiTT156 = "~02~04~71~72~11~12~73~70~81~06~05~86~90~94~96~98~99~92~"
                    strIdTkCheck = GetAttribute(TAX_Utilities_v2.NodeMenu, "ID")
                        
                    If isNewdataBS = False Then
                        If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "01" Then
                                Set xmlNodeCell_temp = TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 20)
                                ParserCellID pGrid, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                                .Col = lCol_temp
                                .Row = lRow_temp
                                '.Formula = ""
                                .value = GetAttribute(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
                                (TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 20), "Value")
                                
                                
                                Set xmlNodeCell_temp = TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 19)
                                ParserCellID pGrid, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                                .Col = lCol_temp
                                .Row = lRow_temp
                                '.Formula = ""
                                '.value = GetAttribute(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 1), "Value")
                                .value = GetAttribute(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
                                (TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 19), "Value")
                            Else
                                If InStr(1, strIdTkhaiTT156, "~" & Trim$(strIdTkCheck) & "~", vbTextCompare) > 0 Then
                                    Set xmlNodeCell_temp = TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 18)
                                    ParserCellID pGrid, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                                    .Col = lCol_temp
                                    .Row = lRow_temp
                                    '.Formula = ""
                                    .value = GetAttribute(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
                                    (TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 18), "Value")
                                    
                                    
                                    Set xmlNodeCell_temp = TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 17)
                                    ParserCellID pGrid, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                                    .Col = lCol_temp
                                    .Row = lRow_temp
                                    
                                    '.Formula = ""
                                    '.value = GetAttribute(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 1), "Value")
                                    .value = GetAttribute(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
                                    (TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 17), "Value")
                                Else
                                    Set xmlNodeCell_temp = TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 7)
                                    ParserCellID pGrid, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                                    .Col = lCol_temp
                                    .Row = lRow_temp
                                    '.Formula = ""
                                    .value = GetAttribute(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
                                    (TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 7), "Value")
                                    
                                    
                                    Set xmlNodeCell_temp = TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 6)
                                    ParserCellID pGrid, GetAttribute(xmlNodeCell_temp, "CellID"), lCol_temp, lRow_temp
                                    .Col = lCol_temp
                                    .Row = lRow_temp
                                    
                                    '.Formula = ""
                                    '.value = GetAttribute(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 1), "Value")
                                    .value = GetAttribute(TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell") _
                                    (TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 6), "Value")
                                End If
                            End If
                    End If
                    '.SetText .ColLetterToNumber("BG"), 23, CStr(format(Date, "dd/mm/yyyy"))
'                    .Col = .ColLetterToNumber("BG")
'                    .Row = 22
'                     SetAttribute TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).nodeFromID("BG_22"), "Value", .Text
'                     .Col = .ColLetterToNumber("BE")
'                    .Row = 18
'                     SetAttribute TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).nodeFromID("BE_18"), "Value", .Text
'                     .Col = .ColLetterToNumber("BD")
'                     .Row = 20
'                     SetAttribute TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).nodeFromID("BD_20"), "Value", .Text
                End With
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "SetupData", Err.Number, Err.Description
End Sub


Public Sub FillData(pGrid As fpSpread, xmlNodeList As MSXML.IXMLDOMNodeList, mCurrentSheet As Integer, blnNewData As Boolean)
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    Dim lRows As Long
    Dim blnHasSetActiveCell As Boolean
    
    With pGrid
     .sheet = mCurrentSheet
    '  mCurrentSheet = lSheet + 1
      blnHasSetActiveCell = False
     .EventEnabled(EventChange) = False
      For Each xmlNodeCell In xmlNodeList
         ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
            If GetAttribute(xmlNodeCell, "FirstCell") = "1" Then
                lRows = GetDynRowCount(pGrid, xmlNodeCell.parentNode)
                InsertRow pGrid, lRow, lRows, True
            End If
       Next
      ' .EventEnabled(EventChange) = True
       
       For Each xmlNodeCell In xmlNodeList
            ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
            .Col = lCol
            .Row = lRow
            If Not .Lock And Not blnHasSetActiveCell Then
                .SetActiveCell lCol, lRow
                blnHasSetActiveCell = True
            End If
                Select Case .CellType
                    Case CellTypeCheckBox
                        ' Check box
                        If UCase(GetAttribute(xmlNodeCell, "Value")) = UCase("x") Then
                            .Text = "1"
                        Else
                            .Text = "0"
                            If GetAttribute(xmlNodeCell, "Value") <> "" And GetAttribute(xmlNodeCell, "Value") <> "0" Then
                                'Set note
                                arrErrCells.Add .sheet & "_" & GetAttribute(xmlNodeCell, "CellID"), .BackColor
                                .CellNote = GetAttribute(GetMessageCellById("0080"), "Msg")
                                .BackColor = 12713215 ' vbRed
                            End If
                            
                            SetAttribute xmlNodeCell, "Value", ""
                        End If
                    Case CellTypeComboBox ', CellTypeEdit, CellTypePic
                        If blnNewData And .Text <> GetAttribute(xmlNodeCell, "Value") Then
                            SetAttribute xmlNodeCell, "Value", .Text
                        Else
                            .Text = GetAttribute(xmlNodeCell, "Value")
                            .Col = lCol
                            .Row = lRow
                            If GetAttribute(xmlNodeCell, "Value") <> .Text Then
                                SetAttribute xmlNodeCell, "Value", .Text
                                '.Text = GetAttribute(xmlNodeCell, "Value")
                                'Set note
                                arrErrCells.Add .sheet & "_" & GetAttribute(xmlNodeCell, "CellID"), .BackColor
                                .CellNote = GetAttribute(GetMessageCellById("0079"), "Msg")
                                .BackColor = 12713215 ' vbRed
                            End If
                        End If
                    Case CellTypePic
                        If blnNewData And .Text <> GetAttribute(xmlNodeCell, "Value") Then
                            SetAttribute xmlNodeCell, "Value", .Text
                        Else
                            .Text = GetAttribute(xmlNodeCell, "Value")
                            .Col = lCol
                            .Row = lRow
                            If GetAttribute(xmlNodeCell, "Value") <> .Text Then
                                SetAttribute xmlNodeCell, "Value", .Text
                                '.Text = GetAttribute(xmlNodeCell, "Value")
                                'Set note
                                arrErrCells.Add .sheet & "_" & GetAttribute(xmlNodeCell, "CellID"), .BackColor
                                .CellNote = GetAttribute(GetMessageCellById("0080"), "Msg")
                                .BackColor = 12713215 ' vbRed
                            End If
                        End If
'*******************************
                    Case CellTypeNumber
'                        If Not .Lock Or (.Lock And .Formula = "") Then
                            If blnNewData And .value <> GetAttribute(xmlNodeCell, "Value") Then
                                SetAttribute xmlNodeCell, "Value", .value
                            Else
                                'Format numeric
                                If Not IsNumeric(GetAttribute(xmlNodeCell, "Value")) Then
                                    arrErrCells.Add .sheet & "_" & GetAttribute(xmlNodeCell, "CellID"), .BackColor
                                    .CellNote = GetAttribute(GetMessageCellById("0080"), "Msg")
                                    .BackColor = 12713215 ' vbRed
                                End If
                                
                                SetAttribute xmlNodeCell, "Value", IIf(Not IsNumeric(GetAttribute(xmlNodeCell, "Value")), "0", GetAttribute(xmlNodeCell, "Value"))
                                
                                'Neu gia tri nam ngoai pham vi
                                'If Not .Lock Then
                                    If Val(GetAttribute(xmlNodeCell, "Value")) > .TypeNumberMax Or Val(GetAttribute(xmlNodeCell, "Value")) < .TypeNumberMin Then
                                        SetAttribute xmlNodeCell, "Value", "0"
                                        'Set note
                                        arrErrCells.Add .sheet & "_" & GetAttribute(xmlNodeCell, "CellID"), .BackColor
                                        .CellNote = GetAttribute(GetMessageCellById("0077"), "Msg") & "[" & .TypeNumberMin & ";" & .TypeNumberMax & "]"
                                        .BackColor = 12713215 ' vbRed
                                    End If
                                'End If
                                
                                .value = GetAttribute(xmlNodeCell, "Value")
                            End If
'                        End If
                    Case CellTypeEdit
                        If blnNewData And .Text <> GetAttribute(xmlNodeCell, "Value") Then
                            SetAttribute xmlNodeCell, "Value", .Text
                        Else
                            .Text = GetAttribute(xmlNodeCell, "Value")
                            .Col = lCol
                            .Row = lRow
                            If GetAttribute(xmlNodeCell, "Value") <> .Text Then
                                SetAttribute xmlNodeCell, "Value", .Text
                                '.Text = GetAttribute(xmlNodeCell, "Value")
                                'Set note
                                arrErrCells.Add .sheet & "_" & GetAttribute(xmlNodeCell, "CellID"), .BackColor
                                .CellNote = GetAttribute(GetMessageCellById("0078"), "Msg") & .TypeMaxEditLen
                                .BackColor = 12713215 ' vbRed
                            End If
                        End If
                '*******************************
                    Case CellTypePercent
                        If Not .Lock Or (.Lock And .Formula = "") Then
                            If blnNewData And .value <> GetAttribute(xmlNodeCell, "Value") Then
                                SetAttribute xmlNodeCell, "Value", .value
                            Else
                            .Text = GetAttribute(xmlNodeCell, "Value")
                            .Col = lCol
                            .Row = lRow
                            If GetAttribute(xmlNodeCell, "Value") <> .Text Then
                                SetAttribute xmlNodeCell, "Value", .Text
                                '.Value = GetAttribute(xmlNodeCell, "Value")
                                'Set note
'                                    arrErrCells.Add .sheet & "_" & GetAttribute(xmlNodeCell, "CellID"), .BackColor
'                                    .CellNote = GetAttribute(GetMessageCellById("0080"), "Msg")
'                                    .BackColor = 12713215 ' vbRed
                            End If
                            End If
                        End If
                    Case Else
                        If blnNewData And .value <> GetAttribute(xmlNodeCell, "Value") Then
                            SetAttribute xmlNodeCell, "Value", .value
                        Else
                            .value = GetAttribute(xmlNodeCell, "Value")
                            .Col = lCol
                            .Row = lRow
                            If GetAttribute(xmlNodeCell, "Value") <> .value Then
                                SetAttribute xmlNodeCell, "Value", .value
                                '.Value = GetAttribute(xmlNodeCell, "Value")
                                'Set note
                                arrErrCells.Add .sheet & "_" & GetAttribute(xmlNodeCell, "CellID"), .BackColor
                                .CellNote = GetAttribute(GetMessageCellById("0080"), "Msg")
                                .BackColor = 12713215 ' vbRed
                            End If
                        End If
                End Select
             Next
   Set xmlNodeCell = Nothing
  End With
       
End Sub

Private Function GetLastfileName() As String
    Dim lngIndex As Long
    Dim fso As New FileSystemObject
    Dim fle As file
    Dim strDataFileName As String
    Dim strFileName As String
    Dim arrStrFileNames() As String
    Dim dNow As Date, dPrevious As Date, dKHBS As Date
    
     If GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") <> "1" Then
                 strDataFileName = "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year
        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then
                 strDataFileName = "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year
        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") <> "1" Then
                 'Data file contain Day from and to.
                 strDataFileName = "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" _
                 & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "")
        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then
                 'Data file contain Day.
                 strDataFileName = "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" _
                 & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year
        Else
                 'Data file not contain Day from and to.
                 strDataFileName = "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" _
                 & TAX_Utilities_v2.Year
             '*********************************
        End If
    
    
    dPrevious = DateSerial(2007, 1, 1)
    
    dKHBS = DateSerial(CInt(Mid$(TAX_Utilities_v2.DateKHBS, 5, 4)), CInt(Mid$(TAX_Utilities_v2.DateKHBS, 3, 2)), CInt(Mid$(TAX_Utilities_v2.DateKHBS, 1, 2)))
    
    For Each fle In fso.GetFolder(GetAbsolutePath(TAX_Utilities_v2.DataFolder)).Files
       If Right$(fle.Name, 4) = ".xml" Then
            If Mid$(fle.Name, 1, Len(fle.Name) - 13) = strDataFileName Then
                strFileName = Mid$(fle.Name, Len(strDataFileName) + 2, 8)
                dNow = DateSerial(CInt(Mid$(strFileName, 5, 4)), CInt(Mid$(strFileName, 3, 2)), CInt(Mid$(strFileName, 1, 2)))
                If dNow > dPrevious And dNow <= dKHBS Then
                    dPrevious = dNow
                    GetLastfileName = strFileName
                End If
            End If
       End If
    Next
    
     If GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") <> "1" Then
                 GetLastfileName = TAX_Utilities_v2.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & "_" & GetLastfileName & ".xml"
        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "ThreeMonth") = "1" Then
                 GetLastfileName = TAX_Utilities_v2.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_0" & TAX_Utilities_v2.ThreeMonths & TAX_Utilities_v2.Year & "_" & GetLastfileName & ".xml"
        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") <> "1" Then
                 'Data file contain Day from and to.
                 GetLastfileName = TAX_Utilities_v2.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" _
                 & TAX_Utilities_v2.Year & "_" & Replace(TAX_Utilities_v2.FirstDay, "/", "") & "_" & Replace(TAX_Utilities_v2.LastDay, "/", "") & "_" & GetLastfileName & ".xml"
        ElseIf GetAttribute(TAX_Utilities_v2.NodeMenu, "Day") = "1" And GetAttribute(TAX_Utilities_v2.NodeMenu, "Month") = "1" Then
                 'Data file contain Day.
                 GetLastfileName = TAX_Utilities_v2.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" _
                 & TAX_Utilities_v2.Day & TAX_Utilities_v2.month & TAX_Utilities_v2.Year & "_" & GetLastfileName & ".xml"
        Else
                 'Data file not contain Day from and to.
                 GetLastfileName = TAX_Utilities_v2.DataFolder & "KHBS_" & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" _
                 & TAX_Utilities_v2.Year & "_" & GetLastfileName & ".xml"
             '*********************************
        End If
    
End Function

Private Sub ResetKHBSData(fpSp As fpSpread, blnExitsData As Boolean)
    On Error GoTo ErrorHandle
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    Dim IsUpdate As Boolean
    Dim xmlNodeC As MSXML.IXMLDOMNode
    Dim xmlNodeH As MSXML.IXMLDOMNode
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim strCellID() As String
    Dim strCellID1 As String
    Dim strValue As String
    
    
    fpSp.ReDraw = False
     For Each xmlNodeReset In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
                fpSp.sheet = 1
                ParserCellID fpSp, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
                fpSp.Col = lCol
                fpSp.Row = lRow
                Select Case fpSp.CellType
'                    Case CellTypeCheckBox
'                        fpSp.Text = vbNullString
'                        SetAttribute xmlNodeReset, "Value", vbNullString
'                    Case CellTypeComboBox
'                        fpSp.Text = vbNullString
'                        SetAttribute xmlNodeReset, "Value", vbNullString
                    Case CellTypeNumber
                        fpSp.value = 0
                        SetAttribute xmlNodeReset, "Value", 0
                    Case Else
''                        fpSp.value = vbNullString
''                        SetAttribute xmlNodeReset, "Value", vbNullString
                End Select
      Next
      If blnExitsData Then
           For Each xmlNodeCells In TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")
                   strCellID = Split(GetAttribute(xmlNodeCells, "CellID"), "_")
                    If strCellID(0) = "BC" Then
                            strCellID1 = Trim(Mid(GetAttribute(xmlNodeCells, "Value"), 100, 20))
                            If strCellID1 <> "" Then
                                    Set xmlNodeC = TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).nodeFromID("BG_" & strCellID(1))
                                    Set xmlNodeH = TAX_Utilities_v2.Data(TAX_Utilities_v2.NodeValidity.childNodes.length - 1).nodeFromID("BF_" & strCellID(1))
                                    strValue = CStr(CDbl(GetAttribute(xmlNodeC, "Value")) - CDbl(GetAttribute(xmlNodeH, "Value")))
                                    ParserCellID fpSp, strCellID1, lCol, lRow
                                    fpSp.sheet = 1
                                    fpSp.Col = lCol
                                    fpSp.Row = lRow
                                    fpSp.value = strValue
                                    
                            End If
                     End If
             
            Next
            Set xmlNodeC = Nothing
            Set xmlNodeH = Nothing
            Set xmlNodeCells = Nothing
            
      End If
      
      
    fpSp.ReDraw = True
    Set xmlNodeReset = Nothing
    
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "ResetKHBSData", Err.Number, Err.Description
End Sub


Private Sub FormatTextPercent(fps As fpSpread, ByVal intSheet As Integer, ByVal lCol As Long, ByVal lRow As Long)
    fps.sheet = intSheet
    fps.Row = lRow
    fps.Col = lCol
    fps.CellType = CellTypeNumber
    ' Set the characters to right
    fps.TypeHAlign = TypeHAlignRight
    fps.TypeVAlign = TypeHAlignCenter
    fps.TypeEditCharSet = TypeEditCharSetNumeric
    fps.TypeNumberMin = 0
    fps.TypeNumberMax = 100
    fps.TypeNumberDecimal = ","
    fps.TypeNumberDecPlaces = 3
    fps.TypePicDefaultText = "...,..."
    fps.TypePicMask = "999,999"
    
End Sub

Private Sub PrintLabelKHBS(idTK As String, fps As fpSpread, ByVal intSheet As Integer)
    With fps
    .sheet = 1
        Select Case idTK
            Case "01"
                .Col = .ColLetterToNumber("CM")
                .Row = 1
            Case "02"
                .Col = .ColLetterToNumber("CI")
                .Row = 1
            Case "04"
                .Col = .ColLetterToNumber("CI")
                .Row = 1
            Case "07"
                .Col = .ColLetterToNumber("CI")
                .Row = 1
            Case "11"
                .Col = .ColLetterToNumber("CI")
                .Row = 1
            Case "12"
                .Col = .ColLetterToNumber("CI")
                .Row = 1
            Case "03"
                .Col = .ColLetterToNumber("AB")
                .Row = 1
            Case "06"
                .Col = .ColLetterToNumber("AI")
                .Row = 1
            Case "09"
                .Col = .ColLetterToNumber("AI")
                .Row = 1
            Case "08"
                .Col = .ColLetterToNumber("AI")
                .Row = 1
            Case "05"
                .Col = .ColLetterToNumber("AI")
                .Row = 1
        End Select
                .BorderStyle = BorderStyleFixedSingle
                .TypeHAlign = TypeHAlignRight
                .TypeVAlign = TypeVAlignTop
                .FontSize = 10
                .FontBold = True
                .Text = GetAttribute(GetMessageCellById("0115"), "Msg")
    End With
End Sub

' Ham get ve thong tin cau truc cua to khai
Public Function getTemplateTk(ByVal strId As String) As String()
    Dim strResult() As String
    Dim tmp As String
    Select Case strId

            ' GTGT
            ' TT28 - 21112011
            ' 01_GTGT / TT156
        Case "01"
            ReDim strResult(3)
            strResult(0) = "D_7~Dynamic_0"
            strResult(1) = "I_23~L_24~J_27~L_27~L_28~J_30~J_31~L_31~J_32~J_33~L_33~J_34~L_34~J_35~J_36~L_36~L_37~L_39~L_40~L_41~L_43~L_44~L_45~L_46~L_47~L_48~Dynamic_0"
            strResult(2) = "D_50~D_51~K_50~K_51~K_53~L_53~O_53~L_14~B_18~D_20~P_20~N_52~Dynamic_0"

            ' 02_GTGT / TT156
        Case "02"
            ReDim strResult(3)
            strResult(0) = "AH_24~Dynamic_0"
            strResult(1) = "CT_40~CT_41~BW_43~CT_43~BW_45~CT_45~BW_46~CT_46~CT_47~CT_48~CT_49~CT_51~CT_52~CT_53~CT_54~Dynamic_0"
            strResult(2) = "V_57~CB_57~V_59~CB_59~C_61~F_61~I_61~K_61~Dynamic_0"

            ' 03_GTGT / TT156
        Case "04"
            ReDim strResult(3)
            strResult(0) = "F_6~Dynamic_0"
            strResult(1) = "Q_36~Q_37~Q_38~Q_39~Q_40~Q_41~Q_42~Dynamic_0"
            strResult(2) = "E_59~O_59~E_61~O_61~C_67~F_67~I_67~L_67~Dynamic_0"

'        Case "95"
'            ReDim strResult(3)
'            strResult(0) = "F_6~Dynamic_0"
'            strResult(1) = "Q_36~Q_37~Q_38~Q_39~Q_40~Q_41~Q_42~Dynamic_0"
'            strResult(2) = "E_59~O_59~E_61~O_61~C_67~F_67~I_67~L_67~Dynamic_0"

            ' 04_GTGT /TT156
        Case "71"
            ReDim strResult(3)
            strResult(0) = "H_14~Dynamic_0"
            strResult(1) = "K_43~Q_43~Z_43~Q_44~Z_44~Q_45~Z_45~Q_46~Z_46~Q_47~Z_47~S_50~O_52~Dynamic_0"
            strResult(2) = "H_63~T_63~H_65~T_65~C_37~F_37~I_37~K_37~L_37~Dynamic_0"

            ' 05_GTGT /TT156
        Case "72"
            ReDim strResult(3)
            strResult(0) = "I_14~Dynamic_0"
            strResult(1) = "K_43~R_43~K_45~R_45~J_48~Dynamic_0"
            strResult(2) = "H_55~R_55~H_57~R_57~C_54~F_54~I_54~J_54~M_54~Dynamic_0"

            ' TNDN
            ' 01A_TNDN\ TT156
        Case "11"
            ReDim strResult(4)
            strResult(0) = "F_19~Dynamic_0"
            strResult(1) = "K_22~K_23~K_24~K_25~K_26~K_27~K_28~K_29~K_30~K_31~K_32~K_33~F_34~K_34~K_35~K_36~K_37~K_38~K_39~K_40~K_41~F_43~H_45~P_45~H_47~H_49~H_51~Dynamic_0"
            strResult(2) = "D_11~D_12~Dynamic_0"
            strResult(3) = "E_54~E_56~J_54~J_56~E_13~G_13~E_57~L_15~Dynamic_0"

            ' 01B_TNDN \ TT156
        Case "12"
            ReDim strResult(3)
            strResult(0) = "E_12~Dynamic_0"
            strResult(1) = "F_6~F_7~K_36~K_37~K_38~K_39~K_40~K_41~K_42~K_43~K_44~K_45~K_46~H_47~K_47~K_48~K_49~K_50~K_51~K_52~K_53~K_54~F_56~H_58~P_58~H_60~H_62~H_64~Dynamic_0"
            strResult(2) = "J_67~J_69~E_67~E_69~E_14~G_14~L_14~Dynamic_0"

            ' 02_TNDN \ TT156
        Case "73"
            ReDim strResult(4)
            strResult(0) = "J_47~Dynamic_0"
            strResult(1) = "AW_61~AW_62~AW_63~AW_64~AW_65~AW_66~AW_67~AW_68~AW_69~AW_72~AW_73~Dynamic_0"
            strResult(2) = "P_54~Q_54~V_34~AL_36~I_40~X_42~AC_42~AH_44~Dynamic_0"
            strResult(3) = "P_89~P_91~AP_89~AP_91~M_54~O_54~AI_54~C_82~Z_16~Dynamic_0"
            ' 04/TNDN \ TT151
        Case "55"
            ReDim strResult(4)
            strResult(0) = "K_24~Dynamic_0"
            strResult(1) = "I_37~U_37~Y_37~AH_37~AT_37~AX_37~BG_37~BS_37~BW_37~CF_37~Dynamic_1"
            strResult(2) = "I_39~Y_39~AH_39~AX_39~BG_39~BW_39~CF_39~Y_28~AX_28~BW_28~CF_28~Dynamic_0"
            strResult(3) = "Q_51~BG_51~Q_53~BG_53~L_28~M_28~N_28~F_28~O_28~Dynamic_0"
            ' 06/TNDN \ TT151
        Case "56"
            ReDim strResult(4)
            strResult(0) = "J_47~Dynamic_0"
            strResult(1) = "V_34~AL_36~I_40~X_42~AC_42~AH_44~Dynamic_0"
            strResult(2) = "AW_61~AW_62~AW_64~AW_65~AW_66~AW_67~AW_68~AW_69~AW_70~AW_71~AW_72~AW_73~AW_74~Dynamic_0"
            strResult(3) = "P_90~P_92~AP_90~AP_92~M_54~N_54~O_54~AI_54~Dynamic_0"
            ' 03_TNDN \TT156
        Case "03"
            ReDim strResult(6)
            strResult(0) = "F_17~Dynamic_0"
            strResult(1) = "D_8~D_10~D_12~G_14~P_14~G_16~Dynamic_0"
            strResult(2) = "O_31~O_33~O_34~O_35~O_36~O_37~O_38~O_39~O_40~O_41~O_42~O_43~O_44~O_45~O_46~O_48~O_49~O_50~O_51~O_52~O_53~O_54~O_55~O_56~O_57~O_58~O_59~O_60~O_61~O_62~O_63~O_64~O_65~O_66~O_67~O_68~O_69~O_70~O_71~O_72~O_73~O_74~O_75~O_76~O_77~O_78~O_79~O_80~Dynamic_0"
            strResult(3) = "F_83~I_85~P_85~I_87~I_89~I_91~E_95~J_95~N_95~E_97~Dynamic_0"
            strResult(4) = "C_103~Dynamic_1"
            strResult(5) = "F_108~N_108~F_110~N_110~E_22~F_22~I_22~J_22~Dynamic_0"

            ' 05_TNDN
            '        Case "14"
            '            ReDim strResult(3)
            '            strResult(0) = "C_7~I_7~Dynamic_0"
            '            strResult(1) = "C_11~D_11~E_11~F_11~G_11~H_11~I_11~J_11~Dynamic_1"
            '            strResult(2) = "I_15~I_16~Dynamic_0"
            ' TNCN
            ' 01A_TNCN_BH \TT28
        Case "46"
           ReDim strResult(3)
           strResult(0) = "G_7~Dynamic_0"
           strResult(1) = "U_40~U_41~U_42~Dynamic_0"
           strResult(2) = "R_44~G_44~G_46~R_46~C_47~F_47~I_47~Dynamic_0"
        ' 01B_TNCN_BH \TT28
        Case "47"
           ReDim strResult(3)
           strResult(0) = "G_7~Dynamic_0"
           strResult(1) = "U_40~U_41~U_42~Dynamic_0"
           strResult(2) = "R_44~G_44~G_46~R_46~C_47~F_47~I_47~Dynamic_0"
        ' 01A_TNCN_SX \TT28
        Case "48"
           ReDim strResult(3)
           strResult(0) = "G_7~Dynamic_0"
           strResult(1) = "U_40~U_41~U_42~Dynamic_0"
           strResult(2) = "R_44~G_44~G_46~R_46~C_47~F_47~I_47~Dynamic_0"
        ' 01B_TNCN_XS \TT28
        Case "49"
            ReDim strResult(3)
            strResult(0) = "G_7~Dynamic_0"
            strResult(1) = "U_40~U_41~U_42~Dynamic_0"
            strResult(2) = "R_44~G_44~G_46~R_46~C_47~F_47~I_47~Dynamic_0"
        ' 02A_TNCN10 \TT28
        Case "15"
            ReDim strResult(3)
            strResult(0) = "I_8~Dynamic_0"
            strResult(1) = "U_38~U_39~U_41~U_42~U_44~U_45~U_46~U_48~U_49~U_50~U_52~U_53~U_54~Dynamic_0"
            strResult(2) = "R_58~R_60~H_58~H_60~C_61~F_61~I_61~Dynamic_0"
        ' 02B_TNCN10 \ TT28
        Case "16"
            ReDim strResult(3)
            strResult(0) = "I_8~Dynamic_0"
            strResult(1) = "U_38~U_39~U_41~U_42~U_44~U_45~U_46~U_48~U_49~U_50~U_52~U_53~U_54~Dynamic_0"
            strResult(2) = "R_58~R_60~H_58~H_60~C_61~F_61~I_61~Dynamic_0"
        ' 03A_TNCN10 \TT28
        Case "50"
            ReDim strResult(4)
            strResult(0) = "H_19~Dynamic_0"
            strResult(1) = "J_7~J_9~F_11~F_13~S_13~F_15~K_15~S_15~Dynamic_0"
            strResult(2) = "U_40~U_41~U_43~U_44~U_46~U_47~U_49~U_50~U_52~U_53~U_55~U_56~Dynamic_0"
            strResult(3) = "R_58~R_60~H_58~H_60~C_64~F_64~I_64~Dynamic_0"
        ' 03B_TNCN10 \TT28
        Case "51"
            ReDim strResult(4)
            strResult(0) = "H_19~Dynamic_0"
            strResult(1) = "J_7~J_9~F_11~F_13~S_13~F_15~K_15~S_15~Dynamic_0"
            strResult(2) = "U_40~U_41~U_43~U_44~U_46~U_47~U_49~U_50~U_52~U_53~U_55~U_56~Dynamic_0"
            strResult(3) = "R_58~R_60~H_58~H_60~C_64~F_64~I_64~Dynamic_0"
'        ' 04A_TNCN
'        Case "39"
'            ReDim strResult(2)
'            strResult(0) = "U_39~U_40~U_41~U_42~U_43~U_44~U_45~U_46~U_47~Dynamic_0"
'            strResult(1) = "R_49~R_51~C_52~F_52~I_52~Dynamic_0"
'        ' 04B_TNCN
'        Case "40"
'            ReDim strResult(2)
'            strResult(0) = "U_39~U_40~U_41~U_42~U_43~U_44~U_45~U_46~U_47~Dynamic_0"
'            strResult(1) = "R_49~R_51~C_52~F_52~I_52~Dynamic_0"
        ' 07_TNCN  \TT28
        Case "36"
            ReDim strResult(3)
            strResult(0) = "I_8~Dynamic_0"
            'strResult(1) = "V_41~R_43~R_44~R_45~R_46~R_47~R_48~R_49~R_50~R_51~R_52~R_53~R_55~R_56~R_57~R_59~R_60~R_61~Dynamic_0"
            strResult(1) = "V_41~R_43~R_44~R_45~R_46~R_47~R_48~R_49~R_50~R_51~R_52~R_53~R_54~R_55~R_57~R_58~R_59~Dynamic_0"
            strResult(2) = "R_66~R_68~H_66~H_68~C_70~F_70~I_70~L_70~Dynamic_0"
        ' 01/KK-BHDC \TT156
        Case "25"
            ReDim strResult(3)
            strResult(0) = "H_8~Dynamic_0"
            strResult(1) = "U_40~U_41~U_42~U_44~U_45~U_47~U_48~U_50~U_51~U_52~U_53~Dynamic_0"
            strResult(2) = "R_58~R_60~H_58~H_60~C_64~F_64~I_64~L_64~Dynamic_0"
            
        ' QT TNCN
        ' 05_TNCN \TT28
        Case "17"
            ReDim strResult(4)
            strResult(0) = "D_20~Dynamic_0"
            strResult(1) = "D_4~D_6~Dynamic_0"
            strResult(2) = "I_36~I_37~I_38~I_39~I_40~I_41~I_42~I_43~I_44~I_45~I_46~I_47~I_48~I_49~I_50~I_51~I_52~I_53~I_54~I_55~I_56~I_57~I_61~I_62~I_63~I_64~I_65~Dynamic_0"
            strResult(3) = "D_69~D_71~M_69~M_71~C_67~F_67~I_67~M_67~N_67~Dynamic_0"
        ' 02_TNCN_BH  \TT28
        Case "42"
            ReDim strResult(3)
            strResult(0) = "D_22~Dynamic_0"
            strResult(1) = "I_39~I_40~I_41~I_42~I_43~Dynamic_0"
            strResult(2) = "M_45~M_47~D_45~D_47~C_48~I_48~Dynamic_0"
        ' 02_TNCN_XS  \TT28
        Case "43"
            ReDim strResult(3)
            strResult(0) = "D_22~Dynamic_0"
            strResult(1) = "I_40~I_41~I_42~I_43~I_44~Dynamic_0"
            strResult(2) = "M_46~M_48~D_46~D_48~C_51~I_51~K_51~L_51~Dynamic_0"
            ' 06_TNCN  \TT28
        Case "59"
            ReDim strResult(3)
            strResult(0) = "D_22~Dynamic_0"
            strResult(1) = "I_41~I_42~I_44~I_45~I_47~I_48~I_50~I_51~I_53~I_54~I_55~I_57~I_58~Dynamic_0"
            strResult(2) = "M_60~M_62~D_60~D_62~C_64~I_64~K_64~L_64~Dynamic_0"
            
            ' 08_TNCN  \TT28
        Case "74"
            ReDim strResult(3)
            strResult(0) = "G_17~Dynamic_0"
            'strResult(1) = "R_18~X_18~K_4~P_4~R_48~R_49~R_50~R_51~R_52~R_53~R_54~R_55~R_56~R_57~R_58~R_59~R_60~Dynamic_0"
            strResult(1) = "R_48~R_49~R_50~R_51~R_52~R_53~R_54~R_55~R_56~R_57~R_58~R_59~R_60~R_61~Dynamic_0"
            strResult(2) = "R_64~R_66~H_64~H_66~C_68~I_68~Dynamic_0"
            ' 08A_TNCN  \TT28
        Case "75"
            ReDim strResult(4)
            strResult(0) = "G_17~Dynamic_0"
            strResult(1) = "R_41~R_42~R_43~R_44~Dynamic_0"
            strResult(2) = "C_51~H_51~L_51~N_51~P_51~R_51~T_51~V_51~X_51~Z_51~Dynamic_1"
            strResult(3) = "R_55~R_57~G_55~G_57~C_59~I_59~Dynamic_0"
            ' 08B_TNCN  \TT28
        Case "76"
            ReDim strResult(4)
            strResult(0) = "G_11~Dynamic_0"
            strResult(1) = "I_40~I_41~I_42~I_43~I_44~I_45~I_46~I_47~I_48~I_49~I_50~I_51~Dynamic_0"
            strResult(2) = "C_57~G_57~K_57~L_57~N_57~P_57~R_57~T_57~W_57~Z_57~Dynamic_1"
            strResult(3) = "M_61~M_63~D_61~D_63~C_38~I_38~K_38~L_38~Dynamic_0"
            
            ' 09_TNCN \TT28
        Case "41"
            ReDim strResult(4)
            strResult(0) = "F_24~Dynamic_0"
            strResult(1) = "F_20~O_20~Dynamic_0"
            strResult(2) = "K_42~K_43~K_44~K_45~K_46~K_47~K_48~K_49~K_50~K_51~K_52~K_53~K_54~K_55~K_56~K_57~K_58~K_59~K_60~K_61~K_62~K_63~K_64~K_65~K_66~K_67~Dynamic_0"
            strResult(3) = "O_69~O_71~F_69~F_71~N_4 ~P_4 ~E_76~K_76~Dynamic_0"
'
'        ' TAIN
        ' 01_TAIN  \TT28
        Case "06"
            ReDim strResult(5)
            strResult(0) = "N_11~Dynamic_0"
            strResult(1) = "D_42~E_42~F_42~G_42~O_42~P_42~S_42~V_42~AA_42~AD_42~AG_42~AJ_42~AN_42~Dynamic_1"
            strResult(2) = "D_46~E_46~F_46~G_46~O_46~P_46~S_46~V_46~AA_46~AD_46~AG_46~AJ_46~AN_46~Dynamic_1"
            strResult(3) = "D_50~E_50~F_50~G_50~O_50~P_50~S_50~V_50~AA_50~AD_50~AG_50~AJ_50~AN_50~Dynamic_1"
            strResult(4) = "W_54~W_56~AJ_54~AJ_56~I_15~J_15~K_15~L_15~M_15~AG_15~AJ_15~AN_15~Dynamic_0"
        ' 02_TAIN \TT28
        Case "77"
            ReDim strResult(4)
            strResult(0) = "K_15~Dynamic_0"
            strResult(1) = "D_40~E_40~F_40~G_40~O_40~P_40~S_40~V_40~AA_40~AD_40~AH_40~AK_40~AO_40~AP_40~AQ_40~Dynamic_1"
            strResult(2) = "D_44~E_44~F_44~G_44~O_44~P_44~S_44~V_44~AA_44~AD_44~AH_44~AK_44~AO_44~AP_44~AQ_44~Dynamic_1"
            strResult(3) = "AG_49~AG_51~I_49~I_51~I_17~M_17~AH_17~AK_17~AO_17~Q_17~R_17~Dynamic_0"
'        ' 03_TAIN
'        Case "08"
'            ReDim strResult(3)
'            strResult(0) = "N_6~AN_8~Dynamic_0"
'            strResult(1) = "D_13~E_13~F_13~G_13~O_13~P_13~S_13~V_13~AA_13~AD_13~AG_13~AJ_13~AN_13~Dynamic_1"
'            strResult(2) = "AJ_17~AJ_19~Dynamic_0"
'
'
         ' TTDB
         ' 01_TTDB
         'dntai  24/05/2011
         'sua theo template cua TT28
        Case "05"
            ReDim strResult(11)
            strResult(0) = "AN_6~AG_6~M_7~Dynamic_0"
            strResult(1) = "N_34~V_36~AA_36~AG_36~AJ_36~AN_36~Dynamic_0"
            strResult(2) = "D_38~E_38~F_38~G_38~O_38~P_38~Q_38~R_38~S_38~V_38~AA_38~AD_38~AG_38~AJ_38~AN_38~Dynamic_1"
            strResult(3) = "V_40~AA_40~AG_40~AJ_40~AN_40~Dynamic_0"
            strResult(4) = "D_42~E_42~F_42~G_42~O_42~P_42~Q_42~R_42~S_42~V_42~AA_42~AD_42~AG_42~AJ_42~AN_42~Dynamic_1"
            strResult(5) = "V_44~Dynamic_0"
            strResult(6) = "D_47~E_47~F_47~G_47~O_47~P_47~Q_47~R_47~S_47~V_47~Dynamic_1"
            strResult(7) = "D_51~E_51~F_51~G_51~O_51~P_51~Q_51~R_51~S_51~V_51~Dynamic_1"
            strResult(8) = "D_55~E_55~F_55~G_55~O_55~P_55~Q_55~R_55~S_55~V_55~Dynamic_1"
            strResult(9) = "V_57~AA_57~AG_57~AJ_57~AN_57~Dynamic_0"
            strResult(10) = "AI_59~U_59~U_61~AI_61~Y_12~AA_12~AC_12~AI_15~AN_15~S_10~T_10~AN_10~L_14~Dynamic_0"
           '
           ' NTNN
           ' 01_NTNN
        Case "70"
            ReDim strResult(5)
            strResult(0) = "Y_21~Dynamic_0"
            strResult(1) = "C_55~L_55~R_55~X_55~AD_5~AI_5~AM_5~AQ_5~AU_5~AY_5~BC_5~BG_5~BM_5~BQ_5~Dynamic_1"
            strResult(2) = "AI_5~AU_4~AU_5~AY_5~BG_5~BM_5~BM_4~BQ_5~BQ_4~Dynamic_0"
            strResult(3) = "AT_5~BC_5~S_33~O_32~Dynamic_0"
            strResult(4) = "O_67~O_69~AX_6~AG_3~C_31~F_31~I_31~BM_3~Dynamic_0"
            
           ' 02_NTNN
        Case "80"
            ReDim strResult(4)
            strResult(0) = "J_18~Dynamic_0"
            strResult(1) = "N_22~AH_22~Dynamic_0"
            strResult(2) = "W_30~AF_30~AO_30~AS_30~W_31~AF_31~AO_31~AS_31~W_32~AF_32~AO_32~AS_32~W_33~AF_33~AO_33~AS_33~W_34~AF_34~AO_34~AS_34~W_35~AF_35~AO_35~AS_35~W_36~AF_36~AO_36~AS_36~W_37~AF_37~AO_37~AS_37~W_38~AF_38~AO_38~AS_38~W_39~AF_39~AO_39~AS_39~W_40~AF_40~AO_40~AS_40~W_41~AF_41~AO_41~AS_41~W_42~AF_42~AO_42~AS_42~Dynamic_0"
            strResult(3) = "P_47~AU_47~P_49~AU_49~M_23~N_23~O_23~Dynamic_0"
            
           ' 03_NTNN
        Case "81"
            ReDim strResult(5)
            strResult(0) = "K_22~Dynamic_0"
            strResult(1) = "C_34~CQ_34~AD_34~AL_34~AU_34~BE_34~BK_34~BU_34~BX_34~CG_34~Dynamic_1"
            strResult(2) = "AU_36~BK_36~BX_36~CG_36~Dynamic_0"
            strResult(3) = "P_26~Q_26~AD_38~CG_26~Dynamic_0"
            strResult(4) = "Q_48~BG_48~Q_50~BG_50~M_26~N_26~O_26~S_26~Dynamic_0"
            ' 04_NTNN
        Case "82"
            ReDim strResult(4)
            strResult(0) = "J_19~Dynamic_0"
            strResult(1) = "N_21~AH_21~Dynamic_0"
            strResult(2) = "T_30~AB_30~AJ_30~AT_30~T_31~AB_31~AJ_31~AT_31~T_32~AB_32~AJ_32~AT_32~T_33~AB_33~AJ_33~AT_33~T_34~AB_34~AJ_34~AT_34~T_35~AB_35~AJ_35~AT_35~T_36~AB_36~AJ_36~AT_36~Dynamic_0"
            strResult(3) = "M_44~AL_44~M_46~AL_46~M_23~N_23~O_23~Dynamic_0"
            '02_PHLP
        Case "88"
            ReDim strResult(4)
            strResult(0) = "D_17~Dynamic_0"
            strResult(1) = "C_49~E_49~H_49~L_49~N_49~Q_49~T_49~X_49~AB_49~AC_49~Dynamic_1"
            strResult(3) = "Q_37~T_37~X_37~Dynamic_0"
            strResult(4) = "H_60~R_60~H_62~R_62~C_35~F_35~I_35~Dynamic_1"
            ' 01_BVMT
        Case "86"
            ReDim strResult(4)
            strResult(0) = "I_17~Q_13~H_15~Dynamic_0"
            strResult(1) = "C_45~J_45~L_45~Q_45~U_45~AC_45~Dynamic_1"
            strResult(2) = "C_49~J_49~L_49~Q_49~U_49~AC_49~Dynamic_1"
            strResult(3) = "H_60~R_60~H_62~R_62~C_35~F_35~I_35~T_37~AB_37~Dynamic_1"
            
            ' 02_BVMT
        Case "87"
            ReDim strResult(4)
            strResult(0) = "H_14~Dynamic_0"
            strResult(1) = "C_45~H_45~J_45~N_45~Q_45~V_45~Y_45~AC_45~Dynamic_1"
            strResult(2) = "C_49~H_49~J_49~N_49~Q_49~V_49~Y_49~AC_49~Dynamic_1"
            strResult(3) = "H_60~R_60~H_62~R_62~C_35~F_35~I_35~K_35~L_35~Q_37~V_37~Y_37~Dynamic_0"
            
            ' 02/TAIN-DK
        Case "89"
           ReDim strResult(4)
            strResult(0) = "P_51~Dynamic_0"
            strResult(1) = "O_30~P_30~Q_30~P_47~P_49~AA_49~Dynamic_0"
            strResult(2) = "AV_55~AV_57~AV_58~AV_59~AV_60~AV_61~AV_62~AV_63~AV_64~AV_65~AV_67~AV_68~AV_69~AV_70~AV_71~AV_72~AV_73~AV_74~AV_75~AV_76~Dynamic_0"
            strResult(3) = "U_92~U_94~AR_92~AR_94~R_30~W_30~O_31~AD_30~AE_30~Dynamic_0"
            '01/KK-TTS
        Case "23"
            ReDim strResult(4)
            strResult(0) = "F_11~Dynamic_0"
            strResult(1) = "P_3~T_3~H_13~R_13~J_16~F_17~L_18~L_19~F_20~F_21~H_22~I_23~I_24~I_25~I_26~Dynamic_0"
            strResult(2) = "C_40~E_40~I_40~L_40~N_40~O_40~P_40~R_40~T_40~U_40~V_40~Dynamic_1"
            strResult(3) = "E_59~O_59~E_61~O_61~C_29~F_29~I_29~G_29~O_29~R_29~Dynamic_0"
        Case Else
            ReDim strResult(1)
            tmp = "null"
            strResult(0) = tmp
    End Select
    getTemplateTk = strResult
End Function

' Sau nay se dem cac chi tieu ma hoa trong ma vach
Public Function GetElementsNoData(xmlCellsNode As MSXML.IXMLDOMNode) As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim lCntElementsNo As Long
    
    For Each xmlCellNode In xmlCellsNode.childNodes
        'If GetAttribute(xmlCellNode, "Encode") = "1" Then
            lCntElementsNo = lCntElementsNo + 1
        'End If
    Next
    GetElementsNoData = lCntElementsNo
End Function
Public Function GetQuyNamTaiChinh(q As Integer, Y As Integer, dNgayTaiChinh As Integer, dThangTaiChinh As Integer, dType As Integer) As Integer
   ' q Quy ke khai
   ' y nam ke khai
   ' dNgayTaiChinh ngay tai chinh lay tren man hinh HTKK
   ' dThangTaiChinh thang tai chinh tren phan thong tin chung HTKK
   ' dType: 0 tra ve quy, 1 tra ve nam
    Dim intYear As Integer, intDay As Integer, intMonth As Integer, result As Integer
   
    intDay = dNgayTaiChinh
    intMonth = (q - 1) * 3 + dThangTaiChinh
    intYear = Y
    If intMonth > 12 Then
        intMonth = intMonth - 12
        intYear = Y + 1
    End If
    If dType = 0 Then
       result = DatePart("Q", DateSerial(intYear, intMonth, intDay))
    Else
       result = Year(DateSerial(intYear, intMonth, intDay))
    End If
    GetQuyNamTaiChinh = result
End Function
'format a day/month/year string as dd/mm/yyyy
'if not able to format, out: vbnullstring
'if able, out a dd/mm string
Public Function Format_ddmmyyyy(str As String) As String
    Dim DD As String, MM As String, YYYY As String, dDate As Date
    
  If str <> "" Or Len(str) > 0 Then
    On Error GoTo e
    DD = Left(str, InStr(str, "/") - 1)
    MM = Mid(str, 4, 2)
    YYYY = Right("0000" & str, 4)
 
    
        If Val(DD) >= 1 And Val(DD) <= 31 Then
            DD = format(DD, "0#")
        Else
            GoTo e
        End If
        
        If Val(MM) >= 1 And Val(MM) <= 12 Then
            MM = format(MM, "0#")
        Else
            GoTo e
        End If
        
        If Val(YYYY) >= 0 And Val(YYYY) <= 9999 Then
            
            If Val(YYYY) >= 0 And Val(YYYY) <= 999 Then YYYY = CStr(2000 + Val(YYYY))
            If Val(YYYY) < 1900 Then GoTo e
            YYYY = format(YYYY, "####")
        Else
            GoTo e
        End If
        
        dDate = format(MM & "/" & DD & "/" & YYYY, "mm/dd/yyyy")
        'Format_ddmm = dd & "/" & mm
        Format_ddmmyyyy = DD & "/" & MM & "/" & YYYY
    End If
    Exit Function
e:
    DisplayMessage "0071", msOKOnly, miCriticalError
    Format_ddmmyyyy = ""
End Function

Public Function GetHanNopTk() As String
    Dim hannop As String
    Dim dNgayCuoiKy As Date
    Dim strarrdate() As String ' su dung cho to khai 02/NTNN va 04/NTNN
    If TAX_Utilities_v2.month <> "" Then
        ' To khai 01/GTGT gia han thang 4,5,6 nam 2012 -> tinh lai han nop
        If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "01" Then
             If strQuy = "TK_THANG" Then
                If (TAX_Utilities_v2.month = 4 Or TAX_Utilities_v2.month = 5 Or TAX_Utilities_v2.month = 6) And TAX_Utilities_v2.Year = 2012 And TAX_Utilities_v2.CheckToKhaiGH = True Then
                    If TAX_Utilities_v2.month = 4 Then
                        hannop = "20/" & "11" & "/" & TAX_Utilities_v2.Year
                    ElseIf TAX_Utilities_v2.month = 5 Then
                        hannop = "20/" & "12" & "/" & TAX_Utilities_v2.Year
                    ElseIf TAX_Utilities_v2.month = 6 Then
                        hannop = "21/" & "01" & "/" & TAX_Utilities_v2.Year + 1
                    End If
                Else
                    ' cac ky ke khai khac van tinh han nop binh thuong
                    If TAX_Utilities_v2.month = 12 Then
                        hannop = "20/" & "01" & "/" & TAX_Utilities_v2.Year + 1
'                    ElseIf TAX_Utilities_v2.month = 4 Then
'                        hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                    Else
                        hannop = "20/" & Right("00" & TAX_Utilities_v2.month + 1, 2) & "/" & TAX_Utilities_v2.Year
                    End If
                End If
            ElseIf strQuy = "TK_QUY" Then
                If Val(TAX_Utilities_v2.ThreeMonths) = 4 Then
                   hannop = "31/" & "01" & "/" & TAX_Utilities_v2.Year + 1
                ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 3 Then
                    hannop = "31/" & "10" & "/" & TAX_Utilities_v2.Year
                ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 2 Then
                    hannop = "31/" & "07" & "/" & TAX_Utilities_v2.Year
                ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 1 Then
                    hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                End If
            End If
        ElseIf GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "02" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "04" _
                Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "71" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "96" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "94" _
                Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "98" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "99" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "92" Then
            If strQuy = "TK_THANG" Then
                 ' cac to khai thang khac van tinh binh thuong
                If TAX_Utilities_v2.month = 12 Then
                    hannop = "20/" & "01" & "/" & TAX_Utilities_v2.Year + 1
'                ElseIf TAX_Utilities_v2.month = 4 Then
'                    hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                Else
                    hannop = "20/" & Right("00" & TAX_Utilities_v2.month + 1, 2) & "/" & TAX_Utilities_v2.Year
                End If
            ElseIf strQuy = "TK_QUY" Then
                If Val(TAX_Utilities_v2.ThreeMonths) = 4 Then
                   hannop = "31/" & "01" & "/" & TAX_Utilities_v2.Year + 1
                ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 3 Then
                    hannop = "31/" & "10" & "/" & TAX_Utilities_v2.Year
                ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 2 Then
                    hannop = "31/" & "07" & "/" & TAX_Utilities_v2.Year
                ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 1 Then
                    hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                End If
            ElseIf strQuy = "TK_LANPS" Then
                hannop = DateAdd("D", 10, DateSerial(CInt(TAX_Utilities_v2.Year), CInt(TAX_Utilities_v2.month), CInt(TAX_Utilities_v2.Day)))
            ElseIf strQuy = "TK_LANXB" Then
                ' dau khi theo lan xuat ban han 35 ngay
                hannop = DateAdd("D", 35, DateSerial(CInt(TAX_Utilities_v2.Year), CInt(TAX_Utilities_v2.month), CInt(TAX_Utilities_v2.Day)))
            End If
        Else
           If GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "72" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "73" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "56" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "70" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "81" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "06" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "05" Or GetAttribute(TAX_Utilities_v2.NodeValidity.parentNode, "ID") = "90" Then
                If strLoaiTKThang_PS = "TK_LANPS" Then
                    hannop = DateAdd("D", 10, DateSerial(CInt(TAX_Utilities_v2.Year), CInt(TAX_Utilities_v2.month), CInt(TAX_Utilities_v2.Day)))
                Else
                    ' cac to khai thang khac van tinh binh thuong
                    If TAX_Utilities_v2.month = 12 Then
                        hannop = "20/" & "01" & "/" & TAX_Utilities_v2.Year + 1
        '            ElseIf TAX_Utilities_v2.month = 4 Then
        '                hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                    Else
                        hannop = "20/" & Right("00" & TAX_Utilities_v2.month + 1, 2) & "/" & TAX_Utilities_v2.Year
                    End If
                End If
           Else
                ' cac to khai thang khac van tinh binh thuong
                If TAX_Utilities_v2.month = 12 Then
                    hannop = "20/" & "01" & "/" & TAX_Utilities_v2.Year + 1
    '            ElseIf TAX_Utilities_v2.month = 4 Then
    '                hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
                Else
                    hannop = "20/" & Right("00" & TAX_Utilities_v2.month + 1, 2) & "/" & TAX_Utilities_v2.Year
                End If
           End If
       End If
    ElseIf TAX_Utilities_v2.ThreeMonths <> "" Then
        If Val(TAX_Utilities_v2.ThreeMonths) = 4 Then
           hannop = "31/" & "01" & "/" & TAX_Utilities_v2.Year + 1
        ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 3 Then
            hannop = "31/" & "10" & "/" & TAX_Utilities_v2.Year
        ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 2 Then
            hannop = "31/" & "07" & "/" & TAX_Utilities_v2.Year
        ElseIf Val(TAX_Utilities_v2.ThreeMonths) = 1 Then
            hannop = "02/" & "05" & "/" & TAX_Utilities_v2.Year
        End If
'                    dNgayCuoiKy = DateAdd("D", 30, GetNgayCuoiQuy(TAX_Utilities_v2.ThreeMonths, CInt(TAX_Utilities_v2.Year), iNgayTaiChinh, iThangTaiChinh))
'                    hannop = format(dNgayCuoiKy, "dd/mm/yyyy")
    Else
        If GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "80" Or GetAttribute(TAX_Utilities_v2.NodeMenu, "ID") = "82" Then
            strarrdate = Split(TAX_Utilities_v2.LastDay, "/")
            dNgayCuoiKy = DateAdd("D", 45, DateSerial(CInt(strarrdate(2)), CInt(strarrdate(1)), CInt(strarrdate(0))))
            hannop = format(dNgayCuoiKy, "dd/mm/yyyy")
        Else
            dNgayCuoiKy = DateAdd("D", 90, NgayCuoiNamTaiChinh(TAX_Utilities_v2.Year, iThangTaiChinh, iNgayTaiChinh))
            hannop = format(dNgayCuoiKy, "dd/mm/yyyy")
        End If
    End If
    
    'Neu vao ngay thu 7 thi cong them 2 ngay,  ngay CN thi cong them mot ngay
    If Weekday(CDate(hannop)) = 7 Then
        hannop = DateAdd("D", 2, CDate(hannop))
        hannop = format(hannop, "dd/mm/yyyy")
    ElseIf Weekday(CDate(hannop)) = 1 Then
        hannop = DateAdd("D", 1, CDate(hannop))
        hannop = format(hannop, "dd/mm/yyyy")
    End If
    GetHanNopTk = hannop
    Exit Function
End Function

Public Function ToDateString(str As String, mmmmYYdd As Boolean) As String
    Dim strArray() As String
    If mmmmYYdd = True Then
        If Len(str) > 10 Then
            If Len(str) > 20 Then
                ToDateString = str
                Exit Function
            Else
                ' format lai dinh dang kieu datetime cua chuan XML
                strArray = Split(Left(str, 10), "/")
                If UBound(strArray) <> 2 Then
                    ToDateString = str
                    Exit Function
                Else
        
                    If Val(strArray(0)) > 0 And Val(strArray(1)) > 0 And Val(strArray(2)) > 0 Then
                        If Val(strArray(0)) <= 31 And Val(strArray(1)) <= 12 And Val(strArray(2)) < 9999 Then
                            ToDateString = strArray(2) & "-" & strArray(1) & "-" & strArray(0) & Right(str, Len(str) - 10)
                            Exit Function
                        End If
                    End If
                End If
            End If
        Else
            ' format theo dinh dang date cua chuan XML
            strArray = Split(str, "/")
    
            If UBound(strArray) <> 2 Then
                ToDateString = str
                Exit Function
            Else
    
                If Val(strArray(0)) > 0 And Val(strArray(1)) > 0 And Val(strArray(2)) > 0 Then
                    If Val(strArray(0)) <= 31 And Val(strArray(1)) <= 12 And Val(strArray(2)) < 9999 Then
                        ToDateString = strArray(2) & "-" & strArray(1) & "-" & strArray(0)
                        Exit Function
                    End If
                End If
            End If
        End If
    Else
        If Len(str) > 10 Then
            ToDateString = str
            Exit Function
        Else
            strArray = Split(str, "-")
    
            If UBound(strArray) <> 2 Then
                ToDateString = str
                Exit Function
            Else
    
                If Val(strArray(0)) > 0 And Val(strArray(1)) > 0 And Val(strArray(2)) > 0 Then
                    If Val(strArray(2)) <= 31 And Val(strArray(1)) <= 12 And Val(strArray(0)) < 9999 Then
                        ToDateString = strArray(2) & "/" & strArray(1) & "/" & strArray(0)
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    ToDateString = str
    Exit Function
End Function

Public Function ToDate(str As String) As Date
     ToDate = DateSerial(Val(Right$(Replace$(str, "/", ""), 4)), Val(Mid$(Replace$(str, "/", ""), 3, 2)), Val(Left$(Replace$(str, "/", ""), 2)))
End Function

Public Function isLocaleDecimalClient() As Boolean
    Dim LocaleDecimal As String
    LocaleDecimal = Mid$(CStr(11 / 10), 2, 1)

    If InStr(1, LocaleDecimal, ",") > 0 Then
        isLocaleDecimalClient = False
    ElseIf InStr(1, LocaleDecimal, ".") > 0 Then
        isLocaleDecimalClient = True
    End If

End Function

Public Sub ParseCell(cellID As String, lCol As Long, lRow As Long)
    Dim cellArray() As String
    cellArray = Split(cellID, "_")

    If UBound(cellArray) = 1 Then
        If Val(cellArray(1)) > 0 Then
            lRow = Val(cellArray(1))
            lCol = frmInterfaces.fpSpread1.ColLetterToNumber(cellArray(0))
        Else
            lRow = 0
            lCol = 0
        End If

    Else
        lRow = 0
        lCol = 0
    End If

End Sub

' Lay ten CQT theo ma
Public Sub GetCQT(ByVal maCQT As String, Optional ByRef TenCQT As String)
Dim arrDanhsach() As String
Dim strDataFileName As String
Dim xmlDomData As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = "..\InterfaceIni\Catalogue_Tinh_Thanh.xml"
    
       If xmlDomData.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If maCQT = arrDanhsach(1) Then
                            TenCQT = arrDanhsach(3)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub


' Lay ten CQT theo ma
Public Sub GetCQT_01GTGT(ByVal maCQT As String, Optional ByRef TenCQT As String)
Dim arrDanhsach() As String
Dim strDataFileName As String
Dim xmlDomData As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = "..\InterfaceIni\Catalogue_Tinh_Thanh.xml"
    
       If xmlDomData.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If maCQT = arrDanhsach(1) And arrDanhsach(0) = "0" Then
                            TenCQT = arrDanhsach(3)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub



Public Function getMST() As String
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim mstDN As Variant
    Dim i As Integer
    Dim xmlDomHeader As New MSXML.DOMDocument
    xmlDomHeader.Load GetAbsolutePath(TAX_Utilities_v2.DataFolder & "Header_01.xml")
    For i = 0 To 12
        mstDN = mstDN & GetAttribute(xmlDomHeader.getElementsByTagName("Cell")(i), "Value")
    Next
    Set xmlDomHeader = Nothing
    getMST = Trim(mstDN)
End Function
