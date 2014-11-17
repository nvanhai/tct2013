Attribute VB_Name = "mdlFunctions"
' Company           : FIS - CMC Software Solution
' Project           : Du an ho tro ke khai thue
' Package           : Interface
' Form, Module
'   or Class name   : mdlFunction
' Descriptions      : public function and variable declare
' Start date        : 10/10/2005 (dd/mm/yyyy)
' Finish date       :
' Coder             : TuyenDS, ThanhDX, TuanLM
' Integrate         :
' Project manager   : ThietKN
' Last modify       :
' Reason of modify  :

Option Explicit

Public Type Quy
    q As Integer
    Y As Integer
    dNgayDauQuy As Date
    dNgayCuoiQuy As Date
End Type
Public strFile(4) As String
Public strNgayTaiChinh As String
Public iNgayTaiChinh As Integer
Public iThangTaiChinh As Integer
Public blnTinhTheoNamTaiChinh As Boolean
Public dNgayDauKy As Date
Public dNgayCuoiKy As Date

Public Const SS_SORT_ORDER_ASCENDING = 1
Public Const APP_VERSION = "9.9.9"
Public Const HTKK_LAST_VERSION = "9.9.9"
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

Const mYear_____ = "T_2"
Const mMonth____ = "T_3"
Const mThreeMonths = "T_4"
Const mTuNgay = "T_6"
Const mDenNgay = "T_7"

Public Type activeForm
    id As String
    showed As Boolean
End Type

Public xmlNodeListMenu As MSXML.IXMLDOMNodeList             ' xml node list for menu
Public xmlHeaderData As New MSXML.DOMDocument               ' xml document for header data
Public xmlSQL As New MSXML.DOMDocument
Public clsDAO As New TAX_Utilities_Svr_New.clsADO
Public arrActiveForm() As activeForm
Public hasActiveForm As Boolean
Public strTaxOfficeId As String                             ' Tax office id
Public strMST As String          ' Tax id
Public strTenGoi As String
Public strDchi As String
Public strNganh As String
Public strMaBPQL As String
Public strDThoai As String
Public strFax As String
Public strTenBpql As String
'vttoan them thong tin dai ly thue
Public strMST_DLT As String
Public strTen_DLT As String
Public strDchi_DLT As String
Public strQHuyen_DLT As String
Public strTTPho_DLT As String
Public strDthoai_DLT As String
Public strFax_DLT As String
Public strMail_DLT As String
Public strSoHD_DLT As String
Public strNgayHD_DLT As String
'end
Public strDBUserName As String                              ' Userid for db QLT
Public strDBPassword As String                              ' Password for db QLT
Public strUserName As String                                ' Name of User
Public strUserID As String                                  ' ID of User
Public TTHTK As String                                      ' Trang thai of to khai
Public LoaiTk As String                                     ' Loai to khai (Chinh thuc,Bsung)

Public LoaiKyKK As Boolean 'True la quy, false la thang

' Su dung de in BB nop cham
Public strPrinterName As String
Public mCurrentSheet As Integer
Public intDataSession As Integer
Public intPrintingSession As Integer
' end
Public strHiddenFormName As String                          ' Save name of hidden form

Public isPITActive As Boolean   ' Kiem tra trang thai active cua PIT

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
        
    strDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(0), "Folder")) & "Header_001.xml"
    xmlHeaderData.Load strDataFileName
    With pGrid
        .Sheet = pSheet
        
        Set xmlNodeListCell = xmlHeaderData.getElementsByTagName("Cell")

        For Each xmlNodeCell In xmlNodeListCell
            ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID2"), lCol, lRow
            .SetText lCol, lRow, GetAttribute(xmlNodeCell, "Value")
        Next

        SetActiveCell pGrid, mYear_____
        .Text = TAX_Utilities_Svr_New.Year
        SetActiveCell pGrid, mMonth____
        .Text = TAX_Utilities_Svr_New.Month
        SetActiveCell pGrid, mThreeMonths
        .Text = TAX_Utilities_Svr_New.ThreeMonths
        SetActiveCell pGrid, mTuNgay
        .Text = TAX_Utilities_Svr_New.FirstDay
        SetActiveCell pGrid, mDenNgay
        .Text = TAX_Utilities_Svr_New.LastDay
        
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
    Dim xmlNodeValidity     As MSXML.IXMLDOMNode
    
    Dim ValidityDate        As Date, StartDate As Date, MaxDate As Date
    Dim idToKhai            As String
    
    idToKhai = GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID")

    If idToKhai = "01" Or idToKhai = "02" Or idToKhai = "04" Or idToKhai = "71" Or idToKhai = "36" Or idToKhai = "68" Or idToKhai = "25" Then
        If LoaiKyKK = False Then

            Select Case TAX_Utilities_Svr_New.Month

                Case "01"
                    ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

                Case "02"
                    ValidityDate = format("28/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

                Case "03"
                    ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

                Case "04"
                    ValidityDate = format("30/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

                Case "05"
                    ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

                Case "06"
                    ValidityDate = format("30/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

                Case "07"
                    ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

                Case "08"
                    ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

                Case "09"
                    ValidityDate = format("30/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

                Case "10"
                    ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

                Case "11"
                    ValidityDate = format("30/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

                Case "12"
                    ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")
            End Select
        
        Else

            Select Case TAX_Utilities_Svr_New.ThreeMonths

                Case "01", "02", "03", "04"
                    ValidityDate = GetNgayCuoiQuy(CInt(TAX_Utilities_Svr_New.ThreeMonths), CInt(TAX_Utilities_Svr_New.Year), iNgayTaiChinh, iThangTaiChinh)
            End Select

        End If

    Else

        If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Month") = "1" Then

        Select Case TAX_Utilities_Svr_New.Month

            Case "01"
                ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

            Case "02"
                ValidityDate = format("28/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

            Case "03"
                ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

            Case "04"
                ValidityDate = format("30/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

            Case "05"
                ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

            Case "06"
                ValidityDate = format("30/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

            Case "07"
                ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

            Case "08"
                ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

            Case "09"
                ValidityDate = format("30/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

            Case "10"
                ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

            Case "11"
                ValidityDate = format("30/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")

            Case "12"
                ValidityDate = format("31/" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "dd/mm/yyyy")
        End Select
        
    ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ThreeMonth") = "1" Then

        Select Case TAX_Utilities_Svr_New.ThreeMonths

            Case "01", "02", "03", "04"
                ValidityDate = GetNgayCuoiQuy(CInt(TAX_Utilities_Svr_New.ThreeMonths), CInt(TAX_Utilities_Svr_New.Year), iNgayTaiChinh, iThangTaiChinh)
        End Select

        ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Year") = "1" Then
            ValidityDate = NgayCuoiNamTaiChinh(CInt(TAX_Utilities_Svr_New.Year), iNgayTaiChinh, iThangTaiChinh)
        Else
            ValidityDate = Date
        End If

    End If
    
    Set xmlNodeListValidity = TAX_Utilities_Svr_New.NodeMenu.selectNodes("Validity")

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
    SaveErrorLog "mdlFunctions", "LoadHeaderData", Err.Number, Err.Description
End Function

''' LoadTemplate description
''' Load a Excel template to grid, the name and the path of MS Excel get from TAX_Utilities_Svr_New.NodeMenu (attribute "InterfaceTemplate")
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
        
    If TAX_Utilities_Svr_New.NodeMenu Is Nothing Then Exit Sub
    'TAX_Utilities_Svr_New.NodeValidity = GetValidityNode
    '**********************
    'ThanhDX added
    If TAX_Utilities_Svr_New.NodeValidity Is Nothing Then
        TAX_Utilities_Svr_New.NodeValidity = GetValidityNode
    End If
    '**********************
    If IsInterface = True Then
        lFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity, "InterfaceTemplate"))
    Else
        lFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity, "ReportTemplate"))
    End If

    With pGrid
        .ImportExcelBook lFileName, vbNullString
        For i = 1 To .SheetCount
            .Sheet = i
            lSheetExist = False
            For Each xmlNodeSheet In TAX_Utilities_Svr_New.NodeValidity.childNodes
                If UCase(GetAttribute(xmlNodeSheet, "ID")) = UCase(.SheetName) Then
'                    lSheetExist = True
                    '*****************
                    'ThanhDX added
                    If GetAttribute(xmlNodeSheet, "Active") <> "0" Then
                        lSheetExist = True
                    End If
                    '*****************
                    Exit For
                End If
            Next
            If lSheetExist = False Then
                If UCase(.SheetName) = UCase("Header") Then
                    LoadHeaderData pGrid, .Sheet
                End If
                .SheetVisible = False
            End If
        Next
    End With
    '***************************************
    'ThanhDX added
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
'    TAX_Utilities_Svr_New.xmlDataReDim (TAX_Utilities_Svr_New.NodeValidity.childNodes.length - 1)
'
'    With pGrid
'        .EventEnabled(EventAllEvents) = False
'        For lSheet = 0 To TAX_Utilities_Svr_New.xmlDataCount
'            'If GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "Active") <> "0" Then
'                .Sheet = lSheet + 1
'
'                TAX_Utilities_Svr_New.Data(lSheet) = New MSXML.DOMDocument
'                TAX_Utilities_Svr_New.Data(lSheet).resolveExternals = True
'                TAX_Utilities_Svr_New.Data(lSheet).validateOnParse = True
'                TAX_Utilities_Svr_New.Data(lSheet).async = False
'                strOriginDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "Folder")) & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
'                If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Year") = "0" Then
'                    strDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "Folder")) & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
'                Else
'                    If Val(TAX_Utilities_Svr_New.Month) <> 0 Then
'                        strDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "Folder")) & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_Svr_New.Month & TAX_Utilities_Svr_New.Year & ".xml"
'                    Else
'                        strDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "Folder")) & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_Svr_New.ThreeMonths & TAX_Utilities_Svr_New.Year & ".xml"
'                    End If
'                End If
'                TAX_Utilities_Svr_New.Data(lSheet).Load strDataFileName
'                If TAX_Utilities_Svr_New.Data(lSheet).parseError.reason <> vbNullString Then
'                    If InStr(1, TAX_Utilities_Svr_New.Data(lSheet).parseError.errorCode, "2146697210") <> 0 Then
'                        TAX_Utilities_Svr_New.Data(lSheet).Load strOriginDataFileName
'                        If TAX_Utilities_Svr_New.Data(lSheet).parseError.reason <> vbNullString Then
'                            MsgBox TAX_Utilities_Svr_New.Data(lSheet).parseError.reason
'                        End If
'                    Else
'                        MsgBox TAX_Utilities_Svr_New.Data(lSheet).parseError.reason
'                    End If
'                End If
'
'                ' If load original data -> not fill
'                Set xmlNodeListCell = TAX_Utilities_Svr_New.Data(lSheet).getElementsByTagName("Cell")
'
'                For Each xmlNodeCell In xmlNodeListCell
'                    ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
'                    If GetAttribute(xmlNodeCell, "FirstCell") = "1" Then
'                        .MaxRows = .MaxRows + 1
'                        .InsertRows lRow, 1
'                        .CopyRowRange lRow - 1, lRow - 1, lRow
'                        '*************
'                        'ThanhDX added
'                        ResetRow pGrid, lRow
'                        '*************
'                    End If
'                    .Col = lCol
'                    .Row = lRow
'                    Select Case .CellType
'                        Case CellTypeCheckBox
'                            ' Check box
'                            If UCase(GetAttribute(xmlNodeCell, "Value")) = UCase("x") Then
'                                .Text = "1"
'                            Else
'                                .Text = "0"
'                            End If
'                        Case CellTypeComboBox, CellTypeDate
'                            .SetText lCol, lRow, GetAttribute(xmlNodeCell, "Value")
'                        Case Else
'
'                            If IsNumeric(GetAttribute(xmlNodeCell, "MaxLen")) Then
'                                .TypeMaxEditLen = Val(GetAttribute(xmlNodeCell, "MaxLen"))
'                            End If
'                            If .CellType = CellTypeNumber And IsNumeric(GetAttribute(xmlNodeCell, "MinValue")) And IsNumeric(GetAttribute(xmlNodeCell, "MaxValue")) Then
'                                .TypeNumberMin = Val(GetAttribute(xmlNodeCell, "MinValue"))
'                                .TypeNumberMax = Val(GetAttribute(xmlNodeCell, "MaxValue"))
'                            End If
'
'                            .Value = GetAttribute(xmlNodeCell, "Value")
'                    End Select
'                    '*************
'                    'ThanhDX added
'                    If .RowHeight(lRow) < .MaxTextRowHeight(lRow) Then _
'                        .RowHeight(lRow) = .MaxTextRowHeight(lRow)
'                    '*************
'                Next
'
'                Set xmlNodeCell = Nothing
'                Set xmlNodeListCell = Nothing
'            'End If
'        Next
'        .EventEnabled(EventAllEvents) = True
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
    Dim blnNewData As Boolean
    Dim strDataFileName As String
    Dim strOriginDataFileName As String

    TAX_Utilities_Svr_New.xmlDataReDim (TAX_Utilities_Svr_New.NodeValidity.childNodes.length - 1)

    With pGrid
        '.EventEnabled(EventAllEvents) = False
        For lSheet = 0 To TAX_Utilities_Svr_New.xmlDataCount
            'If GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "Active") <> "0" Then
                blnNewData = False
                .Sheet = lSheet + 1
                TAX_Utilities_Svr_New.Data(lSheet) = New MSXML.DOMDocument
                TAX_Utilities_Svr_New.Data(lSheet).resolveExternals = True
                TAX_Utilities_Svr_New.Data(lSheet).validateOnParse = True
                TAX_Utilities_Svr_New.Data(lSheet).async = False
                strOriginDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "TemplateFolder")) & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Year") = "0" Then
                    strDataFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                Else
                    If GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Month") = "1" Then
                        strDataFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_Svr_New.Month & TAX_Utilities_Svr_New.Year & ".xml"
                    ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ThreeMonth") = "1" Then
                        strDataFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_Svr_New.ThreeMonths & TAX_Utilities_Svr_New.Year & ".xml"
                    ElseIf GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "Year") = "1" Then
                        strDataFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "DataFile") & "_00" & TAX_Utilities_Svr_New.Year & ".xml"
                    End If
                End If
                TAX_Utilities_Svr_New.Data(lSheet).Load strDataFileName
                If TAX_Utilities_Svr_New.Data(lSheet).parseError.reason <> vbNullString Then
                    If InStr(1, TAX_Utilities_Svr_New.Data(lSheet).parseError.errorCode, "2146697210") <> 0 Then
                        'New data
                        blnNewData = True
                        TAX_Utilities_Svr_New.Data(lSheet).Load strOriginDataFileName
                        If TAX_Utilities_Svr_New.Data(lSheet).parseError.reason <> vbNullString Then
                            MsgBox TAX_Utilities_Svr_New.Data(lSheet).parseError.reason
                        End If
                    Else
                        MsgBox TAX_Utilities_Svr_New.Data(lSheet).parseError.reason
                    End If
                End If

                ' If load original data -> not fill
                Set xmlNodeListCell = TAX_Utilities_Svr_New.Data(lSheet).getElementsByTagName("Cell")

                For Each xmlNodeCell In xmlNodeListCell
                    ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
                    If GetAttribute(xmlNodeCell, "FirstCell") = "1" Then
                        lRows = GetDynRowCount(pGrid, xmlNodeCell.parentNode)
                        InsertRow pGrid, lRow, lRows
'                        .MaxRows = .MaxRows + lRows
'                        .InsertRows lRow, lRows
'                        .CopyRowRange lRow - lRows, lRow - lRows, lRow
                    End If
                    .Col = lCol
                    .Row = lRow
                If GetAttribute(xmlNodeCell, "Receive") <> "0" Then
                    Select Case .CellType
                        Case CellTypeCheckBox
                            ' Check box
                            If UCase(GetAttribute(xmlNodeCell, "Value")) = UCase("x") Or UCase(GetAttribute(xmlNodeCell, "Value")) = "1" Then
                                .Text = "1"
                            Else
                                .Text = "0"
                            End If
                        Case CellTypeComboBox
                            If blnNewData And .Text <> GetAttribute(xmlNodeCell, "Value") Then
                                SetAttribute xmlNodeCell, "Value", .Text
                            Else
                                .Text = GetAttribute(xmlNodeCell, "Value")
                            End If
'                        Case CellTypeDate
'                            .CellType = CellTypeEdit
'                            .SetText lCol, lRow, GetAttribute(xmlNodeCell, "Value")
'                            .CellType = CellTypeDate
'*******************************
'ThanhDX added
'Date: 09/01/2006
                        Case CellTypePic
                            If blnNewData And .Text <> GetAttribute(xmlNodeCell, "Value") Then
                                SetAttribute xmlNodeCell, "Value", .Text
                            Else
                                .SetText lCol, lRow, GetAttribute(xmlNodeCell, "Value")
                            End If
'*******************************
                        Case Else
                            If blnNewData And .Value <> GetAttribute(xmlNodeCell, "Value") Then
                                SetAttribute xmlNodeCell, "Value", .Value
                            Else
                                .Value = GetAttribute(xmlNodeCell, "Value")
                            End If
                    End Select
                  Else
                    UpdateCellReceive pGrid, .Col, .Row, .Value
                  End If
                    .RowHeight(lRow) = 14
                    If .RowHeight(lRow) < .MaxTextRowHeight(lRow) Then
                        .RowHeight(lRow) = .MaxTextRowHeight(lRow)
                    End If
                Next

                Set xmlNodeCell = Nothing
                Set xmlNodeListCell = Nothing
            'End If
        Next
        '.EventEnabled(EventAllEvents) = True
    End With

    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "SetupData", Err.Number, Err.Description
End Sub

'********************************************************
'Descriptions:InsertRow procedure insert range of dynamic rows onto
'             Screen
'Author: ThanhDX
'Date: 24/14/2006
'Input:
'       fpSpread1: fpSpread
'       pRow: row start insert
'       lRows: Count of row will be insert
'********************************************************
Public Sub InsertRow(fpSpread1 As fpSpread, ByVal pRow As Long, lRows As Long)
    On Error GoTo ErrorHandle
    
    Dim i As Long, lBgColor As Long
    Dim lRowCtrl As Long, lColCtrl As Long
    Dim mCurrentSheet As Long
    
    With fpSpread1
        '.Sheet = mCurrentSheet
        .MaxRows = .MaxRows + lRows
        .InsertRows pRow, lRows
        For lRowCtrl = 1 To lRows
        
            .CopyRowRange pRow - lRowCtrl, pRow - lRowCtrl, pRow + lRows - lRowCtrl
            .Row = pRow - lRowCtrl
            For i = 1 To fpSpread1.MaxCols
                '***************************
                'ThanhDX added
                'Date: 26/12/2005
                .Col = i
                lBgColor = .BackColor
                .Row = pRow + lRows - lRowCtrl
                If Not .Lock Then
                    'Set BgColor to inserted cell
                    If lBgColor <> &HC0C0FF Then 'vbRed
                        .BackColor = lBgColor
                    Else
                        .BackColor = vbWhite
                    End If
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
                End If
            Next i
        Next lRowCtrl
    End With
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "mdlFunctions", "InsertRow", Err.Number, Err.Description
End Sub


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
    
    Dim tmpArray() As Variant
    Dim num As Integer
    Dim i As Integer
        
    num = Int(LenB(strData) / numByte) + 1
    
    For i = 1 To num
        ReDim Preserve tmpArray(i)
        tmpArray(i) = MidB(strData, (i - 1) * numByte + 1, numByte)
    Next
    CutStringByNumByte = tmpArray()

Exit Function
ErrHandle:
    SaveErrorLog "mdlFunction", "CutStringByNumByte", Err.Number, Err.Description
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

Public Sub ValidFormatDate(txtDate As MSForms.TextBox, format As String)

    Select Case format
        Case "M"
            If Not ValidNumber(txtDate.Text, 12) Then
                DisplayMessage "0018", msOKOnly
                txtDate.SetFocus
            ElseIf Len(txtDate.Text) = 1 Then
                txtDate.Text = "0" & txtDate.Text
            End If
        Case "Q"
            If Not ValidNumber(txtDate.Text, 4) Then
                DisplayMessage "0018", msOKOnly
                txtDate.SetFocus
            End If
        Case "Y"
            If Not IsNumeric(txtDate.Text) Then
                DisplayMessage "0018", msOKOnly
                txtDate.SetFocus
            ElseIf Len(txtDate.Text) = 3 Then
                If CInt(txtDate.Text) >= 100 Then
                    txtDate.Text = "1" & txtDate.Text
                Else
                    txtDate.Text = "2" & txtDate.Text
                End If
            ElseIf Len(txtDate.Text) = 2 Then
                If CInt(txtDate.Text) >= 80 Then
                    txtDate.Text = "19" & txtDate.Text
                Else
                    txtDate.Text = "20" & txtDate.Text
                End If
            ElseIf Len(txtDate.Text) = 1 Then
                txtDate.Text = "200" & txtDate.Text
            End If
        Case Else
        
    End Select
End Sub

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
    On Error GoTo ErrorHandle
    
    CreateCell = GetAttribute(xmlNodeCell, "Value") & "~"
    
    Exit Function
ErrorHandle:
    SaveErrorLog "mdlFunction", "CreateCell", Err.Number, Err.Description
End Function

Private Function CreateCells(xmlNodeCells As MSXML.IXMLDOMNode) As String
    On Error GoTo ErrorHandle
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    For Each xmlNodeCell In xmlNodeCells.childNodes
        
        CreateCells = CreateCells & CreateCell(xmlNodeCell)
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
    CreateSection = Left(CreateSection, Len(CreateSection) - 1)
    CreateSection = "<S>" & CreateSection & "</S>"
    
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
    CreateSections = "<" & pSheet & ">" & CreateSections & "</" & pSheet & ">"
    
    Exit Function
ErrorHandle:
    SaveErrorLog "mdlFunction", "CreateSections", Err.Number, Err.Description
End Function

'****************************************************
'Description: ResetRow procedure reset all of data in row
'Author:ThanhDX.
'Modify by:
'Date:14/11/2005
'Input: lRow: Row is reset
'Output:
'Return:

'****************************************************

Private Sub ResetRow(fpsGrid As fpSpread, lRow As Long)
Dim lCtrl As Long

With fpsGrid
    .Row = lRow
    For lCtrl = 1 To .MaxCols
        .Col = lCtrl
        .Value = ""
    Next lCtrl
    .RowHeight(lRow) = 14
End With
End Sub

' Ham su dung in BB nop cham
Private Sub ResetRowPrint(ByVal xmlCellNode As MSXML.IXMLDOMNode, fpsGrid As fpSpread, ByVal lRow As Long, ByVal lRows As Long)
    Dim lRowCtrl As Long, lColCtrl As Long
    Dim xmlCellsNode As MSXML.IXMLDOMNode
    Dim xmlTempCellNode As MSXML.IXMLDOMNode
    Dim lngCol As Long, lngRow As Long
    
    Set xmlCellsNode = xmlCellNode.parentNode
    For Each xmlTempCellNode In xmlCellsNode.childNodes
        ParserCellID fpsGrid, GetAttribute(xmlTempCellNode, "CellID"), lngCol, lngRow
        With fpsGrid
            .Col = lngCol
            .Row = lngRow
            If .CellType = CellTypeNumber Then
                .Value = 0
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

'****************************************************
'Description: GetNgayTaiChinh function reset all of data in row
'Author:Gianvd.
'Modify by:
'Date:07/02/2006
'Input: strDate: Date of NgayBatDauNamTaiChinh
'Output:
'Return:

'****************************************************
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

Public Function GetNgayDauQuy(q As Integer, Y As Integer, dNgayTaiChinh As Integer, dThangTaiChinh As Integer) As Date
    Dim intYear As Integer, intDay As Integer, intMonth As Integer
    
    If blnTinhTheoNamTaiChinh And (GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "FinanceYear") = "1") Then
        intDay = dNgayTaiChinh
        intMonth = (q - 1) * 3 + dThangTaiChinh
        intYear = Y
        If intMonth > 12 Then
            intMonth = intMonth - 12
            intYear = Y + 1
        End If
    Else
        intDay = 1
        intYear = Y
        Select Case q
            Case 1
                intMonth = 1
            Case 2
                intMonth = 4
            Case 3
                intMonth = 7
            Case 4
                intMonth = 10
        End Select
        
    End If
    GetNgayDauQuy = DateSerial(intYear, intMonth, intDay)
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
    
        iInterval = DateDiff("D", DateSerial(yTaiChinhDau, mTaiChinhDau, 1), DateSerial(yTaiChinhCuoi, mTaiChinhCuoi, 1)) - 1
        GetNgayCuoiQuy = DateSerial(yTaiChinhDau, mTaiChinhDau, 1) + iInterval
    Else
        GetNgayCuoiQuy = DateSerial(yTaiChinhDau, mTaiChinhDau, 1)
    End If
End Function

Public Function GetNgayDauNam(Y As Integer, dThangTaiChinh As Integer, dNgayTaiChinh As Integer) As Date
    Dim intYear As Integer, intDay As Integer, intMonth As Integer
    
    If blnTinhTheoNamTaiChinh And (GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "FinanceYear") = "1") Then
        intYear = Y
        intMonth = dThangTaiChinh
        intDay = dNgayTaiChinh
    Else
        intDay = 1
        intYear = Y
        intMonth = 1
    End If
    GetNgayDauNam = DateSerial(intYear, intMonth, intDay)
End Function

Function NgayCuoiNamTaiChinh(Y As Integer, dThangTaiChinh As Integer, dNgayTaiChinh As Integer)
    Dim dNgayTC As Date
    
    dNgayTC = DateSerial(Y, dThangTaiChinh, dNgayTaiChinh)
    NgayCuoiNamTaiChinh = DateAdd("M", 12, dNgayTC)
    NgayCuoiNamTaiChinh = DateAdd("D", -1, NgayCuoiNamTaiChinh)
    
End Function

Function hannop() As String
    Dim dNgayCuoiKy As Date
    Dim dHanNop     As Date
    Dim arrDate()   As String
    Dim idToKhai    As Variant

    idToKhai = GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "ID")

    If idToKhai = "01" Or idToKhai = "02" Or idToKhai = "04" Or idToKhai = "71" Or idToKhai = "36" Or idToKhai = "68" Or idToKhai = "25" Then
        If LoaiKyKK = False Then
            If TAX_Utilities_Svr_New.Month = 12 Then
                hannop = "20/" & "01" & "/" & TAX_Utilities_Svr_New.Year + 1
            Else
                hannop = "20/" & Right("0" & TAX_Utilities_Svr_New.Month + 1, 2) & "/" & TAX_Utilities_Svr_New.Year
            End If

        Else

            If TAX_Utilities_Svr_New.ThreeMonths = "04" Then
                If TAX_Utilities_Svr_New.Year = 2013 Then
                    hannop = "06/" & "02" & "/" & TAX_Utilities_Svr_New.Year + 1

                Else
                    hannop = "30/" & "01" & "/" & TAX_Utilities_Svr_New.Year + 1
                    
                End If

            ElseIf TAX_Utilities_Svr_New.ThreeMonths = "03" Then
                '            If TAX_Utilities_Svr_New.Year = 2011 Then
                hannop = "30/" & "10" & "/" & TAX_Utilities_Svr_New.Year
                '            Else
                '                hannop = "01/" & "11" & "/" & TAX_Utilities_Svr_New.Year
                '            End If
            ElseIf TAX_Utilities_Svr_New.ThreeMonths = "02" Then
                hannop = "30/" & "07" & "/" & TAX_Utilities_Svr_New.Year
            ElseIf TAX_Utilities_Svr_New.ThreeMonths = "01" Then
                hannop = "02/" & "05" & "/" & TAX_Utilities_Svr_New.Year
            End If

        End If

    Else

        If TAX_Utilities_Svr_New.Month <> "" Then
            If TAX_Utilities_Svr_New.Month = 12 Then
                hannop = "20/" & "01" & "/" & TAX_Utilities_Svr_New.Year + 1
            Else
                hannop = "20/" & Right("0" & TAX_Utilities_Svr_New.Month + 1, 2) & "/" & TAX_Utilities_Svr_New.Year
            End If

        ElseIf TAX_Utilities_Svr_New.ThreeMonths <> "" Then

            If TAX_Utilities_Svr_New.ThreeMonths = "04" Then
                If TAX_Utilities_Svr_New.Year = 2013 Then
                    hannop = "06/" & "02" & "/" & TAX_Utilities_Svr_New.Year + 1

                Else
                    hannop = "30/" & "01" & "/" & TAX_Utilities_Svr_New.Year + 1
                    
                End If

            ElseIf TAX_Utilities_Svr_New.ThreeMonths = "03" Then
                '            If TAX_Utilities_Svr_New.Year = 2011 Then
                hannop = "30/" & "10" & "/" & TAX_Utilities_Svr_New.Year
                '            Else
                '                hannop = "01/" & "11" & "/" & TAX_Utilities_Svr_New.Year
                '            End If
            ElseIf TAX_Utilities_Svr_New.ThreeMonths = "02" Then
                hannop = "30/" & "07" & "/" & TAX_Utilities_Svr_New.Year
            ElseIf TAX_Utilities_Svr_New.ThreeMonths = "01" Then
                hannop = "02/" & "05" & "/" & TAX_Utilities_Svr_New.Year
            End If

        Else
            dNgayCuoiKy = DateAdd("D", 90, NgayCuoiNamTaiChinh(TAX_Utilities_Svr_New.Year, iThangTaiChinh, iNgayTaiChinh))
            hannop = format(dNgayCuoiKy, "dd/mm/yyyy")
        End If

    End If



    arrDate = Split(hannop, "/")
    dHanNop = DateSerial(CInt(arrDate(2)), CInt(arrDate(1)), CInt(arrDate(0)))

    'Neu vao ngay thu 7 thi cong them 2 ngay,  ngay CN thi cong them mot ngay
    '    If Weekday(CDate(hannop)) = 7 Then
    '        hannop = DateAdd("D", 2, CDate(hannop))
    '        hannop = format(hannop, "dd/mm/yyyy")
    '    ElseIf Weekday(CDate(hannop)) = 1 Then
    '        hannop = DateAdd("D", 1, CDate(hannop))
    '        hannop = format(hannop, "dd/mm/yyyy")
    '    End If
    If Weekday(dHanNop) = 7 Then
        hannop = DateAdd("D", 2, dHanNop)
        hannop = format(hannop, "dd/mm/yyyy")
    ElseIf Weekday(dHanNop) = 1 Then
        hannop = DateAdd("D", 1, dHanNop)
        hannop = format(hannop, "dd/mm/yyyy")
    Else
        hannop = format(dHanNop, "dd/mm/yyyy")
    End If
    
End Function


''' UpdateCell description
''' Update cell value to DOM object when user change cell value
''' Parameter1 pCol     : active column
''' Parameter2 pRow     : active row
''' Parameter3 pValue   : cell value need update
Private Function UpdateCellReceive(fps As fpSpread, ByVal pCol As Long, ByVal pRow As Long, ByVal pValue As String) As Boolean
    On Error GoTo ErrHandle
    
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    GetCellSpan fps, pCol, pRow
    
    Set xmlNodeCell = TAX_Utilities_Svr_New.Data(fps.ActiveSheet - 1).nodeFromID(GetCellID(fps, pCol, pRow))
    
    If GetAttribute(xmlNodeCell, "Value") <> pValue Then
        SetAttribute xmlNodeCell, "Value", pValue
        UpdateCellReceive = True
    End If
    
    Set xmlNodeCell = Nothing
    
    Exit Function
    
ErrHandle:
    SaveErrorLog "mdlFunction", "UpdateCellReceive", Err.Number, Err.Description
End Function

' Ham su dung in BB nop cham
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
    For lSheet = 0 To TAX_Utilities_Svr_New.xmlDataCount
        If GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "Active") <> "0" Then
            .Sheet = lSheet + 1
            
            Set xmlNodeListCell = TAX_Utilities_Svr_New.Data(lSheet).getElementsByTagName("Cell")
    
            For Each xmlNodeCell In xmlNodeListCell
                ParserCellID fpsGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
                If GetAttribute(xmlNodeCell, "FirstCell") = "1" Then
                    GetDynRowCount fpsGrid, xmlNodeCell.parentNode, lRow2s
                    InsertRowPrint fpsGrid, lRow, lRow2s, True
                    ResetRowPrint xmlNodeCell, fpsGrid, lRow, lRow2s
                End If
                .Col = lCol
                .Row = lRow
                
                If GetAttribute(xmlNodeCell, "PageBreak") = "1" Then
                    If Not xmlNodeCell.parentNode.nextSibling Is Nothing Then
                        .RowPageBreak = True
                    Else
                        ' Xu ly rieng cho to quyet toan 09/TNCN (05TNCN->09TNCN)
                        If TAX_Utilities_Svr_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue = "45" Then
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
                If Not IsNullNumber(GetAttribute(xmlNodeCell, "Value")) And GetAttribute(xmlNodeCell, "Value") <> "" And lRow <> 0 And lCol <> 0 Then
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
                            If GetAttribute(TAX_Utilities_Svr_New.NodeValidity.parentNode, "ID") = "17" Or GetAttribute(TAX_Utilities_Svr_New.NodeValidity.parentNode, "ID") = "42" Or GetAttribute(TAX_Utilities_Svr_New.NodeValidity.parentNode, "ID") = "43" Then
                                .TypeNumberDecPlaces = 0
                            Else
                                .TypeNumberDecPlaces = 3 'Len(GetAttribute(xmlNodeCell, "Value")) - InStr(1, GetAttribute(xmlNodeCell, "Value"), ".")
                            End If
                        End If
                    
                    End If
                End If
'end edit
'                    Debug.Print xmlNodeCell.xml
                    .Value = GetAttribute(xmlNodeCell, "Value")
                Else
                    .SetText lCol, lRow, ""
                End If
                
                If GetAttribute(TAX_Utilities_Svr_New.NodeValidity.parentNode, "ID") = "12" And .Col = .ColLetterToNumber("CD") And (.Row = "27" Or .Row = "25") Then
                   If .Text <> "" Then
                        If Len(.Text) <= 2 Then
                            .Text = Left(.Text & ".000", 6)
                        ElseIf Len(.Text) > 2 Then
                            .Text = Left$(.Text & "000", 6)
                        End If
                        
                        If Right(.Text, Len(.Text) - InStr(1, .Text, ".")) = "000" Then
                              .CellType = CellTypeEdit
                              .TypeHAlign = TypeHAlignRight
                              .Text = Left(.Text, Len(.Text) - 4) & "%"
                          Else
                              .CellType = CellTypeEdit
                              .TypeHAlign = TypeHAlignRight
                              .Text = Left(.Text, Len(.Text) - 4) & "," & Right(.Text, 3) & "%"
                          End If
                    End If
                End If
                
                 If GetAttribute(TAX_Utilities_Svr_New.NodeValidity.parentNode, "ID") = "12" And .Col = .ColLetterToNumber("CD") And .Row = "28" And .Text <> "" Then
                    If GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("G_16"), "Value") = "x" Or GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("G_16"), "Value") = "1" Then
                        .CellType = CellTypeEdit
                        .TypeHAlign = TypeHAlignRight
                        .Text = .Text & "%"
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
                            .Text = Left(.Text, Len(.Text) - 4) & "%"
                        Else
                            .CellType = CellTypeEdit
                            .TypeHAlign = TypeHAlignRight
                            .Text = Left(.Text, Len(.Text) - 4) & "," & Right(.Text, 3) & "%"
                        End If
                    End If
                End If
                
                
                
                
'                If GetAttribute(TAX_Utilities_Svr_New.NodeValidity.parentNode, "ID") = "12" And .Col = .ColLetterToNumber("CD") And (.Row = "27" Or .Row = "25") Then
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
'                 If GetAttribute(TAX_Utilities_Svr_New.NodeValidity.parentNode, "ID") = "12" And .Col = .ColLetterToNumber("CD") And .Row = "28" Then
'                    If GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("G_16"), "Value") = "x" Then
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
Public Sub InsertRowPrint(fpSpread1 As fpSpread, ByVal pRow As Long, lRows As Long, Optional blnFillingData As Boolean = False)
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
                        If Not TAX_Utilities_Svr_New.Data(mCurrentSheet - 1).nodeFromID( _
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
                        If Not TAX_Utilities_Svr_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, i, pRow - lRowCtrl)) Is Nothing Then
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
Public Function IsNullNumber(ByVal strValue As String) As Boolean
    strValue = Replace$(strValue, "0", "")
    strValue = Replace$(strValue, ".", "")
    If Trim(strValue) = "" Then IsNullNumber = True
End Function
Public Function GetMessageCellById(ByVal strID As String) As MSXML.IXMLDOMNode
    Dim xmlInforNode As MSXML.IXMLDOMNode
    
    For Each xmlInforNode In TAX_Utilities_Svr_New.NodeMessage
        If GetAttribute(xmlInforNode, "ID") = strID Then
            Set GetMessageCellById = xmlInforNode
            Exit Function
        End If
    Next
End Function
Public Sub PrinterKillDoc()
    Printer.KillDoc
    Printer.PaperSize = vbPRPSA4
End Sub

Public Sub PrinterEndDoc()
    Printer.EndDoc
    Printer.PaperSize = vbPRPSA4
End Sub
Public Sub SetupDataPrint()
    On Error GoTo ErrorHandle

    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim lSheet As Long, lCol As Long, lRow As Long
    Dim lRows As Long
    Dim blnNewData As Boolean
    Dim strDataFileName As String
    Dim strOriginDataFileName As String

    TAX_Utilities_Svr_New.xmlDataReDim (TAX_Utilities_Svr_New.NodeValidity.childNodes.length - 1)

        '.EventEnabled(EventAllEvents) = False
        For lSheet = 0 To TAX_Utilities_Svr_New.xmlDataCount
            'If GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "Active") <> "0" Then
                blnNewData = False
                '.Sheet = lSheet + 1
                TAX_Utilities_Svr_New.Data(lSheet) = New MSXML.DOMDocument
                TAX_Utilities_Svr_New.Data(lSheet).resolveExternals = True
                TAX_Utilities_Svr_New.Data(lSheet).validateOnParse = True
                TAX_Utilities_Svr_New.Data(lSheet).async = False
                strOriginDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "TemplateFolder")) & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                    strDataFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Svr_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                    TAX_Utilities_Svr_New.Data(lSheet).Load strDataFileName
                If TAX_Utilities_Svr_New.Data(lSheet).parseError.reason <> vbNullString Then
                    If InStr(1, TAX_Utilities_Svr_New.Data(lSheet).parseError.errorCode, "2146697210") <> 0 Then
                        'New data
                        blnNewData = True
                        TAX_Utilities_Svr_New.Data(lSheet).Load strOriginDataFileName
                        If TAX_Utilities_Svr_New.Data(lSheet).parseError.reason <> vbNullString Then
                            MsgBox TAX_Utilities_Svr_New.Data(lSheet).parseError.reason
                        End If
                    Else
                        MsgBox TAX_Utilities_Svr_New.Data(lSheet).parseError.reason
                    End If
                End If

                Set xmlNodeCell = Nothing
                Set xmlNodeListCell = Nothing
        Next
    Exit Sub
ErrorHandle:
    SaveErrorLog "mdlFunctions", "SetupData", Err.Number, Err.Description
End Sub


' Get ve ID cua bang data_pkg
Public Function GetDataPkgId() As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim pkgId As Variant
    Dim id As Variant
    Dim noiLamViec As Variant
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    strFileName = spathVat & "\TRAODOI\parm.DBF"
    If fso.FileExists(strFileName) = True Then
        If clsConn.Connected = False Then
            clsConn.CreateConnectionString spathVat & "\TRAODOI\"
            clsConn.Connect
        End If
            'Lay ve noi lam viec
            sSQL = "SELECT prm_value FROM parm " & _
                " WHERE prm_name = 'NOI_LAM_VIEC' "
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 noiLamViec = rs.Fields("prm_value")
            Else
                 noiLamViec = ""
            End If
            pkgId = noiLamViec
            ' lay seq pkg
            sSQL = "SELECT prm_value FROM parm " & _
                " WHERE prm_name = 'EXC_DATA_PKG_SEQ' "
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 id = rs.Fields("prm_value")
                 sSQL = "update parm set prm_value ='" & Val(CStr(id)) + 1 & "' where prm_name = 'EXC_DATA_PKG_SEQ' "
                 clsConn.ExecuteDLL (sSQL)
            Else
                 id = 0
            End If
            pkgId = Trim(CStr(pkgId)) & Trim(CStr(id))
            clsConn.Disconnect
    Else
        pkgId = ""
    End If
    GetDataPkgId = pkgId
End Function

' Get ve Tran_no cua bang tup_exc
Public Function GetTranNo() As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim noiLamViec As Variant
    Dim tupId As Variant
    Dim tranNo As Variant
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    strFileName = spathVat & "\TRAODOI\parm.DBF"
    If fso.FileExists(strFileName) = True Then
        If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & "\TRAODOI\"
        clsConn.Connect
        End If
            'Lay ve noi lam viec
            sSQL = "SELECT prm_value FROM parm " & _
                " WHERE prm_name = 'NOI_LAM_VIEC' "
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 noiLamViec = rs.Fields("prm_value")
            Else
                 noiLamViec = ""
            End If
            tranNo = noiLamViec
            ' Lay seq Tup
            sSQL = "SELECT prm_value FROM parm " & _
                " WHERE prm_name = 'EXC_TRAN_UP_SEQ' "
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 tupId = rs.Fields("prm_value")
                 ' update giai tri tang len
                sSQL = "update parm set prm_value ='" & Val(CStr(tupId)) + 1 & "' where prm_name = 'EXC_TRAN_UP_SEQ' "
                clsConn.ExecuteDLL (sSQL)
            Else
                 tupId = ""
            End If
            tranNo = Trim(CStr(tranNo)) & Trim(CStr(tupId))
            
            clsConn.Disconnect
    Else
        tranNo = ""
    End If
    GetTranNo = tranNo
End Function

' Get ve Tran_no cua bang mup_exc
Public Function GetMupId() As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim noiLamViec As Variant
    Dim seqMupId As Variant
    Dim mupId As Variant
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    strFileName = spathVat & "\TRAODOI\parm.DBF"
    If fso.FileExists(strFileName) = True Then
        If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & "\TRAODOI\"
        clsConn.Connect
        End If
            'Lay ve noi lam viec
            sSQL = "SELECT prm_value FROM parm " & _
                " WHERE prm_name = 'NOI_LAM_VIEC' "
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 noiLamViec = rs.Fields("prm_value")
            Else
                 noiLamViec = ""
            End If
            mupId = noiLamViec
            ' Lay seq Tup
            sSQL = "SELECT prm_value FROM parm " & _
                " WHERE prm_name = 'EXC_MESS_UP_SEQ' "
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 seqMupId = rs.Fields("prm_value")
                    ' update giai tri tang len
                sSQL = "update parm set prm_value ='" & Val(CStr(seqMupId)) + 1 & "' where prm_name = 'EXC_MESS_UP_SEQ' "
                clsConn.ExecuteDLL (sSQL)
            Else
                 seqMupId = ""
            End If
            mupId = Trim(CStr(mupId)) & Trim(CStr(seqMupId))
            clsConn.Disconnect
    Else
        mupId = ""
    End If
    GetMupId = mupId
End Function

' Get thong tin noi lam viec
Public Function GetNoiLamViec() As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim noiLamViec As Variant
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    strFileName = spathVat & "\TRAODOI\parm.DBF"
    If fso.FileExists(strFileName) = True Then
        If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & "\TRAODOI\"
        clsConn.Connect
        End If
            'Lay ve noi lam viec
            sSQL = "SELECT prm_value FROM parm " & _
                " WHERE prm_name = 'NOI_LAM_VIEC' "
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 noiLamViec = rs.Fields("prm_value")
            Else
                 noiLamViec = ""
            End If
            clsConn.Disconnect
    Else
        noiLamViec = ""
    End If
    GetNoiLamViec = Trim(noiLamViec)
End Function

' Get thong tin noi lam viec
Public Function GetNoiNhan(str As String) As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim noiNhan As Variant
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    strFileName = spathVat & "\TRAODOI\parm.DBF"
    If fso.FileExists(strFileName) = True Then
        If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & "\TRAODOI\"
        clsConn.Connect
        End If
            'Lay ve noi lam viec
            sSQL = "SELECT lcn_super FROM LOCA_LST " & _
                " WHERE lcn_code = '" & str & "'"
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 noiNhan = rs.Fields("lcn_super")
            Else
                 noiNhan = ""
            End If
            clsConn.Disconnect
    Else
        noiNhan = ""
    End If
    GetNoiNhan = Trim(noiNhan)
End Function

' Get thong tin tns_code
Public Function GetTnsCode(str As String) As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim tnsCode As Variant
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    strFileName = spathVat & "\TRAODOI\tab_lst.DBF"
    If fso.FileExists(strFileName) = True Then
        If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & "\TRAODOI\"
        clsConn.Connect
        End If
            sSQL = "SELECT tsn_code FROM tab_lst " & _
                " WHERE tab_code = '" & str & "'"
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 tnsCode = rs.Fields("tsn_code")
            Else
                 tnsCode = ""
            End If
            clsConn.Disconnect
    Else
        tnsCode = ""
    End If
    GetTnsCode = Trim(tnsCode)
End Function

' Change ma tk to tab_code
Public Function changeTK2TabCode(str As String) As String
   Dim tabCode As String
   Select Case str
        Case "02A_TNCN10"
             tabCode = "029"
        Case "02B_TNCN10"
             tabCode = "031"
        Case "03A_TNCN10"
             tabCode = "033"
        Case "03B_TNCN10"
             tabCode = "035"
        Case "07_TNCN10"
             tabCode = "041"
        Case Else
            MsgBox ("To khai khong ton tai")
    End Select
    changeTK2TabCode = tabCode
End Function

' Get ve cac pkd_id gui len bi loi
Public Function GetPkgIDErr() As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim pkgIDErr As Variant
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    strFileName = spathVat & "\TRAODOI\data_pkg.DBF"
    pkgIDErr = "('')"
    If fso.FileExists(strFileName) = True Then
        If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & "\TRAODOI\"
        clsConn.Connect
        End If
            sSQL = "SELECT id FROM data_pkg " & _
                " WHERE pkg_type = '1' and curr_sta= '01'"
            Set rs = clsConn.Execute(sSQL)
            If rs.Fields.Count > 0 Then
                 pkgIDErr = "'"
                 Do While Not rs.EOF
                    pkgIDErr = pkgIDErr & Trim(rs.Fields(0).Value) & "','"
                    rs.MoveNext
                 Loop
                 pkgIDErr = Left$(pkgIDErr, Len(pkgIDErr) - 2)
                 pkgIDErr = "(" & pkgIDErr & ")"
            End If
            clsConn.Disconnect
    End If
    GetPkgIDErr = Trim(pkgIDErr)
End Function


' Kiem tra activ PIT
Public Function checkActivePIT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO
    strFileName = spathVat & "\NTK_TG\tmp_parm.DBF"

    Dim resultPIT As Boolean
    On Error GoTo ErrHandle
    resultPIT = False
    'connect to database NTK_TG
    If fso.FileExists(strFileName) = True Then
        If clsConn.Connected = False Then
            clsConn.CreateConnectionString spathVat & "\NTK_TG\"
            clsConn.Connect
        End If
        sSQL = "SELECT prm_value FROM tmp_parm " & _
                " WHERE prm_name = 'NTK.PIT_ACTIVE' "
        Set rs = clsConn.Execute(sSQL)
        If Not rs Is Nothing Then
            If rs.Fields.Count > 0 Then
                If Trim(rs.Fields(0).Value) = "1" Then
                    resultPIT = True
                Else
                    resultPIT = False
                End If
            End If
        End If
    End If
    checkActivePIT = resultPIT
    Exit Function
ErrHandle:
    SaveErrorLog "mdlFunctions", "checkActivePIT", Err.Number, Err.Description
End Function

' THong tin AC
Public Function changeMaToKhai(strID As String) As String

    ' Cac mau an chi
    If strID = "64" Then changeMaToKhai = "01_TBAC"
    If strID = "65" Then changeMaToKhai = "01_AC"
    If strID = "66" Then changeMaToKhai = "BC21_AC"
    If strID = "67" Then changeMaToKhai = "03_TBAC"
    If strID = "68" Then changeMaToKhai = "BC26_AC"
    If strID = "18" Then changeMaToKhai = "BC26_AC_SL"
    If strID = "27" Then changeMaToKhai = "01_BK_BC26_AC"
    If strID = "91" Then changeMaToKhai = "04_TBAC"
    If strID = "07" Then changeMaToKhai = "01_TBAC_BLP"
    If strID = "13" Then changeMaToKhai = "01_AC_BLP"
    If strID = "09" Then changeMaToKhai = "BC21_AC_BLP"
    If strID = "10" Then changeMaToKhai = "03_TBAC_BLP"
    If strID = "14" Then changeMaToKhai = "BC26_AC_BLP"
End Function

'lay ten CQT tu ma CQT
Public Sub GetTenCQT(ByVal id As String, Optional ByRef TenTN As String)
Dim arrDanhsach() As String
Dim strDataFileName As String
Dim xmlDOMdata As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = "..\InterfaceTemplates\Catalogue_Tinh_Thanh.xml"
    
       If xmlDOMdata.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDOMdata.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If id = arrDanhsach(1) Then
                            TenTN = arrDanhsach(3)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub


