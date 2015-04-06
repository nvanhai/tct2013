Attribute VB_Name = "mdlFunctions"
Option Explicit

Public Type Quy
    q As Integer
    y As Integer
    dNgayDauQuy As Date
    dNgayCuoiQuy As Date
End Type

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
Const mTuNgay = "T_5"
Const mDenNgay = "T_6"

Public Type activeForm
    ID As String
    showed As Boolean
End Type

Public xmlNodeListMenu As MSXML.IXMLDOMNodeList             ' xml node list for menu
Public xmlHeaderData As New MSXML.DOMDocument               ' xml document for header data
Public xmlSQL As New MSXML.DOMDocument
Public clsDAO As New TAX_Utilities_Srv_New.clsADO
Public arrActiveForm() As activeForm
Public hasActiveForm As Boolean
Public strTaxOfficeId As String                             ' Tax office id
Public strMST As String                                     ' Tax id
Public strDBUserName As String                              ' Userid for db QLT
Public strDBPassword As String                              ' Password for db QLT
Public strUserName As String                                ' Name of User
Public strUserID As String                                ' ID of User

Public LoaiKyKK As Boolean 'True la quy, false la thang

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
        
    strDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(0), "Folder")) & "Header_001.xml"
    xmlHeaderData.Load strDataFileName
    With pGrid
        .Sheet = pSheet
        
        Set xmlNodeListCell = xmlHeaderData.getElementsByTagName("Cell")

        For Each xmlNodeCell In xmlNodeListCell
            ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID2"), lCol, lRow
            .SetText lCol, lRow, GetAttribute(xmlNodeCell, "Value")
        Next

        SetActiveCell pGrid, mYear_____
        .Text = TAX_Utilities_Srv_New.Year
        SetActiveCell pGrid, mMonth____
        .Text = TAX_Utilities_Srv_New.Month
        SetActiveCell pGrid, mThreeMonths
        .Text = TAX_Utilities_Srv_New.ThreeMonths
        SetActiveCell pGrid, mTuNgay
        .Text = TAX_Utilities_Srv_New.FirstDay
        SetActiveCell pGrid, mDenNgay
        .Text = TAX_Utilities_Srv_New.LastDay
        
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
    
    Dim ValidityDate        As Date, StartDate As Date, MaxDate As Date
    Dim idToKhai            As String
    
    idToKhai = GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ID")
    
    'thang/quy
    If idToKhai = "01" Or idToKhai = "02" Or idToKhai = "25" Or idToKhai = "04" Or idToKhai = "71" Or idToKhai = "36" Or idToKhai = "68" Or idToKhai = "94" Then
        If LoaiKyKK = False Then

            Select Case TAX_Utilities_Srv_New.Month

                Case "01"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "02"
                    ValidityDate = format("28/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "03"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "04"
                    ValidityDate = format("30/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "05"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "06"
                    ValidityDate = format("30/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "07"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "08"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "09"
                    ValidityDate = format("30/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "10"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "11"
                    ValidityDate = format("30/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "12"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")
            End Select
        
        Else

            Select Case TAX_Utilities_Srv_New.ThreeMonths

                Case "01", "02", "03", "04"
                    ValidityDate = GetNgayCuoiQuy(CInt(TAX_Utilities_Srv_New.ThreeMonths), CInt(TAX_Utilities_Srv_New.Year), iNgayTaiChinh, iThangTaiChinh)
            End Select

        End If

    Else

        If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then

            Select Case TAX_Utilities_Srv_New.Month

                Case "01"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "02"
                    ValidityDate = format("28/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "03"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "04"
                    ValidityDate = format("30/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "05"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "06"
                    ValidityDate = format("30/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "07"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "08"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "09"
                    ValidityDate = format("30/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "10"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "11"
                    ValidityDate = format("30/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")

                Case "12"
                    ValidityDate = format("31/" & TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "dd/mm/yyyy")
            End Select
        
        ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = "1" Then
            Select Case TAX_Utilities_Srv_New.ThreeMonths
                Case "01", "02", "03", "04"
                    ValidityDate = GetNgayCuoiQuy(CInt(TAX_Utilities_Srv_New.ThreeMonths), CInt(TAX_Utilities_Srv_New.Year), iNgayTaiChinh, iThangTaiChinh)
            End Select
        ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = "1" Then
            ValidityDate = NgayCuoiNamTaiChinh(CInt(TAX_Utilities_Srv_New.Year), iNgayTaiChinh, iThangTaiChinh)
        Else
            ValidityDate = Date
        End If

    End If
    
    Set xmlNodeListValidity = TAX_Utilities_Srv_New.NodeMenu.selectNodes("Validity")
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
''' Load a Excel template to grid, the name and the path of MS Excel get from TAX_Utilities_Srv_New.NodeMenu (attribute "InterfaceTemplate")
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
        
    If TAX_Utilities_Srv_New.NodeMenu Is Nothing Then Exit Sub
    'TAX_Utilities_Srv_New.NodeValidity = GetValidityNode
    '**********************
    If TAX_Utilities_Srv_New.NodeValidity Is Nothing Then
        TAX_Utilities_Srv_New.NodeValidity = GetValidityNode
    End If
    '**********************
    If IsInterface = True Then
        lFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "InterfaceTemplate"))
    Else
        lFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity, "ReportTemplate"))
    End If

    With pGrid
        .ImportExcelBook lFileName, vbNullString
        For i = 1 To .SheetCount
            .Sheet = i
            lSheetExist = False
            For Each xmlNodeSheet In TAX_Utilities_Srv_New.NodeValidity.childNodes
                If UCase(GetAttribute(xmlNodeSheet, "ID")) = UCase(.SheetName) Then
'                    lSheetExist = True
                    '*****************
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
'    TAX_Utilities_Srv_New.xmlDataReDim (TAX_Utilities_Srv_New.NodeValidity.childNodes.length - 1)
'
'    With pGrid
'        .EventEnabled(EventAllEvents) = False
'        For lSheet = 0 To TAX_Utilities_Srv_New.xmlDataCount
'            'If GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "Active") <> "0" Then
'                .Sheet = lSheet + 1
'
'                TAX_Utilities_Srv_New.Data(lSheet) = New MSXML.DOMDocument
'                TAX_Utilities_Srv_New.Data(lSheet).resolveExternals = True
'                TAX_Utilities_Srv_New.Data(lSheet).validateOnParse = True
'                TAX_Utilities_Srv_New.Data(lSheet).async = False
'                strOriginDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "Folder")) & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
'                If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = "0" Then
'                    strDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "Folder")) & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
'                Else
'                    If Val(TAX_Utilities_Srv_New.Month) <> 0 Then
'                        strDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "Folder")) & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_Srv_New.Month & TAX_Utilities_Srv_New.Year & ".xml"
'                    Else
'                        strDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "Folder")) & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_Srv_New.ThreeMonths & TAX_Utilities_Srv_New.Year & ".xml"
'                    End If
'                End If
'                TAX_Utilities_Srv_New.Data(lSheet).Load strDataFileName
'                If TAX_Utilities_Srv_New.Data(lSheet).parseError.reason <> vbNullString Then
'                    If InStr(1, TAX_Utilities_Srv_New.Data(lSheet).parseError.reason, "The system cannot locate the object specified.") <> 0 Then
'                        TAX_Utilities_Srv_New.Data(lSheet).Load strOriginDataFileName
'                        If TAX_Utilities_Srv_New.Data(lSheet).parseError.reason <> vbNullString Then
'                            MsgBox TAX_Utilities_Srv_New.Data(lSheet).parseError.reason
'                        End If
'                    Else
'                        MsgBox TAX_Utilities_Srv_New.Data(lSheet).parseError.reason
'                    End If
'                End If
'
'                ' If load original data -> not fill
'                Set xmlNodeListCell = TAX_Utilities_Srv_New.Data(lSheet).getElementsByTagName("Cell")
'
'                For Each xmlNodeCell In xmlNodeListCell
'                    ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
'                    If GetAttribute(xmlNodeCell, "FirstCell") = "1" Then
'                        .MaxRows = .MaxRows + 1
'                        .InsertRows lRow, 1
'                        .CopyRowRange lRow - 1, lRow - 1, lRow
'                        '*************
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
    
    

    TAX_Utilities_Srv_New.xmlDataReDim (TAX_Utilities_Srv_New.NodeValidity.childNodes.length - 1)
    'TAX_Utilities_Srv_New.xmlDataReDim (TAX_Utilities_Srv_New.NodeValidity.childNodes.length)
    With pGrid
        .EventEnabled(EventAllEvents) = False
        For lSheet = 0 To TAX_Utilities_Srv_New.xmlDataCount
            'If GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "Active") <> "0" Then
                blnNewData = False
                .Sheet = lSheet + 1
                TAX_Utilities_Srv_New.Data(lSheet) = New MSXML.DOMDocument
                TAX_Utilities_Srv_New.Data(lSheet).resolveExternals = True
                TAX_Utilities_Srv_New.Data(lSheet).validateOnParse = True
                TAX_Utilities_Srv_New.Data(lSheet).async = False
                strOriginDataFileName = GetAbsolutePath(GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "TemplateFolder")) & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = vbNullString Or GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = "0" Then
                    strDataFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & ".xml"
                Else
                    If GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Month") = "1" Then
                        strDataFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_Srv_New.Month & TAX_Utilities_Srv_New.Year & ".xml"
                    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "ThreeMonth") = "1" Then
                        strDataFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_" & TAX_Utilities_Srv_New.ThreeMonths & TAX_Utilities_Srv_New.Year & ".xml"
                    ElseIf GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "Year") = "1" Then
                        strDataFileName = GetAbsolutePath("..\DataFiles\") & GetAttribute(TAX_Utilities_Srv_New.NodeValidity.childNodes(lSheet), "DataFile") & "_00" & TAX_Utilities_Srv_New.Year & ".xml"
                    End If
                End If
                TAX_Utilities_Srv_New.Data(lSheet).Load strDataFileName
                If TAX_Utilities_Srv_New.Data(lSheet).parseError.reason <> vbNullString Then
                    If InStr(1, TAX_Utilities_Srv_New.Data(lSheet).parseError.errorCode, "2146697210") <> 0 Then
                        'New data
                        blnNewData = True
                        TAX_Utilities_Srv_New.Data(lSheet).Load strOriginDataFileName
                        If TAX_Utilities_Srv_New.Data(lSheet).parseError.reason <> vbNullString Then
                            MsgBox TAX_Utilities_Srv_New.Data(lSheet).parseError.reason
                        End If
                    Else
                        MsgBox TAX_Utilities_Srv_New.Data(lSheet).parseError.reason
                    End If
                End If

                ' If load original data -> not fill
                Set xmlNodeListCell = TAX_Utilities_Srv_New.Data(lSheet).getElementsByTagName("Cell")

                For Each xmlNodeCell In xmlNodeListCell
                    ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
                    If GetAttribute(xmlNodeCell, "FirstCell") = "1" Then
                        lRows = GetDynRowCount(pGrid, xmlNodeCell.parentNode)
                        InsertRow pGrid, lRow, lRows
                    End If
                    .Col = lCol
                    .Row = lRow
                 If GetAttribute(xmlNodeCell, "Receive") <> "0" Then
                    Select Case .CellType
                        Case CellTypeCheckBox
                            ' Check box
                            If UCase(GetAttribute(xmlNodeCell, "Value")) = UCase("x") Or UCase(GetAttribute(xmlNodeCell, "Value")) = UCase("1") Then
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

                        Case CellTypePic
                            If blnNewData And .Text <> GetAttribute(xmlNodeCell, "Value") Then
                                SetAttribute xmlNodeCell, "Value", .Text
                            Else
                                .SetText lCol, lRow, GetAttribute(xmlNodeCell, "Value")
                            End If

                        Case Else
                            If blnNewData And .Value <> GetAttribute(xmlNodeCell, "Value") Then
                                SetAttribute xmlNodeCell, "Value", .Value
                            Else
                                .Value = GetAttribute(xmlNodeCell, "Value")
                            End If
                    End Select
                  Else
                    UpdateCellReceive pGrid, lSheet, .Col, .Row, .Value
                  End If
                    
                    .RowHeight(lRow) = 14
                    If .RowHeight(lRow) < .MaxTextRowHeight(lRow) Then
                        .RowHeight(lRow) = .MaxTextRowHeight(lRow)
                    End If
                Next

                Set xmlNodeCell = Nothing
                Set xmlNodeListCell = Nothing
        Next
        .EventEnabled(EventAllEvents) = True
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
        If arrActiveForm(i).ID = pID Then
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
    GetQuyHienTai.y = Year(dNgayHienTai)
    GetQuyHienTai.dNgayDauQuy = GetNgayDauQuy(GetQuyHienTai.q, GetQuyHienTai.y, dNgayTaiChinh, dThangTaiChinh)
    GetQuyHienTai.dNgayCuoiQuy = GetNgayCuoiQuy(GetQuyHienTai.q, GetQuyHienTai.y, dNgayTaiChinh, dThangTaiChinh)
End Function

Public Function GetNgayDauQuy(q As Integer, y As Integer, dNgayTaiChinh As Integer, dThangTaiChinh As Integer) As Date
    Dim intYear As Integer, intDay As Integer, intMonth As Integer
    
    If blnTinhTheoNamTaiChinh And (GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "FinanceYear") = "1") Then
        intDay = dNgayTaiChinh
        intMonth = (q - 1) * 3 + dThangTaiChinh
        intYear = y
        If intMonth > 12 Then
            intMonth = intMonth - 12
            intYear = y + 1
        End If
    Else
        intDay = 1
        intYear = y
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
Public Function GetNgayCuoiQuy(q As Integer, y As Integer, dNgayTaiChinh As Integer, dThangTaiChinh As Integer) As Date
    Dim mTaiChinhDau As Integer
    Dim mTaiChinhCuoi As Integer
    Dim yTaiChinhDau As Integer
    Dim yTaiChinhCuoi As Integer
    Dim iInterval As Integer
    
    mTaiChinhDau = (q - 1) * 3 + dThangTaiChinh + 2 'Thang cuoi quy
    If dNgayTaiChinh = 1 Then
        mTaiChinhCuoi = mTaiChinhDau + 1 'Thang dau quy sau
        yTaiChinhDau = y
        yTaiChinhCuoi = y
        If mTaiChinhDau > 12 Then
            mTaiChinhDau = mTaiChinhDau - 12
            yTaiChinhDau = y + 1
        End If
        If mTaiChinhCuoi > 12 Then
            mTaiChinhCuoi = mTaiChinhCuoi - 12
            yTaiChinhCuoi = y + 1
        End If
    
        iInterval = DateDiff("D", DateSerial(yTaiChinhDau, mTaiChinhDau, 1), DateSerial(yTaiChinhCuoi, mTaiChinhCuoi, 1)) - 1
        GetNgayCuoiQuy = DateSerial(yTaiChinhDau, mTaiChinhDau, 1) + iInterval
    Else
        GetNgayCuoiQuy = DateSerial(yTaiChinhDau, mTaiChinhDau, 1)
    End If
End Function

Public Function GetNgayDauNam(y As Integer, dThangTaiChinh As Integer, dNgayTaiChinh As Integer) As Date
    Dim intYear As Integer, intDay As Integer, intMonth As Integer
    
    If blnTinhTheoNamTaiChinh And (GetAttribute(TAX_Utilities_Srv_New.NodeMenu, "FinanceYear") = "1") Then
        intYear = y
        intMonth = dThangTaiChinh
        intDay = dNgayTaiChinh
    Else
        intDay = 1
        intYear = y
        intMonth = 1
    End If
    GetNgayDauNam = DateSerial(intYear, intMonth, intDay)
End Function

Function NgayCuoiNamTaiChinh(y As Integer, dThangTaiChinh As Integer, dNgayTaiChinh As Integer)
    Dim dNgayTC As Date
    
    dNgayTC = DateSerial(y, dThangTaiChinh, dNgayTaiChinh)
    NgayCuoiNamTaiChinh = DateAdd("M", 12, dNgayTC)
    NgayCuoiNamTaiChinh = DateAdd("D", -1, NgayCuoiNamTaiChinh)
    
End Function

Public Function changeMaToKhai(strID As String) As String
    If strID = "01" Then changeMaToKhai = "01_GTGT13"
    If strID = "02" Then changeMaToKhai = "02_GTGT13"
    If strID = "04" Then changeMaToKhai = "03_GTGT13"
    If strID = "07" Then changeMaToKhai = "01_TBAC_BLP"
    If strID = "11" Then changeMaToKhai = "01A_TNDN13"
    If strID = "12" Then changeMaToKhai = "01B_TNDN13"
    If strID = "03" Then changeMaToKhai = "03_TNDN14"
    If strID = "06" Then changeMaToKhai = "01_TAIN13"
    'If strID = "09" Then changeMaToKhai = "02_TAIN"
    If strID = "08" Then changeMaToKhai = "03_TAIN"
    If strID = "05" Then changeMaToKhai = "01_TTDB11"
    
    If strID = "73" Then changeMaToKhai = "02_TNDN14"
    
    If strID = "70" Then changeMaToKhai = "01_NTNN"
    If strID = "80" Then changeMaToKhai = "02_NTNN14"
    If strID = "81" Then changeMaToKhai = "03_NTNN11"
    If strID = "82" Then changeMaToKhai = "04_NTNN14"
    
    If strID = "71" Then changeMaToKhai = "04_GTGT13"
    If strID = "72" Then changeMaToKhai = "05_GTGT11"
    If strID = "86" Then changeMaToKhai = "01_BVMT11"
    If strID = "90" Then changeMaToKhai = "01_TBVMT13"
    If strID = "87" Then changeMaToKhai = "02_BVMT14"
    If strID = "77" Then changeMaToKhai = "02_TAIN14"
       
    If strID = "46" Then changeMaToKhai = "01A_TNCN_BH11"
    If strID = "47" Then changeMaToKhai = "01B_TNCN_BH11"
    If strID = "48" Then changeMaToKhai = "01A_TNCN_XS11"
    If strID = "49" Then changeMaToKhai = "01B_TNCN_XS11"
    
    If strID = "15" Then changeMaToKhai = "02A_TNCN11"
    If strID = "16" Then changeMaToKhai = "02B_TNCN11"
    If strID = "53" Then changeMaToKhai = "02A_TNCN"
    If strID = "37" Then changeMaToKhai = "02B_TNCN"
    
    If strID = "50" Then changeMaToKhai = "03A_TNCN13"
    If strID = "51" Then changeMaToKhai = "03B_TNCN13"
    If strID = "54" Then changeMaToKhai = "03A_TNCN"
    If strID = "38" Then changeMaToKhai = "03B_TNCN"
    
    If strID = "39" Then changeMaToKhai = "04A_TNCN"
    If strID = "40" Then changeMaToKhai = "04B_TNCN"
    
    If strID = "36" Then changeMaToKhai = "07_TNCN11"
    
    If strID = "74" Then changeMaToKhai = "08_TNCN11"
    If strID = "75" Then changeMaToKhai = "08A_TNCN11"
    
    If strID = "17" Then changeMaToKhai = "05_TNCN11"
    If strID = "59" Then changeMaToKhai = "06_TNCN14"
    If strID = "41" Then changeMaToKhai = "09_TNCN14"
    If strID = "76" Then changeMaToKhai = "08B_TNCN11"
    If strID = "42" Then changeMaToKhai = "02_TNCN_BH11"
    If strID = "43" Then changeMaToKhai = "02_TNCN_XS11"
    
    If strID = "18" Then changeMaToKhai = "BC26_AC_SL"
    If strID = "69" Then changeMaToKhai = "15_BCTC10"
    'update v3.2.0
    If strID = "19" Then changeMaToKhai = "48_BCTC13"
    If strID = "20" Then changeMaToKhai = "16_BCTC"
    If strID = "21" Then changeMaToKhai = "99_BCTC"
    
    
    ' nvhai
    ' Xu ly nhac cac BCTC in bang HTKK 2.1.0
    
'    If strID = "55" Then changeMaToKhai = "15_01_CDKT"
'    If strID = "56" Then changeMaToKhai = "15_02_SXKD"
    If strID = "57" Then changeMaToKhai = "15_03_LCTTTT"
    If strID = "58" Then changeMaToKhai = "15_04_LCTTGT"
    
    'If strID = "24" Then changeMaToKhai = "48_01_CDKT"
    If strID = "24" Then changeMaToKhai = "01_BCTL_DK"
    
    'If strID = "25" Then changeMaToKhai = "48_02_SXKD"
    If strID = "25" Then changeMaToKhai = "01_TNCN_BHDC13"
    
    'If strID = "26" Then changeMaToKhai = "48_03_LCTTTT"
    If strID = "27" Then changeMaToKhai = "01_BK_BC26_AC"
    If strID = "28" Then changeMaToKhai = "16_01_CDKT"
    If strID = "29" Then changeMaToKhai = "16_02_SXKD"
    If strID = "30" Then changeMaToKhai = "16_03_LCTTTT"
    If strID = "31" Then changeMaToKhai = "16_04_LCTTGT"
    If strID = "32" Then changeMaToKhai = "99_01_CDKT"
    If strID = "33" Then changeMaToKhai = "99_02_SXKD"
    If strID = "34" Then changeMaToKhai = "99_03_LCTTTT"
    If strID = "35" Then changeMaToKhai = "99_04_LCTTGT"
    
    ' Cac mau an chi
    If strID = "64" Then changeMaToKhai = "01_TBAC"
    If strID = "65" Then changeMaToKhai = "01_AC"
    If strID = "66" Then changeMaToKhai = "BC21_AC"
    If strID = "67" Then changeMaToKhai = "03_TBAC"
    If strID = "13" Then changeMaToKhai = "01_AC_BLP"
    If strID = "10" Then changeMaToKhai = "03_TBAC_BLP"
    If strID = "68" Then changeMaToKhai = "BC26_AC"
    If strID = "91" Then changeMaToKhai = "04_TBAC"
        ' Cac mau bien lai
    If strID = "07" Then changeMaToKhai = "01_TBAC_BLP"
    If strID = "13" Then changeMaToKhai = "01_AC_BLP"
    If strID = "09" Then changeMaToKhai = "BC21_AC_BLP"
    If strID = "10" Then changeMaToKhai = "03_TBAC_BLP"
    If strID = "14" Then changeMaToKhai = "BC26_AC_BLP" ' 05_TNDN=>BC26_AC_BLP
    
    'Mau moi V3.2.0
    If strID = "94" Then changeMaToKhai = "01_TD_GTGT13"
    If strID = "98" Then changeMaToKhai = "01A_TNDN_DK13"
    If strID = "22" Then changeMaToKhai = "95_BCTC"
    If strID = "95" Then changeMaToKhai = "03B_GTGT11"
    If strID = "23" Then changeMaToKhai = "01_TNCN_TTS"
    
    'To khai dau khi thuy dien
    If strID = "92" Then changeMaToKhai = "01_TAIN_DK"
    If strID = "98" Then changeMaToKhai = "01A_TNDN_DK"
    If strID = "99" Then changeMaToKhai = "01B_TNDN_DK"
    If strID = "96" Then changeMaToKhai = "03_TD_TAIN"
    If strID = "94" Then changeMaToKhai = "01_TD_GTGT"
    If strID = "93" Then changeMaToKhai = "02_TNDN_DK"
    If strID = "89" Then changeMaToKhai = "02_TAIN_DK"
    
    If strID = "85" Then changeMaToKhai = "01_PHLP"
    If strID = "88" Then changeMaToKhai = "02_PHLP"
    If strID = "97" Then changeMaToKhai = "03A_TD_TAIN"
    If strID = "26" Then changeMaToKhai = "02_TNCN_BHDC"
    
    If strID = "84" Then changeMaToKhai = "01_MBAI"
End Function

' Ham change sang ma cua QLT
Public Function changeMaToKhaiQLT(strID As String, isLanPS, LoaiKyKK) As String
    changeMaToKhaiQLT = ""
    
    ' To khai 01_GTGT
    ' Voi to khai GTGT LoaiKyKK As Boolean 'True la quy, false la thang
    If strID = "01" And LoaiKyKK = False Then
        changeMaToKhaiQLT = "14"
    ElseIf strID = "01" And LoaiKyKK = True Then
        changeMaToKhaiQLT = "83"
    End If
    
    ' To khai 02_GTGT
    If strID = "02" And LoaiKyKK = False Then
        changeMaToKhaiQLT = "68"
    ElseIf strID = "02" And LoaiKyKK = True Then
        changeMaToKhaiQLT = "84"
    End If
    
    'Khong chan cap to khai 03/GTGT va 04/GTGT cu & moi
    ' To khai 03_GTGT
    ' TODO cap nhat them ID cu
    If strID = "04" And LoaiKyKK = False Then
        changeMaToKhaiQLT = "02,A1,96,31" ' 02,03
    ElseIf strID = "04" And LoaiKyKK = True Then
        changeMaToKhaiQLT = "85,A2,97,88"
    End If
    
    'Khong chan cap to khai 03/GTGT va 04/GTGT cu & moi
    ' To khai 04_GTGT
'    If strID = "71" And isLanPS = True Then
'        changeMaToKhaiQLT = "98"
    If strID = "71" And LoaiKyKK = False Then
        changeMaToKhaiQLT = "96,31,02,A1"
    ElseIf strID = "71" And LoaiKyKK = True Then
        changeMaToKhaiQLT = "97,88,85,A2"
    End If
    
    ' To khai 05_GTGT
    If strID = "72" And isLanPS = True Then
        changeMaToKhaiQLT = "36"
    ElseIf strID = "72" And isLanPS = False Then
        changeMaToKhaiQLT = "32"
    End If
    
    ' To khai 01A_TNDN
    If strID = "11" Then changeMaToKhaiQLT = "37"
    
    ' To khai 01B_TNDN
    If strID = "12" Then changeMaToKhaiQLT = "26"
    
    ' To khai 02_TNDN
    If strID = "73" And isLanPS = True Then
        changeMaToKhaiQLT = "64"
    ElseIf strID = "73" And isLanPS = False Then
        changeMaToKhaiQLT = "67"
    End If
    
    ' To khai 01_TTDB
    If strID = "05" And isLanPS = True Then
        changeMaToKhaiQLT = "40"
    ElseIf strID = "05" And isLanPS = False Then
        changeMaToKhaiQLT = "25"
    End If
    
    ' To khai 01_TAIN
    If strID = "06" And isLanPS = True Then
        changeMaToKhaiQLT = "92"
    ElseIf strID = "06" And isLanPS = False Then
        changeMaToKhaiQLT = "24"
    End If

    ' To khai 01_NTNN
    If strID = "70" And isLanPS = True Then
        changeMaToKhaiQLT = "46"
    ElseIf strID = "70" And isLanPS = False Then
        changeMaToKhaiQLT = "27"
    End If
    
    ' To khai 03_NTNN
    If strID = "81" And isLanPS = True Then
        changeMaToKhaiQLT = "70"
    ElseIf strID = "81" And isLanPS = False Then
        changeMaToKhaiQLT = "69"
    End If
    
    ' To khai 01/TBVMT
    If strID = "90" And isLanPS = True Then
        changeMaToKhaiQLT = "93"
    ElseIf strID = "90" And isLanPS = False Then
        changeMaToKhaiQLT = "91"
    End If
    
    ' To khai 01A_TNCN_BH11
    If strID = "46" Then changeMaToKhaiQLT = "96"
    ' To khai 01B_TNCN_BH11
    If strID = "47" Then changeMaToKhaiQLT = "97"
    ' To khai 01A_TNCN_XS11
    If strID = "48" Then changeMaToKhaiQLT = "98"
    ' To khai 01B_TNCN_XS11
    If strID = "49" Then changeMaToKhaiQLT = "99"
    ' To khai 02A_TNCN11
    If strID = "15" Then changeMaToKhaiQLT = "29"
    ' To khai 02B_TNCN11
    If strID = "16" Then changeMaToKhaiQLT = "30"
    ' To khai 03A_TNCN11
    If strID = "50" Then changeMaToKhaiQLT = "21"
    ' To khai 03B_TNCN11
    If strID = "51" Then changeMaToKhaiQLT = "60"
    ' To khai 07_TNCN11
    If strID = "36" Then changeMaToKhaiQLT = "19"
    
    'TODO........
'    '--Cap nhat to khai thuy dien, dau khi 18/03/2010 ver 3.2.2
'    '--01A/TNDN-DK
'    If (strID = "98") Then changeMaToKhaiQLT = ""
'
'    '--01B/TNDN-DK
'    If (strID = "") Then changeMaToKhaiQLT = ""
'
'    '--01/TAIN-DK
'    If (strID = "") Then changeMaToKhaiQLT = ""
'
'    '--01/TD-GTGT
'    If (strID = "") Then changeMaToKhaiQLT = ""
'
'    '--03/TD-TAIN
'    If (strID = "") Then changeMaToKhaiQLT = ""
'
'    '--01/BCTL-DK
'    If (strID = "") Then changeMaToKhaiQLT = ""
End Function


''' UpdateCell description
''' Update cell value to DOM object when user change cell value
''' Parameter1 pCol     : active column
''' Parameter2 pRow     : active row
''' Parameter3 pValue   : cell value need update
Private Function UpdateCellReceive(fps As fpSpread, sSheet As Long, ByVal pCol As Long, ByVal pRow As Long, ByVal pValue As String) As Boolean
    On Error GoTo ErrHandle
    
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    GetCellSpan fps, pCol, pRow
    
    Set xmlNodeCell = TAX_Utilities_Srv_New.Data(sSheet).nodeFromID(GetCellID(fps, pCol, pRow))
    
    If GetAttribute(xmlNodeCell, "Value") <> pValue Then
        SetAttribute xmlNodeCell, "Value", pValue
        UpdateCellReceive = True
    End If
    
    Set xmlNodeCell = Nothing
    
    Exit Function
    
ErrHandle:
    SaveErrorLog "mdlFunction", "UpdateCellReceive", Err.Number, Err.Description
End Function


' Get ve ID cua bang data_pkg
Public Function GetDataPkgId() As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim pkgId As Variant
    Dim ID As Variant
    Dim noiLamViec As Variant
    Dim clsConn As New TAX_Utilities_Srv_New.clsADO
    If clsConn.Connected = False Then
        'clsConn.CreateConnectionString [MSDAORA.1], "QLT", "TKB", "TKB"
        clsConn.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
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
        sSQL = "select exc_data_pkg_seq.nextval prm_value from dual"
        Set rs = clsConn.Execute(sSQL)
        If Not rs Is Nothing Then
             ID = rs.Fields("prm_value")
        Else
             ID = 0
        End If
        pkgId = Trim(CStr(pkgId)) & Trim(CStr(ID))
        clsConn.Disconnect
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
    Dim clsConn As New TAX_Utilities_Srv_New.clsADO
    If clsConn.Connected = False Then
        'clsConn.CreateConnectionString [MSDAORA.1], "QLT", "TKB", "TKB"
        clsConn.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
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
    sSQL = "select EXC_TRAN_UP_SEQ.nextval prm_value from dual"
    Set rs = clsConn.Execute(sSQL)
    If Not rs Is Nothing Then
         tupId = rs.Fields("prm_value")
    Else
         tupId = ""
    End If
    tranNo = Trim(CStr(tranNo)) & Trim(CStr(tupId))
    clsConn.Disconnect
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
    Dim clsConn As New TAX_Utilities_Srv_New.clsADO
    If clsConn.Connected = False Then
        'clsConn.CreateConnectionString [MSDAORA.1], "QLT", "TKB", "TKB"
        clsConn.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
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
    sSQL = "select EXC_MESS_UP_SEQ.nextval prm_value from dual"
    Set rs = clsConn.Execute(sSQL)
    If Not rs Is Nothing Then
         seqMupId = rs.Fields("prm_value")
    Else
         seqMupId = ""
    End If
    mupId = Trim(CStr(mupId)) & Trim(CStr(seqMupId))
    clsConn.Disconnect
    GetMupId = mupId
End Function

' Get thong tin noi lam viec
Public Function GetNoiLamViec() As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim noiLamViec As Variant
    Dim clsConn As New TAX_Utilities_Srv_New.clsADO
    If clsConn.Connected = False Then
        'clsConn.CreateConnectionString [MSDAORA.1], "QLT", "TKB", "TKB"
        clsConn.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
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
    GetNoiLamViec = Trim(noiLamViec)
End Function

' Get thong tin noi lam viec
Public Function GetNoiNhan(str As String) As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim noiNhan As Variant
    Dim clsConn As New TAX_Utilities_Srv_New.clsADO
    If clsConn.Connected = False Then
        'clsConn.CreateConnectionString [MSDAORA.1], "QLT", "TKB", "TKB"
        clsConn.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
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
    GetNoiNhan = Trim(noiNhan)
End Function

' Get thong tin tns_code
Public Function GetTnsCode(str As String) As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileName As String
    Dim tnsCode As Variant
    Dim clsConn As New TAX_Utilities_Srv_New.clsADO
    If clsConn.Connected = False Then
        'clsConn.CreateConnectionString [MSDAORA.1], "QLT", "TKB", "TKB"
        clsConn.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
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
    Dim clsConn As New TAX_Utilities_Srv_New.clsADO
    pkgIDErr = "('')"
    If clsConn.Connected = False Then
        'clsConn.CreateConnectionString [MSDAORA.1], "QLT", "TKB", "TKB"
        clsConn.CreateConnectionString [MSDAORA.1], "QLT", strDBUserName, strDBPassword
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
    GetPkgIDErr = Trim(pkgIDErr)
End Function

Public Sub DataDM(ByVal ID As String, Optional ByRef TenTN As String)
Dim arrDanhsach() As String
Dim strDataFileName As String
Dim xmlDOMdata As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = "..\InterfaceTemplates\xml\Catalogue_Tinh_Thanh.xml"
    
       If xmlDOMdata.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDOMdata.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If ID = arrDanhsach(1) Then
                            TenTN = arrDanhsach(3)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub


