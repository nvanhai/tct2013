VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_04GTGT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Private Const S1TxtL13Row = 13
Private Const S1TxtL13Col = "L"
Private Const S1TxtL14Row = 14
Private Const S1TxtL14Col = "L"
Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
'Dim xNodeData0 As MSXML.IXMLDOMNode
Dim xNode As MSXML.IXMLDOMNode
Dim arrData() As String
Public loaiKHBS As String
Public Strloaitk As String
Public StrSolanBosung As String
Public FlagThaydoi As String

Private Const Row_MSTDN = 14
Private Const Col_MSTDN = "H"

Private Const colE = "E"

Public strMaNhom As String
Public strMaNhomAdd As String
Public lViTri As Long


' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

' pl 01-1_GTGT
'Private Const sheetPl01_1_GTGT = 2
Private Const sheetPl01_KHBS_GTGT = 2
Private Const rowStartKCTPL01_1 = 8
Private Const rowStartThue_0_PL01_1 = 13
Private Const rowStartThue_5_PL01_1 = 18
Private Const rowStartThue_10_PL01_1 = 23
Private Const rowStart_Khong_TH_PL01_1 = 28
Private Const colD = "D"
Private Const colF = "F"
Private Const colG = "G"
Private Const colH = "H"
Private Const colI = "I"
Private Const colJ = "J"
Private Const colK = "K"
Private Const colL = "L"
Private Const colM = "M"
Private Const colN = "N"
Private Const colO = "O"
Private Const colBG = "BG"
Private Const rowHidden53 = 53
Private Const rowHidden54 = 54

Public strTkThangQuy As String
Public dNgayDauQuy As Date
'Dinh nghia type luu du  Row
Public Type MY_ROW
    tenCT As String
    cot1 As Double ' Thue suat 0%
    cot2 As Double ' Thue suat 5%
    cot3 As Double ' Thue suat 10%
    Cot4 As Integer ' mas so
    Cot5 As Integer ' Vitri
End Type
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    With fps
        SetDateFormat fps, 1, 65, .ColLetterToNumber("T"), DDMMYYYY
        SetDateFormat fps, sheetPl01_KHBS_GTGT, 46, .ColLetterToNumber("BG"), DDMMYYYY
        SetDateFormat fps, sheetPl01_KHBS_GTGT, 31, .ColLetterToNumber("BE"), DDMMYYYY
        
        .Sheet = 1
        .Row = 65
        .Col = .ColLetterToNumber("T")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft

'        .Sheet = .SheetCount - 1 'KHBS
'        .Row = 46
'        .Col = .ColLetterToNumber("BG")
'        .Text = Format(Date, "dd/mm/yyyy")
'        .TypeHAlign = TypeHAlignLeft
        
        
        
        '###PL01_1
'        .Sheet = sheetPl01_1_GTGT 'PL01_1
'            .Row = 8
'            .Col = .ColLetterToNumber("C")
'            .Text = ""
'            .CellTag = .CellTag & "~1"
'            .Row = 13
'            .Text = ""
'            .CellTag = .CellTag & "~2"
'
'
'            .Row = 18
'            .Text = ""
'            .CellTag = .CellTag & "~3"
'
'
'            .Row = 23
'            .Text = ""
'            .CellTag = .CellTag & "~4"
'
'
'            .Row = 28
'            .Text = ""
'            .CellTag = .CellTag & "~5"
'
'
'        SetDateFormat fps, sheetPl01_1_GTGT, 8, .ColLetterToNumber(colG), DDMMYYYY
'        SetDateFormat fps, sheetPl01_1_GTGT, 13, .ColLetterToNumber(colG), DDMMYYYY
'        SetDateFormat fps, sheetPl01_1_GTGT, 18, .ColLetterToNumber(colG), DDMMYYYY
'        SetDateFormat fps, sheetPl01_1_GTGT, 23, .ColLetterToNumber(colG), DDMMYYYY
'        SetDateFormat fps, sheetPl01_1_GTGT, 28, .ColLetterToNumber(colG), DDMMYYYY
    End With
'    ' set ma hoa don cho PL 01-1, 01-2
'    setMaHoaDon
    ' set CQT ra quyet dinh hoan
    setCQTQuanLyHoanThue fps
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim i As Long
    Dim intCtrl As Integer
    Dim strId As Variant
    Dim xmlNode As MSXML.IXMLDOMNode
    With fps
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        'update MSTDN luôn bang MST tren phan TT chung
        UpdateMST fps, Col_MSTDN, Row_MSTDN
        
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("C")
            .Row = 37
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 37
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 37
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("C")
            .Row = 37
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 37
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 37
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        
        ' set loai tk thang/ quy
         If strTkThangQuy = "TK_THANG" Then
            .Col = .ColLetterToNumber("K")
            .Row = 37
            .Text = "0"
            UpdateCell fps, .Col, .Row, .Text
            
            .Col = .ColLetterToNumber("L")
            .Row = 37
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf strTkThangQuy = "TK_QUY" Then
            .Col = .ColLetterToNumber("K")
            .Row = 37
            .Text = "1"
            UpdateCell fps, .Col, .Row, .Text
            
            .Col = .ColLetterToNumber("L")
            .Row = 37
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf strTkThangQuy = "TK_LANPS" Then
            .Col = .ColLetterToNumber("K")
            .Row = 37
            .Text = "2"
            UpdateCell fps, .Col, .Row, .Text
            
            .Col = .ColLetterToNumber("L")
            .Row = 37
            .Text = TAX_Utilities_v1.Day & "/" & TAX_Utilities_v1.Month & "/" & TAX_Utilities_v1.Year
            UpdateCell fps, .Col, .Row, .Text
        End If
            'end
        
'        .Sheet = sheetPl01_1_GTGT 'PL1
'            .ActiveSheet = .Sheet
'                .Row = 1
'                .Col = .ColLetterToNumber("F")
'                .CellType = CellTypeButton
'                .TypeButtonTextColor = RGB(0, 0, 255)
'                .TypeButtonColor = &H8000000F
'                .TypeButtonAlign = TypeButtonAlignRight
'                .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
                
        fps.EventEnabled(EventAllEvents) = True
        
        mCurrentSheet = .Sheet
    End With
    
  '  CellChange fps.ColLetterToNumber(S1TxtL13Col), S1TxtL13Row
End Sub
Public Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
'Dim strId As Variant
'Dim iCol As Long, iRow As Long
'Dim intIndexCombo As Integer, i As Integer
'Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
'Dim itemp As Variant, temp As Variant
'Dim cValue As Variant
'
'Dim tempRow As MY_ROW
'Dim arrRow() As MY_ROW, arrRow2() As MY_ROW, arrRow3() As MY_ROW, arrRow4() As MY_ROW
'Dim tempValue As Variant
'Dim varTemp As Variant
'
'    With fps
'        If .ActiveSheet = 1 Then
'            '...
'        ElseIf .ActiveSheet = sheetPl01_1_GTGT Then
'        .Col = Col
'        .Row = Row
'        varTemp = .Text
'         If Col = .ColLetterToNumber("C") And .Text <> "" Then
'            .EventEnabled(EventChange) = False
'            .Col = .ColLetterToNumber(colD)
'            .Text = varTemp
'            UpdateCell fps, .Col, .Row, .Text
'            .EventEnabled(EventChange) = True
'         End If
'         End If
'     End With
End Sub

Private Sub fps_EditChange(ByVal Col As Long, ByVal Row As Long)
    'CellChange2 Col, Row
End Sub

Private Sub fps_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim i As Integer, iCol As Long, iRow As Long
     With fps
       If KeyCode = vbKeyF6 Then
            iCol = .ActiveCol
            iRow = .ActiveRow
            mblKeyF6 = True
            fps_Change iCol, iRow
            mblKeyF6 = False
        End If
     End With
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim i As Integer, iCol As Long, iRow As Long
'Dim countRowDyn As Long, coundtRowDel As Long
'Dim arrRow() As MY_ROW
'Dim tempRow As MY_ROW
'Dim tempValue As Variant, tempGroup As Variant, tempMaNhom As Variant
'Dim tCol As Long, tRow As Long
'    With fps
'
'        iCol = .ActiveCol
'        iRow = .ActiveRow
'        GetCellSpan fps, iCol, iRow
''       ' Neu group II hoac III, IV thi khong them dong khi bam F5 hoac F6
''       If Not isAddRow(iCol, iRow) Then
''            Exit Sub
''       End If
'       fps.EventEnabled(EventAllEvents) = False
'       If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
'            If .ActiveSheet = sheetPl01_1_GTGT Then
'                    mCurrentSheet = sheetPl01_1_GTGT
'                    i = 1
'                    .Col = .ColLetterToNumber("B")
'                    .Row = 8
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("B")
'                         .Row = i + 8
'                         i = i + 1
'                    Loop Until .Text = "aa"
'                    i = 1
'                   .Row = .Row + 4
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("L")
'                         .Text = "0%"
'                         UpdateCell fps, .Col, .Row, .Text
'                         .Col = .ColLetterToNumber("B")
'                         .Row = 1 + .Row
'                         i = i + 1
'                    Loop Until .Text = "bb"
'                    i = 1
'                   .Row = .Row + 4
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("L")
'                         .Text = "5%"
'                         UpdateCell fps, .Col, .Row, .Text
'                         .Col = .ColLetterToNumber("B")
'                         .Row = 1 + .Row
'                         i = i + 1
'                    Loop Until .Text = "cc"
'                    i = 1
'                   .Row = .Row + 4
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("L")
'                         .Text = "10%"
'                         UpdateCell fps, .Col, .Row, .Text
'                         .Col = .ColLetterToNumber("B")
'                         .Row = 1 + .Row
'                         i = i + 1
'                    Loop Until .Text = "dd"
'
'                                        i = 1
'                   .Row = .Row + 4
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("B")
'                         .Row = 1 + .Row
'                         i = i + 1
'                    Loop Until .Text = "ee"
'                    .SetActiveCell iCol, iRow
'                    .Row = iRow
'                    .Col = iCol
'
'                    .SetActiveCell iCol, iRow
'                    'CellChange2
'                    'CellChange3
'                    'LockCellBySheet
'            End If
'        End If
'        fps.EventEnabled(EventAllEvents) = True
'    End With

End Sub


Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim tCol As Long, tRow As Long
    Dim cCol As Long, cRow As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    With fps
        mCurrentSheet = .Sheet
        .EventEnabled(EventAllEvents) = False
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("T") And Row = 65 Then
                .GetText Col, Row, varTemp
                  If varTemp <> "" And varTemp <> "../../...." Then
                      If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                          .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                          .TypeHAlign = TypeHAlignLeft
                      Else
                               '.SetFocus
                               .SetActiveCell Col, Row
                      End If
                  Else
                   .SetText Col, Row, ""
                  End If
                 .Col = Col
                 .Row = Row
                 UpdateCell fps, .Col, .Row, .Text
            End If
        End If
        
'        If .ActiveSheet = sheetPl01_1_GTGT Then
'                   If Col = .ColLetterToNumber(colG) Then
'                       .GetText Col, Row, varTemp
'                       If varTemp <> "" And varTemp <> "../../...." Then
'                           If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
'                               .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                           Else
'                               .SetActiveCell Col, Row
'                           End If
'                       Else
'                        .SetText Col, Row, ""
'                       End If
'                       .Col = Col
'                       .Row = Row
'                      UpdateCell fps, .Col, .Row, .Text
'                   End If
'
'                   'chuyen cot ky hieu ve chu cai in hoa
'                   If Col = .ColLetterToNumber("E") Then
'                        .GetText Col, Row, varTemp
'                        .SetText Col, Row, UCase(varTemp)
'                        .Col = Col
'                        .Row = Row
'                        UpdateCell fps, .Col, .Row, .Text
'                   End If
'        End If
        
        'KHBS
'       If .ActiveSheet = .SheetCount - 1 Then
'         Dim count1 As Integer
'            .Row = 9
'            .Col = .ColLetterToNumber("B")
'            Do While .value <> "bb"
'                count1 = count1 + 1
'                .Row = .Row + 1
'            Loop
'                If Col = .ColLetterToNumber("BG") And Row = 23 + count1 - 5 Then
'                       .GetText Col, Row, varTemp
'                       If varTemp <> "" And varTemp <> "../../...." Then
'                           If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
'                               .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                               .TypeHAlign = TypeHAlignLeft
'                           Else
'                               .SetFocus
'                               .SetActiveCell Col, Row
'                           End If
'                       Else
'                        .SetText Col, Row, ""
'                       End If
'                      .Col = Col
'                      .Row = Row
'                      UpdateCell fps, .Col, .Row, .Text
'                   End If
'                .Sheet = .SheetCount - 1
'        End If
        
'        If .ActiveSheet = sheetPl01_1_GTGT Then
'                   If Col = .ColLetterToNumber(colE) Then
'                       .GetText Col, Row, varTemp
'                       If varTemp <> "" And varTemp <> "../../...." Then
'                           If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
'                               .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                           Else
'                               .SetActiveCell Col, Row
'                           End If
'                       Else
'                        .SetText Col, Row, ""
'                       End If
'                       .Col = Col
'                       .Row = Row
'                      UpdateCell fps, .Col, .Row, .Text
'                   End If
'        End If
        
      .EventEnabled(EventAllEvents) = True
    End With
    
    'CellChange Col, Row
End Sub


'Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
'    Dim vL13, vG6 As Variant
'    Dim strG6Previous As String
'    Dim varTemp As Variant
'    Dim temp As Variant
'    Dim cCol As Long, cRow As Long
'    With fps
'        .EventEnabled(EventAllEvents) = False
'        .Sheet = .ActiveSheet
'        mCurrentSheet = .ActiveSheet
'        GetCellSpan fps, Col, Row
'
'        If .ActiveSheet = 1 Then
'            If Row = 12 And Col = .ColLetterToNumber("P") Then
'                    .GetText Col, Row, temp
'                    If Val(temp) <= 0 Then
'                        .Col = .ColLetterToNumber("P")
'                        .Row = 13
'                        .value = 0
'                        .value = 0
'                        UpdateCell fps, .Col, .Row, .value
'                    Else
'                       .Col = .ColLetterToNumber("P")
'                       .Row = 13
'                       UpdateCell fps, .Col, .Row, .value
'                    End If
'            End If
'            If Row = 12 And Col = .ColLetterToNumber("M") Then
'                    .GetText Col, Row, temp
'                    If Val(temp) <= 0 Then
'                        .Col = .ColLetterToNumber("M")
'                        .Row = 13
'                        .value = 0
'                        .value = 0
'                        UpdateCell fps, .Col, .Row, .value
'                    Else
'                       .Col = .ColLetterToNumber("M")
'                       .Row = 13
'                       UpdateCell fps, .Col, .Row, .value
'                    End If
'            End If
'
'      End If
'
'      'KHBS
'        If .ActiveSheet = .SheetCount - 1 Then
'            mCurrentSheet = .Sheet
'
'        End If
'
'        .EventEnabled(EventAllEvents) = True
'    End With
'End Sub

'Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
'Dim strId As Variant
'Dim iCol As Long, iRow As Long
'Dim intIndexCombo As Integer, i As Integer
'Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
'Dim itemp As Variant, temp As Variant
'Dim cValue As Variant
'
'Dim intIndexCuc As Integer, intIndexChiCuc As Integer
'Dim varIdCucThue As Variant
'Dim xmlDomData As New MSXML.DOMDocument
'Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
'Dim xmlNode As MSXML.IXMLDOMNode
'Dim arrDanhsach() As String
'Dim tenCucThue As String
'Dim maCucThue As String
'Dim tenChiCucThue As String
'Dim maChiCucThue As String
'
'Dim loaiCq As String
'Dim maLoaiCq As String
'
'With fps
''   If .ActiveSheet = .SheetCount - 1 Then
''            'lay ve dong combobox
''            Set xmlCellNode = TAX_Utilities_v1.Data(.SheetCount - 2).nodeFromID(GetCellID(fps, .ColLetterToNumber("BI"), 5))
''            Set xmlCellNode = xmlCellNode.parentNode
''            Set xmlCellNode = xmlCellNode.parentNode
''            Set xmlCellNode = xmlCellNode.parentNode.lastChild
''            Set xmlCellNode = xmlCellNode.firstChild
''            Set xmlCellNode = xmlCellNode.firstChild
''            For i = 1 To 6
''                Set xmlCellNode = xmlCellNode.nextSibling
''            Next i
''            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
''
''
''            .Sheet = .SheetCount - 1
''            mCurrentSheet = .SheetCount - 1
''
''            .Col = Col
''            .Row = Row
''
''            If Col = .ColLetterToNumber("BE") And Row = iRow Then 'Khi chon Combo
''                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
''                intIndexCuc = .TypeComboBoxCurSel
''                .Col = .ColLetterToNumber("BI")
''                .Row = Row
''                .TypeComboBoxCurSel = intIndexCuc
''                'Nhap lai Combo cot Q
''                .Col = .ColLetterToNumber("BI")
''                UpdateCell fps, .Col, .Row, .Text
''                ' Lay Id cua cuc thue vua chon o tren
''                .GetText .Col, .Row, varIdCucThue
''
''                If varIdCucThue <> "" Or varIdCucThue <> vbNullString Then
''                    ' Lay thong tin cho chi cuc thue
''                    varIdCucThue = Left$(varIdCucThue, 3)
''
''                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
''                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
''                        For Each xmlNode In xmlNodeListCell
''                            If GetAttribute(xmlNode, "Value") <> "" Then
''                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
''
''                                If arrDanhsach(0) = "0" And arrDanhsach(2) = varIdCucThue Then
''                                    tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
''                                    maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
''                                End If
''                            End If
''                        Next
''                        Set xmlDomData = Nothing
''                        Set xmlNodeListCell = Nothing
''                        Set xmlNode = Nothing
''                    End If
''
''                    .Row = iRow + 2
''                    .Col = .ColLetterToNumber("BE")
''                    .TypeComboBoxList = tenChiCucThue
''
''                    .Row = iRow + 2
''                    .Col = .ColLetterToNumber("BI")
''                    .TypeComboBoxList = maChiCucThue
''                End If
''            End If
''
''            If Col = .ColLetterToNumber("BE") And Row = iRow + 2 Then 'Khi chon Combo
''                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
''                intIndexChiCuc = .TypeComboBoxCurSel
''                .Col = .ColLetterToNumber("BI")
''                .Row = Row
''                .TypeComboBoxCurSel = intIndexChiCuc
''                'Nhap lai Combo cot Q
''                .Col = .ColLetterToNumber("BI")
''                UpdateCell fps, .Col, .Row, .Text
''            End If
''    End If
'End With
'    TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
'End Sub


'Calling this function whenever we don't need this class anymore.
'No parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        
        CheckDynamicError 'Set Exception Error on cells of interface
            
        UpdateSheet
            
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

' ham lay ve gia tri quy
Private Function convertThang2Quy(ByVal vMonth As Integer) As Integer
    If vMonth >= 1 And vMonth <= 3 Then
        convertThang2Quy = 1
    ElseIf vMonth >= 4 And vMonth <= 6 Then
        convertThang2Quy = 2
    ElseIf vMonth >= 7 And vMonth <= 9 Then
        convertThang2Quy = 3
    Else
        convertThang2Quy = 4
    End If
    Exit Function
End Function

'Private Function CheckValidNgayHD(vdate As String) As Boolean
'
'   Dim dd As String, mm As String, yyyy As String, dDate As Date, dDateNow As Date
'   Dim sDateNow As String
'   On Error GoTo e
'    dd = Right("00" & Left(vdate, InStr(vdate, "/") - 1), 2)
'    mm = Right("00" & Mid(vdate, 4, 2), 2)
'    yyyy = Right("0000" & vdate, 4)
'    sDateNow = "01/" & TAX_Utilities_v1.Month & "/" & TAX_Utilities_v1.Year
'    dDateNow = Format(CDate(sDateNow), "dd/mm/yyyy")
'    dDate = Format(CDate(vdate), "dd/mm/yyyy")
'        If CInt(yyyy) > CInt(TAX_Utilities_v1.Year) Then
'            CheckValidNgayHD = False
'            Exit Function
'        End If
'        If DateAdd("m", 5, dDate) < dDateNow Then
'        CheckValidNgayHD = False
'            Exit Function
'        End If
'   CheckValidNgayHD = True
'   Exit Function
'
'e:
'    CheckValidNgayHD = False
' End Function
 
'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    Dim i As Long
    Dim vDoanhSo As Variant, vGTGT As Variant
    Dim vTarget15 As Variant, vTarget16 As Variant
    Dim vTarget25 As Variant, vTarget26 As Variant
    Dim vTarget17 As Variant, vTarget18 As Variant, vTarget19 As Variant, vTarget20 As Variant, vTarget21 As Variant, vTarget22 As Variant, vTarget23 As Variant, vTarget24 As Variant
    Dim vTarget38 As Variant, vTarget39 As Variant
    Dim vKyHieuHD As Variant, vSoHD As Variant, vNgayHD As Variant, vTennguoimua As Variant, vMST As Variant
    Dim vMaHoaDon As Variant, vKyHieuMauHoaDon As Variant
    
    Dim vErrMsgKyHieuHD As Variant, vErrMsgSoHD As Variant, vErrMsgNgayHD As Variant, vErrMsgTennguoimua As Variant, vErrMsgMST As Variant
    Dim vErrMsg As Variant
    Dim vErrKyHieuMauHoaDon As Variant
    
    Dim iFlag5 As Integer, iFlag6 As Integer
    Dim iFlagvKyHieuHD As Integer, iFlagvSoHD As Integer, iFlagvNgayHD As Integer, iFlagvTennguoimua As Integer, iFlagvMst As Integer, iFlagvMst_PL6 As Integer
    Dim iFlagvKyHieuMauHD As Integer
    Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    Dim iCurrentSheet As Integer, strCheck As String
    Dim iWhiteRow As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer
    Dim arrCellTag() As String
    Dim strCellTag As String
    Dim strFormula As String
    
    Dim flagTenCSPT As Integer
    Dim flagDiaChi As Integer
    Dim vErrorTenCSPT, vErrorDiaChi As Variant
    Dim blCheck_S2 As Boolean
    iFlag5 = 0
    iFlag6 = 0
    iFlagvKyHieuHD = 0
    iFlagvSoHD = 0
    iFlagvNgayHD = 0
    iFlagvTennguoimua = 0
    iFocusFlag = 0
    
    flagTenCSPT = 0
    flagDiaChi = 0
    Dim vErrorSoChungTu As Variant
    Dim vErrMsgThuesuat As Variant, vThuesuat As Variant
    Dim iFlagvThuesuat As Integer
    Dim arrDate() As String
    Dim vdate As Date
    Dim vNgayHt As Date
    'Dim vErrorSoChungTu As Variant
    Dim vErrorNgayNop As Variant
    
    Dim strTemp As String, result As String, strLoai As String
    Dim iFlagMauSoHD As Integer, iFlagSoLienHD As Integer
    Dim errMauSoHD As Variant, errSoLienHD As Variant
    Dim errEmty As Variant
    Dim iFlagEmty As Integer
    
    Dim vErrorNgayPHHD As Variant
    Dim iflagNgayPHHD As Integer
    
    With fps
    '.Visible = False
    '.ReDraw = False
    iCurrentSheet = .Sheet
     
     .Sheet = .SheetCount
     .GetText .ColLetterToNumber("E"), 17, vErrMsgKyHieuHD
     .GetText .ColLetterToNumber("E"), 18, vErrMsgSoHD
     .GetText .ColLetterToNumber("E"), 19, vErrMsgNgayHD
     .GetText .ColLetterToNumber("E"), 20, vErrMsgTennguoimua
     .GetText .ColLetterToNumber("E"), 21, vErrMsgMST
     
     .GetText .ColLetterToNumber("E"), 22, vErrorNgayPHHD
     .GetText .ColLetterToNumber("E"), 23, errMauSoHD
     .GetText .ColLetterToNumber("E"), 24, errSoLienHD
     .GetText .ColLetterToNumber("E"), 25, errEmty

    CheckErrorMST
    
'    fps.EventEnabled(EventAllEvents) = False
'    .Sheet = sheetPl01_1_GTGT
'    blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    If blCheck_S2 = True Then
'        For i = 8 To .MaxRows
'            iWhiteRow = 1
'            .Row = i
'            .Col = .ColLetterToNumber("C")
'            arrCellTag = Split(.CellTag, "~")
'            If UBound(arrCellTag) > 0 Then
'            strCellTag = arrCellTag(1)
'            '.GetText .ColLetterToNumber("C"), .Row, vMaHoaDon
'            '.GetText .ColLetterToNumber("D"), .Row, vKyHieuMauHoaDon
'            '.GetText .ColLetterToNumber("E"), .Row, vKyHieuHD
'            .GetText .ColLetterToNumber("F"), .Row, vSoHD
'            .GetText .ColLetterToNumber("G"), .Row, vNgayHD
'            .GetText .ColLetterToNumber("H"), .Row, vTennguoimua
'            .GetText .ColLetterToNumber("K"), .Row, vDoanhSo
'            .GetText .ColLetterToNumber("I"), .Row, vMST
'
'
'            If Trim(vKyHieuHD) <> "" Or Trim(vSoHD) <> "" Or Trim(vNgayHD) <> "" Or Trim(vTennguoimua) <> "" Or Trim(vDoanhSo) <> "0" Or Trim(vMaHoaDon) <> "" Or Trim(vKyHieuMauHoaDon) <> "" Then iWhiteRow = 0
'                Select Case strCellTag
'                    Case "1", "2", "3", "4", "5"
'                        ' kiem tra ky hieu mau HD
''                        If Trim(vKyHieuMauHoaDon) = "" And iWhiteRow = 0 Then
''                            If Trim$(vMaHoaDon) = "TT120" Then
''                                .Col = .ColLetterToNumber("D")
''                                .CellNote = ""  'static
''                                .BackColor = mNonErrorColor
''                            Else
''                                .Col = .ColLetterToNumber("D")
''                                .CellNote = errEmty  'static
''                                .BackColor = mErrorColor
''                                iFlagEmty = 1
''                            End If
''                        Else
''                            strLoai = getLoaiHD(CStr(vMaHoaDon))
''                            result = CheckMauSoHD(CStr(vKyHieuMauHoaDon), strLoai, CStr(vMaHoaDon))
''                            If result = "1" And iWhiteRow = 0 Then
''                                .Col = .ColLetterToNumber("D")
''                                .CellNote = errMauSoHD  'static
''                                .BackColor = mErrorColor
''                                iFlagMauSoHD = 1
''                            ElseIf result = "2" And iWhiteRow = 0 Then
''                                .Col = .ColLetterToNumber("D")
''                                .CellNote = errSoLienHD  'static
''                                .BackColor = mErrorColor
''                                iFlagSoLienHD = 1
''                            Else
''                                .Col = .ColLetterToNumber("D")
''                                .CellNote = ""  'static
''                                .BackColor = mNonErrorColor
''                            End If
''                        End If
'                        ' end
'
'
''                        If Trim(vKyHieuHD) = "" And iWhiteRow = 0 Then
''                            If Trim$(vMaHoaDon) = "TT120" Then
''                                .Col = .ColLetterToNumber("E")
''                                .CellNote = ""
''                                .BackColor = mNonErrorColor
''                            Else
''                                .Col = .ColLetterToNumber("E")
''                                .CellNote = vErrMsgKyHieuHD  'static
''                                .BackColor = mErrorColor
''                                iFlagvKyHieuHD = 1
''                            End If
''                        Else
''                            .Col = .ColLetterToNumber("E")
''                            .CellNote = ""
''                            .BackColor = mNonErrorColor
''                        End If
'
'                        If Trim(vSoHD) = "" And iWhiteRow = 0 Then
'                            .Col = .ColLetterToNumber("F")
'                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                            .CellNote = vErrMsgSoHD  'static
'                            .BackColor = mErrorColor
'                            iFlagvSoHD = 1
'                            'capture the first error cell
'                            If iFocusFlag = 0 Then
'                                strFocusSheetName = "'" & .SheetName & "'"
'                                strFocusCol = str(.Col)
'                                strFocusRow = str(.Row)
'                                iFocusFlag = 1
'                            End If
'                        Else
'                            .Col = .ColLetterToNumber("F")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
'                        End If
'
'                        If Trim(vNgayHD) = "" And iWhiteRow = 0 Then
'                            .Col = .ColLetterToNumber("G")
'                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                            .CellNote = vErrMsgNgayHD  'static
'                            .BackColor = mErrorColor
'                            iFlagvNgayHD = 1
''                            'capture the first error cell
'                            If iFocusFlag = 0 Then
'                                strFocusSheetName = "'" & .SheetName & "'"
'                                strFocusCol = str(.Col)
'                                strFocusRow = str(.Row)
'                                iFocusFlag = 1
'                            End If
'
'                         ElseIf Trim(vNgayHD) <> "" And iWhiteRow = 0 Then
''                            ' To khai quy
''                            If strTkThangQuy = "TK_QUY" Then
''                                vNgayHt = Date
''                                arrDate = Split(vNgayHD, "/")
''                                vdate = DateSerial(CInt(arrDate(2)), CInt(arrDate(1)), CInt(arrDate(0)))
''                                If CInt(Right(vNgayHD, 4)) > CInt(TAX_Utilities_v1.Year) Then
''                                    .Col = .ColLetterToNumber("G")
''                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                    .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
''                                    .BackColor = mErrorColor
''                                    iFlagvNgayHD = 1
''                                    If iFocusFlag = 0 Then
''                                        strFocusSheetName = "'" & .SheetName & "'"
''                                        strFocusCol = str(.Col)
''                                        strFocusRow = str(.Row)
''                                        iFocusFlag = 1
''                                    End If
''                                 ElseIf CInt(Right(vNgayHD, 4)) = CInt(TAX_Utilities_v1.Year) And convertThang2Quy(CInt(Right("00" & Mid(vNgayHD, 4, 2), 2))) > CInt(TAX_Utilities_v1.ThreeMonths) Then
''                                    .Col = .ColLetterToNumber("G")
''                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                    .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
''                                    .BackColor = mErrorColor
''                                    iFlagvNgayHD = 1
''                                    If iFocusFlag = 0 Then
''                                        strFocusSheetName = "'" & .SheetName & "'"
''                                        strFocusCol = str(.Col)
''                                        strFocusRow = str(.Row)
''                                        iFocusFlag = 1
''                                    End If
''                                 ElseIf DateDiff("D", vNgayHt, vdate) > 0 Then
''                                    .Col = .ColLetterToNumber("G")
''                                    .CellNote = GetAttribute(GetMessageCellById("0273"), "Msg") 'static
''                                    .BackColor = mErrorColor
''                                    iFlagvNgayHD = 1
''                                     If iFocusFlag = 0 Then
''                                        strFocusSheetName = "'" & .SheetName & "'"
''                                        strFocusCol = str(.Col)
''                                        strFocusRow = str(.Row)
''                                        iFocusFlag = 1
''                                      End If
''                                 Else
''                                    .Col = .ColLetterToNumber("G")
''                                    .CellNote = ""
''                                    .BackColor = mNonErrorColor
''                                End If
''                            Else
''                                If CInt(Right(vNgayHD, 4)) > CInt(TAX_Utilities_v1.Year) Then
''                                    .Col = .ColLetterToNumber("G")
''                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                    .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
''                                    .BackColor = mErrorColor
''                                    iFlagvNgayHD = 1
'''                                    If iFocusFlag = 0 Then
'''                                        strFocusSheetName = "'" & .SheetName & "'"
'''                                        strFocusCol = str(.Col)
'''                                        strFocusRow = str(.Row)
'''                                        iFocusFlag = 1
'''                                    End If
''                                 ElseIf CInt(Right(vNgayHD, 4)) = CInt(TAX_Utilities_v1.Year) And CInt(Right("00" & Mid(vNgayHD, 4, 2), 2)) > CInt(TAX_Utilities_v1.Month) Then
''                                    .Col = .ColLetterToNumber("G")
''                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                    .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
''                                    .BackColor = mErrorColor
''                                    iFlagvNgayHD = 1
'''                                    If iFocusFlag = 0 Then
'''                                        strFocusSheetName = "'" & .SheetName & "'"
'''                                        strFocusCol = str(.Col)
'''                                        strFocusRow = str(.Row)
'''                                        iFocusFlag = 1
'''                                    End If
''                                 ElseIf (CheckValidNgayHD(CStr(vNgayHD)) = False) Then
''                                    .Col = .ColLetterToNumber("G")
''                                    .CellNote = GetAttribute(GetMessageCellById("0124"), "Msg") 'static
''                                    .BackColor = mAlertColor
''                                    iFlagvNgayHD = 0
'''                                     If iFocusFlag = 0 Then
'''                                        strFocusSheetName = "'" & .SheetName & "'"
'''                                        strFocusCol = str(.Col)
'''                                        strFocusRow = str(.Row)
'''                                        iFocusFlag = 1
'''                                      End If
''                                 Else
''                                    .Col = .ColLetterToNumber("G")
''                                    .CellNote = ""
''                                    .BackColor = mNonErrorColor
''                                End If
''                            End If
'                            If strTkThangQuy = "TK_QUY" Then
'                                 If CheckNgayInKyKK(CStr(vNgayHD), TAX_Utilities_v1.ThreeMonths, TAX_Utilities_v1.Year, True) = False Then
'                                        .Col = .ColLetterToNumber("G")
'                                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                        .CellNote = vErrorNgayPHHD   'static
'                                        .BackColor = mErrorColor
'                                        iflagNgayPHHD = 1
'                                        'capture the first error cell
''                                        If iFocusFlag = 0 Then
''                                            strFocusSheetName = "'" & .SheetName & "'"
''                                            strFocusCol = str(.Col)
''                                            strFocusRow = str(.Row)
''                                            iFocusFlag = 1
''                                        End If
'                                    Else
'                                         .Col = .ColLetterToNumber("G")
'                                         .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                         .CellNote = ""   'static
'                                         .BackColor = mNonErrorColor
'                                   End If
'                             ElseIf strTkThangQuy = "TK_LANPS" Then
'                                   arrDate = Split(vNgayHD, "/")
'                                   vdate = DateSerial(CInt(arrDate(2)), CInt(arrDate(1)), CInt(arrDate(0)))
'                                   If DateDiff("D", vdate, DateSerial(CInt(TAX_Utilities_v1.Year), CInt(TAX_Utilities_v1.Month), CInt(TAX_Utilities_v1.Day))) <> 0 Then
'                                        .Col = .ColLetterToNumber("G")
'                                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                        .CellNote = vErrorNgayPHHD   'static
'                                        .BackColor = mErrorColor
'                                        iflagNgayPHHD = 1
'                                        'capture the first error cell
''                                        If iFocusFlag = 0 Then
''                                            strFocusSheetName = "'" & .SheetName & "'"
''                                            strFocusCol = str(.Col)
''                                            strFocusRow = str(.Row)
''                                            iFocusFlag = 1
''                                        End If
'                                    Else
'                                         .Col = .ColLetterToNumber("G")
'                                         .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                         .CellNote = ""   'static
'                                         .BackColor = mNonErrorColor
'                                   End If
'                             Else
'                                   If CheckNgayInKyKK(CStr(vNgayHD), TAX_Utilities_v1.Month, TAX_Utilities_v1.Year, False) = False Then
'                                        .Col = .ColLetterToNumber("G")
'                                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                        .CellNote = vErrorNgayPHHD   'static
'                                        .BackColor = mErrorColor
'                                        iflagNgayPHHD = 1
'                                        'capture the first error cell
''                                        If iFocusFlag = 0 Then
''                                            strFocusSheetName = "'" & .SheetName & "'"
''                                            strFocusCol = str(.Col)
''                                            strFocusRow = str(.Row)
''                                            iFocusFlag = 1
''                                        End If
'                                    Else
'                                         .Col = .ColLetterToNumber("G")
'                                         .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                         .CellNote = ""   'static
'                                         .BackColor = mNonErrorColor
'                                   End If
'                                End If
'
'                        Else
'                                .Col = .ColLetterToNumber("G")
'                                .CellNote = ""
'                                .BackColor = mNonErrorColor
'                        End If
'
'                        If Trim(vTennguoimua) = "" And iWhiteRow = 0 Then
'                            .Col = .ColLetterToNumber("H")
'                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                            .CellNote = vErrMsgTennguoimua  'static
'                            .BackColor = mErrorColor
'                            iFlagvTennguoimua = 1
'
''                            'capture the first error cell
'                            If iFocusFlag = 0 Then
'                                strFocusSheetName = "'" & .SheetName & "'"
'                                strFocusCol = str(.Col)
'                                strFocusRow = str(.Row)
'                                iFocusFlag = 1
'                            End If
'                        Else
'                            .Col = .ColLetterToNumber("H")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
'                        End If
'          ' Check MST
'                        If Trim(vMST) <> vbNullString Then
''                           vMST = Replace(vMST, " ", "")
''                           vMST = Replace(vMST, "-", "")
''                            'vMST = Left(vMST, 10)
'
'                           strCheck = ""
'
'                           If Len(vMST) = 14 And InStr(vMST, "-") > 0 Then
'                               vMST = Left$(vMST, 10) & Right$(vMST, 3)
'                            End If
'                            vMST = Replace(vMST, "-", "")
'
'                           If Len(vMST) = 10 Or Len(vMST) = 13 Then
'                                strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
'                           End If
'
'                           If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
'                                .Col = .ColLetterToNumber("I")
'                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                .CellNote = vErrMsgMST   'static
'                                .BackColor = mErrorColor
'                                iFlagvMst = 1
'
''                                'capture the first error cell
'                                If iFocusFlag = 0 Then
'                                    strFocusSheetName = "'" & .SheetName & "'"
'                                    strFocusCol = str(.Col)
'                                    strFocusRow = str(.Row)
'                                    iFocusFlag = 1
'                                End If
'                           Else
'                                .Col = .ColLetterToNumber("I")
'                                .CellNote = ""
'                                .BackColor = mNonErrorColor
'                            End If
'                        Else
'                            .Col = .ColLetterToNumber("I")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
'                        End If
'                End Select
'           End If
'        Next
'     End If
     
    .Sheet = .SheetCount
    .SetText .ColLetterToNumber("B"), 17, IIf(iFlagvKyHieuHD = 1 And blCheck_S2 = True, "0", "1")
    .SetText .ColLetterToNumber("B"), 18, IIf(iFlagvSoHD = 1 And blCheck_S2 = True, "0", "1")
    .SetText .ColLetterToNumber("B"), 19, IIf(iFlagvNgayHD = 1 And blCheck_S2 = True, "0", "1")
    .SetText .ColLetterToNumber("B"), 20, IIf(iFlagvTennguoimua = 1 And blCheck_S2 = True, "0", "1")
    .SetText .ColLetterToNumber("B"), 21, IIf(iFlagvMst = 1 And blCheck_S2 = True, "0", "1")
    
    .SetText .ColLetterToNumber("B"), 22, IIf(iflagNgayPHHD = 1 And blCheck_S2 = True, "0", "1")
    .SetText .ColLetterToNumber("B"), 23, IIf(iFlagMauSoHD = 1 And blCheck_S2 = True, "0", "1")
    .SetText .ColLetterToNumber("B"), 24, IIf(iFlagSoLienHD = 1 And blCheck_S2 = True, "0", "1")
    .SetText .ColLetterToNumber("B"), 25, IIf(iFlagEmty = 1 And blCheck_S2 = True, "0", "1")
    
    .Sheet = .SheetCount
    .GetText .ColLetterToNumber("B"), 17, vTarget21
    .GetText .ColLetterToNumber("B"), 18, vTarget22
    .GetText .ColLetterToNumber("B"), 19, vTarget23
    .GetText .ColLetterToNumber("B"), 20, vTarget24
    .GetText .ColLetterToNumber("B"), 21, vTarget25
    
'    If blCheck_S2 Then
'        .SetText .ColLetterToNumber("B"), 17, IIf(vTarget21 = "1", "1", "0")
'        .SetText .ColLetterToNumber("B"), 18, IIf(vTarget22 = "1", "1", "0")
'        .SetText .ColLetterToNumber("B"), 19, IIf(vTarget23 = "1", "1", "0")
'        .SetText .ColLetterToNumber("B"), 20, IIf(vTarget24 = "1", "1", "0")
'        .SetText .ColLetterToNumber("B"), 21, IIf(vTarget25 = "1", "1", "0")
'     Else
'        .SetText .ColLetterToNumber("B"), 17, "1"
'        .SetText .ColLetterToNumber("B"), 18, "1"
'        .SetText .ColLetterToNumber("B"), 19, "1"
'        .SetText .ColLetterToNumber("B"), 20, "1"
'        .SetText .ColLetterToNumber("B"), 21, "1"
'        .SetText .ColLetterToNumber("B"), 22, "1"
'        .SetText .ColLetterToNumber("B"), 23, "1"
'        .SetText .ColLetterToNumber("B"), 24, "1"
'        .SetText .ColLetterToNumber("B"), 25, "1"
'     End If
    .Sheet = iCurrentSheet
    End With
    
End Sub
Sub CheckErrorMST()
    Dim vError1       As Variant, vError2 As Variant, vError3 As Variant
    Dim MST1          As Variant, MST2 As Variant, MST3 As Variant, MST4 As Variant
    Dim MST5          As Variant, MST6 As Variant, MST7 As Variant, MST8 As Variant
    Dim MST9          As Variant, MST10 As Variant, MST11 As Variant, MST12 As Variant, MST13 As Variant
    Dim iCurrentSheet As Integer, strCheck As String
    Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer

    With fps
        
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 12, vError1
        .GetText .ColLetterToNumber("E"), 13, vError2
        .GetText .ColLetterToNumber("E"), 14, vError3
        
        .GetText .ColLetterToNumber(SxMST1Col), SxMST1Row, MST1
        .GetText .ColLetterToNumber(SxMST2Col), SxMST2Row, MST2
        .GetText .ColLetterToNumber(SxMST3Col), SxMST3Row, MST3
        .GetText .ColLetterToNumber(SxMST4Col), SxMST4Row, MST4
        .GetText .ColLetterToNumber(SxMST5Col), SxMST5Row, MST5
        .GetText .ColLetterToNumber(SxMST6Col), SxMST6Row, MST6
        .GetText .ColLetterToNumber(SxMST7Col), SxMST7Row, MST7
        .GetText .ColLetterToNumber(SxMST8Col), SxMST8Row, MST8
        .GetText .ColLetterToNumber(SxMST9Col), SxMST9Row, MST9
        .GetText .ColLetterToNumber(SxMST10Col), SxMST10Row, MST10
        .GetText .ColLetterToNumber(SxMST11Col), SxMST11Row, MST11
        .GetText .ColLetterToNumber(SxMST12Col), SxMST12Row, MST12
        .GetText .ColLetterToNumber(SxMST13Col), SxMST13Row, MST13
        
        strCheck = CheckTaxCode(MST1, MST2, MST3, MST4, MST5, MST6, MST7, MST8, MST9, MST10, MST11, MST12, MST13)

        iFlagTaxCode1 = CInt(Left(strCheck, 1))
        strCheck = Right(strCheck, Len(strCheck) - 1)

        iFlagTaxCode2 = CInt(Left(strCheck, 1))
        strCheck = Right(strCheck, Len(strCheck) - 1)

        iFlagTaxCode3 = CInt(strCheck)
        
        If iFlagTaxCode1 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 12
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 12
            .Text = "1"
        End If
        
        If iFlagTaxCode2 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 13
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 13
            .Text = "1"
        End If
        
        If iFlagTaxCode3 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 14
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 14
            .Text = "1"
        End If
        
        .Sheet = 1
        .Col = .ColLetterToNumber("F")
        .Row = 5
        .CellNote = ""
        .BackColor = mFormColor

        If iFlagTaxCode1 = 1 Then
            .CellNote = .CellNote & "> " & vError1
            .BackColor = mErrorColor
        End If

        If iFlagTaxCode2 = 1 Then
            .CellNote = .CellNote & IIf(Trim(.CellNote) = "", "", vbCrLf) & "> " & vError2
            .BackColor = mErrorColor
        End If

        If iFlagTaxCode3 = 1 Then
            .CellNote = .CellNote & IIf(Trim(.CellNote) = "", "", vbCrLf) & "> " & vError3
            .BackColor = mErrorColor
        End If

        .Sheet = iCurrentSheet
    End With

End Sub

Public Sub SetActiveSheet() 'ByRef xmlNodeValid As MSXML.IXMLDOMNode)
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    Dim strValue3 As String
    Dim strValue4 As Variant
    
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
'            Case "01_1"
'                blnNullValue = True
'                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 2)
'                    strValue2 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 3)
'                    strValue3 = GetAttribute(xmlCellNode, "Value")
'                'Set active voi dong 5
'
'                fps.Sheet = sheetPl01_1_GTGT
'                fps.GetText fps.ColLetterToNumber("G"), 8, strValue4
'
'                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Or Val(strValue4) <> 0 Then
'                            blnNullValue = False
'                    End If
'
'                    If blnNullValue Then
'                        SetAttribute xmlNodeValid, "Active", "0"
'                    End If
             Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell").length - 4)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
                            blnNullValue = False
                    End If

                    If blnNullValue Then
                        'SetAttribute xmlNodeValid, "Active", "0"
                        'FlagThaydoi = "1"
                    End If
        End Select
    Next
    
    Set xmlCellNode = Nothing
    Set xmlNodeValid = Nothing
End Sub
Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean
    Dim totalCell  As Long, totalCell1  As Long
    Dim countRowReset As Long, countRowReset1 As Long
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    totalCell = TAX_Utilities_v1.Data(0).getElementsByTagName("Cell").length
    totalCell1 = TAX_Utilities_v1.Data(fps.SheetCount - 2).getElementsByTagName("Cell").length
    countRowReset = (totalCell - 26) / 6
    countRowReset1 = (totalCell1 - 8) / 6
    With fps
        For Each xmlNodeReset In TAX_Utilities_v1.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            .Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            If ((lRow < 43 And lRow <> 25) Or lRow > 62 + countRowReset) And mCurrentSheet = 1 Then
            ElseIf (lRow >= 20 + countRowReset1) And mCurrentSheet = (fps.SheetCount - 1) Then
            Else
                .Col = lCol
                .Row = lRow
                
                Select Case .CellType
                    Case CellTypeCheckBox
                        .Text = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                    Case CellTypeComboBox
                        .Text = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                    Case CellTypeNumber
                        .value = 0
                        UpdateCell fps, lCol, lRow, "0"
                    Case Else
                        .value = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                End Select
                'mAdjustData = IIf(IsUpdate = True, IsUpdate, mAdjustData)
                TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
            End If
        Next
    End With
End Function

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()
    Dim i As Integer
    With fps
       ' Set so thu tu
        ' group 2
         fps.EventEnabled(EventAllEvents) = False
'         ' Group 1
'         If mCurrentSheet = 1 Then
'             .Sheet = 1
'             i = 1
'             .Col = .ColLetterToNumber("B")
'             .Row = 45
'             Do
'                  .Text = "1." & Trim(str(i))
'                  .Col = .ColLetterToNumber("B")
'                  .Row = i + 45
'                  i = i + 1
'             Loop Until .Text = "aa"
'             i = 1
'            .Row = .Row + 3
'             Do
'                  .Text = "2." & Trim(str(i))
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "bb"
'
'             ' Group 3
'             i = 1
'            .Row = .Row + 3
'             Do
'                  .Text = "3." & Trim(str(i))
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "cc"
'             ' Group 4
'             i = 1
'            .Row = .Row + 3
'             Do
'                  .Text = "4." & Trim(str(i))
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "dd"
'        End If
        fps.EventEnabled(EventAllEvents) = True
        
        mCurrentSheet = .Sheet
    End With
End Sub

Public Sub Prepare_KHBS()
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    Dim strPath As String
    Dim strComboName As String
    Dim strComboMCT As String
    Dim strCombo As String
    
    mblKhoiDauToKhai = True
    fps.Sheet = fps.SheetCount - 1
    fps.ActiveSheet = fps.Sheet
    
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT.xml"
    
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "ID") = GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") Then
                Exit For
            End If
        Next
    End If
    
    
    Set xmlNodeListItem = xmlNode.childNodes
    
    For Each xmlNodeItem In xmlNodeListItem
                strCombo = strCombo + CPab(GetAttribute(xmlNodeItem, "Caption"), 100) + CPab(GetAttribute(xmlNodeItem, "CellID"), 20) + GetAttribute(xmlNodeItem, "MCT") + Chr$(9)
                strComboName = strComboName + GetAttribute(xmlNodeItem, "Caption") + Chr$(9)
                strComboMCT = strComboMCT + GetAttribute(xmlNodeItem, "MCT") + Chr$(9)
    Next
     With fps
        
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        .Col = .ColLetterToNumber("BC")
        .Row = 9
        .TypeComboBoxList = strCombo
        .Row = 13
        .TypeComboBoxList = strCombo
        .Col = .ColLetterToNumber("BD")
        .Row = 9
        .TypeComboBoxList = strComboName
        .Row = 13
        .TypeComboBoxList = strComboName
        .Col = .ColLetterToNumber("BE")
        .Row = 9
        .TypeComboBoxList = strComboMCT
        .Row = 13
        .TypeComboBoxList = strComboMCT
    End With
    
End Sub

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

Public Sub UpdateChangeKHBS()
    Dim varTemp As Variant
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    With fps
        .Sheet = 1
        For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
            ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
            .GetText lCol, lRow, varTemp
            UpdateLastKHBSCell fps, lCol, lRow, varTemp
        Next
        .Sheet = mCurrentSheet
    End With
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

' Insert them 1 row
Private Sub InsertNode(ByVal pCol As Long, ByVal pRow As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCell As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    Dim lLRowBound As Long, lURowBound As Long
    Dim lRow2s As Long, lRows As Long
    
    ' Get cellspan for merge cell on interface templates
    GetCellSpan fps, pCol, pRow
    
    Set xmlNodeCells = TAX_Utilities_v1.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fps, pCol, pRow)).parentNode
    
    lRows = GetDynRowCount(fps, xmlNodeCells, lRow2s, lLRowBound, lURowBound)
    
    'If Not xmlNodeCells.nextSibling Is Nothing Then GoTo EXIT_SUB
    If GetAttribute(xmlNodeCells.parentNode, "Dynamic") <> "1" Then GoTo EXIT_SUB
    If Val(GetAttribute(xmlNodeCells.parentNode, "MaxRows")) = xmlNodeCells.parentNode.childNodes.length Then GoTo EXIT_SUB
    
    ' insert new row on grid
    InsertRow fps, lURowBound + 1, lRows
    'fpSpread1.SetFocus
    
    ' increase value of row in xmlDocument
    IncreaseRowInDOM fps, TAX_Utilities_v1.Data(mCurrentSheet - 1), lURowBound + 1, lRows, lRow2s
    'IncreaseRowInDOM lURowBound + 1, lRows, lRow2s

    Set xmlNodeNewCells = xmlNodeCells.cloneNode(True)
    For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
        ' Set new ID for node (CellID)
        ParserCellID fps, GetAttribute(xmlNodeNewCell, "CellID"), lCol, lRow
        SetAttribute xmlNodeNewCell, "CellID", GetCellID(fps, lCol, lRow + lRows)
                
        ' Set first cell = 1
        SetAttribute xmlNodeNewCell, "FirstCell", "1"
        
        ' Reset empty value for new node
        fps.Col = lCol
        fps.Row = lRow
        Select Case fps.CellType
            Case CellTypeNumber
                SetAttribute xmlNodeNewCell, "Value", "0"
            Case Else
                SetAttribute xmlNodeNewCell, "Value", vbNullString
        End Select
        
        ' Set new ID for node (CellID2)
        ParserCellID fps, GetAttribute(xmlNodeNewCell, "CellID2"), lCol, lRow
        SetAttribute xmlNodeNewCell, "CellID2", GetCellID(fps, lCol, lRow + lRow2s)
    Next
    
    ' Insert new node to DOM object
    If Not xmlNodeCells.nextSibling Is Nothing Then
        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, xmlNodeCells.nextSibling
    Else
        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, Null
    End If
'   **********************************
'    added
'   Date: 12/04/06
    'mAdjustData = True
    TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
'   **********************************
EXIT_SUB:
    Set xmlNodeNewCell = Nothing
    Set xmlNodeNewCells = Nothing
    Set xmlNodeCells = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "04_GTGT", "InsertNode", Err.number, Err.Description
End Sub


Private Function GetDynRowCount(pGrid As fpSpread, xmlNodeCells As MSXML.IXMLDOMNode, Optional ByRef lReportRows As Long, Optional ByRef lMinRow As Long, Optional lMaxRow As Long)
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


Private Sub InsertRow(fpSpread1 As fpSpread, ByVal pRow As Long, lRows As Long, Optional blnFillingData As Boolean = False)
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
                        If Not TAX_Utilities_v1.Data(mCurrentSheet - 1).nodeFromID( _
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
                        If Not TAX_Utilities_v1.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, i, pRow - lRowCtrl)) Is Nothing Then
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
    SaveErrorLog "04_GTGT", "InsertRow", Err.number, Err.Description
End Sub

Private Sub IncreaseRowInDOM(fpSpread1 As fpSpread, xmlDomData As MSXML.DOMDocument, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
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
    SaveErrorLog "04_GTGT", "IncreaseRowInDOM", Err.number, Err.Description
End Sub


Private Sub DeleteNode(ByVal intSheet As Integer, ByVal pCol As Long, ByVal pRow As Long, Optional ByVal blnForce As Boolean = True)
    On Error GoTo ErrorHandle
    Dim lLRowBound As Long, lURowBound As Long
    Dim lRow2s As Long, lRows As Long
    Dim xmlNodeCells As MSXML.IXMLDOMNode, xmlNodeTemp As MSXML.IXMLDOMNode
    Dim xmlNodeTemp2 As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    
    GetCellSpan fps, pCol, pRow
    
    Set xmlNodeCells = TAX_Utilities_v1.Data(intSheet - 1).nodeFromID(GetCellID(fps, pCol, pRow)).parentNode
    
    lRows = GetDynRowCount(fps, xmlNodeCells, lRow2s, lLRowBound, lURowBound)

    If GetAttribute(xmlNodeCells.parentNode, "Dynamic") <> "1" Then GoTo EXIT_SUB
    
'*********************************************************
' added
'Date: 01/03/06
    'Check whether user want to delete
'    If lRows > 1 And blnForce And Not IsEmptyValue(xmlNodeCells) Then
'        If DisplayMessage("0075", msYesNo, miQuestion, , mrYes) = mrNo Then
'            Exit Sub
'        End If
'    End If
'*********************************************************
        
    If xmlNodeCells.parentNode.childNodes.length <= 1 Then
        ClearRows xmlNodeCells
        TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
        GoTo EXIT_SUB
    Else
        If GetAttribute(xmlNodeCells.firstChild, "FirstCell") = "0" Then
            'Set FirstCell attr to "0" for next Cells node
            Set xmlNodeCells = xmlNodeCells.nextSibling
            SetAttribute xmlNodeCells.firstChild, "FirstCell", "0"
            Set xmlNodeCells = xmlNodeCells.previousSibling
        End If
    End If
    
    'If xmlNodeCells.xml = xmlNodeCells.parentNode.childNodes(0).xml Then GoTo EXIT_SUB
    
    'Jump active cell to prevous section
    'fpSpread1.SetActiveCell fpSpread1.ActiveCol, fpSpread1.ActiveRow - lRows
    
    ' Delete curent row on Form
    DeleteRow intSheet, lLRowBound, lRows
    'fpSpread1.SetFocus
    
    xmlNodeCells.parentNode.removeChild xmlNodeCells
    
    ' Decrease value of row in xmlDocument
    DecreaseRowInDOM intSheet, lLRowBound + 1, lRows, lRow2s
    
    fps.Col = fps.ActiveCol
    fps.Row = fps.ActiveRow
    If fps.Lock = True Then
        Do
            fps.Row = fps.Row - 1
        Loop Until (Not fps.Lock) Or (fps.Row = 1)
        fps.SetActiveCell fps.Col, fps.Row
    End If
'   ************************************
'    added
'   Date: 12/04/06
    TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
'   ************************************
EXIT_SUB:
    Set xmlNodeCells = Nothing
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "04_GTGT", "DeleteNode", Err.number, Err.Description
End Sub


Private Sub ClearRows(xmlCellsNode As MSXML.IXMLDOMNode) '(ByVal lRow As Long, ByVal lRows As Long)
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode

    With fps
        For Each xmlCellNode In xmlCellsNode.childNodes
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
            .Sheet = mCurrentSheet
            .Col = lCol
            .Row = lRow
            If .Lock = False Or (.Lock = True And .Formula = vbNullString) Then
                Select Case .CellType
                    Case CellTypeNumber
                        .Text = "0"
                        SetAttribute xmlCellNode, "Value", "0"
                    Case Else
                        .Text = ""
                        SetAttribute xmlCellNode, "Value", ""
                End Select
           End If
                
            
            If .CellNote <> vbNullString Then
                .CellNote = vbNullString
                .BackColor = vbWhite
            End If
            '*******************************
        Next
    End With
End Sub


Private Sub DeleteRow(ByVal intSheet As Integer, ByVal pRow As Long, ByVal lRows As Long)
    On Error GoTo ErrorHandle
    
    With fps
        .EventEnabled(EventChange) = False
        .ReDraw = False
        '.Visible = False
        .Sheet = intSheet
        .DeleteRows pRow, lRows
        .MaxRows = .MaxRows - lRows
        '.Visible = True
        .ReDraw = True
        .EventEnabled(EventChange) = True
    End With
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "04_GTGT", "DeleteRow", Err.number, Err.Description
End Sub


Private Sub DecreaseRowInDOM(ByVal intSheet As Integer, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, i As Long
    
    If TAX_Utilities_v1.Data(intSheet - 1) Is Nothing Then Exit Sub
    Set xmlNodeListCell = TAX_Utilities_v1.Data(intSheet - 1).getElementsByTagName("Cell")
    
    For i = 0 To xmlNodeListCell.length - 1
        ParserCellID fps, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
        If lRow >= pRow Then
            ' Decrease value of row attribute - 1 "CellID"
            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fps, lCol, lRow - lRows)
            
            ' Decrease value of row attribute - 1 "CellID2"
            ParserCellID fps, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fps, lCol, lRow - lRow2s)
        End If
    Next
    
    Set xmlNodeListCell = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "04_GTGT", "DecreaseRowInDOM", Err.number, Err.Description
End Sub


' Kiem tra xem neu la group II, III, IV ma bam F5, F6 se khong them dong
Public Function isAddRow(ByVal iCol As Integer, ByVal iRow As Integer) As Boolean
    Dim varTemp  As Variant
    With fps
        .Sheet = 1
        ' Group 1
        .Col = .ColLetterToNumber("AF")
        .Row = iRow
        .GetText .Col, .Row, varTemp
        If Trim$(CStr(varTemp)) = "F" Then
            isAddRow = False
        Else
            isAddRow = True
        End If
    End With
End Function


Public Sub UpdateSheet()
    Dim varTemp As Variant
    Dim temp As Integer
    Dim blCheck_S2 As Boolean
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    
    ' xu ly phan bo ma hoa don ban 321 in datafile 320
    Dim temp1 As Variant
    Dim xmlCellNodeData1 As MSXML.IXMLDOMNode
    Dim xmlCellNodeData2 As MSXML.IXMLDOMNode
    Dim xmlCellNodeData3 As MSXML.IXMLDOMNode
    Dim xmlCellNodeData4 As MSXML.IXMLDOMNode
    ' end
    
    With fps
        .Sheet = 1
        temp = mCurrentSheet
        mCurrentSheet = 1
        For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
            ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
            .GetText lCol, lRow, varTemp
            UpdateCell fps, lCol, lRow, varTemp
        Next
         
         
'        .Sheet = 2
'        mCurrentSheet = .Sheet
'        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'        If blCheck_S2 = True Then
'            For Each xmlCellNodeData In TAX_Utilities_v1.Data(.Sheet - 1).getElementsByTagName("Cell")
'                    ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'                    .GetText lCol, lRow, varTemp
'                    UpdateCell fps, lCol, lRow, varTemp
'            Next
'
'            ' xu ly truong hop datafile ban 320 se set encode cell dau tien = 0 (bo ma hoa don)
'            Set xmlCellNodeData4 = TAX_Utilities_v1.Data(1).getElementsByTagName("Sections").Item(0)
'            temp1 = GetAttribute(xmlCellNodeData4, "Version")
'            If Val(temp1) = 320 Then
'                For Each xmlCellNodeData1 In TAX_Utilities_v1.Data(1).getElementsByTagName("Section")
'                    If GetAttribute(xmlCellNodeData1, "Dynamic") = "1" Then
'                        For Each xmlCellNodeData2 In xmlCellNodeData1.childNodes
'                            Set xmlCellNodeData3 = xmlCellNodeData2.firstChild
'                            If GetAttribute(xmlCellNodeData3, "FirstCell") <> vbNullString Then
'                                SetAttribute xmlCellNodeData3, "Encode", "0"
'                            End If
'                        Next
'                    End If
'                Next
'            End If
'
'            ' end xu ly data data
'        End If
          
        ' Update lai sheet KHBS
        .Sheet = .SheetCount - 1
        mCurrentSheet = .Sheet
        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(.Sheet - 1).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
         
  
        mCurrentSheet = temp
    End With
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

' Dieu chinh to khai BS
Public Function getValueTK(strFileLoadData As String) As String

    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode

    'Dim strOriginDataFileName As String
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim result As String
    Dim iDom As Integer
    Dim MY As Date

    iDom = 0
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        If .Load(strFileLoadData) = True Then
            Set xmlNode = .nodeFromID("Z_43")
            result = "CT23_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("Z_44")
            result = result & "~" & "CT25_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("Z_45")
            result = result & "~" & "CT27_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("Z_46")
            result = result & "~" & "CT29_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_04GTGT", "getValueTK", Err.number, Err.Description
End Function

Public Function getDieuChinhGiam(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue, strDieuChinhGiam, strDieuChinhTang As String
    Dim tempArrOld() As String
    Dim tempArrCurr() As String
    Dim tempArrOldValue() As String
    Dim tempArrCurrValue() As String
    Dim i, j As Integer
    Dim strMaCT As String
    Dim arrMapCT() As String ' Luu gia tri cua dong chi tieu
    Dim arrMapCTValue() As String ' Luu gia tri cua cac cell chi tieu
    strMaCT = getMapCT
    With fps
                .Sheet = 1
                .GetText .ColLetterToNumber("Z"), 43, varTemp
                strCurrentValue = "CT23_" & Trim(varTemp)
                .GetText .ColLetterToNumber("Z"), 44, varTemp
                strCurrentValue = strCurrentValue & "~CT25_" & Trim(varTemp)
                .GetText .ColLetterToNumber("Z"), 45, varTemp
                strCurrentValue = strCurrentValue & "~CT27_" & Trim(varTemp)
                .GetText .ColLetterToNumber("Z"), 46, varTemp
                strCurrentValue = strCurrentValue & "~CT29_" & Trim(varTemp)
                ' tinh chenh lech giua 2 lan dieu chinh
                tempArrOld = Split(strOldValue, "~")
                tempArrCurr = Split(strCurrentValue, "~")
                arrMapCT = Split(strMaCT, "~")
                strDieuChinhGiam = ""
                strDieuChinhTang = ""
                For i = 0 To UBound(tempArrOld)
                    tempArrOldValue = Split(tempArrOld(i), "_")
                    tempArrCurrValue = Split(tempArrCurr(i), "_")
                    ' lay cac gia tri cua tung cot trong sheet KHBS
                    For j = 0 To UBound(arrMapCT)
                        arrMapCTValue = Split(arrMapCT(j), "_")
                        ' phan tu so 0 la ma chi tieu
                        If arrMapCTValue(0) = Mid(tempArrOldValue(0), 3) Then
                            strMaCT = arrMapCT(j)
                            Exit For
                        End If
                    Next j
                    
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
                        ' dieu chinh giam
                        strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
                        ' dieu chinh tang
                        strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If

                Next i
                If strDieuChinhGiam <> "" Then
                    strDieuChinhGiam = Mid(strDieuChinhGiam, 2, Len(strDieuChinhGiam) - 1)
                End If
                If strDieuChinhTang <> "" Then
                    strDieuChinhTang = Mid(strDieuChinhTang, 2, Len(strDieuChinhTang) - 1)
                End If
                getDieuChinhGiam = strDieuChinhGiam & "###" & strDieuChinhTang
      End With
End Function
Function getMapCT() As String
    On Error GoTo ErrHandle
    
    Dim tempValue As Variant
    Dim strPath As String
    Dim i As Integer
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT_TT28.xml"
    
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "ID") = GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") Then
                Exit For
            End If
        Next
    End If
    
    Set xmlNodeListItem = xmlNode.childNodes
    
    For Each xmlNodeItem In xmlNodeListItem
         tempValue = tempValue & "~" & GetAttribute(xmlNodeItem, "MCT") & "_" & GetAttribute(xmlNodeItem, "Caption")
    Next
    tempValue = Mid(tempValue, 2)
    getMapCT = tempValue
    Exit Function
ErrHandle:
    SaveErrorLog "cls_04GTGT", "getMapCT", Err.number, Err.Description
End Function

' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("O"), 52, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function

' Lay ve chi tieu 34 o to khai lan truoc
Public Function getValueCTDC(strFileLoadData As String) As String

    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode

    'Dim strOriginDataFileName As String
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim result As String
    Dim iDom As Integer
    Dim MY As Date

    iDom = 0
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        If .Load(strFileLoadData) = True Then
            Set xmlNode = .nodeFromID("O_52")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_04GTGT", "getValueCTDC", Err.number, Err.Description
End Function
' Ham set lai gia tri cua cong thuc khi tai bang ke 01_1
Sub CellChange2(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
    Dim varTemp As Variant
    Dim cellTagArr() As String
    With fps
        .Sheet = .ActiveSheet
        If .Sheet = 2 Then
             If Col = .ColLetterToNumber("K") Then
                ' chi tieu 21
                .Sheet = 1
                .Col = .ColLetterToNumber("K")
                .Row = 43
                .Formula = "K48"
                 ' chi tieu 22
                .Sheet = 1
                .Col = .ColLetterToNumber("Q")
                .Row = 43
                .Formula = "L48"
                 ' chi tieu 24
                .Sheet = 1
                .Col = .ColLetterToNumber("Q")
                .Row = 44
                .Formula = "M48"
                 ' chi tieu 26
                .Sheet = 1
                .Col = .ColLetterToNumber("Q")
                .Row = 45
                .Formula = "N48"
                 ' chi tieu 28
                .Sheet = 1
                .Col = .ColLetterToNumber("Q")
                .Row = 46
                .Formula = "O48"
            End If
        End If
    End With
End Sub


' Ham set lai gia tri cua cong thuc khi tai bang ke 01_1
'Sub CellChange4()
'    Dim varTemp As Variant
'    Dim cellTagArr() As String
'    With fps
'        .Sheet = .ActiveSheet
'        If .Sheet = 2 Then
'                ' chi tieu 21
'                .Sheet = 1
'                .Col = .ColLetterToNumber("K")
'                .Row = 43
'                .Formula = "K48"
'                 ' chi tieu 22
'                .Sheet = 1
'                .Col = .ColLetterToNumber("Q")
'                .Row = 43
'                .Formula = "L48"
'                 ' chi tieu 24
'                .Sheet = 1
'                .Col = .ColLetterToNumber("Q")
'                .Row = 44
'                .Formula = "M48"
'                 ' chi tieu 26
'                .Sheet = 1
'                .Col = .ColLetterToNumber("Q")
'                .Row = 45
'                .Formula = "N48"
'                 ' chi tieu 28
'                .Sheet = 1
'                .Col = .ColLetterToNumber("Q")
'                .Row = 46
'                .Formula = "O48"
'        End If
'    End With
'End Sub

'Sub resetFormula()
'     With fps
'         ' chi tieu 21
'        .Sheet = 1
'        .Col = .ColLetterToNumber("K")
'        .Row = 43
'        .Formula = ""
'         ' chi tieu 22
'        .Sheet = 1
'        .Col = .ColLetterToNumber("R")
'        .Row = 43
'        .Formula = ""
'         ' chi tieu 24
'        .Sheet = 1
'        .Col = .ColLetterToNumber("R")
'        .Row = 44
'        .Formula = ""
'         ' chi tieu 26
'        .Sheet = 1
'        .Col = .ColLetterToNumber("R")
'        .Row = 45
'        .Formula = ""
'         ' chi tieu 28
'        .Sheet = 1
'        .Col = .ColLetterToNumber("R")
'        .Row = 46
'        .Formula = ""
'    End With
'End Sub

'When cell's data is changed or a control at a cell is handled, This sub is called
'parameter Col: the column of cell that rise event
'parameter Row: the row of cell that rise event
'parameter f: the flag to mark some special option
'**: when this sub is on executing, all events of fpspread will be disabled until the end of the sub
'Public Sub LockCellBySheet()
'    Dim i As Integer
'    Dim a As Double, b As Double, c As Double
'    Dim arrRow As Variant, arrCol As Variant
'    Dim arrCellTag() As String
'    Dim cCol As Long, cRow As Long
'    Dim varTemp As Variant
'    Dim xmlCellNode As MSXML.IXMLDOMNode
'    Dim iCol  As Long, iRow As Long
'    Dim strValue1  As Variant, strValue2 As Variant
'With fps
'    'Because of unclearly-rised events, this function is here for turning it off
'    'at the end of this sub, this event will be turned on
'    .EventEnabled(EventAllEvents) = False
'
'
'    ' Kiem tra neu co phu luc moi set cong thuc tong hop so lieu tu tinh
'    Dim tempCurrSheet As Integer
'    Dim ct21 As Variant
'    Dim blCheck_S1 As Boolean
'    Dim blCheck_S2 As Boolean
'    Dim blCheck_S5 As Boolean
'    Dim blCheck_S4A As Boolean
'    Dim blCheck_S4B As Boolean
'    'blCheck_S1 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    'blCheck_S5 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(6).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    'blCheck_S4A = IIf(TAX_Utilities_v1.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    'blCheck_S4B = IIf(TAX_Utilities_v1.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    tempCurrSheet = mCurrentSheet
'    .Sheet = 1
'    ' Chi tieu 21
'    .Col = .ColLetterToNumber("I")
'    .Row = 23
'    ct21 = .Text
'            If blCheck_S2 = True And ct21 = "0" Then
'                .Sheet = 1
'                ' Chi tieu 23
'                .Col = .ColLetterToNumber("J")
'                .Row = 27
'                .Lock = True
'                ' Chi tieu 24
'                .Col = .ColLetterToNumber("L")
'                .Row = 27
'                .Lock = True
'            ElseIf (blCheck_S2 = False Or blCheck_S2 = True) And ct21 = "1" Then
'                .Sheet = 1
'                ' Chi tieu 23
'                .Col = .ColLetterToNumber("J")
'                .Row = 27
'                .Lock = True
'                ' Chi tieu 24
'                .Col = .ColLetterToNumber("L")
'                .Row = 27
'                .Lock = True
'            Else
'                .Sheet = 1
'                ' Chi tieu 23
'                .Col = .ColLetterToNumber("J")
'                .Row = 27
'                .Lock = False
'                ' Chi tieu 24
'                .Col = .ColLetterToNumber("L")
'                .Row = 27
'                .Lock = False
'            End If
'    .EventEnabled(EventAllEvents) = True
'End With
'End Sub

' Ham set lai gia tri cua cong thuc khi tai bang ke 01_2
'Sub CellChange3()
''    Dim blCheck_S4A As Boolean
''    Dim blCheck_S4B As Boolean
''
''    blCheck_S4A = IIf(TAX_Utilities_v1.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
''    blCheck_S4B = IIf(TAX_Utilities_v1.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
''
''    ' Tong hop len bang ke 01_4A
''        With fps
''        .Sheet = .ActiveSheet
''        Select Case .Sheet
''            Case 2 ' Phu luc 01_1
''                If blCheck_S4A = True Then
''                    .Sheet = 5
''                    ' chi tieu tong doanh so ban ra tren phu luc 1
''                    .Col = .ColLetterToNumber("D")
''                    .Row = 16
''                    .Formula = "C24"
''                    ' chi tieu T?ng doanh thu hàng hóa, d?ch v? bán ra ch?u thu? GTGT
''                    .Sheet = 5
''                    .Col = .ColLetterToNumber("D")
''                    .Row = 17
''                    .Formula = "D24"
''                End If
''         End Select
''    End With
'End Sub
'******************************
'Set trang thai Flag KHBS active
Public Sub SetActiveKHBS() 'ByRef xmlNodeValid As MSXML.IXMLDOMNode)
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    Dim strValue3 As String
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
                Case "KHBS"
                    blnNullValue = True
                    'Check value of last node in Data
                    Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell") _
                        (TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell").length > 31 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
                            blnNullValue = False
                    End If

                    If blnNullValue Then
                        SetAttribute xmlNodeValid, "Active", "0"
                        FlagThaydoi = "1"
                    End If
        End Select
    Next
    
    Set xmlCellNode = Nothing
    Set xmlNodeValid = Nothing
End Sub
' set ma hoa don
'Private Sub setMaHoaDon()
'    Dim strDataFileName As String
'    Dim arrDanhsach() As String
'    Dim strComboHien As String
'    Dim strCombo As String
'    Dim MSTDN As String
'    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
'    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
'    Dim xmlNode As MSXML.IXMLDOMNode
'
'    strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_loai_HD.xml")
'    ' Lay danh muc loai hoa don
'    If xmlDomData.Load(strDataFileName) Then
'        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
'        For Each xmlNode In xmlNodeListCell
'            If GetAttribute(xmlNode, "Value") <> "" Then
'                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
'                strCombo = strCombo + CPab(arrDanhsach(0), 10) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(2), 200) + Chr$(9)
'                strComboHien = strComboHien + arrDanhsach(1) + Chr$(9)
'            End If
'        Next
'    End If
'
'     With fps
'        .Sheet = sheetPl01_1_GTGT
'        ' Set gia tri ma HD
'        ' Group khong chiu thue
'        .Col = .ColLetterToNumber("C")
'        .Row = 8
'        .TypeComboBoxList = strComboHien
'        'Group 0%
'        .Col = .ColLetterToNumber("C")
'        .Row = 13
'        .TypeComboBoxList = strComboHien
'        ' Group 5%
'        .Col = .ColLetterToNumber("C")
'        .Row = 18
'        .TypeComboBoxList = strComboHien
'        ' Group 10%
'        .Col = .ColLetterToNumber("C")
'        .Row = 23
'        .TypeComboBoxList = strComboHien
'        ' Group khong phai tong hop len Tk 01/GTGT
'        .Col = .ColLetterToNumber("C")
'        .Row = 28
'        .TypeComboBoxList = strComboHien
'
'    End With
'End Sub

' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 13
    denRowDL = 18
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "20_30_2"
    ' end
End Sub

Public Sub FinishImport()
'Dim i As Integer, iCol As Long, iRow As Long
'Dim varTemp As Variant
'With fps
'fps.EventEnabled(EventAllEvents) = False
'    If .ActiveSheet = 2 Then
'             i = 1
'             .Col = .ColLetterToNumber("B")
'             .Row = 8
'             Do
'                  .Text = str(i)
''                         .Col = .ColLetterToNumber("K")
''                         .Text = "0"
''                         UpdateCell fps, .Col, .Row, .Text
'                   ' chuyen ky hieu thanh chu hoa
'                   .Col = .ColLetterToNumber("E")
'                   .GetText .Col, .Row, varTemp
'                   .Text = UCase$(CStr(varTemp))
'                   UpdateCell fps, .Col, .Row, .Text
'                   ' end
'                  .Col = .ColLetterToNumber("B")
'                  .Row = i + 8
'                  i = i + 1
'             Loop Until .Text = "aa"
'             i = 1
'            .Row = .Row + 4
'             Do
'                  .Text = str(i)
'
''                         .Col = .ColLetterToNumber("M")
''                         .Text = "0"
''                         UpdateCell fps, .Col, .Row, .Text
'
'                  ' chuyen ky hieu thanh chu hoa
'                  .Col = .ColLetterToNumber("E")
'                   .GetText .Col, .Row, varTemp
'                   .Text = UCase$(CStr(varTemp))
'                   UpdateCell fps, .Col, .Row, .Text
'                  ' end
'
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "bb"
'             i = 1
'            .Row = .Row + 4
'             Do
'                  .Text = str(i)
''                         .Col = .ColLetterToNumber("J")
''                         .TypeHAlign = TypeHAlignCenter
''                         .Text = "5%"
''                         UpdateCell fps, .Col, .Row, .Text
'
'                 ' chuyen ky hieu thanh chu hoa
'                  .Col = .ColLetterToNumber("E")
'                   .GetText .Col, .Row, varTemp
'                   .Text = UCase$(CStr(varTemp))
'                   UpdateCell fps, .Col, .Row, .Text
'                  ' end
'
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "cc"
'             i = 1
'            .Row = .Row + 4
'             Do
'                  .Text = str(i)
''                         .Col = .ColLetterToNumber("J")
''                         .TypeHAlign = TypeHAlignCenter
''                         .Text = "10%"
''                         UpdateCell fps, .Col, .Row, .Text
'                   ' chuyen ky hieu thanh chu hoa
'                  .Col = .ColLetterToNumber("E")
'                   .GetText .Col, .Row, varTemp
'                   .Text = UCase$(CStr(varTemp))
'                   UpdateCell fps, .Col, .Row, .Text
'                  ' end
'
'                  .Col = .ColLetterToNumber("B")
'
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "dd"
'             i = 1
'             .Row = .Row + 4
'             Do
'                  .Text = str(i)
''                         .Col = .ColLetterToNumber("J")
''                         .TypeHAlign = TypeHAlignCenter
''                         .Text = "10%"
''                         UpdateCell fps, .Col, .Row, .Text
'
'                  ' chuyen ky hieu thanh chu hoa
'                  .Col = .ColLetterToNumber("E")
'                   .GetText .Col, .Row, varTemp
'                   .Text = UCase$(CStr(varTemp))
'                   UpdateCell fps, .Col, .Row, .Text
'                  ' end
'
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "ee"
'             .SetActiveCell iCol, iRow
'             .Row = iRow
'             .Col = iCol
'             CellChange4
'        End If
'        fps.EventEnabled(EventAllEvents) = True
'        End With
End Sub
