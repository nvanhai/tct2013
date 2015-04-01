VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_02GTGT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Gia Tri Gia Tang" interface sheets
'this Class is belong to TAX_Business_v1 project which will be compline to DLL
'Coder: HTPHUONG
Option Explicit

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Private Const S1ChkRow = 6
Private Const S1ChkCol = "I"
Private Const S1Txt11Row = 7
Private Const S1Txt11Col = "L"
Private Const S4TxtF2row = 2
Private Const S4TxtF2Col = "F"

Private Const Row_MSTDN = 24
Private Const Col_MSTDN = "AH"

Private Const SECOND_SHEET = 2
Private Const THIRD_SHEET = 3
Private Const FOURTH_SHEET = 4
Private Const THIRD_SHEET_CELL1 = "D_19"
Private Const FOURTH_SHEET_CELL1 = "E_19"


Private Const rowNguoiKy = 57
Private Const rowNgayKy = 59
Private Const rowHiddenBsLan = 61
Private Const rowHangHoaMuaVao = 43

Private Const rowCt21 = 40
Private Const rowCt21a = 41
Private Const rowCt22 = 43
Private Const rowCt23 = 43
Private Const rowCt24 = 45
Private Const rowCt28a = 48
Private Const rowCt29 = 49
Private Const rowCt30 = 52
Private Const rowCt30a = 51
Private Const rowCt31 = 53
Private Const rowCt32 = 54
Private Const rowHidden62 = 62


Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
'Dim xNodeData0 As MSXML.IXMLDOMNode
Dim xNode As MSXML.IXMLDOMNode
Dim arrData() As String
Public Strloaitk As String
Public StrSolanBosung As String
Public loaiKHBS As String
Public FlagThaydoi As String

' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

Public strTkThangQuy As String
Public dNgayDauQuy As Date


'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'no parameter
Public Sub Prepare1()
'Dim strComboHien As Variant
With fps
    'danh sach  list cua combobox
   ' strComboHien = "0%" + Chr$(9) + "5%" + Chr$(9) + "10%" + Chr$(9)
    '--
'    SetDateFormat fps, 2, 8, .ColLetterToNumber("G"), DDMMYYYY
'    SetDateFormat fps, 3, 8, .ColLetterToNumber("G"), DDMMYYYY
'    SetDateFormat fps, 4, 8, .ColLetterToNumber("G"), DDMMYYYY
'    SetDateFormat fps, 5, 8, .ColLetterToNumber("G"), DDMMYYYY
'    SetDateFormat fps, 6, 8, .ColLetterToNumber("G"), DDMMYYYY
'    SetDateFormat fps, 7, 8, .ColLetterToNumber("G"), DDMMYYYY
'    SetDateFormat fps, 8, 8, .ColLetterToNumber("G"), DDMMYYYY
'    SetDateFormat fps, 9, 8, .ColLetterToNumber("G"), DDMMYYYY
'    SetDateFormat fps, 10, 8, .ColLetterToNumber("G"), DDMMYYYY
'    SetDateFormat fps, 11, 8, .ColLetterToNumber("G"), DDMMYYYY
    
    SetDateFormat fps, 1, rowNgayKy, .ColLetterToNumber("CB"), DDMMYYYY
    .Sheet = 1 'To khai GTGT
        .Row = rowNgayKy
        .Col = .ColLetterToNumber("CB")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
'    .Sheet = 2 'PL01_1
'        .Row = 8
'        .Col = .ColLetterToNumber("C")
'        .Text = ""
'        .CellTag = .CellTag & "~1"
'        .Col = .ColLetterToNumber("J")
'        .Row = .Row
''        .TypeComboBoxList = strComboHien
''        SetDateFormat fps, 2, 8, .ColLetterToNumber("E"), DDMMYYYY
        
'    .Sheet = 3 'PL01_1
'        .Row = 8
'        .Col = .ColLetterToNumber("C")
'        .Text = ""
'        .CellTag = .CellTag & "~1"
'        .Col = .ColLetterToNumber("J")
'        .Row = .Row
'    .Sheet = 4 'PL01_1
'        .Row = 8
'        .Col = .ColLetterToNumber("C")
'        .Text = ""
'        .CellTag = .CellTag & "~1"
'        .Col = .ColLetterToNumber("J")
'        .Row = .Row
'    .Sheet = 5 'PL01_1
'        .Row = 8
'        .Col = .ColLetterToNumber("C")
'        .Text = ""
'        .CellTag = .CellTag & "~1"
'        .Col = .ColLetterToNumber("J")
'        .Row = .Row
'    .Sheet = 6 'PL01_1
'        .Row = 8
'        .Col = .ColLetterToNumber("C")
'        .Text = ""
'        .CellTag = .CellTag & "~1"
'        .Col = .ColLetterToNumber("J")
'        .Row = .Row
'    .Sheet = 7 'PL01_1
'        .Row = 8
'        .Col = .ColLetterToNumber("C")
'        .Text = ""
'        .CellTag = .CellTag & "~1"
'        .Col = .ColLetterToNumber("J")
'        .Row = .Row
'    .Sheet = 8 'PL01_1
'        .Row = 8
'        .Col = .ColLetterToNumber("C")
'        .Text = ""
'        .CellTag = .CellTag & "~1"
'        .Col = .ColLetterToNumber("J")
'        .Row = .Row
'    .Sheet = 9 'PL01_1
'        .Row = 8
'        .Col = .ColLetterToNumber("C")
'        .Text = ""
'        .CellTag = .CellTag & "~1"
'        .Col = .ColLetterToNumber("J")
'        .Row = .Row
'
'    .Sheet = 10 'PL01_1
'        .Row = 8
'        .Col = .ColLetterToNumber("C")
'        .Text = ""
'        .CellTag = .CellTag & "~1"
'        .Col = .ColLetterToNumber("J")
'        .Row = .Row
'     .Sheet = 11 'PL01_1
'        .Row = 8
'        .Col = .ColLetterToNumber("C")
'        .Text = ""
'        .CellTag = .CellTag & "~1"
'        .Col = .ColLetterToNumber("J")
'        .Row = .Row
    
    SetDateFormat fps, .SheetCount - 1, 31, .ColLetterToNumber("BE"), DDMMYYYY
    SetDateFormat fps, .SheetCount - 1, 46, .ColLetterToNumber("BG"), DDMMYYYY
    .Sheet = .SheetCount - 1 'KHBS
    .Row = 46
    .Col = .ColLetterToNumber("BG")
    .Text = Format(Date, "dd/mm/yyyy")
    .TypeHAlign = TypeHAlignLeft
    
End With
    ' set co quan thue quan ly
    setCQTQuanLy
'    ' set ma hoa don
'    setMaHoaDon
    
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'no parameter
Public Sub Prepare2()
    Dim i As Long
    Dim intCtrl As Integer
    Dim strId As Variant
    Dim xmlNode As MSXML.IXMLDOMNode
    With fps
        'vttoan: set trang thai to khai
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        
'        CellChange .ColLetterToNumber("BW"), rowHangHoaMuaVao
        'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
        mCurrentSheet = .Sheet
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("C")
            .Row = rowHiddenBsLan
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = rowHiddenBsLan
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = rowHiddenBsLan
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("C")
            .Row = rowHiddenBsLan
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = rowHiddenBsLan
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = rowHiddenBsLan
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        ' set loai tk thang/ quy
         If strTkThangQuy = "TK_THANG" Then
            .Col = .ColLetterToNumber("K")
            .Row = 61
            .Text = "0"
            UpdateCell fps, .Col, .Row, .Text
        ElseIf strTkThangQuy = "TK_QUY" Then
            .Col = .ColLetterToNumber("K")
            .Row = 61
            .Text = "1"
            UpdateCell fps, .Col, .Row, .Text
        End If
            'end
        
        .EventEnabled(EventAllEvents) = True
        ' end
        
        
        
    'dhdang sua yc them nhieu Sheet PL 01_1
'    Dim k As Variant
'    For k = 2 To 11
'        .Sheet = k
'        .ActiveSheet = .Sheet
'        .Row = 1
'        .Col = .ColLetterToNumber("F")
'        .CellType = CellTypeButton
'        .TypeButtonTextColor = RGB(0, 0, 255)
'        .TypeButtonColor = &H8000000F
'        .TypeButtonAlign = TypeButtonAlignRight
'        .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
''
''        .Row = 1
''        .Col = .ColLetterToNumber("E")
''        .Lock = True
'    Next
    
'     If Not TAX_Utilities_v1.DataKHBS Is Nothing Then
'        Set xmlNode = TAX_Utilities_v1.Data(0).nodeFromID(GetCellID(fps, fps.ColLetterToNumber("L"), 17))
'        vThuephainop = GetAttribute(xmlNode, "Value")
'        Set xmlNode = Nothing
'
'         .Sheet = .SheetCount - 1 'KHBS
'            Dim cCol As Long, cRow As Long
'          '  .ActiveSheet = .Sheet
'            Set xmlNode = TAX_Utilities_v1.Data(.SheetCount - 2).childNodes(2).childNodes(2).lastChild.lastChild
'            mCurrentSheet = .SheetCount - 1
'            fps.Sheet = mCurrentSheet
'            If GetAttribute(xmlNode, "Value") <> "" Then
'                If GetAttribute(xmlNode, "Value") < 0 Then
'                    ParserCellID fps, GetAttribute(xmlNode, "CellID"), cCol, cRow
'                    .Col = .ColLetterToNumber("BE")
'                    .Row = cRow + 3
'                    .value = 0
'                    UpdateCell fps, .Col, .Row, .value
'                End If
'            End If
'            i = 1
'            .Col = .ColLetterToNumber("B")
'            .Row = 9
'            Do
'              .GetText .ColLetterToNumber("BC"), .Row, strId
'              If strId <> "" Then
'                .Col = .ColLetterToNumber("BF")
'                .CellTag = .Row & "~" & Trim(Mid(strId, 100, 20))
'
'                    If Trim(Mid(strId, 100, 20)) = "L_18" Or Trim(Mid(strId, 100, 20)) = "L_19" Then
'                        .Col = .ColLetterToNumber("BH")
'                        .Row = .Row
'                        .Formula = "BG" & .Row & "-" & "BF" & .Row
'                    Else
'                        .Col = .ColLetterToNumber("BH")
'                        .Row = .Row
'                        .Formula = "BF" & .Row & "-" & "BG" & .Row
'                    End If
'
'              End If
'              .Col = .ColLetterToNumber("B")
'              .Row = .Row + 1
'            Loop Until .Text = "bb"
'
'            Set xmlNode = TAX_Utilities_v1.Data(.SheetCount - 2).childNodes(2).firstChild
'            SetAttribute xmlNode, "DateKHBS", TAX_Utilities_v1.DateKHBS
'            SetAttribute xmlNode, "loaiKHBS", loaiKHBS
'            If loaiKHBS = "frmKHBS_BS" Then
'                .SetText .ColLetterToNumber("B"), 5, GetAttribute(GetMessageCellById("0115"), "Msg") & _
'                GetAttribute(GetMessageCellById("0117"), "Msg") & Mid(TAX_Utilities_v1.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 5, 4)
'            Else
'                .SetText .ColLetterToNumber("B"), 5, GetAttribute(GetMessageCellById("0116"), "Msg") & _
'                GetAttribute(GetMessageCellById("0117"), "Msg") & Mid(TAX_Utilities_v1.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 5, 4)
'            End If
'            .EventEnabled(EventAllEvents) = False
'            .Sheet = 1
'            mCurrentSheet = 1
'            .Col = .ColLetterToNumber("CB")
'            .Row = rowNguoiKy
'            .Formula = "TRIM(Header!T5)"
'            UpdateCell fps, .Col, .Row, .Text
'            .Sheet = 1 'To khai GTGT
'            .Row = rowNgayKy
'            .Col = .ColLetterToNumber("CB")
'            .Text = Format(Date, "dd/mm/yyyy")
'            .TypeHAlign = TypeHAlignLeft
'            .EventEnabled(EventAllEvents) = True
'
'             ReDim arrData(TAX_Utilities_v1.DataKHBS.getElementsByTagName("Cell").length)
'            Dim j As Integer
'            For Each xNode In TAX_Utilities_v1.DataKHBS.getElementsByTagName("Cell")
'                arrData(j) = GetAttribute(xNode, "CellID") & ":" & GetAttribute(xNode, "Value")
'                j = j + 1
'            Next
'            Set xNode = Nothing
'        End If
    End With
 
    setLast25Target
    
    
End Sub

'inner class event: Change
Public Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
Dim str25Previous As String
Dim varTemp As Variant
Dim strId As Variant
Dim iCol As Long, iRow As Long
Dim cCol As Long, cRow As Long
Dim cValue As Variant, temp As Variant
Dim cValueG As Variant
Dim cTagF As String
Dim cValueF As Variant
Dim i As Integer
Dim xmlNode As MSXML.IXMLDOMNode
Dim xmlNodeH15 As MSXML.IXMLDOMNode
With fps
    If .ActiveSheet = 1 Then
    mCurrentSheet = .Sheet
    fps.EventEnabled(EventAllEvents) = False
        'Check value of 11
        GetCellSpan fps, Col, Row
        If Col = .ColLetterToNumber("L") And Row = 8 Then
            .GetText Col, Row, varTemp
            If GetLast25Previous(str25Previous) Then
                If CStr(varTemp) <> str25Previous Then
                    .Visible = True
                    DisplayMessage "0060", msOKOnly, miWarning
                    fps.SetFocus
                End If
            End If
        End If
        
        .GetText .ColLetterToNumber("L"), 17, varTemp
'           If loaiKHBS = "" Then
'                If varTemp <= 0 Then
'                    .Row = 18
'                    .Col = .ColLetterToNumber("L")
'                    .Text = "0"
'                    .Lock = True
'                    UpdateCell fps, .Col, .Row, .Text
'                    .Col = .ColLetterToNumber("L")
'                    .Row = 19
'                    .Text = "0"
'                    .Lock = True
'                    UpdateCell fps, .Col, .Row, .Text
'                Else
'                    .Col = .ColLetterToNumber("L")
'                    .Row = 18
'                    .Lock = False
'                    .Row = 19
'                    .Lock = False
'                End If
'
'            End If
       fps.EventEnabled(EventAllEvents) = True
        
    End If
    
'    If .ActiveSheet = 2 Or .ActiveSheet = 3 Or .ActiveSheet = 4 Or .ActiveSheet = 5 Or .ActiveSheet = 6 Or .ActiveSheet = 7 Or .ActiveSheet = 8 Or .ActiveSheet = 9 Or .ActiveSheet = 10 Or .ActiveSheet = 11 Then
'        Dim xmlCellNode As MSXML.IXMLDOMNode
'        Dim strValue1  As Variant, strValue2 As Variant
'        fps.EventEnabled(EventAllEvents) = False
'           mCurrentSheet = .Sheet
'            Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
'                          (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 2)
'            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'            .Col = iCol
'            .Row = iRow
'            strValue1 = .value
''            UpdateCell fps, .Col, .Row, .Value
'            Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
'                          (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 1)
'            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'            .Col = iCol
'            .Row = iRow
'            strValue2 = .value
''            UpdateCell fps, .Col, .Row, .Value
'            If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                   fps.EventEnabled(EventAllEvents) = False
'                    .Sheet = 1
'                    mCurrentSheet = .Sheet
'                    'Chi tieu [11]
'                    .GetText .ColLetterToNumber("C"), 24, varTemp
'                    .Col = .ColLetterToNumber("J")
'                    .Row = 10
'                    .value = varTemp
'                    UpdateCell fps, .Col, .Row, .value
'                    'Chi tieu [12]
'                    .GetText .ColLetterToNumber("D"), 24, varTemp
'                    .Col = .ColLetterToNumber("L")
'                    .Row = 10
'                    .value = varTemp
'                    UpdateCell fps, .Col, .Row, .value
'                   .Sheet = .ActiveSheet
'                    mCurrentSheet = .Sheet
'                    fps.EventEnabled(EventAllEvents) = True
'              End If
'
'             ' set ma hoa don
'           .Col = Col
'            .Row = Row
'            varTemp = .Text
'             If Col = .ColLetterToNumber("C") And .Text <> "" Then
'                .EventEnabled(EventChange) = False
'                .Col = .ColLetterToNumber("D")
'                .Text = varTemp
'                .EventEnabled(EventChange) = True
'             End If
'
'      fps.EventEnabled(EventAllEvents) = True
'    End If
    
    
'    If .ActiveSheet = .SheetCount - 1 Then
'         fps.EventEnabled(EventAllEvents) = False
'        'Check value of 11
''        GetCellSpan fps, Col, Row
''        If blGiaTriCu And (Col = .ColLetterToNumber("BD") Or Col = .ColLetterToNumber("BE")) Then
''                 DisplayMessage "0110", msOKOnly, miWarning
''                .Row = RowGiaTriCu
''                .Col = .ColLetterToNumber("BD")
''                .Text = ""
''                UpdateCell fps, .Col, .Row, .Text
''                .Col = .ColLetterToNumber("BE")
''                .Text = ""
''                UpdateCell fps, .Col, .Row, .Text
''                .Col = .ColLetterToNumber("BF")
''                .Text = ""
''                UpdateCell fps, .Col, .Row, .Text
''                .Col = .ColLetterToNumber("BG")
''                .Text = ""
''                UpdateCell fps, .Col, .Row, .Text
''                .Col = .ColLetterToNumber("BH")
''                .Text = ""
''                UpdateCell fps, .Col, .Row, .Text
''                blGiaTriCu = False
''                fps.SetFocus
''                fps.EventEnabled(EventAllEvents) = True
''                Exit Sub
''         End If
'
'        'get gia tri da ke khai
''        .GetText .ColLetterToNumber("BC"), Row, strId
''          If strId <> "" Then
''                Dim arr() As String
''                 For i = 0 To UBound(arrData)
''                     arr() = Split(arrData(i), ":")
''                     If arr(0) = Trim(Mid(strId, 100, 20)) Then
''                         cValue = arr(1)
''                         Exit For
''                     End If
''                 Next
''               .Col = .ColLetterToNumber("BF")
''               .Row = Row
''               cTagF = .CellTag
''               If cTagF <> "" Then
''                         For i = 0 To UBound(arrData)
''                               arr() = Split(arrData(i), ":")
''                               If arr(0) = Right(cTagF, Len(cTagF) - InStr(cTagF, "~")) Then
''                                   cValueF = arr(1)
''                                   Exit For
''                               End If
''                           Next
''                End If
''
''               .CellTag = .Row & "~" & Trim(Mid(strId, 100, 20))
''               .value = cValue
''               UpdateCell fps, .Col, .Row, .value
''               .Col = .ColLetterToNumber("BG")
''               .Row = Row
''               UpdateCell fps, .Col, .Row, .value
''               .Col = .ColLetterToNumber("BH")
''               .Row = Row
''               UpdateCell fps, .Col, .Row, .value
''
''                If Trim(Mid(strId, 100, 20)) = "L_18" Or Trim(Mid(strId, 100, 20)) = "L_19" Then
''                    .Col = .ColLetterToNumber("BH")
''                    .Row = Row
''                    .Formula = "BG" & .Row & "-" & "BF" & .Row
''                     UpdateCell fps, .Col, .Row, .value
''               Else
''                    .Col = .ColLetterToNumber("BH")
''                    .Row = Row
''                    .Formula = "BF" & .Row & "-" & "BG" & .Row
''                     UpdateCell fps, .Col, .Row, .value
''               End If
''
''               If loaiKHBS = "frmKHBS_TT" Then
''                    .GetText .ColLetterToNumber("BG"), Row, cValueG
''               Else
''                    .GetText .ColLetterToNumber("BH"), Row, cValueG
''               End If
''
''              If cValueG <> "" Then
''                      ParserCellID fps, Trim(Mid(strId, 100, 20)), cCol, cRow
''                      If (Trim(Mid(strId, 100, 20)) = "L_11" Or Trim(Mid(strId, 100, 20)) = "L_12") And loaiKHBS = "frmKHBS_BS" Then
''                            cValueG = -(cValueG)
''                      End If
''                      .Sheet = 1
''                      .SetText cCol, cRow, cValueG
''                      If loaiKHBS = "frmKHBS_TT" Then
''                        UpdateKHBSCell fps, cCol, cRow, cValueG
''                      Else
''                        UpdateKHBSCell fps, cCol, cRow, cValueG
''                      End If
''              End If
''
''' Set lai gia tri cu cua sheet 01
''               .Sheet = .SheetCount - 1
''               .Col = .ColLetterToNumber("BF")
''               .Row = Row
''
''               If cTagF <> "" And cTagF <> .CellTag Then
''                  If Left(cTagF, InStr(cTagF, "~") - 1) = .Row Then
''                   ' .GetText .ColLetterToNumber("F"), Row, temp
''                    ParserCellID fps, Right(cTagF, Len(cTagF) - InStr(cTagF, "~")), iCol, iRow
''                    UpdateKHBSCell fps, iCol, iRow, cValueF
''                    .Sheet = 1
''                    If loaiKHBS = "frmKHBS_TT" Then
''                        .SetText iCol, iRow, cValueF
''                    Else
''                        .SetText iCol, iRow, 0
''                    End If
''                    .Sheet = .SheetCount - 1
''                    If loaiKHBS = "frmKHBS_TT" Then
''                        .GetText .ColLetterToNumber("BG"), Row, varTemp
''                    Else
''                        .GetText .ColLetterToNumber("BH"), Row, varTemp
''                    End If
''                    If varTemp <> "" Then
''                        ParserCellID fps, Trim(Mid(strId, 100, 20)), iCol, iRow
''                        If loaiKHBS = "frmKHBS_TT" Then
''                            UpdateKHBSCell fps, cCol, cRow, varTemp
''                        Else
''                            UpdateKHBSCell fps, cCol, cRow, varTemp + cValueF
''                        End If
''                       ' UpdateKHBSCell fps, iCol, iRow, varTemp
''                        .Sheet = 1
''                        .SetText iCol, iRow, varTemp
''                        .Sheet = .SheetCount - 1
''                    End If
''                   End If
''               End If
'''Xoa dong
''               If mblKeyF6 = True Then
''                ParserCellID fps, Right(cTagF, Len(cTagF) - InStr(cTagF, "~")), iCol, iRow
''                UpdateKHBSCell fps, iCol, iRow, cValueF
''                .Sheet = 1
''                .SetText iCol, iRow, 0
''                .Sheet = .SheetCount - 1
''                .Col = .ColLetterToNumber("BF")
''                .Row = Row
''                .value = 0
''                .Col = .ColLetterToNumber("BG")
''                .Row = Row
''                .value = 0
''                .Col = .ColLetterToNumber("BH")
''                .Row = Row
''                .value = 0
''               End If
''
''              mCurrentSheet = .Sheet
''              .Sheet = .SheetCount - 1
''            Set xmlNodeH15 = TAX_Utilities_v1.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
''                 ParserCellID fps, GetAttribute(xmlNodeH15, "CellID"), cCol, cRow
''                .Col = cCol
''                .Row = cRow
''                 UpdateCell fps, .Col, .Row, .value
''
''                If .value < 0 Then
''                    .Col = .ColLetterToNumber("BE")
''                    .Row = cRow + 3
''                    .value = 0
''                    UpdateCell fps, .Col, .Row, 0
''                 Else
''                    .Col = .ColLetterToNumber("BE")
''                    .Row = cRow + 3
''                    UpdateCell fps, .Col, .Row, .value
''                End If
''                .Sheet = mCurrentSheet
''          End If
'            fps.EventEnabled(EventAllEvents) = True
'       End If
     mblKhoiDauToKhai = False
  End With
    
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
'Dim tCol As Long, tRow As Long
'Dim strId As Variant, Col7 As Variant, intIndexCombo As Integer
'Dim NguoiKyCol As Long, NguoiKyRow As Long
'    With fps
'        iCol = .ActiveCol
'        iRow = .ActiveRow
'        GetCellSpan fps, iCol, iRow
'
'        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
''             If .ActiveSheet = 2 Then
''                    fps.EventEnabled(EventAllEvents) = True
''                    i = 1
''                    .Col = .ColLetterToNumber("B")
''                    .Row = 8
''                    Do
''                         .Text = str(i)
''                         .Col = .ColLetterToNumber("B")
''                         .Row = i + 8
''                         i = i + 1
''                    Loop Until .Text = "aa"
''                    .SetActiveCell iCol, iRow
''                    .Row = iRow
''                    .Col = iCol
''                    .SetActiveCell iCol, iRow
''
''                    CellChange iCol, iRow
''                    fps.EventEnabled(EventAllEvents) = False
''            End If
''            If .ActiveSheet = .SheetCount - 1 Then
''                 fps_Change iCol, iRow
''                 fps.EventEnabled(EventAllEvents) = False
''                    i = 1
''                    .Col = .ColLetterToNumber("B")
''                    .Row = 9
''                    Do
''                         .Text = str(i)
''                         .Col = .ColLetterToNumber("B")
''                         .Row = i + 9
''                         i = i + 1
''                    Loop Until .Text = "aa"
''                    i = 1
''                   .Row = .Row + 3
''                    Do
''                         .Text = str(i)
''                         .Col = .ColLetterToNumber("B")
''                         .Row = 1 + .Row
''                         i = i + 1
''                    Loop Until .Text = "bb"
''                    .SetActiveCell iCol, iRow
''                    .Row = iRow
''                    .Col = iCol
''                    .SetActiveCell iCol, iRow
''                 fps.EventEnabled(EventAllEvents) = True
''            End If
'        End If
'        fps.EventEnabled(EventAllEvents) = True
'    End With

End Sub

'inner class event: Leavecell
Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim tCol As Long, tRow As Long
    Dim cCol As Long, cRow As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    With fps
        mCurrentSheet = .Sheet
        .EventEnabled(EventAllEvents) = False
         If .ActiveSheet = 1 Then
             If Col = .ColLetterToNumber("CB") And Row = rowNgayKy Then
                      .GetText Col, Row, varTemp
                        If varTemp <> "" And varTemp <> "../../...." Then
                            If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                                .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                                .TypeHAlign = TypeHAlignLeft
                            Else
                               .SetFocus
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
'        If .ActiveSheet = 2 Or .ActiveSheet = 3 Or .ActiveSheet = 4 Or .ActiveSheet = 5 Or .ActiveSheet = 6 Or .ActiveSheet = 7 _
'        Or .ActiveSheet = 8 Or .ActiveSheet = 9 Or .ActiveSheet = 10 Or .ActiveSheet = 11 Then
'            If Col = .ColLetterToNumber("G") Then
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
'                .Col = Col
'                .Row = Row
'                UpdateCell fps, .Col, .Row, .Text
'            End If
'            'chuyen cot ky hieu ve chu cai in hoa
'            If Col = .ColLetterToNumber("E") Then
'                 .GetText Col, Row, varTemp
'                 .SetText Col, Row, UCase(varTemp)
'                 .Col = Col
'                 .Row = Row
'                 UpdateCell fps, .Col, .Row, .Text
'            End If
'
'        End If
'KHBS
'        If .ActiveSheet = .SheetCount - 1 Then
'         Dim count1 As Integer
'            .Row = 9
'            .Col = .ColLetterToNumber("B")
'            Do While .value <> "bb"
'                count1 = count1 + 1
'                .Row = .Row + 1
'            Loop
'                  If Col = .ColLetterToNumber("BG") And Row = 46 + count1 - 5 Then
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
'
'        End If
      .EventEnabled(EventAllEvents) = True
    End With
    
    'CellChange Col, Row
End Sub
Private Sub CheckDynamicError()
    'Dim vDoanhSo As Variant, vGTGT As Variant
    Dim vTarget15 As Variant, vTarget16 As Variant
    'Dim vSoHD As Variant, vNgayHD As Variant, vTennguoimua As Variant, vMST As Variant, ErrMsg1NgayPL01 As Variant, ErrMsg2NgayPl01 As Variant
    'Dim vErrMsgMST As Variant, vErr As Variant
    'Dim iFlagNgay1Pl01 As Variant
    'Dim iFlagvKyHieuHD As Integer, iFlagvSoHD As Integer, iFlagvNgayHD As Integer, iFlagvTennguoimua As Integer, iFlagvMst As Integer
    Dim iFlagEmtyCT21 As Integer, iFlagEmtyCT29, iFlagEmtyCT31 As Integer
    Dim iCurrentSheet As Integer
    Dim iFlagNgayKy As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer
        
    Dim strFormula As String
    Dim Errct21 As Variant
    Dim Errct29 As Variant
    Dim Errct31 As Variant
    Dim Errct22 As Variant
    Dim Errct23 As Variant
    Dim errNgayKy As Variant
    
    
    Dim ctCT21 As Variant
    Dim ct29 As Variant
    Dim ctCT29 As Variant
    Dim ct30 As Variant
    Dim CT30a As Variant
    Dim ct31 As Variant
    Dim ct21 As Variant
    Dim ct32 As Variant
    
    
    Dim ngayHienTai As Variant
    Dim strDateNHT As Variant
    Dim CheckNgayKy As Variant
    Dim DateNgayKy As Variant
    
    
    'Dim ErrThueGTGT As Variant
    'iFlagvKyHieuHD = 0
    'iFlagvSoHD = 0
    'iFlagvNgayHD = 0
    'iFlagvTennguoimua = 0
    iFocusFlag = 0
'    Dim blCheck_S2(10) As Boolean
'    Dim blCheck_S2_flag As Boolean
'    blCheck_S2_flag = False
    'Dim vErrMsgThuesuat As Variant
'    Dim iFlagvThuesuat As Integer
    
    'Dim vErrorPhatHanhHD As Variant
    'Dim iflagNgayPhatHanhHD As Integer
    
    'Dim errMauSoHD As Variant, errSoLienHD As Variant
    'Dim errEmty As Variant
    'Dim iFlagEmty As Integer
'    Dim vGhiChu As Variant
    
    'Dim iFlagvThueGTGT As Integer
    'Dim vErrThueGTGT As Variant
    
    
    
    With fps
    iCurrentSheet = .Sheet
    .Sheet = .SheetCount
    CheckErrorMST
         .Sheet = .SheetCount
         .GetText .ColLetterToNumber("E"), 27, errNgayKy
         .GetText .ColLetterToNumber("E"), 25, Errct29
         .GetText .ColLetterToNumber("E"), 30, Errct31
         .GetText .ColLetterToNumber("E"), 26, Errct21
         .GetText .ColLetterToNumber("E"), 15, Errct22
         .GetText .ColLetterToNumber("E"), 16, Errct23
         '.GetText .ColLetterToNumber("E"), 17, vErr
         '.GetText .ColLetterToNumber("E"), 21, vErrMsgMST
         '.GetText .ColLetterToNumber("E"), 24, vErrMsgThuesuat
         '.GetText .ColLetterToNumber("E"), 28, ErrMsg1NgayPL01
         '.GetText .ColLetterToNumber("E"), 29, ErrMsg2NgayPl01
         '.GetText .ColLetterToNumber("E"), 23, ErrThueGTGT
         
         
        '.GetText .ColLetterToNumber("E"), 34, vErrorPhatHanhHD
'        .GetText .ColLetterToNumber("E"), 35, errMauSoHD
'        .GetText .ColLetterToNumber("E"), 36, errSoLienHD
        '.GetText .ColLetterToNumber("E"), 37, errEmty
        '.GetText .ColLetterToNumber("E"), 38, vErrThueGTGT
         'vttoan: bat rang buoc ct21, 30 ,31, 22,23
'             .Sheet = 2
'            .GetText .ColLetterToNumber("N"), 2, ct08pl
'            .GetText .ColLetterToNumber("N"), 3, ct10pl
'            .GetText .ColLetterToNumber("N"), 1, quyPhatHanh
'            .GetText .ColLetterToNumber("N"), 4, namPhatHanh
            .Sheet = 1
            .GetText .ColLetterToNumber("CT"), rowCt21, ct21
            .GetText .ColLetterToNumber("DO"), rowCt21, ctCT21
            .GetText .ColLetterToNumber("CT"), rowCt29, ct29
            .GetText .ColLetterToNumber("DO"), rowCt29, ctCT29
            .GetText .ColLetterToNumber("CT"), rowCt30, ct30
            
            .GetText .ColLetterToNumber("CT"), rowCt30a, CT30a
            
            .GetText .ColLetterToNumber("CT"), rowCt31, ct31
            .GetText .ColLetterToNumber("CT"), rowCt32, ct32
            .GetText .ColLetterToNumber("CB"), rowNgayKy, DateNgayKy
            'chek ngay ky " cho phep sua nhung ko duoc lon hon ngay hien tai"
            ngayHienTai = DateSerial(Mid(Trim(DateNgayKy), 7, 4), Mid(Trim(DateNgayKy), 4, 2), Mid(Trim(DateNgayKy), 1, 2))
            strDateNHT = DateTime.Now
            CheckNgayKy = DateDiff("d", strDateNHT, ngayHienTai)
                If CheckNgayKy > 0 Then
                    .Col = .ColLetterToNumber("CB")
                    .Row = rowNgayKy
                    .CellNote = errNgayKy
                    .BackColor = mErrorColor
                    iFlagNgayKy = 1
                Else
                    .Col = .ColLetterToNumber("CB")
                    .Row = rowNgayKy
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
               If Val(ct21) <> Val(ctCT21) Then
                    .Col = .ColLetterToNumber("CT")
                    .Row = rowCt21
                    .CellNote = Errct21
                    .BackColor = mAlertColor
                    iFlagEmtyCT21 = 1
                Else
                    .Col = .ColLetterToNumber("CT")
                    .Row = rowCt21
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
             'ct30, ct31, 30a
                If Val(ct30) > Val(ct29) - Val(CT30a) Then
                    .Col = .ColLetterToNumber("CT")
                    .Row = rowCt30
                    .CellNote = Errct29
                    .BackColor = mErrorColor
                    iFlagEmtyCT29 = 1
                Else
                    .Col = .ColLetterToNumber("CT")
                    .Row = rowCt30
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                If Val(ct31) > Val(ct29) Then
                    .Col = .ColLetterToNumber("CT")
                    .Row = rowCt31
                    .CellNote = Errct31
                    .BackColor = mErrorColor
                    iFlagEmtyCT31 = 1
                Else
                    .Col = .ColLetterToNumber("CT")
                    .Row = rowCt31
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            
    'iFlagNgay1Pl01 = 0
    'Check hop le cac Sheet PL01_1
'    Dim k As Variant
'    Dim h As Variant
'    h = 0
'    For k = 2 To 11
'    .Sheet = k
'     blCheck_S2(h) = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'     If blCheck_S2(h) = True Then
'            blCheck_S2_flag = True
'            iWhiteRow = 1
'            .Row = 8
'            .Col = .ColLetterToNumber("B")
'            Do
'                .GetText .ColLetterToNumber("F"), .Row, vSoHD
'                .GetText .ColLetterToNumber("G"), .Row, vNgayHD
'                .GetText .ColLetterToNumber("H"), .Row, vTennguoimua
'                .GetText .ColLetterToNumber("K"), .Row, vDoanhSo
'                .GetText .ColLetterToNumber("M"), .Row, vGTGT
'                .GetText .ColLetterToNumber("I"), .Row, vMST
'                .GetText .ColLetterToNumber("N"), .Row, vGhiChu
'
'             If Trim(vSoHD) <> "" Or Trim(vNgayHD) <> "" Or Trim(vTennguoimua) <> "" Or Trim(vDoanhSo) <> "0" Or Trim(vGTGT) <> "0" Or Trim(vGhiChu) <> "" Then iWhiteRow = 0 Else iWhiteRow = 1
'                'vttoan: bat ngay, thang, nam phat hanh theo tai lieu
'                 ' ngay bat dau phat hanh tren phu luc 01 <= ky tinh thue
'                 If Trim(vNgayHD) = "" And iWhiteRow = 0 Then
'                    .Col = .ColLetterToNumber("G")
'                    .CellNote = vErr
'                    .BackColor = mErrorColor
'                    iFlagvNgayHD = 1
'                 ElseIf Trim(vNgayHD) <> "" Then
'                    ' Kiem tra ngay phat hanh nho hon ngay cuoi ky ke khai
'                             If strTkThangQuy = "TK_QUY" Then
'                                 If DateDiff("D", getNgayCuoiKyKK(TAX_Utilities_v1.ThreeMonths, TAX_Utilities_v1.Year, True), Format(vNgayHD, "dd/mm/yyyy")) > 0 Then
'                                     .Col = .ColLetterToNumber("G")
'                                     .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                     .CellNote = vErrorPhatHanhHD   'static
'                                     .BackColor = mErrorColor
'                                     iflagNgayPhatHanhHD = 1
'                                     'capture the first error cell
''                                     If iFocusFlag = 0 Then
''                                         strFocusSheetName = "'" & .SheetName & "'"
''                                         strFocusCol = str(.Col)
''                                         strFocusRow = str(.Row)
''                                         iFocusFlag = 1
''                                     End If
'                                 Else
'                                      .Col = .ColLetterToNumber("G")
'                                      .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                      .CellNote = ""   'static
'                                      .BackColor = mNonErrorColor
'                                End If
'                             Else
'                                If DateDiff("D", getNgayCuoiKyKK(TAX_Utilities_v1.Month, TAX_Utilities_v1.Year, False), Format(vNgayHD, "dd/mm/yyyy")) > 0 Then
'                                     .Col = .ColLetterToNumber("G")
'                                     .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                     .CellNote = vErrorPhatHanhHD   'static
'                                     .BackColor = mErrorColor
'                                     iflagNgayPhatHanhHD = 1
'                                     'capture the first error cell
''                                     If iFocusFlag = 0 Then
''                                         strFocusSheetName = "'" & .SheetName & "'"
''                                         strFocusCol = str(.Col)
''                                         strFocusRow = str(.Row)
''                                         iFocusFlag = 1
''                                     End If
'                                 Else
'                                      .Col = .ColLetterToNumber("G")
'                                      .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                      .CellNote = ""   'static
'                                      .BackColor = mNonErrorColor
'                                End If
'                             End If
'                End If
'
'                 ' so
'                 If Trim(vSoHD) = "" And iWhiteRow = 0 Then
'                     .Col = .ColLetterToNumber("F")
'                     .CellNote = vErr
'                     .BackColor = mErrorColor
'                     iFlagvSoHD = 1
'                 Else
'                     .Col = .ColLetterToNumber("F")
'                     .CellNote = ""
'                     .BackColor = mNonErrorColor
'                 End If
'                 ' ten gnuoi mua
'                 If Trim(vTennguoimua) = "" And iWhiteRow = 0 Then
'                     .Col = .ColLetterToNumber("H")
'                     .CellNote = vErr
'                     .BackColor = mErrorColor
'                     iFlagvTennguoimua = 1
'                  Else
'                     .Col = .ColLetterToNumber("H")
'                     .CellNote = ""
'                     .BackColor = mNonErrorColor
'                 End If
'
'                 ' check thue va doanh so
'                 If Val(vGTGT) <> 0 And Val(vGTGT) >= Val(vDoanhSo) Then
'                        .Col = .ColLetterToNumber("M")
'                        .CellNote = vErrThueGTGT
'                        '.BackColor = mErrorColor
'                        .BackColor = mAlertColor
'                        'iFlagvThueGTGT = 1
'                 Else
'                        .Col = .ColLetterToNumber("M")
'                        .CellNote = ""
'                        .BackColor = mNonErrorColor
'                 End If
'
'
'                 'Check MST
'                 If Trim(vMST) <> vbNullString Then
'                     strCheck = ""
'
'                     If Len(vMST) = 14 And InStr(vMST, "-") > 0 Then
'                        vMST = Left$(vMST, 10) & Right$(vMST, 3)
'                     End If
'                     vMST = Replace(vMST, "-", "")
'
'                     If Len(vMST) = 10 Or Len(vMST) = 13 Then
'                         strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
'                         If Len(vMST) = 10 And strCheck <> "000" Then
'                             .Col = .ColLetterToNumber("I")
'                             .CellNote = vErrMsgMST
'                             .BackColor = mErrorColor
'                             iFlagvMst = 1
'                         ElseIf Len(vMST) = 13 And (strCheck <> "000" Or IsNumeric(vMST) = False) Then
'                             .Col = .ColLetterToNumber("I")
'                             .CellNote = vErrMsgMST
'                             .BackColor = mErrorColor
'                             iFlagvMst = 1
'                         Else
'                             .Col = .ColLetterToNumber("I")
'                             .CellNote = ""
'                             .BackColor = mNonErrorColor
'                         End If
'                     ElseIf Len(vMST) <> 10 Or Len(vMST) <> 13 Then
'                         .Col = .ColLetterToNumber("I")
'                         .CellNote = vErrMsgMST
'                         .BackColor = mErrorColor
'                         iFlagvMst = 1
'                     End If
'                 Else
'                     .Col = .ColLetterToNumber("AA")
'                     .CellNote = ""
'                     .BackColor = mNonErrorColor
'                 End If
'
'            .Col = .ColLetterToNumber("B")
'            .Row = .Row + 1
'             Loop Until UCase(.Text) = "AA"
'        End If
'    h = h + 1
'  Next
            .Sheet = .SheetCount
            '.SetText .ColLetterToNumber("B"), 17, IIf(iFlagvKyHieuHD = 1, "0", "1")
            '.SetText .ColLetterToNumber("B"), 18, IIf(iFlagvSoHD = 1, "0", "1")
            '.SetText .ColLetterToNumber("B"), 19, IIf(iFlagvNgayHD = 1, "0", "1")
            '.SetText .ColLetterToNumber("B"), 20, IIf(iFlagvTennguoimua = 1, "0", "1")
            '.SetText .ColLetterToNumber("B"), 21, IIf(iFlagvMst = 1, "0", "1")
            '.SetText .ColLetterToNumber("B"), 24, IIf(iFlagvThuesuat = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 25, IIf(iFlagEmtyCT29 = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 30, IIf(iFlagEmtyCT31 = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 26, IIf(iFlagEmtyCT21 = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 27, IIf(iFlagNgayKy = 1, "0", "1")
            '.SetText .ColLetterToNumber("B"), 27, IIf(iFlagGTGT = 1, "0", "1")
            '.SetText .ColLetterToNumber("B"), 28, IIf(iFlagNgay1Pl01 = 1, "0", "1")
            
            '.SetText .ColLetterToNumber("B"), 34, IIf(iflagNgayPhatHanhHD = 1, "0", "1")
'            .SetText .ColLetterToNumber("B"), 35, IIf(iFlagMauSoHD = 1, "0", "1")
'            .SetText .ColLetterToNumber("B"), 36, IIf(iFlagSoLienHD = 1, "0", "1")
            '.SetText .ColLetterToNumber("B"), 37, IIf(iFlagEmty = 1, "0", "1")
            '.SetText .ColLetterToNumber("B"), 38, IIf(iFlagvThueGTGT = 1, "0", "1")
            
            
            
            If iFocusFlag = 1 Then
                ' setcell for focus
                strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
                .Col = .ColLetterToNumber("C")
                .Row = 17
                .Formula = strFormula
                .Row = 18
                .Formula = strFormula
                .Row = 19
                .Formula = strFormula
                .Row = 20
                .Formula = strFormula
                .Row = 21
                .Formula = strFormula
                .Row = 29
                .Formula = strFormula
            End If
         
            .GetText .ColLetterToNumber("B"), 15, vTarget15
            .GetText .ColLetterToNumber("B"), 16, vTarget16
         
'            If blCheck_S2_flag Then
'                .SetText .ColLetterToNumber("B"), 15, IIf(vTarget15 = "1", "1", "0")
'                .SetText .ColLetterToNumber("B"), 16, IIf(vTarget16 = "1", "1", "0")
'            Else
'               .SetText .ColLetterToNumber("B"), 15, "1"
'               .SetText .ColLetterToNumber("B"), 16, "1"
'            End If
            .Sheet = iCurrentSheet
    End With

End Sub
'This function is used to support report
'some special data that we want to save to data files
Public Sub SaveExceptionData()
    Dim i As Long, DataText As Variant
    Dim arrRow As Variant, arrCol As Variant
    Dim arrCellTag() As String
    With fps
        .Sheet = 1
        mCurrentSheet = .Sheet
        'J10 , L10, L16, L17, L20
        arrRow = Array(10, 10, 16, 17, 20)
        arrCol = Array("J", "L", "L", "L", "L")
        For i = 0 To UBound(arrRow)
            .Row = arrRow(i)
            .Col = .ColLetterToNumber(arrCol(i))
            UpdateCell fps, .Col, .Row, .value
        Next


    End With
                  
End Sub

'Calling this function whenever we don't need this class anymore.
'It will do some special task with special Error/ special Data
'no parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim vTemp As Variant
    Dim iCol As Long, iRow As Long
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        
        SaveExceptionData 'Save Exception Data to DOM
         
'         mCurrentSheet = 2
'        .Sheet = mCurrentSheet
'
'        Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
'        (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 1)
'         ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'         .GetText iCol, iRow, vTemp
'         UpdateCell fps, iCol, iRow, vTemp
'
'         Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
'        (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 2)
'          ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'         .GetText iCol, iRow, vTemp
'         UpdateCell fps, iCol, iRow, vTemp
         
        CheckDynamicError 'Set Exception Error on cells of interface
        UpdateSheet
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub


Private Sub setLast25Target()
    
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    'Dim strOriginDataFileName As String
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim str25Target As String
    Dim iDom As Integer
    Dim MY As Date
    
    iDom = 0
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        If strTkThangQuy = "TK_THANG" Then
        MY = DateSerial(TAX_Utilities_v1.Year, TAX_Utilities_v1.Month, 1)
        
        MY = DateAdd("M", -1, MY)
        'strOriginDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(1), "DataFile") & ".xml"
        strCurrentDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(iDom), "DataFile") & "_" & TAX_Utilities_v1.Month & TAX_Utilities_v1.Year & ".xml"
        '***************************************
        'ThanhDX added
        'Date: 14/04/06
        
        'strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(GetValidityNode("02", Format(Month(MY), "0#"), , Format(Year(MY), "####")).childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        '***************************************
        Else
            MY = dNgayDauQuy
            
            MY = DateAdd("D", -1, MY)
            'strOriginDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(1), "DataFile") & ".xml"
            strCurrentDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(iDom), "DataFile") & "_Q0" & TAX_Utilities_v1.ThreeMonths & TAX_Utilities_v1.Year & ".xml"
            '***************************************
            'strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(iDom), "DataFile") & "_Q0" & DatePart("Q", MY) & DatePart("YYYY", MY) & ".xml"
            '***************************************
        End If
        
        If xmlDomCurrentData.Load(strCurrentDataFileName) = True Then
            GoTo Tidy_up
        End If
        If .Load(strDataFileName) = True Then
            Set xmlNode = .nodeFromID("CT_" & rowCt32) 'L_20: col & row of 25 target
            str25Target = GetAttribute(xmlNode, "Value")
            Set xmlNode = Nothing
            
            If Trim(str25Target) <> "" And Trim(str25Target) <> "0" Then
                fps.Sheet = 1
                fps.Col = fps.ColLetterToNumber("CT")
                fps.Row = rowCt21
                fps.Col = fps.ColLetterToNumber("DO")
                fps.Row = rowCt21
                If Trim(fps.Text) = "0" Or Trim(fps.Text) = "" Then
                    fps.Text = str25Target
                    mCurrentSheet = 1
                    UpdateCell fps, fps.Col, fps.Row, fps.value
                    fps.Col = fps.ColLetterToNumber("CT")
                    fps.Row = rowCt21
                    UpdateCell fps, fps.Col, fps.Row, fps.value
                    fps.Col = fps.ColLetterToNumber("DO")
                    fps.Row = rowCt21
                    UpdateCell fps, fps.Col, fps.Row, fps.value
                End If
            End If
            
        ElseIf .parseError.reason <> vbNullString Then
            If InStr(1, .parseError.errorCode, "2146697210") <> 0 Then
                'file data of last month does not exist
            Else
                MsgBox .parseError.reason
                Exit Sub
            End If
        End If
        
    End With
    
Tidy_up:
    Set xmlDomLastData = Nothing
    Set xmlDomCurrentData = Nothing
    Set xmlNodeListCell = Nothing
    Set xmlNode = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_02GTGT", "setLast25Target", Err.number, Err.Description
End Sub

Sub CheckErrorMST()
    Dim vError1 As Variant, vError2 As Variant, vError3 As Variant
    Dim MST1 As Variant, MST2 As Variant, MST3 As Variant, MST4 As Variant
    Dim MST5 As Variant, MST6 As Variant, MST7 As Variant, MST8 As Variant
    Dim MST9 As Variant, MST10 As Variant, MST11 As Variant, MST12 As Variant, MST13 As Variant
    Dim iCurrentSheet As Integer, strCheck As String
    Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    With fps
        
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 22, vError1
        .GetText .ColLetterToNumber("E"), 23, vError2
        .GetText .ColLetterToNumber("E"), 24, vError3
        
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
            .Row = 22
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 22
            .Text = "1"
        End If
        
        If iFlagTaxCode2 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 23
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 23
            .Text = "1"
        End If
        
        
        If iFlagTaxCode3 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 24
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 24
            .Text = "1"
        End If
        
        .Sheet = 1
        .Col = .ColLetterToNumber("D")
        .Row = 3
        .CellNote = ""
        '.BackColor = mFormColor
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

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
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

'With fps
'   If .ActiveSheet = .SheetCount - 1 Then
'            'lay ve dong combobox
'            Set xmlCellNode = TAX_Utilities_v1.Data(.SheetCount - 2).nodeFromID(GetCellID(fps, .ColLetterToNumber("BI"), 5))
'            Set xmlCellNode = xmlCellNode.parentNode
'            Set xmlCellNode = xmlCellNode.parentNode
'            Set xmlCellNode = xmlCellNode.parentNode.lastChild
'            Set xmlCellNode = xmlCellNode.firstChild
'            Set xmlCellNode = xmlCellNode.firstChild
'            For i = 1 To 6
'                Set xmlCellNode = xmlCellNode.nextSibling
'            Next i
'            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'
'
'            .Sheet = .SheetCount - 1
'            mCurrentSheet = .SheetCount - 1
'
'            .Col = Col
'            .Row = Row
'
'            If Col = .ColLetterToNumber("BE") And Row = iRow Then 'Khi chon Combo
'                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
'                intIndexCuc = .TypeComboBoxCurSel
'                .Col = .ColLetterToNumber("BI")
'                .Row = Row
'                .TypeComboBoxCurSel = intIndexCuc
'                'Nhap lai Combo cot Q
'                .Col = .ColLetterToNumber("BI")
'                UpdateCell fps, .Col, .Row, .Text
'                ' Lay Id cua cuc thue vua chon o tren
'                .GetText .Col, .Row, varIdCucThue
'
'                If varIdCucThue <> "" Or varIdCucThue <> vbNullString Then
'                    ' Lay thong tin cho chi cuc thue
'                    varIdCucThue = Left$(varIdCucThue, 3)
'
'                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
'                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
'                        For Each xmlNode In xmlNodeListCell
'                            If GetAttribute(xmlNode, "Value") <> "" Then
'                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
'
'                                If arrDanhsach(0) = "0" And arrDanhsach(2) = varIdCucThue Then
'                                    tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
'                                    maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
'                                End If
'                            End If
'                        Next
'                        Set xmlDomData = Nothing
'                        Set xmlNodeListCell = Nothing
'                        Set xmlNode = Nothing
'                    End If
'
'                    .Row = iRow + 2
'                    .Col = .ColLetterToNumber("BE")
'                    .TypeComboBoxList = tenChiCucThue
'
'                    .Row = iRow + 2
'                    .Col = .ColLetterToNumber("BI")
'                    .TypeComboBoxList = maChiCucThue
'                End If
'            End If
'
'            If Col = .ColLetterToNumber("BE") And Row = iRow + 2 Then 'Khi chon Combo
'                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
'                intIndexChiCuc = .TypeComboBoxCurSel
'                .Col = .ColLetterToNumber("BI")
'                .Row = Row
'                .TypeComboBoxCurSel = intIndexChiCuc
'                'Nhap lai Combo cot Q
'                .Col = .ColLetterToNumber("BI")
'                UpdateCell fps, .Col, .Row, .Text
'            End If
'
'
''         If Col = .ColLetterToNumber("BD") And .Text <> "" Then
''            intIndexCombo = .TypeComboBoxCurSel
''            .Col = .ColLetterToNumber("BC")
''            .TypeComboBoxCurSel = intIndexCombo
''            'Update combo C
''            UpdateCell fps, .Col, .Row, .Text
''            'Chon lai combo cua colF
''            .Col = .ColLetterToNumber("BE")
''            .GetText .ColLetterToNumber("BC"), Row, strId
''            .Text = Trim(Right(strId, 10))
''            UpdateCell fps, .Col, .Row, .Text
''        End If
''         If Col = .ColLetterToNumber("BE") And .Text <> "" Then
''            intIndexCombo = .TypeComboBoxCurSel
''            .Col = .ColLetterToNumber("BC")
''            .TypeComboBoxCurSel = intIndexCombo
''            'Update combo C
''            UpdateCell fps, .Col, .Row, .Text
''            'Chon lai combo cua colF
''            .Col = .ColLetterToNumber("BD")
''            .GetText .ColLetterToNumber("BC"), Row, strId
''            .Text = Trim(Left(strId, 100))
''            UpdateCell fps, .Col, .Row, .Text
''        End If
'
'
''        Set xmlCellNode = TAX_Utilities_v1.Data(.SheetCount - 2).nodeFromID(GetCellID(fps, Col, Row))
''        Set xmlCellNode = xmlCellNode.parentNode.parentNode.firstChild.firstChild
''        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
''        blGiaTriCu = False
''        If Col = fps.ColLetterToNumber("BD") Or Col = fps.ColLetterToNumber("BE") Then
''            fps.Col = .ColLetterToNumber("BC")
''            fps.Row = Row
''            temp = fps.Text
''            Set xmlCellsNode = xmlCellNode.parentNode
''            i = 1
''            .Col = .ColLetterToNumber("B")
''            ParserCellID fps, "BD_9", iCol, iRow
''            .Row = iRow
''            Do
''                 .Col = .ColLetterToNumber("BC")
''                 itemp = .Text
''                 If .Row <> Row Then
''                    If itemp = temp Then 'Neu co thi dat co blGiaTriCu = True
''                        RowGiaTriCu = Row
''                        ColGiaTriCu = .ColLetterToNumber("BD")
''                        .Row = Row
''                        .Col = .ColLetterToNumber("BD")
''                        blGiaTriCu = True
''                        Exit Sub
''                    End If
''                 End If
''                 .Col = .ColLetterToNumber("B")
''                 .Row = i + iRow
''                 i = i + 1
''            Loop Until .Text = "bb"
''       End If
'
'    End If
    'CellChange Col, Row
'End With
'    TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
End Sub



'******************************
'Description: Set active properties in validity node
Public Sub SetActiveSheet() 'ByRef xmlNodeValid As MSXML.IXMLDOMNode)
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
'             Case "01_2"
'                blnNullValue = True
'                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 2)
'                    strValue2 = GetAttribute(xmlCellNode, "Value")
'
'                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                            blnNullValue = False
'                    End If
'
'                    If blnNullValue Then
'                        SetAttribute xmlNodeValid, "Active", "0"
'                    End If
'                Case "01_2_1"
'                blnNullValue = True
'                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_v1.Data(2).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(2).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_v1.Data(2).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(2).getElementsByTagName("Cell").length - 2)
'                    strValue2 = GetAttribute(xmlCellNode, "Value")
'
'                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                            blnNullValue = False
'                    End If
'
'                    If blnNullValue Then
'                        SetAttribute xmlNodeValid, "Active", "0"
'                    End If
'                Case "01_2_2"
'                blnNullValue = True
'                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_v1.Data(3).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(3).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_v1.Data(3).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(3).getElementsByTagName("Cell").length - 2)
'                    strValue2 = GetAttribute(xmlCellNode, "Value")
'
'                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                            blnNullValue = False
'                    End If
'
'                    If blnNullValue Then
'                        SetAttribute xmlNodeValid, "Active", "0"
'                    End If
'                Case "01_2_3"
'                blnNullValue = True
'                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_v1.Data(4).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(4).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_v1.Data(4).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(4).getElementsByTagName("Cell").length - 2)
'                    strValue2 = GetAttribute(xmlCellNode, "Value")
'
'                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                            blnNullValue = False
'                    End If
'
'                    If blnNullValue Then
'                        SetAttribute xmlNodeValid, "Active", "0"
'                    End If
'                Case "01_2_4"
'                blnNullValue = True
'                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_v1.Data(5).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(5).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_v1.Data(5).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(5).getElementsByTagName("Cell").length - 2)
'                    strValue2 = GetAttribute(xmlCellNode, "Value")
'
'                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                            blnNullValue = False
'                    End If
'
'                    If blnNullValue Then
'                        SetAttribute xmlNodeValid, "Active", "0"
'                    End If
'                Case "01_2_5"
'                blnNullValue = True
'                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_v1.Data(6).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(6).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_v1.Data(6).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(6).getElementsByTagName("Cell").length - 2)
'                    strValue2 = GetAttribute(xmlCellNode, "Value")
'
'                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                            blnNullValue = False
'                    End If
'
'                    If blnNullValue Then
'                        SetAttribute xmlNodeValid, "Active", "0"
'                    End If
'                Case "01_2_6"
'                blnNullValue = True
'                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_v1.Data(7).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(7).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_v1.Data(7).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(7).getElementsByTagName("Cell").length - 2)
'                    strValue2 = GetAttribute(xmlCellNode, "Value")
'
'                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                            blnNullValue = False
'                    End If
'
'                    If blnNullValue Then
'                        SetAttribute xmlNodeValid, "Active", "0"
'                    End If
'                    Case "01_2_7"
'                blnNullValue = True
'                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_v1.Data(8).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(8).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_v1.Data(8).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(8).getElementsByTagName("Cell").length - 2)
'                    strValue2 = GetAttribute(xmlCellNode, "Value")
'
'                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                            blnNullValue = False
'                    End If
'
'                    If blnNullValue Then
'                        SetAttribute xmlNodeValid, "Active", "0"
'                    End If
'                Case "01_2_8"
'                blnNullValue = True
'                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_v1.Data(9).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(9).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_v1.Data(9).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(9).getElementsByTagName("Cell").length - 2)
'                    strValue2 = GetAttribute(xmlCellNode, "Value")
'
'                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                            blnNullValue = False
'                    End If
'
'                    If blnNullValue Then
'                        SetAttribute xmlNodeValid, "Active", "0"
'                    End If
'                Case "01_2_9"
'                blnNullValue = True
'                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_v1.Data(10).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(10).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_v1.Data(10).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v1.Data(10).getElementsByTagName("Cell").length - 2)
'                    strValue2 = GetAttribute(xmlCellNode, "Value")
'
'                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                            blnNullValue = False
'                    End If
'
'                    If blnNullValue Then
'                        SetAttribute xmlNodeValid, "Active", "0"
'                    End If
                Case "KHBS"
                    blnNullValue = True
                    'Check value of last node in Data
                    Set xmlCellNode = TAX_Utilities_v1.Data(fps.SheetCount - 2).getElementsByTagName("Cell") _
                        (TAX_Utilities_v1.Data(fps.SheetCount - 2).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Then
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

'******************************
'Description: Reset all error satus in sheet contain dynamic error.
'Author: ThanhDX
'Date: 31/12/2005

'******************************
Public Sub ResetErrorCells()
'    Dim lRow As Long, lCtrl As Long
'    Dim lBColumnColor As Long, lCColumnColor As Long
'    Dim lDColumnColor As Long, lEColumnColor As Long
'    Dim lFColumnColor As Long
'    Dim intCurrentSheet As Integer
'
'    With fps
'    intCurrentSheet = .Sheet
'     'dhdang sua yc them nhieu Sheet PL 01_1
'    Dim k As Variant
'    For k = 2 To 11
'        .Sheet = k
'        For lCtrl = 1 To .MaxRows
'            .Col = .ColLetterToNumber("C")
'            .Row = lCtrl
'            If Not .Lock Then
'                .CellNote = ""
'                .BackColor = vbWhite
'            End If
'
'            .Col = .ColLetterToNumber("D")
'            .Row = lCtrl
'            If Not .Lock Then
'                .CellNote = ""
'                .BackColor = vbWhite
'            End If
'
'            .Col = .ColLetterToNumber("E")
'            .Row = lCtrl
'            If Not .Lock Then
'                .CellNote = ""
'                .BackColor = vbWhite
'            End If
'
'            .Col = .ColLetterToNumber("F")
'            .Row = lCtrl
'            If Not .Lock Then
'                .CellNote = ""
'                .BackColor = vbWhite
'            End If
'            .Col = .ColLetterToNumber("G")
'            .Row = lCtrl
'            If Not .Lock Then
'                .CellNote = ""
'                .BackColor = vbWhite
'            End If
'        Next lCtrl
'        .Sheet = intCurrentSheet
'    Next
'End With
End Sub

Private Function GetLast25Previous(ByRef strValue As String) As Boolean
     
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim strDataFileName As String
    Dim str43Target As String
    Dim iDom As Integer
    Dim MY As Date
    
    iDom = 0
    GetLast25Previous = True
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        If strTkThangQuy = "TK_THANG" Then
        MY = DateSerial(TAX_Utilities_v1.Year, TAX_Utilities_v1.Month, 1)
        MY = DateAdd("M", -1, MY)
        
        '***************************************
        'ThanhDX added
        'Date: 14/04/06
        
        'strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(GetValidityNode("02", Format(Month(MY), "0#"), , Format(Year(MY), "####")).childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        '***************************************
        Else
            MY = dNgayDauQuy
            
            MY = DateAdd("D", -1, MY)
            
            'strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(iDom), "DataFile") & "_Q0" & DatePart("Q", MY) & DatePart("YYYY", MY) & ".xml"
            '***************************************
        End If
                
        If .Load(strDataFileName) = True Then
            Set xmlNode = .nodeFromID("CT_" & rowCt32) 'L_20: col & row of 25 target
            strValue = GetAttribute(xmlNode, "Value")
            Set xmlNode = Nothing
        ElseIf .parseError.reason <> vbNullString Then
            If InStr(1, .parseError.errorCode, "2146697210") <> 0 Then
                'file data of last month does not exist
            Else
                MsgBox .parseError.reason
            End If
            GetLast25Previous = False
        End If
        
    End With
    
    Set xmlDomLastData = Nothing
    Set xmlNode = Nothing
    
    Exit Function
    
ErrorHandle:
    SaveErrorLog "cls_02GTGT", "GetLast25Previous", Err.number, Err.Description
End Function

'*******************************************************
'Description: ResetData procedure reset specified cells
'Author:ThanhDX
'Date:23/01/2006

'*******************************************************
Public Sub ResetData()
'    Dim intCurSheet As Integer, intCtrl As Long
'    Dim lCol As Long, lRow As Long
'    Dim arrStrCellUnLock() As String
'    Dim strCell As String
'
'    'Array of reset cells
'    arrStrCellUnLock = Split("J_11,J_12,L_11,L_12,J_20,J_22,J_23,L_23,J_24,L_24", ",")
'
'    With fps
'        intCurSheet = .Sheet
'        .Sheet = 1
'        mCurrentSheet = .Sheet
'        For intCtrl = 0 To UBound(arrStrCellUnLock)
'            strCell = arrStrCellUnLock(intCtrl)
'            ParserCellID fps, strCell, lCol, lRow
'            .Col = lCol
'            .Row = lRow
'            .Lock = False
'        Next
'        .Sheet = intCurSheet
'        mCurrentSheet = .Sheet
'    End With
End Sub

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()
'        'baclt added
'        If fps.ActiveSheet = 4 Then
'            mCurrentSheet = 4
'            fps.Sheet = mCurrentSheet
'                With fps
'                    .Col = fps.ColLetterToNumber(S4TxtF2Col)
'                    .Row = S4TxtF2row
'                    .Text = TAX_Utilities_v1.Year - 1
'                    .TypeHAlign = TypeHAlignLeft
'                    UpdateCell fps, .Col, .Row, .Text
'                End With
'        End If
End Sub

Public Sub Prepare_KHBS()
'    Dim xmlDocument As New MSXML.DOMDocument
'    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
'    Dim xmlNode As MSXML.IXMLDOMNode
'    Dim xmlNodeItem As MSXML.IXMLDOMNode
'    Dim strPath As String
'    Dim strComboName As String
'    Dim strComboMCT As String
'    Dim strCombo As String
'
'    mblKhoiDauToKhai = True
'    fps.Sheet = fps.SheetCount - 1
'    fps.ActiveSheet = fps.Sheet
'
'    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT.xml"
'
'    If xmlDocument.Load(strPath) Then
'        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
'        For Each xmlNode In xmlNodeListItem
'            If GetAttribute(xmlNode, "ID") = GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") Then
'                Exit For
'            End If
'        Next
'    End If
'
'
'    Set xmlNodeListItem = xmlNode.childNodes
'
'    For Each xmlNodeItem In xmlNodeListItem
'                strCombo = strCombo + CPab(GetAttribute(xmlNodeItem, "Caption"), 100) + CPab(GetAttribute(xmlNodeItem, "CellID"), 20) + GetAttribute(xmlNodeItem, "MCT") + Chr$(9)
'                strComboName = strComboName + GetAttribute(xmlNodeItem, "Caption") + Chr$(9)
'                strComboMCT = strComboMCT + GetAttribute(xmlNodeItem, "MCT") + Chr$(9)
'    Next
'     With fps
'
'        .Sheet = .ActiveSheet
'        mCurrentSheet = .Sheet
'        .Col = .ColLetterToNumber("BC")
'        .Row = 9
'        .TypeComboBoxList = strCombo
'        .Row = 13
'        .TypeComboBoxList = strCombo
'        .Col = .ColLetterToNumber("BD")
'        .Row = 9
'        .TypeComboBoxList = strComboName
'        .Row = 13
'        .TypeComboBoxList = strComboName
'        .Col = .ColLetterToNumber("BE")
'        .Row = 9
'        .TypeComboBoxList = strComboMCT
'        .Row = 13
'        .TypeComboBoxList = strComboMCT
'    End With
    
End Sub




Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
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
'        If DateAdd("m", 6, dDate) < dDateNow Then
'        CheckValidNgayHD = False
'            Exit Function
'        End If
'   CheckValidNgayHD = True
'   Exit Function
'
'e:
'    CheckValidNgayHD = False
' End Function


Public Sub UpdateChangeKHBS()
'    Dim varTemp As Variant
'    Dim lCol As Long, lRow As Long
'    Dim xmlCellNode As MSXML.IXMLDOMNodeList
'    Dim xmlCellNodeData As MSXML.IXMLDOMNode
'    With fps
'        .Sheet = 1
'        For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
'            ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'            .GetText lCol, lRow, varTemp
'            UpdateLastKHBSCell fps, lCol, lRow, varTemp
'        Next
'        .Sheet = mCurrentSheet
'    End With
'    Set xmlCellNodeData = Nothing
'    Set xmlCellNode = Nothing
End Sub


Public Sub FinishImport()
'Dim i As Integer, iCol As Long, iRow As Long
'Dim varTemp As Variant
' With fps
'    fps.EventEnabled(EventAllEvents) = False
'    If .ActiveSheet = 2 Or .ActiveSheet = 3 Or .ActiveSheet = 4 Or .ActiveSheet = 5 Or .ActiveSheet = 6 Or .ActiveSheet = 7 _
'        Or .ActiveSheet = 8 Or .ActiveSheet = 9 Or .ActiveSheet = 10 Or .ActiveSheet = 11 Then
'                fps.EventEnabled(EventAllEvents) = False
'                i = 1
'                .Col = .ColLetterToNumber("B")
'                .Row = 8
'                Do
'                     .Text = str(i)
'                     ' chuyen ky hieu thanh chu hoa
'                    .Col = .ColLetterToNumber("E")
'                    .GetText .Col, .Row, varTemp
'                    .Text = UCase$(CStr(varTemp))
'                    UpdateCell fps, .Col, .Row, .Text
'                    ' end
'                     .Col = .ColLetterToNumber("B")
'                     .Row = i + 8
'                     i = i + 1
'                Loop Until .Text = "aa"
'                .SetActiveCell iCol, iRow
'                .Row = iRow
'                .Col = iCol
'                .SetActiveCell iCol, iRow
'                fps.EventEnabled(EventAllEvents) = False
'
'                Dim xmlCellNode As MSXML.IXMLDOMNode
'                Dim strValue1  As Variant, strValue2 As Variant
'                Dim varTemp1 As Variant, varTemp2 As Variant
'                       mCurrentSheet = .Sheet
'                        Set xmlCellNode = TAX_Utilities_v1.Data(.ActiveSheet - 1).getElementsByTagName("Cell") _
'                                      (TAX_Utilities_v1.Data(.ActiveSheet - 1).getElementsByTagName("Cell").length - 2)
'                        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'                        .Col = iCol
'                        .Row = iRow
'                        strValue1 = .value
'                        UpdateCell fps, .Col, .Row, .value
'                        Set xmlCellNode = TAX_Utilities_v1.Data(.ActiveSheet - 1).getElementsByTagName("Cell") _
'                                      (TAX_Utilities_v1.Data(.ActiveSheet - 1).getElementsByTagName("Cell").length - 1)
'                        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'                        .Col = iCol
'                        .Row = iRow
'                        strValue2 = .value
'                        UpdateCell fps, .Col, .Row, .value
'                        If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                                 fps.EventEnabled(EventAllEvents) = False
'                                .Sheet = 1
'                                mCurrentSheet = .Sheet
'                                'Chi tieu [11]
'                                .GetText .ColLetterToNumber("C"), 24, varTemp1
'                                .Col = .ColLetterToNumber("J")
'                                .Row = 10
'                                .value = varTemp1
'                                UpdateCell fps, .Col, .Row, .value
'                                'Chi tieu [12]
'                                .GetText .ColLetterToNumber("D"), 24, varTemp2
'                                .Col = .ColLetterToNumber("L")
'                                .Row = 10
'                                .value = varTemp2
'                                UpdateCell fps, .Col, .Row, .value
'                               .Sheet = .ActiveSheet
'                                mCurrentSheet = .Sheet
'                                fps.EventEnabled(EventAllEvents) = True
'                          End If
'
'                    CellChange iCol, iRow
'    End If
'    fps.EventEnabled(EventAllEvents) = True
' End With
End Sub

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
            '[21], [21a], [23], [28a], [30], [30a], [31]
            Set xmlNode = .nodeFromID("CT_" & rowCt21)
            result = "CT21_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("CT_" & rowCt21a)
            result = result & "~" & "CT21a_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("CT_" & rowCt23)
            result = result & "~" & "CT23_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("CT_" & rowCt28a)
            result = result & "~" & "CT28a_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("CT_" & rowCt30)
            result = result & "~" & "CT30_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("CT_" & rowCt30a)
            result = result & "~" & "CT30a_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("CT_" & rowCt31)
            result = result & "~" & "CT31_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_02GTGT", "getValueTK", Err.number, Err.Description
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
                '[21], [21a], [23], [28a], [30], [30a], [31]
                .GetText .ColLetterToNumber("CT"), rowCt21, varTemp
                strCurrentValue = "CT21_" & Trim(varTemp)
                .GetText .ColLetterToNumber("CT"), rowCt21a, varTemp
                strCurrentValue = strCurrentValue & "~CT21a_" & Trim(varTemp)
                .GetText .ColLetterToNumber("CT"), rowCt23, varTemp
                strCurrentValue = strCurrentValue & "~CT23_" & Trim(varTemp)
                .GetText .ColLetterToNumber("CT"), rowCt28a, varTemp
                strCurrentValue = strCurrentValue & "~CT28a_" & Trim(varTemp)
                .GetText .ColLetterToNumber("CT"), rowCt30, varTemp
                strCurrentValue = strCurrentValue & "~CT30_" & Trim(varTemp)
                .GetText .ColLetterToNumber("CT"), rowCt30a, varTemp
                strCurrentValue = strCurrentValue & "~CT30a_" & Trim(varTemp)
                .GetText .ColLetterToNumber("CT"), rowCt31, varTemp
                strCurrentValue = strCurrentValue & "~CT31_" & Trim(varTemp)
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
                'If Left(strMaCT, 3) = "21_" Or Left(strMaCT, 3) = "23_" Or Left(strMaCT, 3) = "28a" Or Left(strMaCT, 3) = "30_" Or Left(strMaCT, 3) = "31_" Then
                '[21], [21a], [23], [28a], [30], [30a], [31]
                If Left(strMaCT, 3) = "21_" Or Left(strMaCT, 3) = "21a" Or Left(strMaCT, 3) = "23_" Or Left(strMaCT, 3) = "28a" Or Left(strMaCT, 3) = "30_" Or Left(strMaCT, 3) = "30a" Or Left(strMaCT, 3) = "31_" Then
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
                        ' dieu chinh giam
                        strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
                        ' dieu chinh tang
                        strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                Else
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
                        ' dieu chinh giam
                        strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
                        ' dieu chinh tang
                        strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
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
    SaveErrorLog "02_GTGT", "getMapCT", Err.number, Err.Description
End Function


'When cell's data is changed or a control at a cell is handled, This sub is called
'parameter Col: the column of cell that rise event
'parameter Row: the row of cell that rise event
'parameter f: the flag to mark some special option
'**: when this sub is on executing, all events of fpspread will be disabled until the end of the sub
'Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
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
'    .Sheet = .ActiveSheet
'    ' Kiem tra neu co phu luc moi set cong thuc tong hop so lieu tu tinh
'    Dim tempCurrSheet As Integer
'    Dim blCheck_S2 As Boolean
'    Dim blCheck_S2_1 As Boolean
'    Dim blCheck_S2_2 As Boolean
'    Dim blCheck_S2_3 As Boolean
'    Dim blCheck_S2_4 As Boolean
'    Dim blCheck_S2_5 As Boolean
'    Dim blCheck_S2_6 As Boolean
'    Dim blCheck_S2_7 As Boolean
'    Dim blCheck_S2_8 As Boolean
'    Dim blCheck_S2_9 As Boolean
'    blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S2_1 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S2_2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(3).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S2_3 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S2_4 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S2_5 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(6).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S2_6 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(7).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S2_7 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(8).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S2_8 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(9).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S2_9 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(10).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'
'    tempCurrSheet = mCurrentSheet
'
'    If (blCheck_S2 = True) Or (blCheck_S2_1 = True) Or (blCheck_S2_2 = True) Or (blCheck_S2_3 = True) Or (blCheck_S2_4 = True) Or (blCheck_S2_5 = True) Or (blCheck_S2_6 = True) Or (blCheck_S2_7 = True) Or (blCheck_S2_8 = True) Or (blCheck_S2_9 = True) Then
'        .Sheet = 1
'        ' Chi tieu 22
'        .Col = .ColLetterToNumber("BW")
'        .Row = rowCt22
'        .Lock = True
'        .Formula = "H" & rowHidden62
'        ' Chi tieu 23
'        .Col = .ColLetterToNumber("CT")
'        .Row = rowCt23
'        .Lock = True
'        .Formula = "S" & rowHidden62
'
'    Else
'        .Sheet = 1
'        ' chi tieu 22
'        .Col = .ColLetterToNumber("BW")
'        .Row = rowCt22
'        .Lock = False
'        .Formula = ""
'        ' chi tieu 23
'        .Col = .ColLetterToNumber("CT")
'        .Row = rowCt23
'        .Lock = False
'        .Formula = ""
'
'    End If
'
'    mCurrentSheet = tempCurrSheet
'
'    .EventEnabled(EventAllEvents) = True
'End With
'End Sub
Public Sub UpdateSheet()
    Dim varTemp As Variant
    Dim temp As Integer
    Dim blCheck_S2 As Boolean
    Dim blCheck_S8 As Boolean
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

        'dhdang sua yc them nhieu Sheet PL 01_1
'    Dim k As Variant
'    For k = 2 To 11
'        .Sheet = k
'        mCurrentSheet = k
'        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'        If blCheck_S2 = True Then
'            For Each xmlCellNodeData In TAX_Utilities_v1.Data(k - 1).getElementsByTagName("Cell")
'                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'                .GetText lCol, lRow, varTemp
'                UpdateCell fps, lCol, lRow, varTemp
'            Next
'
'            ' xu ly truong hop datafile ban 320 se set encode cell dau tien = 0 (bo ma hoa don)
'            Set xmlCellNodeData4 = TAX_Utilities_v1.Data(k - 1).getElementsByTagName("Sections").Item(0)
'            temp1 = GetAttribute(xmlCellNodeData4, "Version")
'            If Val(temp1) = 320 Then
'                For Each xmlCellNodeData1 In TAX_Utilities_v1.Data(k - 1).getElementsByTagName("Section")
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
'         End If
'    Next
'        .Sheet = .SheetCount - 1
'        mCurrentSheet = .Sheet
'        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'        If blCheck_S2 = True Then
'            For Each xmlCellNodeData In TAX_Utilities_v1.Data(2).getElementsByTagName("Cell")
'                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'                .GetText lCol, lRow, varTemp
'                UpdateCell fps, lCol, lRow, varTemp
'            Next
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

' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("CT"), rowCt32, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function

' Lay ve chi tieu 32 o to khai lan truoc
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
            Set xmlNode = .nodeFromID("CT_" & rowCt32)
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_02GTGT", "getValueCTDC", Err.number, Err.Description
End Function
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
                    Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
                        (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length > 31 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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

' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 22
    denRowDL = 32
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "20_30_2"
    ' end
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

' set co quan thue ra quyet dinh hoan thue
Private Sub setCQTQuanLy()
        Dim xmlDomData As New MSXML.DOMDocument
        Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
        Dim xmlNode As MSXML.IXMLDOMNode
        Dim arrDanhsach() As String
        Dim tenCucThue As String
        Dim maCucThue As String
        Dim tenChiCucThue As String
        Dim maChiCucThue As String
        
        Dim loaiCq As String
        Dim maLoaiCq As String
        
        
        
        
        If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    
                    If arrDanhsach(0) = "1" Then
                        tenCucThue = tenCucThue + arrDanhsach(3) + Chr$(9)
                        maCucThue = maCucThue + arrDanhsach(1) + Chr$(9)
                    End If
                    'If arrDanhsach(0) = "0" Then
                        tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
                        maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
                    'End If
                End If
            Next
            Set xmlDomData = Nothing
            Set xmlNodeListCell = Nothing
            Set xmlNode = Nothing
        End If
        
        ' setloai
        If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_DM_LoaiCq.xml")) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    loaiCq = loaiCq + arrDanhsach(1) + Chr$(9)
                    maLoaiCq = maLoaiCq + arrDanhsach(0) + Chr$(9)
                End If
            Next
            Set xmlDomData = Nothing
            Set xmlNodeListCell = Nothing
            Set xmlNode = Nothing
        End If
        
        With fps
            .Sheet = .SheetCount - 1 'set thong tin CQL
            '.EventEnabled(EventAllEvents) = False
            
            .Row = 33
            .Col = .ColLetterToNumber("BE")
            .TypeComboBoxList = tenCucThue
            
            .Row = 33
            .Col = .ColLetterToNumber("BI")
            .TypeComboBoxList = maCucThue
            
            .Row = 35
            .Col = .ColLetterToNumber("BE")
            .TypeComboBoxList = tenChiCucThue
    
            .Row = 35
            .Col = .ColLetterToNumber("BI")
            .TypeComboBoxList = maChiCucThue
                                
            '.EventEnabled(EventAllEvents) = True
        End With
End Sub


'' set ma hoa don
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
'    Dim i As Integer
'    Dim strTemp As String
'
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
'                strTemp = strTemp & arrDanhsach(0) & "_" & arrDanhsach(1) & "~"
'            End If
'        Next
'    End If
'
'    ' set ma loai HD
'    'arrLoaiHd = Split(Left$(strTemp, Len(strTemp) - 1), "~")
'
'
'     With fps
'        For i = 2 To 11
'            .Sheet = i
'            ' Set gia tri ma HD
'            ' Group khong chiu thue
'            .Col = .ColLetterToNumber("C")
'            .Row = 8
'            .TypeComboBoxList = strComboHien
'            'Group 0%
'            .Col = .ColLetterToNumber("C")
'            .Row = 13
'            .TypeComboBoxList = strComboHien
'            ' Group 5%
'            .Col = .ColLetterToNumber("C")
'            .Row = 18
'            .TypeComboBoxList = strComboHien
'            ' Group 10%
'            .Col = .ColLetterToNumber("C")
'            .Row = 23
'            .TypeComboBoxList = strComboHien
'            ' Group khong phai tong hop len Tk 01/GTGT
'            .Col = .ColLetterToNumber("C")
'            .Row = 28
'            .TypeComboBoxList = strComboHien
'        Next i
'    End With
'End Sub



