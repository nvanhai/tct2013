VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01GTGT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Gia Tri Gia Tang" interface sheets
'this Class is belong to TAX_Business_v1 project which will be compline to DLL
'Coder: HTPhuong
Option Explicit

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Private Const S1ChkRow = 23
Private Const S1ChkCol = "I"
Private Const S1Txt11Row = 24
Private Const S1Txt11Col = "L"
Private Const S4TxtF2row = 2
Private Const S4TxtF2Col = "F"

Private Const Row_MSTDN = 7
Private Const Col_MSTDN = "D"

Private Const THIRD_SHEET = 3
Private Const FOURTH_SHEET = 4
Private Const THIRD_SHEET_CELL1 = "D_20"
Private Const FOURTH_SHEET_CELL1 = "E_19"
Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
'Dim xNodeData0 As MSXML.IXMLDOMNode
Dim xNode As MSXML.IXMLDOMNode
Dim arrData() As String

Dim arrLoaiHd() As String

Public Strloaitk As String
Public loaiKHBS As String
Public StrSolanBosung As String
Public FlagThaydoi As String
Public strLoaiNNKD As String

Public strTkThangQuy As String
Public dNgayDauQuy As Date

' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

' khai bao toa do cac cell va row
Private Const sheetTK01_GTGT = 1
Private Const sheetPl01_1_GTGT = 2
Private Const sheetPl01_2_GTGT = 3
'Private Const sheetPl01_3_GTGT = 4
'Private Const sheetPl01_4A_GTGT = 5
'Private Const sheetPl01_4B_GTGT = 6
Private Const sheetPl01_5_GTGT = 4
Private Const sheetPl01_6_GTGT = 5
'Private Const sheetPl01_7_GTGT = 9

Private Const sheetPl01_1TD_GTGT = 6
Private Const sheetPl01_2TD_GTGT = 7

Private Const sheetPl01_KHBS_GTGT = 8

Private Const rowHidden53 = 53
Private Const rowHidden54 = 54
Private Const rowHidden52 = 52
Private Const rowNgayKy = 50
Private Const rowNgayKyKhbs = 23

Private Const rowStartPL01_3 = 10
Private Const rowStartPL01_5 = 6
Private Const rowStartPL01_6 = 8
Private Const rowStartPL01_6_B = 8
Private Const rowStartPL01_7 = 8
' pl 01-1_GTGT
Private Const rowStartKCTPL01_1 = 8
Private Const rowStartThue_0_PL01_1 = 13
Private Const rowStartThue_5_PL01_1 = 18
Private Const rowStartThue_10_PL01_1 = 23
Private Const rowStart_Khong_TH_PL01_1 = 28
'pl 01-2_GTGT




Private Const colD = "D"
Private Const colE = "E"
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

'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'no parameter
Public Sub Prepare1()
Dim strComboHien As Variant
Dim ThangKK As Integer
Dim NamKK As Integer
Dim Thongtu As String
Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
Dim strDataFileName As String
Dim xmlNodeListItems As MSXML.IXMLDOMNodeList
Dim strCboLyDo As String
Dim strCboLyDoAn As String
Dim xmlNode As MSXML.IXMLDOMNode
Dim i As Integer
Dim fldList() As String



    With fps
    '  'danh sach  list cua combobox
    '    strComboHien = "0%" + Chr$(9) + "5%" + Chr$(9) + "10%" + Chr$(9)
        '--
        SetDateFormat fps, sheetTK01_GTGT, rowNgayKy, .ColLetterToNumber(colK), DDMMYYYY
        
'        ' set cho PL 01-3/GTGT
'        SetDateFormat fps, sheetPl01_3_GTGT, rowStartPL01_3, .ColLetterToNumber(colD), DDMMYYYY
'        SetDateFormat fps, sheetPl01_3_GTGT, rowStartPL01_3, .ColLetterToNumber(colJ), DDMMYYYY
'        SetDateFormat fps, sheetPl01_3_GTGT, rowStartPL01_3, .ColLetterToNumber(colN), DDMMYYYY
'        SetDateFormat fps, sheetPl01_3_GTGT, rowStartPL01_3, .ColLetterToNumber("R"), DDMMYYYY
'        SetDateFormat fps, sheetPl01_3_GTGT, rowStartPL01_3, .ColLetterToNumber("V"), DDMMYYYY
'        SetDateFormat fps, sheetPl01_3_GTGT, rowStartPL01_3, .ColLetterToNumber("Z"), DDMMYYYY
'        SetDateFormat fps, sheetPl01_3_GTGT, rowStartPL01_3, .ColLetterToNumber("AD"), DDMMYYYY
        
        
        
        SetDateFormat fps, sheetPl01_5_GTGT, rowStartPL01_5, .ColLetterToNumber(colD), DDMMYYYY
        
        'SetDateFormat fps, sheetPl01_KHBS_GTGT, rowNgayKyKhbs, .ColLetterToNumber(colBG), DDMMYYYY
        SetDateFormat fps, .SheetCount - 1, 31, .ColLetterToNumber("BE"), DDMMYYYY
        SetDateFormat fps, .SheetCount - 1, 46, .ColLetterToNumber("BG"), DDMMYYYY
        .Sheet = sheetTK01_GTGT 'To khai GTGT
            .Row = rowNgayKy
            .Col = .ColLetterToNumber(colK)
            .Text = Format(Date, "dd/mm/yyyy")
            .TypeHAlign = TypeHAlignLeft
            
        .Sheet = sheetPl01_1_GTGT 'PL01_1
            .Row = 8
            .Col = .ColLetterToNumber("C")
            .Text = ""
            .CellTag = .CellTag & "~1"
            .Row = 13
            .Text = ""
            .CellTag = .CellTag & "~2"
    
        
            .Row = 18
            .Text = ""
            .CellTag = .CellTag & "~3"
    
        
            .Row = 23
            .Text = ""
            .CellTag = .CellTag & "~4"
            
                    
'            .Row = 28
'            .Text = ""
'            .CellTag = .CellTag & "~5"
            
    
        
        .Sheet = sheetPl01_2_GTGT ''PL01_1
            .Row = 8
            .Col = .ColLetterToNumber("C")
            .Text = ""
            .CellTag = .CellTag & "~1"
            
'            .Row = 13
'            .Text = ""
'            .CellTag = .CellTag & "~2"
            
            .Row = 18
            .Text = ""
            .CellTag = .CellTag & "~3"
            
            .Row = 23
            .Text = ""
            .CellTag = .CellTag & "~4"
            
'            .Row = 28
'            .Text = ""
'            .CellTag = .CellTag & "~5"
            
'         ' set gia tri mac dinh cho PL 01_4B
'         .Sheet = sheetPl01_4B_GTGT
'         .Row = 2
'         .Col = .ColLetterToNumber("F")
'         .Text = TAX_Utilities_v1.Year - 1
'         ' end
       
        .Sheet = sheetPl01_KHBS_GTGT 'Pl KHBS
            .Row = rowNgayKyKhbs
            .Col = .ColLetterToNumber(colBG)
            .Text = Format(Date, "dd/mm/yyyy")
            .TypeHAlign = TypeHAlignLeft
       
        SetDateFormat fps, sheetPl01_1_GTGT, 8, .ColLetterToNumber(colG), DDMMYYYY
        SetDateFormat fps, sheetPl01_1_GTGT, 13, .ColLetterToNumber(colG), DDMMYYYY
        SetDateFormat fps, sheetPl01_1_GTGT, 18, .ColLetterToNumber(colG), DDMMYYYY
        SetDateFormat fps, sheetPl01_1_GTGT, 23, .ColLetterToNumber(colG), DDMMYYYY
'        SetDateFormat fps, sheetPl01_1_GTGT, 28, .ColLetterToNumber(colG), DDMMYYYY
        
        SetDateFormat fps, sheetPl01_2_GTGT, 8, .ColLetterToNumber(colG), DDMMYYYY
'        SetDateFormat fps, sheetPl01_2_GTGT, 13, .ColLetterToNumber(colG), DDMMYYYY
        SetDateFormat fps, sheetPl01_2_GTGT, 18, .ColLetterToNumber(colG), DDMMYYYY
        SetDateFormat fps, sheetPl01_2_GTGT, 23, .ColLetterToNumber(colG), DDMMYYYY
'        SetDateFormat fps, sheetPl01_2_GTGT, 28, .ColLetterToNumber(colG), DDMMYYYY
        
        
        'SetDateFormat fps, 7, 16, .ColLetterToNumber("O"), DDMMYYYY
        
        ' Set gia tri check gia han cho to khai 01/GTGT
        .Sheet = sheetTK01_GTGT
        ThangKK = TAX_Utilities_v1.Month
        NamKK = TAX_Utilities_v1.Year
        If NamKK = 2012 And (ThangKK = 4 Or ThangKK = 5 Or ThangKK = 6) Then
            Thongtu = GetAttribute(GetMessageCellById("0262"), "Msg")
            .SetText .ColLetterToNumber("C"), 18, Thongtu
            
            ' Ky dc gia han moi set gia tri ly do gia han
            strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_DM_Lydo_GH.xml")
            If xmlDomData.Load(strDataFileName) Then
                 Set xmlNodeListItems = xmlDomData.getElementsByTagName("Item")
                 For Each xmlNode In xmlNodeListItems
                     fldList = Split(GetAttribute(xmlNode, "Value"), "###")
                     strCboLyDo = strCboLyDo + fldList(1) + Chr$(9)
                     strCboLyDoAn = strCboLyDoAn + fldList(0) + Chr$(9)
                 Next
             End If
                  'set ly do gia han
                .Row = 20
                .Col = .ColLetterToNumber("D")
                .TypeComboBoxList = strCboLyDo
                ' set Ma ly do gia han
                .Col = .ColLetterToNumber("P")
                .TypeComboBoxList = strCboLyDoAn
        Else
            Thongtu = GetAttribute(GetMessageCellById("0263"), "Msg")
            .SetText .ColLetterToNumber("C"), 18, Thongtu
            ' Ky dc gia han moi set gia tri ly do gia han
            strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_DM_Lydo_GH_TT16.xml")
            If xmlDomData.Load(strDataFileName) Then
                 Set xmlNodeListItems = xmlDomData.getElementsByTagName("Item")
                 For Each xmlNode In xmlNodeListItems
                     fldList = Split(GetAttribute(xmlNode, "Value"), "###")
                     strCboLyDo = strCboLyDo + fldList(1) + Chr$(9)
                     strCboLyDoAn = strCboLyDoAn + fldList(0) + Chr$(9)
                 Next
             End If
                  'set ly do gia han
            .Row = 20
            .Col = .ColLetterToNumber("D")
            .TypeComboBoxList = strCboLyDo
            ' set Ma ly do gia han
            .Col = .ColLetterToNumber("P")
            .TypeComboBoxList = strCboLyDoAn
        End If
    
    End With
    ' set thong tin CQT quan ly
    setCQTQuanLy
'    ' set ma hoa don cho PL 01-1, 01-2
'    setMaHoaDon


End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'no parameter
Public Sub Prepare2()
    Dim i As Long, DataText As Variant
    Dim arrCellTag() As String
    Dim arrRow As Variant, arrCol As Variant
    Dim arrRowChange As Variant, arrColChange As Variant
    Dim intCtrl As Integer
    Dim strId As Variant
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim tempCurrSheet As Integer
    With fps

    mblKhoiDauToKhai = True
        .EventEnabled(EventAllEvents) = True
        tempCurrSheet = mCurrentSheet
        .Sheet = sheetTK01_GTGT 'To khai GTGT
            .ActiveSheet = .Sheet
            mCurrentSheet = 1
            'CellChange .ColLetterToNumber(S1ChkCol), S1ChkRow
            LockCellBySheet
            CellLock .ColLetterToNumber(S1ChkCol), S1ChkRow
            mblKhoiDauToKhai = False
            'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
            
            'Check rang buôc CT 40 va 42
            .Row = 44
            .Col = .ColLetterToNumber("L")
            If Val(.value) <> 0 Then
                .Row = 46
                .Col = .ColLetterToNumber("L")
                .Lock = True
            Else
                .Row = 46
                .Col = .ColLetterToNumber("L")
                .Lock = False
            End If
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber(colK)
            .Row = rowHidden52
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber(colL)
            .Row = rowHidden52
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber(colO)
            .Row = rowHidden52
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber(colK)
            .Row = rowHidden52
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber(colL)
            .Row = rowHidden52
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber(colO)
            .Row = rowHidden52
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        'ma NNKD
        .Col = .ColLetterToNumber(colL)
        .Row = 14
        .Text = strLoaiNNKD
        UpdateCell fps, .Col, .Row, .Text
        
        ' Neu check gia han moi cho phep chon GT
        .Col = .ColLetterToNumber("B")
        .Row = 18
        If .Text = "1" Or UCase$(.Text) = "X" Then
            .Col = .ColLetterToNumber(colD)
            .Row = 20
            .Lock = False
        Else
            .Col = .ColLetterToNumber(colD)
            .Row = 20
            .Lock = True
        End If
        ' End set gia han
        
        ' set loai tk thang/ quy
         If strTkThangQuy = "TK_THANG" Then
            .Col = .ColLetterToNumber("N")
            .Row = 52
            .Text = "0"
            UpdateCell fps, .Col, .Row, .Text
        ElseIf strTkThangQuy = "TK_QUY" Then
            .Col = .ColLetterToNumber("N")
            .Row = 52
            .Text = "1"
            UpdateCell fps, .Col, .Row, .Text
        End If
            'end
            
        .Sheet = sheetPl01_1_GTGT 'PL1
            .ActiveSheet = .Sheet
                .Row = 1
                .Col = .ColLetterToNumber("F")
                .CellType = CellTypeButton
                .TypeButtonTextColor = RGB(0, 0, 255)
                .TypeButtonColor = &H8000000F
                .TypeButtonAlign = TypeButtonAlignRight
                .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
        .Sheet = sheetPl01_2_GTGT 'PL2
             .ActiveSheet = .Sheet
                .Row = 1
                .Col = .ColLetterToNumber("F")
                .CellType = CellTypeButton
                .TypeButtonTextColor = RGB(0, 0, 255)
                .TypeButtonColor = &H8000000F
                .TypeButtonAlign = TypeButtonAlignRight
                .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
                
'        .Sheet = sheetPl01_3_GTGT 'PL3
'             .ActiveSheet = .Sheet
'                .Row = 1
'                .Col = .ColLetterToNumber("C")
'                .CellType = CellTypeButton
'                .TypeButtonTextColor = RGB(0, 0, 255)
'                .TypeButtonColor = &H8000000F
'                .TypeButtonAlign = TypeButtonAlignRight
'                .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
                
        ' set lock CQT phan bo
        .Sheet = sheetPl01_2TD_GTGT
        Dim vChon As Variant
        If .SheetVisible = True Then
        .Row = 42
         Do
                .GetText .ColLetterToNumber("G"), .Row, vChon
                If vChon = 1 Then
                    .Col = .ColLetterToNumber("P")
                    .Lock = True
                End If
                
                .Row = .Row + 1
                .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
         End If
                
         .ActiveSheet = 1
         mCurrentSheet = tempCurrSheet
         .EventEnabled(EventAllEvents) = True
    End With
      
    'Get last [43] target and assign to current [11] target
    setLast43Target
    
End Sub

'inner class event: Change
Public Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
Dim str43Previous As String
Dim varTemp As Variant
Dim xmlNodeH15 As MSXML.IXMLDOMNode
Dim strId As Variant
Dim iCol As Long, iRow As Long
Dim cCol As Long, cRow As Long
Dim cValue As Variant, temp As Variant
Dim cValueG As Variant
Dim cTagF As String
Dim cValueF As Variant
Dim i As Integer



Dim dCT25 As Double

With fps
    If .ActiveSheet = sheetTK01_GTGT Then
        'Check value of 11
        GetCellSpan fps, Col, Row
        If Col = .ColLetterToNumber(colL) And Row = 24 Then
            .GetText Col, Row, varTemp
            If GetLast43Previous(str43Previous) Then
                If CStr(varTemp) <> str43Previous Then
                    .Visible = True
                    DisplayMessage "0060", msOKOnly, miWarning
                    fps.SetFocus
                End If
            End If
        End If
'    ElseIf .ActiveSheet = sheetPl01_1_GTGT Then
'        .Col = Col
'        .Row = Row
'        varTemp = .Text
'         If Col = .ColLetterToNumber("C") And .Text <> "" Then
'            .EventEnabled(EventChange) = False
'            .Col = .ColLetterToNumber(colD)
'            .Text = varTemp
'            .EventEnabled(EventChange) = True
'         End If
'    ElseIf .ActiveSheet = sheetPl01_2_GTGT Then
'        .Col = Col
'        .Row = Row
'        varTemp = .Text
'         If Col = .ColLetterToNumber("C") And .Text <> "" Then
'            .EventEnabled(EventChange) = False
'            .Col = .ColLetterToNumber(colD)
'            .Text = varTemp
'            .EventEnabled(EventChange) = True
'         End If
    End If
    
 mblKhoiDauToKhai = False
    
End With
    LockCellBySheet
    setChangeComb Col, Row
    CellChange Col, Row
End Sub

'inner class event: ButtonClicked
Public Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim lResult As Integer
    Dim blTemp As Boolean
    
    fps.EventEnabled(EventAllEvents) = False
    
    If Col = fps.ColLetterToNumber(S1ChkCol) And Row = S1ChkRow Then
       If mblKhoiDauToKhai = False Then
            fps.Sheet = fps.ActiveSheet
            mCurrentSheet = fps.Sheet
            fps.Row = Row
            fps.Col = Col
            blTemp = TAX_Utilities_v1.AdjustData(fps.ActiveSheet - 1)
            If fps.value = 1 Then
                lResult = DisplayMessage("0034", msYesNo, miQuestion, , mrNo)
                If lResult = 3 Then '<>6
                    fps.value = 0
                    UpdateCell fps, fps.Col, fps.Row, IIf(fps.value = 1, "x", vbNullString) 'fps.Value
                    TAX_Utilities_v1.AdjustData(fps.ActiveSheet - 1) = blTemp
                    fps.EventEnabled(EventAllEvents) = True
                    Exit Sub
                Else
                    UpdateCell fps, fps.Col, fps.Row, IIf(fps.value = 1, "x", vbNullString) 'fps.Value
                    TAX_Utilities_v1.AdjustData(fps.ActiveSheet - 1) = True
                End If
            ElseIf fps.value = 0 Then
                UpdateCell fps, fps.Col, fps.Row, IIf(fps.value = 1, "x", vbNullString)
                TAX_Utilities_v1.AdjustData(fps.ActiveSheet - 1) = True
                ' Khong check se lay lai gia tri cua phu luc len
                SetValueOCheck
                
            End If
        End If
        
        CellLock Col, Row
    End If
    
    ' set gia lock cell ly do gia han
    If fps.ActiveSheet = sheetTK01_GTGT Then
            ' Neu check gia han moi cho phep chon GT
            If Col = fps.ColLetterToNumber("B") And Row = 18 Then
                fps.Col = fps.ColLetterToNumber("B")
                fps.Row = 18
                If fps.Text = "1" Or UCase$(fps.Text) = "X" Then
                    fps.Col = fps.ColLetterToNumber("D")
                    fps.Row = 20
                    fps.Lock = False
                    If Strloaitk = "TKBS" Then
                        TAX_Utilities_v1.CheckToKhaiGH = True
                        TinhHanNop_KHBS fps
                        fps.Sheet = 1
                    End If
                Else
                    fps.Col = fps.ColLetterToNumber("D")
                    fps.Row = 20
                    fps.Lock = True
                    fps.SetText fps.ColLetterToNumber("D"), 20, ""
                    fps.SetText fps.ColLetterToNumber("P"), 20, "00"
                    UpdateCell fps, fps.ColLetterToNumber("P"), 20, "00"
                    UpdateCell fps, fps.ColLetterToNumber("D"), 20, ""
                    If Strloaitk = "TKBS" Then
                        TAX_Utilities_v1.CheckToKhaiGH = False
                        TinhHanNop_KHBS fps
                        fps.Sheet = 1
                    End If
                End If
            End If
    End If
    
    fps.EventEnabled(EventAllEvents) = True
    
     With fps

            If Col = .ColLetterToNumber("G") And Row >= 42 And .ActiveSheet = sheetPl01_2TD_GTGT Then
                   fps.EventEnabled(EventAllEvents) = False
                   .Sheet = sheetPl01_2TD_GTGT
                   .Row = Row
                   .Col = .ColLetterToNumber("G")
                   If (.value = "1") Then
                       .Row = Row
                       ' chi tieu STT
                       .Col = .ColLetterToNumber("P")
                       .Lock = True
                       
                       .Text = ""
                       UpdateCell fps, .Col, .Row, .Text
                       
                       .Col = .ColLetterToNumber("AH")
                       .Text = ""
                       UpdateCell fps, .Col, .Row, .Text
                   Else
                       .Row = Row
                       ' chi tieu STT
                       .Col = .ColLetterToNumber("P")
                       .Lock = False
                      
                   End If
                   fps.EventEnabled(EventAllEvents) = True
               End If
    End With
    
    'CellChange Col, Row
    
End Sub

Private Sub fps_EditChange(ByVal Col As Long, ByVal Row As Long)
    CellChange1 Col, Row
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
Private Sub fps_KeyPress(KeyAscii As Integer)
     With fps
        If .ActiveSheet = 6 Then
            If .Col = .ColLetterToNumber("F") And .Row = 2 Then
                If (KeyAscii < 48 And KeyAscii <> 3 And KeyAscii <> 8 And KeyAscii <> 22 And KeyAscii <> 13) Or KeyAscii > 57 Then
                    KeyAscii = 0
                End If
            End If
        End If
    End With
    
    
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, iCol As Long, iRow As Long
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tCol As Long, tRow As Long
Dim strId As Variant, Col7 As Variant, intIndexCombo As Integer
Dim varTemp As Variant
Dim NguoiKyCol As Long, NguoiKyRow As Long
    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        fps.EventEnabled(EventAllEvents) = False
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = sheetPl01_1_GTGT Then
                    mCurrentSheet = sheetPl01_1_GTGT
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 8
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 8
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("L")
'                         .Text = "0%"
'                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("L")
'                         .Text = "5%"
'                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "cc"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("L")
'                         .Text = "10%"
'                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "dd"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "ee"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                     
                    .SetActiveCell iCol, iRow
                    CellChange2
                    CellChange3
                    LockCellBySheet
            End If
            If .ActiveSheet = sheetPl01_2_GTGT Then
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 8
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 8
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "cc"
                    
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "dd"
                    
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "ee"
                    
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
                    CellChange2
                    CellChange3
                    LockCellBySheet
            End If
            
            
'            If .ActiveSheet = sheetPl01_3_GTGT Then
'                    i = 1
'                    .Col = .ColLetterToNumber("B")
'                    .Row = rowStartPL01_3
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("B")
'                         .Row = i + rowStartPL01_3
'                         i = i + 1
'                    Loop Until .Text = "aa"
'                    .SetActiveCell iCol, iRow
'
'                    .Row = iRow
'                    .Col = iCol
'                    .SetActiveCell iCol, iRow
'            End If
            
            
'            If .ActiveSheet = sheetPl01_7_GTGT Then
'                    i = 1
'                    .Col = .ColLetterToNumber("B")
'                    .Row = 8
'                    Do
'                         .Text = str(i)
'                         .GetText .ColLetterToNumber("D"), .Row, varTemp
'                         If varTemp = vbNullString Or Trim(varTemp) = "" Then
'                            .SetText .ColLetterToNumber("D"), .Row, TAX_Utilities_v1.Convert("ChiÕc", TCVN, UNICODE)
'                         End If
'                         .Col = .ColLetterToNumber("B")
'                         .Row = i + 8
'                         i = i + 1
'                    Loop Until .Text = "aa"
'                    i = 1
'                   .Row = .Row + 4
'                     Do
'                          .Text = str(i)
'                                 .GetText .ColLetterToNumber("D"), .Row, varTemp
'                                 If varTemp = vbNullString Or Trim(varTemp) = "" Then
'                                    .SetText .ColLetterToNumber("D"), .Row, TAX_Utilities_v1.Convert("ChiÕc", TCVN, UNICODE)
'                                 End If
'                          .Col = .ColLetterToNumber("B")
'                          .Row = 1 + .Row
'                          i = i + 1
'                     Loop Until .Text = "bb"
'                     i = 1
'                    .Row = .Row + 4
'                    .SetActiveCell iCol, iRow
'                    .Row = iRow
'                    .Col = iCol
'                    .SetActiveCell iCol, iRow
'            End If
            If .ActiveSheet = sheetPl01_5_GTGT Then
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = rowStartPL01_5
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + rowStartPL01_5
                         i = i + 1
                    Loop Until .Text = "aa"
                    .SetActiveCell iCol, iRow

                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
            End If
            If .ActiveSheet = sheetPl01_6_GTGT Then
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 8
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 8
                         i = i + 1
                    Loop Until .Text = "aa"
                    .SetActiveCell iCol, iRow

                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
            End If
            
            
            ' set lock CQT phan bo
            If .ActiveSheet = sheetPl01_2TD_GTGT Then
                .Sheet = sheetPl01_2TD_GTGT
                Dim vChon As Variant
                If .SheetVisible = True Then
                .Row = 42
                 Do
                        .GetText .ColLetterToNumber("G"), .Row, vChon
                        If vChon = 1 Then
                            .Col = .ColLetterToNumber("P")
                            .Lock = True
                        Else
                            .Col = .ColLetterToNumber("P")
                            .Lock = False
                        End If
                        
                        .Row = .Row + 1
                        .Col = .ColLetterToNumber("B")
                    Loop Until .Text = "aa"
                 End If
            End If
        End If
  
      fps.EventEnabled(EventAllEvents) = True
    End With
    
End Sub


'inner class event: Leavecell
Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim tCol As Long, tRow As Long
    Dim cCol As Long, cRow As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim dCT25 As Double
    Dim temp As Variant
    
    With fps
    mCurrentSheet = .Sheet
    .EventEnabled(EventAllEvents) = False
        If .ActiveSheet = sheetTK01_GTGT Then
             If Col = .ColLetterToNumber(colK) And Row = rowNgayKy Then
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
        If .ActiveSheet = sheetPl01_1_GTGT Or .ActiveSheet = sheetPl01_2_GTGT Then
                   If Col = .ColLetterToNumber(colG) Then
                       .GetText Col, Row, varTemp
                       If varTemp <> "" And varTemp <> "../../...." Then
                           If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                               .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                           Else
                               .SetActiveCell Col, Row
                           End If
                       Else
                        .SetText Col, Row, ""
                       End If
                       .Col = Col
                       .Row = Row
                      UpdateCell fps, .Col, .Row, .Text
                   End If
                   
                   'chuyen cot ky hieu ve chu cai in hoa
                   If Col = .ColLetterToNumber("E") Then
                        .GetText Col, Row, varTemp
                        .SetText Col, Row, UCase(varTemp)
                        .Col = Col
                        .Row = Row
                        UpdateCell fps, .Col, .Row, .Text
                   End If
        End If
        
'        ' Phu luc 01-3/GTGT
'         If .ActiveSheet = sheetPl01_3_GTGT Then
'                   If Col = .ColLetterToNumber(colD) Or Col = .ColLetterToNumber(colJ) Or Col = .ColLetterToNumber(colN) Or Col = .ColLetterToNumber("R") Or Col = .ColLetterToNumber("V") Or Col = .ColLetterToNumber("Z") Or Col = .ColLetterToNumber("AD") Then
'                       .GetText Col, Row, varTemp
'                       If varTemp <> "" And varTemp <> "../../...." Then
'                           If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
'                               .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                           Else
'                               .SetFocus
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
'
'        End If
        
        
        ' end
        If .ActiveSheet = sheetPl01_5_GTGT Then
                   If Col = .ColLetterToNumber(colD) Then
                       .GetText Col, Row, varTemp
                       If varTemp <> "" And varTemp <> "../../...." Then
                           If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                               .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
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
        If .ActiveSheet = sheetPl01_KHBS_GTGT Then
         Dim count1 As Integer
            .Row = 9
            .Col = .ColLetterToNumber("B")
            Do While .value <> "bb"
                count1 = count1 + 1
                .Row = .Row + 1
            Loop
                  If Col = .ColLetterToNumber(colBG) And Row = rowNgayKyKhbs + count1 - 5 Then
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
                .Sheet = .SheetCount - 1
             
        End If
        .EventEnabled(EventAllEvents) = True
End With
    'CellChange Col, Row
    LockCellBySheet
End Sub

'When cell's data is changed or a control at a cell is handled, This sub is called
'parameter Col: the column of cell that rise event
'parameter Row: the row of cell that rise event
'parameter f: the flag to mark some special option
'**: when this sub is on executing, all events of fpspread will be disabled until the end of the sub
Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
    Dim i As Integer
    Dim a As Double, b As Double, c As Double
    Dim arrRow As Variant, arrCol As Variant
    Dim arrCellTag() As String
    Dim cCol As Long, cRow As Long
    Dim varTemp As Variant
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim iCol  As Long, iRow As Long
    Dim strValue1  As Variant, strValue2 As Variant
With fps
    'Because of unclearly-rised events, this function is here for turning it off
    'at the end of this sub, this event will be turned on
    .EventEnabled(EventAllEvents) = False

    
    ' Kiem tra neu co phu luc moi set cong thuc tong hop so lieu tu tinh
    Dim tempCurrSheet As Integer
    Dim ct21 As Variant
    Dim blCheck_S1 As Boolean
    Dim blCheck_S2 As Boolean
    Dim blCheck_S5 As Boolean
    Dim blCheck_S4A As Boolean
    Dim blCheck_S4B As Boolean
    blCheck_S1 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl01_1_GTGT - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl01_2_GTGT - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S5 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl01_5_GTGT - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    tempCurrSheet = mCurrentSheet
    .Sheet = 1
    ' Chi tieu 21
    .Col = .ColLetterToNumber("I")
    .Row = 23
    ct21 = .Text
        
    .Sheet = .ActiveSheet
    Select Case .Sheet
        Case 1 ' To khai GTGT
            mCurrentSheet = .Sheet
            GetCellSpan fps, Col, Row
            If Row = 44 And Col = .ColLetterToNumber("L") Then
             'Check rang buôc CT 40 va 42
                .Row = 44
                .Col = .ColLetterToNumber("L")
                If Val(.value) <> 0 Then
                    .Row = 46
                    .Col = .ColLetterToNumber("L")
                    .Lock = True
                Else
                    .Row = 46
                    .Col = .ColLetterToNumber("L")
                    .Lock = False
                End If
            End If
        Case 3 'PL2
    End Select
    
    .EventEnabled(EventAllEvents) = True
End With
End Sub

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
    
    iFlag5 = 0
    iFlag6 = 0
    iFlagvKyHieuHD = 0
    iFlagvSoHD = 0
    iFlagvNgayHD = 0
    iFlagvTennguoimua = 0
    iFocusFlag = 0
    
    flagTenCSPT = 0
    flagDiaChi = 0
    
    Dim vErrMsgThuesuat As Variant, vThuesuat As Variant
    Dim iFlagvThuesuat As Integer
    
    ' pl5
    Dim vSoChungTu As Variant
    Dim vNgayNop As Variant
    Dim vNoiNop As Variant
    Dim vCoQuanQL As Variant
    Dim vSoTienDN As Variant
    'pl3
    Dim vLoaiXe As Variant
    Dim vDVtinh As Variant
    Dim vTongsoxe As Variant
    Dim vSoLuongxe As Variant
    Dim vGiaxe As Variant
    
    '/
    Dim iflagSoChungTu As Integer
    Dim iflag_pl_3 As Integer
    Dim iflagNgayNop As Integer
    iflag_pl_3 = 0
    iflagSoChungTu = 0
    iflagNgayNop = 0
    Dim vErrorSoChungTu As Variant
    Dim vError_pl_3 As Variant
    Dim vErrorNgayNop As Variant
    ' end pl5
    
    
    
    Dim blCheck_S2 As Boolean
    Dim blCheck_S3 As Boolean
    Dim blCheck_S4 As Boolean
    Dim blCheck_S5 As Boolean
     Dim blCheck_S7 As Boolean
    Dim blCheck_S8 As Boolean
    
    ' Kiem tra check gia han
    Dim vCheckGH, vLyDoGH, vErrorCheckGH As Variant
    Dim iflagCheckGH As Integer
    
    Dim arrDate() As String
    Dim vdate As Date
    Dim vNgayHt As Date
    
    Dim countThuSuat0 As Long
    
        ' Khai bao bien PL 01-3
    Dim soHopDongXK, ngayHopDongXK, ngoaiTeXK, tienVNDXK, hinhThucTTXK, thoHanTTXK
    Dim toKhaiHangHoaXK, ngayHangHoaXK, ngoaiTeHangHoaXK, tienVNDHangHoaXK
    Dim soHoaDonXK, ngayHoaDonXK, ngoaiTeHoaDonXK, tienVNDHoaDonXK
    Dim soChungTu, ngayChungTu, ngoaiTeChungTu, tienVNDChungTu
    Dim soVanBan, ngayVanBan, ngoaiTeVanBan, tienVNDVanBan
    Dim soHDNK, ngayHDNK, ngoaiTeHDNK, tienVNDHDNK
    Dim soHangHoaNK, ngayHangHoaNK, ngoaiTeHangHoaNK, tienVNDHangHoaNK
    Dim chungTuTTbutru, ghiChu
    Dim vErrorNgayHDXK, vErrorNgayDKTKXK, vErrorNgayHoaDonXK, vErrorNgayChungTuTT, vErrorNgayVBXN, vErrorNgayHDNK, vErrorNgayTKNK
    Dim iflagNgayHopDongXK As Integer, iflagNgayDangKyTKXK As Integer, iflagNgayHoaDonXK As Integer, iflagNgayChungTuTT As Integer, iflagNgayVBXN As Integer, iflagNgayHDNK As Integer, iflagNgayTKNK As Integer

    ' PL 01-1
    Dim vErrorNgayPHHD As Variant
    Dim iflagNgayPHHD As Integer
    
    Dim vErrorNgayPHHD_1 As Variant
    Dim iflagNgayPHHD_1 As Integer
    
    Dim strTemp As String, result As String, strLoai As String
    Dim iFlagMauSoHD As Integer, iFlagSoLienHD As Integer
    Dim errMauSoHD As Variant, errSoLienHD As Variant
    Dim errEmty As Variant
    Dim iFlagEmty As Integer
    ' PL 01-2
    Dim vErrorPhatHanhHD As Variant
    Dim iflagNgayPhatHanhHD As Integer
    
    Dim vErrThueGTGT As Variant
    Dim iFlagThueGTGT As Integer
    
    With fps
    
    '.Visible = False
    '.ReDraw = False
    iCurrentSheet = .Sheet
    .Sheet = .SheetCount
    CheckErrorMST
    
    .Sheet = 1

     .Sheet = .SheetCount
     .GetText .ColLetterToNumber("E"), 27, vErrMsgKyHieuHD
     .GetText .ColLetterToNumber("E"), 28, vErrMsgSoHD
     .GetText .ColLetterToNumber("E"), 29, vErrMsgNgayHD
     .GetText .ColLetterToNumber("E"), 30, vErrMsgTennguoimua
     .GetText .ColLetterToNumber("E"), 31, vErrMsgMST
     .GetText .ColLetterToNumber("E"), 37, vErrMsgThuesuat
     
     .GetText .ColLetterToNumber("E"), 37, vErrMsgThuesuat
     .GetText .ColLetterToNumber("E"), 57, vErrorSoChungTu
     .GetText .ColLetterToNumber("E"), 58, vError_pl_3
     
     .GetText .ColLetterToNumber("E"), 55, vErrorNgayNop
     
     .GetText .ColLetterToNumber("E"), 62, vErrorTenCSPT
     .GetText .ColLetterToNumber("E"), 63, vErrorDiaChi
     
     .GetText .ColLetterToNumber("E"), 64, vErrorCheckGH
     
     .GetText .ColLetterToNumber("E"), 70, vErrorNgayHDXK
     .GetText .ColLetterToNumber("E"), 71, vErrorNgayDKTKXK
     .GetText .ColLetterToNumber("E"), 72, vErrorNgayHoaDonXK
     .GetText .ColLetterToNumber("E"), 73, vErrorNgayChungTuTT
     .GetText .ColLetterToNumber("E"), 74, vErrorNgayVBXN
     .GetText .ColLetterToNumber("E"), 75, vErrorNgayHDNK
     .GetText .ColLetterToNumber("E"), 76, vErrorNgayTKNK
     
     .GetText .ColLetterToNumber("E"), 77, vErrorNgayPHHD
     .GetText .ColLetterToNumber("E"), 78, vErrorPhatHanhHD
     .GetText .ColLetterToNumber("E"), 79, errMauSoHD
     .GetText .ColLetterToNumber("E"), 80, errSoLienHD
     .GetText .ColLetterToNumber("E"), 81, errEmty
     .GetText .ColLetterToNumber("E"), 88, vErrThueGTGT
     
     .GetText .ColLetterToNumber("E"), 89, vErrorNgayPHHD_1
     
     
     ' Kiem tra check gia han
     .Sheet = 1
     .GetText .ColLetterToNumber("B"), 18, vCheckGH
     .GetText .ColLetterToNumber("D"), 20, vLyDoGH
     'If TAX_Utilities_v1.Year = 2012 And (TAX_Utilities_v1.Month = 4 Or TAX_Utilities_v1.Month = 5 Or TAX_Utilities_v1.Month = 6) Then
        If Trim(vCheckGH) = "1" Or Trim$(UCase(vCheckGH)) = "X" Then
            If Trim$(vLyDoGH) = "" Then
                .Col = .ColLetterToNumber("D")
                .Row = 20
                .CellNote = vErrorCheckGH  'static
                .BackColor = mErrorColor
                iflagCheckGH = 1
            Else
                .Col = .ColLetterToNumber("D")
                .Row = 20
                .CellNote = ""  'static
                .BackColor = mNonErrorColor
                iflagCheckGH = 0
            End If
        Else
                .Col = .ColLetterToNumber("D")
                .Row = 20
                .CellNote = ""  'static
                .BackColor = mNonErrorColor
                iflagCheckGH = 0
        End If
        
    .Sheet = .SheetCount
    .SetText .ColLetterToNumber("B"), 64, IIf(iflagCheckGH = 1, "0", "1")
     'End If
     ' end kiem tra check gia han
    
    ' if sheet invisible then set value="1"
    'countThuSuat0 = 0
    .Sheet = sheetPl01_1_GTGT
    blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    If blCheck_S2 = True Then
        For i = 8 To .MaxRows
            iWhiteRow = 1
            .Row = i
            .Col = .ColLetterToNumber("C")
            arrCellTag = Split(.CellTag, "~")
            If UBound(arrCellTag) > 0 Then
            strCellTag = arrCellTag(1)
            '.GetText .ColLetterToNumber("C"), .Row, vMaHoaDon
            '.GetText .ColLetterToNumber("D"), .Row, vKyHieuMauHoaDon
            
            '.GetText .ColLetterToNumber("E"), .Row, vKyHieuHD
            .GetText .ColLetterToNumber("F"), .Row, vSoHD
            .GetText .ColLetterToNumber("G"), .Row, vNgayHD
            .GetText .ColLetterToNumber("H"), .Row, vTennguoimua
            .GetText .ColLetterToNumber("K"), .Row, vDoanhSo
            .GetText .ColLetterToNumber("M"), .Row, vGTGT
            .GetText .ColLetterToNumber("I"), .Row, vMST
            
            
            'If Trim(vKyHieuHD) <> "" Or Trim(vSoHD) <> "" Or Trim(vNgayHD) <> "" Or Trim(vTennguoimua) <> "" Or Trim(vDoanhSo) <> "0" Or Trim(vGTGT) <> "0" Or Trim(vMaHoaDon) <> "" Or Trim(vKyHieuMauHoaDon) <> "" Then iWhiteRow = 0
            If Trim(vSoHD) <> "" Or Trim(vNgayHD) <> "" Or Trim(vTennguoimua) <> "" Or Trim(vDoanhSo) <> "0" Or Trim(vGTGT) <> "0" Then iWhiteRow = 0
                Select Case strCellTag
                    Case "1", "2", "3", "4", "5"
                        ' kiem tra ky hieu mau HD
'                        If Trim(vKyHieuMauHoaDon) = "" And iWhiteRow = 0 Then
'                            If Trim(vMaHoaDon) = "TT120" Then
'                                .Col = .ColLetterToNumber("D")
'                                .CellNote = ""
'                                .BackColor = mNonErrorColor
'                            Else
'                                .Col = .ColLetterToNumber("D")
'                                .CellNote = errEmty  'static
'                                .BackColor = mErrorColor
'                                iFlagEmty = 1
'                            End If
'                        Else
'                            strLoai = getLoaiHD(CStr(vMaHoaDon))
'                            result = CheckMauSoHD(CStr(vKyHieuMauHoaDon), strLoai, CStr(vMaHoaDon))
'                            If result = "1" And iWhiteRow = 0 Then
'                                .Col = .ColLetterToNumber("D")
'                                .CellNote = errMauSoHD  'static
'                                .BackColor = mErrorColor
'                                iFlagMauSoHD = 1
'                            ElseIf result = "2" And iWhiteRow = 0 Then
'                                .Col = .ColLetterToNumber("D")
'                                .CellNote = errSoLienHD  'static
'                                .BackColor = mErrorColor
'                                iFlagSoLienHD = 1
'                            Else
'                                .Col = .ColLetterToNumber("D")
'                                .CellNote = ""  'static
'                                .BackColor = mNonErrorColor
'                            End If
'                        End If
                        ' end
                    
'                        If Trim(vKyHieuHD) = "" And iWhiteRow = 0 Then
'                            If Trim(vMaHoaDon) = "TT120" Then
'                                .Col = .ColLetterToNumber("E")
'                                .CellNote = ""  'static
'                                .BackColor = mNonErrorColor
'                            Else
'                                .Col = .ColLetterToNumber("E")
'                                .CellNote = vErrMsgKyHieuHD  'static
'                                .BackColor = mErrorColor
'                                iFlagvKyHieuHD = 1
'                            End If
'                            'capture the first error cell
'    '                        If iFocusFlag = 0 Then
'    '                            strFocusSheetName = "'" & .SheetName & "'"
'    '                            strFocusCol = str(.Col)
'    '                            strFocusRow = str(.Row)
'    '                            iFocusFlag = 1
'    '                        End If
'                        Else
'                            .Col = .ColLetterToNumber("E")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
'                        End If
                        
                        If Trim(vSoHD) = "" And iWhiteRow = 0 Then
                            .Col = .ColLetterToNumber("F")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgSoHD  'static
                            .BackColor = mErrorColor
                            iFlagvSoHD = 1
                            'capture the first error cell
'                            If iFocusFlag = 0 Then
'                                strFocusSheetName = "'" & .SheetName & "'"
'                                strFocusCol = str(.Col)
'                                strFocusRow = str(.Row)
'                                iFocusFlag = 1
'                            End If
                        Else
'                            ' Kiem tra co du lieu dong thue suat 0% thi bat buoc nhap PL 01-3
'                            If strCellTag = "2" And Trim(vSoHD) <> "" Then
'                                countThuSuat0 = countThuSuat0 + 1
'                            End If
                            .Col = .ColLetterToNumber("F")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                        
                        If Trim(vNgayHD) = "" And iWhiteRow = 0 Then
                            .Col = .ColLetterToNumber("G")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgNgayHD  'static
                            .BackColor = mErrorColor
                            iFlagvNgayHD = 1
                            'capture the first error cell
'                            If iFocusFlag = 0 Then
'                                strFocusSheetName = "'" & .SheetName & "'"
'                                strFocusCol = str(.Col)
'                                strFocusRow = str(.Row)
'                                iFocusFlag = 1
'                            End If
                        
                         ElseIf Trim(vNgayHD) <> "" And iWhiteRow = 0 Then
                            ' To khai quy
                            'If strTkThangQuy = "TK_QUY" Then
'                                vNgayHt = Date
'                                arrDate = Split(vNgayHD, "/")
'                                vdate = DateSerial(CInt(arrDate(2)), CInt(arrDate(1)), CInt(arrDate(0)))
'                                If CInt(Right(vNgayHD, 4)) > CInt(TAX_Utilities_v1.Year) Then
'                                    .Col = .ColLetterToNumber("G")
'                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                    .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
'                                    .BackColor = mErrorColor
'                                    iFlagvNgayHD = 1
'                                    If iFocusFlag = 0 Then
'                                        strFocusSheetName = "'" & .SheetName & "'"
'                                        strFocusCol = str(.Col)
'                                        strFocusRow = str(.Row)
'                                        iFocusFlag = 1
'                                    End If
'                                 ElseIf CInt(Right(vNgayHD, 4)) = CInt(TAX_Utilities_v1.Year) And convertThang2Quy(CInt(Right("00" & Mid(vNgayHD, 4, 2), 2))) > CInt(TAX_Utilities_v1.ThreeMonths) Then
'                                    .Col = .ColLetterToNumber("G")
'                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                    .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
'                                    .BackColor = mErrorColor
'                                    iFlagvNgayHD = 1
'                                    If iFocusFlag = 0 Then
'                                        strFocusSheetName = "'" & .SheetName & "'"
'                                        strFocusCol = str(.Col)
'                                        strFocusRow = str(.Row)
'                                        iFocusFlag = 1
'                                    End If
'                                 ElseIf DateDiff("D", vNgayHt, vdate) > 0 Then
'                                    .Col = .ColLetterToNumber("G")
'                                    .CellNote = GetAttribute(GetMessageCellById("0273"), "Msg") 'static
'                                    .BackColor = mErrorColor
'                                    iFlagvNgayHD = 1
'                                     If iFocusFlag = 0 Then
'                                        strFocusSheetName = "'" & .SheetName & "'"
'                                        strFocusCol = str(.Col)
'                                        strFocusRow = str(.Row)
'                                        iFocusFlag = 1
'                                      End If
'                                 Else
'                                    .Col = .ColLetterToNumber("G")
'                                    .CellNote = ""
'                                    .BackColor = mNonErrorColor
'                                End If
'                            Else
'                                If CInt(Right(vNgayHD, 4)) > CInt(TAX_Utilities_v1.Year) Then
'                                    .Col = .ColLetterToNumber("G")
'                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                    .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
'                                    .BackColor = mErrorColor
'                                    iFlagvNgayHD = 1
'                                    If iFocusFlag = 0 Then
'                                        strFocusSheetName = "'" & .SheetName & "'"
'                                        strFocusCol = str(.Col)
'                                        strFocusRow = str(.Row)
'                                        iFocusFlag = 1
'                                    End If
'                                 ElseIf CInt(Right(vNgayHD, 4)) = CInt(TAX_Utilities_v1.Year) And CInt(Right("00" & Mid(vNgayHD, 4, 2), 2)) > CInt(TAX_Utilities_v1.Month) Then
'                                    .Col = .ColLetterToNumber("G")
'                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                    .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
'                                    .BackColor = mErrorColor
'                                    iFlagvNgayHD = 1
'                                    If iFocusFlag = 0 Then
'                                        strFocusSheetName = "'" & .SheetName & "'"
'                                        strFocusCol = str(.Col)
'                                        strFocusRow = str(.Row)
'                                        iFocusFlag = 1
'                                    End If
'                                 ElseIf (CheckValidNgayHD(CStr(vNgayHD)) = False) Then
'                                    .Col = .ColLetterToNumber("G")
'                                    .CellNote = GetAttribute(GetMessageCellById("0124"), "Msg") 'static
'                                    .BackColor = mAlertColor
'                                    iFlagvNgayHD = 0
'                                     If iFocusFlag = 0 Then
'                                        strFocusSheetName = "'" & .SheetName & "'"
'                                        strFocusCol = str(.Col)
'                                        strFocusRow = str(.Row)
'                                        iFocusFlag = 1
'                                      End If
'                                 Else
'                                    .Col = .ColLetterToNumber("G")
'                                    .CellNote = ""
'                                    .BackColor = mNonErrorColor
'                                End If
'                            End If
                            
                            If strTkThangQuy = "TK_QUY" Then
                                If strCellTag = "3" Or strCellTag = "4" Then
                                     If CheckNgayInKyKK(CStr(vNgayHD), TAX_Utilities_v1.ThreeMonths, TAX_Utilities_v1.Year, True) = False Then
                                            .Col = .ColLetterToNumber("G")
                                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                            .CellNote = vErrorNgayPHHD   'static
                                            .BackColor = mErrorColor
                                            iflagNgayPHHD = 1
                                            'capture the first error cell
    '                                        If iFocusFlag = 0 Then
    '                                            strFocusSheetName = "'" & .SheetName & "'"
    '                                            strFocusCol = str(.Col)
    '                                            strFocusRow = str(.Row)
    '                                            iFocusFlag = 1
    '                                        End If
                                        Else
                                             .Col = .ColLetterToNumber("G")
                                             .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                             .CellNote = ""   'static
                                             .BackColor = mNonErrorColor
                                       End If
                                  Else
                                       If CheckNgayTruocKyKK(CStr(vNgayHD), TAX_Utilities_v1.ThreeMonths, TAX_Utilities_v1.Year, True) = False Then
                                            .Col = .ColLetterToNumber("G")
                                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                            .CellNote = vErrorNgayPHHD_1   'static
                                            .BackColor = mErrorColor
                                            iflagNgayPHHD_1 = 1
                                            'capture the first error cell
    '                                        If iFocusFlag = 0 Then
    '                                            strFocusSheetName = "'" & .SheetName & "'"
    '                                            strFocusCol = str(.Col)
    '                                            strFocusRow = str(.Row)
    '                                            iFocusFlag = 1
    '                                        End If
                                        Else
                                             .Col = .ColLetterToNumber("G")
                                             .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                             .CellNote = ""   'static
                                             .BackColor = mNonErrorColor
                                       End If
                                  End If
                                Else
                                    If strCellTag = "3" Or strCellTag = "4" Then
                                           If CheckNgayInKyKK(CStr(vNgayHD), TAX_Utilities_v1.Month, TAX_Utilities_v1.Year, False) = False Then
                                                .Col = .ColLetterToNumber("G")
                                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                                .CellNote = vErrorNgayPHHD   'static
                                                .BackColor = mErrorColor
                                                iflagNgayPHHD = 1
                                                'capture the first error cell
        '                                        If iFocusFlag = 0 Then
        '                                            strFocusSheetName = "'" & .SheetName & "'"
        '                                            strFocusCol = str(.Col)
        '                                            strFocusRow = str(.Row)
        '                                            iFocusFlag = 1
        '                                        End If
                                            Else
                                                 .Col = .ColLetterToNumber("G")
                                                 .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                                 .CellNote = ""   'static
                                                 .BackColor = mNonErrorColor
                                           End If
                                      Else
                                           If CheckNgayTruocKyKK(CStr(vNgayHD), TAX_Utilities_v1.Month, TAX_Utilities_v1.Year, False) = False Then
                                                .Col = .ColLetterToNumber("G")
                                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                                .CellNote = vErrorNgayPHHD_1   'static
                                                .BackColor = mErrorColor
                                                iflagNgayPHHD_1 = 1
                                                'capture the first error cell
        '                                        If iFocusFlag = 0 Then
        '                                            strFocusSheetName = "'" & .SheetName & "'"
        '                                            strFocusCol = str(.Col)
        '                                            strFocusRow = str(.Row)
        '                                            iFocusFlag = 1
        '                                        End If
                                            Else
                                                 .Col = .ColLetterToNumber("G")
                                                 .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                                 .CellNote = ""   'static
                                                 .BackColor = mNonErrorColor
                                           End If
                                      End If
                                End If

                        Else
                                .Col = .ColLetterToNumber("G")
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                        End If
                        
                        If Trim(vTennguoimua) = "" And iWhiteRow = 0 Then
                            .Col = .ColLetterToNumber("H")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgTennguoimua  'static
                            .BackColor = mErrorColor
                            iFlagvTennguoimua = 1
                            
                            'capture the first error cell
'                            If iFocusFlag = 0 Then
'                                strFocusSheetName = "'" & .SheetName & "'"
'                                strFocusCol = str(.Col)
'                                strFocusRow = str(.Row)
'                                iFocusFlag = 1
'                            End If
                        Else
                            .Col = .ColLetterToNumber("H")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
          ' Check MST
                        If Trim(vMST) <> vbNullString Then
                           vMST = Replace(vMST, " ", "")
                           
                            If Len(vMST) = 14 And InStr(vMST, "-") > 0 Then
                                vMST = Left$(vMST, 10) & Right$(vMST, 3)
                            End If
                           
                           vMST = Replace(vMST, "-", "")
                            'vMST = Left(vMST, 10)
    
                           strCheck = ""
                           If Len(vMST) = 10 Or Len(vMST) = 13 Then
                                strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                           End If
                           
                           If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                                .Col = .ColLetterToNumber("I")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrMsgMST   'static
                                .BackColor = mErrorColor
                                iFlagvMst = 1
                                
                                'capture the first error cell
'                                If iFocusFlag = 0 Then
'                                    strFocusSheetName = "'" & .SheetName & "'"
'                                    strFocusCol = str(.Col)
'                                    strFocusRow = str(.Row)
'                                    iFocusFlag = 1
'                                End If
                            End If
                        Else
                            .Col = .ColLetterToNumber("I")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                        
                        If strCellTag = "3" Then
                            If Trim(vDoanhSo) <> "0" And Trim(vGTGT) <> "0" Then
                                If (Abs(Round((vDoanhSo * 0.05), 0) - Round(vGTGT, 0)) > 100000) Or (Abs(Round((vDoanhSo * 0.05), 0) - Round(vGTGT, 0)) > Round(Abs(vDoanhSo) * 0.05 * 0.0001, 0)) Then
                                    .Col = .ColLetterToNumber("M")
                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                    .CellNote = GetAttribute(GetMessageCellById("0118"), "Msg") 'static
                                    .BackColor = mAlertColor
                                Else
                                    .Col = .ColLetterToNumber("M")
                                    .CellNote = ""
                                    .BackColor = mNonErrorColor
                                End If
                            End If
                        End If
                        If strCellTag = "4" Then
                          If Trim(vDoanhSo) <> "0" Or Trim(vGTGT) <> "0" Then
                                If (Abs(Round((vDoanhSo * 0.1), 0) - Round(vGTGT, 0)) > 100000) Or (Abs(Round((vDoanhSo * 0.1), 0) - Round(vGTGT, 0)) > Round(Abs(vDoanhSo) * 0.1 * 0.0001, 0)) Then
                                    .Col = .ColLetterToNumber("M")
                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                    .CellNote = GetAttribute(GetMessageCellById("0118"), "Msg") 'static
                                    .BackColor = mAlertColor
                                 Else
                                    .Col = .ColLetterToNumber("M")
                                    .CellNote = ""
                                    .BackColor = mNonErrorColor
                                End If
                            End If
                        End If
                        
                        
                        
                End Select
           End If
        Next
        
'            ' kiem tra bat buoc nhap PL 01-3
'            If countThuSuat0 > 0 And IIf(TAX_Utilities_v1.NodeValidity.childNodes(3).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) = False Then
'                .Sheet = 2
'                .ActiveSheet = 2
'                If Not .Visible Then
'                    .Visible = True
'                    DisplayMessage "0290", msOKOnly, miWarning
'                    .Visible = False
'                Else
'                    DisplayMessage "0290", msOKOnly, miWarning
'                End If
'                strFocusSheetName = "'" & .SheetName & "'"
'                strFocusCol = str(.Col)
'                strFocusRow = str(.Row)
'                iFocusFlag = 1
'            End If
     End If
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 27, IIf(iFlagvKyHieuHD = 1 And blCheck_S2 = True, "0", "1")
            .SetText .ColLetterToNumber("B"), 28, IIf(iFlagvSoHD = 1 And blCheck_S2 = True, "0", "1")
            .SetText .ColLetterToNumber("B"), 29, IIf(iFlagvNgayHD = 1 And blCheck_S2 = True, "0", "1")
            .SetText .ColLetterToNumber("B"), 30, IIf(iFlagvTennguoimua = 1 And blCheck_S2 = True, "0", "1")
            .SetText .ColLetterToNumber("B"), 31, IIf(iFlagvMst = 1 And blCheck_S2 = True, "0", "1")
            
            .SetText .ColLetterToNumber("B"), 77, IIf(iflagNgayPHHD = 1 And blCheck_S2 = True, "0", "1")
            .SetText .ColLetterToNumber("B"), 79, IIf(iFlagMauSoHD = 1 And blCheck_S2 = True, "0", "1")
            .SetText .ColLetterToNumber("B"), 80, IIf(iFlagSoLienHD = 1 And blCheck_S2 = True, "0", "1")
            .SetText .ColLetterToNumber("B"), 81, IIf(iFlagEmty = 1 And blCheck_S2 = True, "0", "1")
            .SetText .ColLetterToNumber("B"), 89, IIf(iflagNgayPHHD_1 = 1 And blCheck_S2 = True, "0", "1")
'
'            If countThuSuat0 > 0 And IIf(TAX_Utilities_v1.NodeValidity.childNodes(3).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) = False Then
'                .SetText .ColLetterToNumber("B"), 69, "0"
'            Else
'                .SetText .ColLetterToNumber("B"), 69, "1"
'            End If
'            If iFocusFlag = 1 Then
'                ' setcell for focus
'                strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
'                .Col = .ColLetterToNumber("C")
'                .Row = 27
'                .Formula = strFormula
'                .Row = 28
'                .Formula = strFormula
'                .Row = 29
'                .Formula = strFormula
'                .Row = 30
'                .Formula = strFormula
'                .Row = 31
'                .Formula = strFormula
'            End If
         
         'Check MST
         
         
         .Sheet = .SheetCount
    '     .GetText .ColLetterToNumber("B"), 19, vTarget17
    '     .GetText .ColLetterToNumber("B"), 20, vTarget18
         .GetText .ColLetterToNumber("B"), 21, vTarget21
         .GetText .ColLetterToNumber("B"), 22, vTarget22
         .GetText .ColLetterToNumber("B"), 23, vTarget23
         .GetText .ColLetterToNumber("B"), 24, vTarget24
         .GetText .ColLetterToNumber("B"), 25, vTarget25
         .GetText .ColLetterToNumber("B"), 26, vTarget26
         
         If blCheck_S2 Then
            .SetText .ColLetterToNumber("B"), 21, IIf(vTarget21 = "1", "1", "0")
            .SetText .ColLetterToNumber("B"), 22, IIf(vTarget22 = "1", "1", "0")
            .SetText .ColLetterToNumber("B"), 23, IIf(vTarget23 = "1", "1", "0")
            .SetText .ColLetterToNumber("B"), 24, IIf(vTarget24 = "1", "1", "0")
            .SetText .ColLetterToNumber("B"), 25, IIf(vTarget25 = "1", "1", "0")
            .SetText .ColLetterToNumber("B"), 26, IIf(vTarget26 = "1", "1", "0")
          Else
           .SetText .ColLetterToNumber("B"), 21, "1"
           .SetText .ColLetterToNumber("B"), 22, "1"
           .SetText .ColLetterToNumber("B"), 23, "1"
           .SetText .ColLetterToNumber("B"), 24, "1"
           .SetText .ColLetterToNumber("B"), 25, "1"
           .SetText .ColLetterToNumber("B"), 26, "1"
          End If
  
    
    .Sheet = sheetPl01_2_GTGT
    blCheck_S3 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
   If blCheck_S3 = True Then
        For i = 8 To .MaxRows
            iWhiteRow = 1
            .Row = i
            .Col = .ColLetterToNumber("C")
            arrCellTag = Split(.CellTag, "~")
            If UBound(arrCellTag) > 0 Then
            strCellTag = arrCellTag(1)
            '.GetText .ColLetterToNumber("C"), .Row, vMaHoaDon
            '.GetText .ColLetterToNumber("D"), .Row, vKyHieuMauHoaDon
            '.GetText .ColLetterToNumber("E"), .Row, vKyHieuHD
            .GetText .ColLetterToNumber("F"), .Row, vSoHD
            .GetText .ColLetterToNumber("G"), .Row, vNgayHD
            .GetText .ColLetterToNumber("H"), .Row, vTennguoimua
            .GetText .ColLetterToNumber("K"), .Row, vDoanhSo
            .GetText .ColLetterToNumber("M"), .Row, vGTGT
            .GetText .ColLetterToNumber("I"), .Row, vMST
            '.GetText .ColLetterToNumber("L"), .Row, vThuesuat
    
            'If Trim(vKyHieuHD) <> "" Or Trim(vSoHD) <> "" Or Trim(vNgayHD) <> "" Or Trim(vTennguoimua) <> "" Or Trim(vDoanhSo) <> "0" Or Trim(vGTGT) <> "0" Or Trim(vMaHoaDon) <> "" Or Trim(vKyHieuMauHoaDon) <> "" Then iWhiteRow = 0
            If Trim(vSoHD) <> "" Or Trim(vNgayHD) <> "" Or Trim(vTennguoimua) <> "" Or Trim(vDoanhSo) <> "0" Or Trim(vGTGT) <> "0" Then iWhiteRow = 0
                Select Case strCellTag
                    Case "1", "2", "3", "4", "5"
                        ' kiem tra ky hieu mau HD
'                        If Trim(vKyHieuMauHoaDon) = "" And iWhiteRow = 0 Then
'                            If Trim(vMaHoaDon) = "TT120" Then
'                                .Col = .ColLetterToNumber("D")
'                                .CellNote = ""  'static
'                                .BackColor = mNonErrorColor
'                            Else
'                                .Col = .ColLetterToNumber("D")
'                                .CellNote = errEmty  'static
'                                .BackColor = mErrorColor
'                                iFlagEmty = 1
'                            End If
'                        Else
'                            strLoai = getLoaiHD(CStr(vMaHoaDon))
'                            result = CheckMauSoHD(CStr(vKyHieuMauHoaDon), strLoai, CStr(vMaHoaDon))
'                            If result = "1" And iWhiteRow = 0 Then
'                                .Col = .ColLetterToNumber("D")
'                                .CellNote = errMauSoHD  'static
'                                .BackColor = mErrorColor
'                                iFlagMauSoHD = 1
'                            ElseIf result = "2" And iWhiteRow = 0 Then
'                                .Col = .ColLetterToNumber("D")
'                                .CellNote = errSoLienHD  'static
'                                .BackColor = mErrorColor
'                                iFlagSoLienHD = 1
'                            Else
'                                .Col = .ColLetterToNumber("D")
'                                .CellNote = ""  'static
'                                .BackColor = mNonErrorColor
'                            End If
'                        End If
                        ' end
                    
                    
'                        If Trim(vKyHieuHD) = "" And iWhiteRow = 0 Then
'                            If Trim(vMaHoaDon) = "TT120" Then
'                                 .Col = .ColLetterToNumber("E")
'                                .CellNote = ""
'                                .BackColor = mNonErrorColor
'                            Else
'                                .Col = .ColLetterToNumber("E")
'                                .CellNote = vErrMsgKyHieuHD  'static
'                                .BackColor = mErrorColor
'                                iFlagvKyHieuHD = 1
'                            End If
'                            'capture the first error cell
'    '                        If iFocusFlag = 0 Then
'    '                            strFocusSheetName = "'" & .SheetName & "'"
'    '                            strFocusCol = str(.Col)
'    '                            strFocusRow = str(.Row)
'    '                            iFocusFlag = 1
'    '                        End If
'                        Else
'                            .Col = .ColLetterToNumber("E")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
'                        End If
                        
                        If Trim(vSoHD) = "" And iWhiteRow = 0 Then
                            .Col = .ColLetterToNumber("F")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgSoHD  'static
                            .BackColor = mErrorColor
                            iFlagvSoHD = 1
                            'capture the first error cell
'                            If iFocusFlag = 0 Then
'                                strFocusSheetName = "'" & .SheetName & "'"
'                                strFocusCol = str(.Col)
'                                strFocusRow = str(.Row)
'                                iFocusFlag = 1
'                            End If
                        Else
                            .Col = .ColLetterToNumber("F")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                        
                        If Trim(vNgayHD) = "" And iWhiteRow = 0 Then
                            .Col = .ColLetterToNumber("G")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgNgayHD  'static
                            .BackColor = mErrorColor
                            iFlagvNgayHD = 1
                            
                            'capture the first error cell
'                            If iFocusFlag = 0 Then
'                                strFocusSheetName = "'" & .SheetName & "'"
'                                strFocusCol = str(.Col)
'                                strFocusRow = str(.Row)
'                                iFocusFlag = 1
'                            End If
                        ElseIf Trim(vNgayHD) <> "" And iWhiteRow = 0 Then
                            ' To khai quy
'                            If strTkThangQuy = "TK_QUY" Then
'                                vNgayHt = Date
'                                arrDate = Split(vNgayHD, "/")
'                                vdate = DateSerial(CInt(arrDate(2)), CInt(arrDate(1)), CInt(arrDate(0)))
'                                If CInt(Right(vNgayHD, 4)) > CInt(TAX_Utilities_v1.Year) Then
'                                       .Col = .ColLetterToNumber("G")
'                                       .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                       .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
'                                       .BackColor = mErrorColor
'                                       iFlagvNgayHD = 1
'                                       If iFocusFlag = 0 Then
'                                           strFocusSheetName = "'" & .SheetName & "'"
'                                           strFocusCol = str(.Col)
'                                           strFocusRow = str(.Row)
'                                           iFocusFlag = 1
'                                       End If
'                                    ElseIf CInt(Right(vNgayHD, 4)) = CInt(TAX_Utilities_v1.Year) And convertThang2Quy(CInt(Right("00" & Mid(vNgayHD, 4, 2), 2))) > CInt(TAX_Utilities_v1.ThreeMonths) Then
'                                       .Col = .ColLetterToNumber("G")
'                                       .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                       .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
'                                       .BackColor = mErrorColor
'                                       iFlagvNgayHD = 1
'                                       If iFocusFlag = 0 Then
'                                           strFocusSheetName = "'" & .SheetName & "'"
'                                           strFocusCol = str(.Col)
'                                           strFocusRow = str(.Row)
'                                           iFocusFlag = 1
'                                       End If
'                                    ElseIf (DateDiff("D", vNgayHt, vdate) > 0) And strCellTag <> "2" Then
'                                       .Col = .ColLetterToNumber("G")
'                                       .CellNote = GetAttribute(GetMessageCellById("0273"), "Msg") 'static
'                                       .BackColor = mErrorColor
'                                       iFlagvNgayHD = 1
'                                        If iFocusFlag = 0 Then
'                                           strFocusSheetName = "'" & .SheetName & "'"
'                                           strFocusCol = str(.Col)
'                                           strFocusRow = str(.Row)
'                                           iFocusFlag = 1
'                                         End If
'                                    Else
'                                       .Col = .ColLetterToNumber("G")
'                                       .CellNote = ""
'                                       .BackColor = mNonErrorColor
'                                   End If
'                            Else
'                                If CInt(Right(vNgayHD, 4)) > CInt(TAX_Utilities_v1.Year) Then
'                                       .Col = .ColLetterToNumber("G")
'                                       .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                       .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
'                                       .BackColor = mErrorColor
'                                       iFlagvNgayHD = 1
'                                       If iFocusFlag = 0 Then
'                                           strFocusSheetName = "'" & .SheetName & "'"
'                                           strFocusCol = str(.Col)
'                                           strFocusRow = str(.Row)
'                                           iFocusFlag = 1
'                                       End If
'                                    ElseIf CInt(Right(vNgayHD, 4)) = CInt(TAX_Utilities_v1.Year) And CInt(Right("00" & Mid(vNgayHD, 4, 2), 2)) > CInt(TAX_Utilities_v1.Month) Then
'                                       .Col = .ColLetterToNumber("G")
'                                       .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                       .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
'                                       .BackColor = mErrorColor
'                                       iFlagvNgayHD = 1
'                                       If iFocusFlag = 0 Then
'                                           strFocusSheetName = "'" & .SheetName & "'"
'                                           strFocusCol = str(.Col)
'                                           strFocusRow = str(.Row)
'                                           iFocusFlag = 1
'                                       End If
'                                    ElseIf (CheckValidNgayHD(CStr(vNgayHD)) = False) And strCellTag <> "2" Then
'                                       .Col = .ColLetterToNumber("G")
'                                       .CellNote = GetAttribute(GetMessageCellById("0124"), "Msg") 'static
'                                       .BackColor = mAlertColor
'                                        If iFocusFlag = 0 Then
'                                           strFocusSheetName = "'" & .SheetName & "'"
'                                           strFocusCol = str(.Col)
'                                           strFocusRow = str(.Row)
'                                           iFocusFlag = 1
'                                         End If
'                                    Else
'                                       .Col = .ColLetterToNumber("G")
'                                       .CellNote = ""
'                                       .BackColor = mNonErrorColor
'                                   End If
'                              End If
                             ' Kiem tra ngay phat hanh nho hon ngay cuoi ky ke khai
                             If strTkThangQuy = "TK_QUY" Then
                                 If DateDiff("D", getNgayCuoiKyKK(TAX_Utilities_v1.ThreeMonths, TAX_Utilities_v1.Year, True), Format(vNgayHD, "dd/mm/yyyy")) > 0 Then
                                     .Col = .ColLetterToNumber("G")
                                     .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                     .CellNote = vErrorPhatHanhHD   'static
                                     .BackColor = mErrorColor
                                     iflagNgayPhatHanhHD = 1
                                     'capture the first error cell
'                                     If iFocusFlag = 0 Then
'                                         strFocusSheetName = "'" & .SheetName & "'"
'                                         strFocusCol = str(.Col)
'                                         strFocusRow = str(.Row)
'                                         iFocusFlag = 1
'                                     End If
                                 Else
                                      .Col = .ColLetterToNumber("G")
                                      .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                      .CellNote = ""   'static
                                      .BackColor = mNonErrorColor
                                End If
                             Else
                                If DateDiff("D", getNgayCuoiKyKK(TAX_Utilities_v1.Month, TAX_Utilities_v1.Year, False), Format(vNgayHD, "dd/mm/yyyy")) > 0 Then
                                     .Col = .ColLetterToNumber("G")
                                     .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                     .CellNote = vErrorPhatHanhHD   'static
                                     .BackColor = mErrorColor
                                     iflagNgayPhatHanhHD = 1
                                     'capture the first error cell
'                                     If iFocusFlag = 0 Then
'                                         strFocusSheetName = "'" & .SheetName & "'"
'                                         strFocusCol = str(.Col)
'                                         strFocusRow = str(.Row)
'                                         iFocusFlag = 1
'                                     End If
                                 Else
                                      .Col = .ColLetterToNumber("G")
                                      .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                      .CellNote = ""   'static
                                      .BackColor = mNonErrorColor
                                End If
                             End If
                            
    '                        If CheckValidNgayHD(CStr(vNgayHD)) = False Then
    '                           .Col = .ColLetterToNumber("E")
    '                           .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    '                           .CellNote = GetAttribute(GetMessageCellById("0112"), "Msg") 'static
    '                           .BackColor = mErrorColor
    '                           iFlagvNgayHD = 1
    '                            'capture the first error cell
    '                            If iFocusFlag = 0 Then
    '                                strFocusSheetName = "'" & .SheetName & "'"
    '                                strFocusCol = str(.Col)
    '                                strFocusRow = str(.Row)
    '                                iFocusFlag = 1
    '                            End If
    '                         Else
    '                            .Col = .ColLetterToNumber("E")
    '                            .CellNote = ""
    '                            .BackColor = mNonErrorColor
    '                         End If
                        Else
                            .Col = .ColLetterToNumber("G")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                        
                        If Trim(vTennguoimua) = "" And iWhiteRow = 0 Then
                            .Col = .ColLetterToNumber("H")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgTennguoimua  'static
                            .BackColor = mErrorColor
                            iFlagvTennguoimua = 1
                            
                            'capture the first error cell
'                            If iFocusFlag = 0 Then
'                                strFocusSheetName = "'" & .SheetName & "'"
'                                strFocusCol = str(.Col)
'                                strFocusRow = str(.Row)
'                                iFocusFlag = 1
'                            End If
                        Else
                            .Col = .ColLetterToNumber("H")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
             'Check MST
                        If Trim(vMST) <> vbNullString Then
                        vMST = Replace(vMST, " ", "")
                        
                        If Len(vMST) = 14 And InStr(vMST, "-") > 0 Then
                             vMST = Left$(vMST, 10) & Right$(vMST, 3)
                        End If
                        
                        vMST = Replace(vMST, "-", "")
                         'vMST = Left(vMST, 10)
                        strCheck = ""
                           If Len(vMST) = 10 Or Len(vMST) = 13 Then
                                strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                           End If
                           If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                                .Col = .ColLetterToNumber("I")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrMsgMST   'static
                                .BackColor = mErrorColor
                                iFlagvMst = 1
                                
                                'capture the first error cell
                                If iFocusFlag = 0 Then
                                    strFocusSheetName = "'" & .SheetName & "'"
                                    strFocusCol = str(.Col)
                                    strFocusRow = str(.Row)
                                    iFocusFlag = 1
                                End If
                            End If
                        Else
                            .Col = .ColLetterToNumber("I")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                        
'                        If CDbl(vThuesuat) = 0 Or CDbl(vThuesuat) = 5 Or CDbl(vThuesuat) = 10 Then
'                            .Col = .ColLetterToNumber("L")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
'                        Else
'                            .Col = .ColLetterToNumber("L")
'                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                            .CellNote = vErrMsgThuesuat   'static
'                            .BackColor = mErrorColor
'                            iFlagvThuesuat = 1
'                            'capture the first error cell
'                            If iFocusFlag = 0 Then
'                                strFocusSheetName = "'" & .SheetName & "'"
'                                strFocusCol = str(.Col)
'                                strFocusRow = str(.Row)
'                                iFocusFlag = 1
'                            End If
'                        End If
                        
                        ' kiem tra nguong cua thue GTGT
                        If strCellTag = "1" Or strCellTag = "2" Or strCellTag = "3" Or strCellTag = "4" Or strCellTag = "5" Then
                            If Val(vGTGT) <> 0 And Val(vGTGT) >= Val(vDoanhSo) Then
                                .Col = .ColLetterToNumber("M")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrThueGTGT
                                '.BackColor = mErrorColor
                                .BackColor = mAlertColor
                                'iFlagThueGTGT = 1
                            Else
                                .Col = .ColLetterToNumber("M")
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            End If
                        End If
                        
                End Select
        
           End If
        Next
    End If
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 32, IIf(iFlagvKyHieuHD = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 33, IIf(iFlagvSoHD = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 34, IIf(iFlagvNgayHD = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 35, IIf(iFlagvTennguoimua = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 36, IIf(iFlagvMst = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 37, IIf(iFlagvThuesuat = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 78, IIf(iflagNgayPhatHanhHD = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 79, IIf(iFlagMauSoHD = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 80, IIf(iFlagSoLienHD = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 81, IIf(iFlagEmty = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 88, IIf(iFlagThueGTGT = 1 And blCheck_S3 = True, "0", "1")
'        If iFocusFlag = 1 Then
'            ' setcell for focus
'
'            strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
'            .Col = .ColLetterToNumber("C")
'            .Row = 32
'            .Formula = strFormula
'            .Row = 33
'            .Formula = strFormula
'            .Row = 34
'            .Formula = strFormula
'            .Row = 35
'            .Formula = strFormula
'            .Row = 36
'            .Formula = strFormula
'            .Row = 37
'            .Formula = strFormula
'
'        End If
     
     .Sheet = .SheetCount
     .GetText .ColLetterToNumber("B"), 19, vTarget19
     .GetText .ColLetterToNumber("B"), 20, vTarget20
    
     
     
     If blCheck_S3 Then
        .SetText .ColLetterToNumber("B"), 19, IIf(vTarget19 = "1", "1", "0")
        .SetText .ColLetterToNumber("B"), 20, IIf(vTarget20 = "1", "1", "0")
      Else
       .SetText .ColLetterToNumber("B"), 19, "1"
       .SetText .ColLetterToNumber("B"), 20, "1"
      End If
    
    .GetText .ColLetterToNumber("B"), 38, vTarget38
    .GetText .ColLetterToNumber("B"), 39, vTarget39
   ' .GetText .ColLetterToNumber("E"), 40, vErrMsg
  
'    ' Phu luc 01-3
'    .Sheet = 4
'    blCheck_S4 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    If blCheck_S4 = True Then
'        .Col = .ColLetterToNumber("B")
'        .Row = 10
'        i = 1
'        Do
'              iWhiteRow = 1
'              .GetText .ColLetterToNumber("C"), .Row, soHopDongXK
'              .GetText .ColLetterToNumber("D"), .Row, ngayHopDongXK
'              .GetText .ColLetterToNumber("E"), .Row, ngoaiTeXK
'              .GetText .ColLetterToNumber("F"), .Row, tienVNDXK
'
'              .GetText .ColLetterToNumber("G"), .Row, hinhThucTTXK
'              .GetText .ColLetterToNumber("H"), .Row, thoHanTTXK
'              .GetText .ColLetterToNumber("I"), .Row, toKhaiHangHoaXK
'              .GetText .ColLetterToNumber("J"), .Row, ngayHangHoaXK
'              .GetText .ColLetterToNumber("K"), .Row, ngoaiTeHangHoaXK
'              .GetText .ColLetterToNumber("L"), .Row, tienVNDHangHoaXK
'
'              .GetText .ColLetterToNumber("M"), .Row, soHoaDonXK
'              .GetText .ColLetterToNumber("N"), .Row, ngayHoaDonXK
'              .GetText .ColLetterToNumber("O"), .Row, ngoaiTeHoaDonXK
'              .GetText .ColLetterToNumber("P"), .Row, tienVNDHoaDonXK
'
'              .GetText .ColLetterToNumber("Q"), .Row, soChungTu
'              .GetText .ColLetterToNumber("R"), .Row, ngayChungTu
'              .GetText .ColLetterToNumber("S"), .Row, ngoaiTeChungTu
'              .GetText .ColLetterToNumber("T"), .Row, tienVNDChungTu
'
'              .GetText .ColLetterToNumber("U"), .Row, soVanBan
'              .GetText .ColLetterToNumber("V"), .Row, ngayVanBan
'              .GetText .ColLetterToNumber("W"), .Row, ngoaiTeVanBan
'              .GetText .ColLetterToNumber("X"), .Row, tienVNDVanBan
'
'              .GetText .ColLetterToNumber("Y"), .Row, soHDNK
'              .GetText .ColLetterToNumber("Z"), .Row, ngayHDNK
'              .GetText .ColLetterToNumber("AA"), .Row, ngoaiTeHDNK
'              .GetText .ColLetterToNumber("AB"), .Row, tienVNDHDNK
'
'              .GetText .ColLetterToNumber("AC"), .Row, soHangHoaNK
'              .GetText .ColLetterToNumber("AD"), .Row, ngayHangHoaNK
'              .GetText .ColLetterToNumber("AE"), .Row, ngoaiTeHangHoaNK
'              .GetText .ColLetterToNumber("AF"), .Row, tienVNDHangHoaNK
'              .GetText .ColLetterToNumber("AG"), .Row, chungTuTTbutru
'              .GetText .ColLetterToNumber("AH"), .Row, ghiChu
'
'              If Trim$(soHopDongXK) <> "" Or Trim$(ngayHopDongXK) <> "" Or Trim$(ngoaiTeXK) <> "" Or Val(tienVNDXK) > 0 Or Trim$(hinhThucTTXK) <> "" _
'              Or Trim$(thoHanTTXK) <> "" Or Trim$(toKhaiHangHoaXK) <> "" Or Trim$(ngayHangHoaXK) <> "" Or Trim$(ngoaiTeHangHoaXK) <> "" Or Val(tienVNDHangHoaXK) > 0 _
'              Or Trim$(soHoaDonXK) <> "" Or Trim$(ngayHoaDonXK) <> "" Or Trim$(ngoaiTeHoaDonXK) <> "" Or Val(tienVNDHoaDonXK) > 0 _
'              Or Trim$(soChungTu) <> "" Or Trim$(ngayChungTu) <> "" Or Trim$(ngoaiTeChungTu) <> "" Or Val(tienVNDChungTu) > 0 _
'              Or Trim$(soVanBan) <> "" Or Trim$(ngayVanBan) <> "" Or Trim$(ngoaiTeVanBan) <> "" Or Val(tienVNDVanBan) > 0 _
'              Or Trim$(soHDNK) <> "" Or Trim$(ngayHDNK) <> "" Or Trim$(ngoaiTeHDNK) <> "" Or Val(tienVNDHDNK) > 0 _
'              Or Trim$(soHangHoaNK) <> "" Or Trim$(ngayHangHoaNK) <> "" Or Trim$(ngoaiTeHangHoaNK) <> "" Or Val(tienVNDHangHoaNK) > 0 Or Trim$(chungTuTTbutru) <> "" Or Trim$(ghiChu) <> "" Then iWhiteRow = 0
'
'              If iWhiteRow = 0 Then
'                ' kiem tra ngay hop dong xuat khau
''                    If (Trim(ngayHopDongXK) = "" Or ngayHopDongXK = vbNullString) Then
''                          .Col = .ColLetterToNumber("D")
''                          .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                          .CellNote = ""   'static
''                          .BackColor = mNonErrorColor
''                    Else
''                          If DateDiff("D", Date, Format(ngayHopDongXK, "dd/mm/yyyy")) > 0 Then
''                                 .Col = .ColLetterToNumber("D")
''                                 .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                 .CellNote = vErrorNgayHDXK   'static
''                                 .BackColor = mErrorColor
''                                 iflagNgayHopDongXK = 1
''                                 'capture the first error cell
''                                 If iFocusFlag = 0 Then
''                                     strFocusSheetName = "'" & .SheetName & "'"
''                                     strFocusCol = str(.Col)
''                                     strFocusRow = str(.Row)
''                                     iFocusFlag = 1
''                                 End If
''                             Else
''                                  .Col = .ColLetterToNumber("D")
''                                  .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                  .CellNote = ""   'static
''                                  .BackColor = mNonErrorColor
''                            End If
''                    End If
'
'                    'Kiem tra ngay dang ky to khai xuat khau
''                     If (Trim(ngayHangHoaXK) = "" Or ngayHangHoaXK = vbNullString) Then
''                          .Col = .ColLetterToNumber("J")
''                          .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                          .CellNote = ""   'static
''                          .BackColor = mNonErrorColor
''                    Else
''                        If strTkThangQuy = "TK_QUY" Then
''                          If CheckNgayInKyKK(CStr(ngayHangHoaXK), TAX_Utilities_v1.ThreeMonths, TAX_Utilities_v1.Year, True) = False Then
''                                 .Col = .ColLetterToNumber("J")
''                                 .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                 .CellNote = vErrorNgayDKTKXK   'static
''                                 .BackColor = mErrorColor
''                                 iflagNgayDangKyTKXK = 1
''                                 'capture the first error cell
''                                 If iFocusFlag = 0 Then
''                                     strFocusSheetName = "'" & .SheetName & "'"
''                                     strFocusCol = str(.Col)
''                                     strFocusRow = str(.Row)
''                                     iFocusFlag = 1
''                                 End If
''                             Else
''                                  .Col = .ColLetterToNumber("J")
''                                  .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                  .CellNote = ""   'static
''                                  .BackColor = mNonErrorColor
''                            End If
''                         Else
''                            If CheckNgayInKyKK(CStr(ngayHangHoaXK), TAX_Utilities_v1.Month, TAX_Utilities_v1.Year, False) = False Then
''                                 .Col = .ColLetterToNumber("J")
''                                 .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                 .CellNote = vErrorNgayDKTKXK   'static
''                                 .BackColor = mErrorColor
''                                 iflagNgayDangKyTKXK = 1
''                                 'capture the first error cell
''                                 If iFocusFlag = 0 Then
''                                     strFocusSheetName = "'" & .SheetName & "'"
''                                     strFocusCol = str(.Col)
''                                     strFocusRow = str(.Row)
''                                     iFocusFlag = 1
''                                 End If
''                             Else
''                                  .Col = .ColLetterToNumber("J")
''                                  .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                  .CellNote = ""   'static
''                                  .BackColor = mNonErrorColor
''                            End If
''                         End If
''                    End If
'
'                    ' Kiem tra ngay hoa don xuat khau
'                    If (Trim(ngayHoaDonXK) = "" Or ngayHoaDonXK = vbNullString) Then
'                          .Col = .ColLetterToNumber("N")
'                          .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                          .CellNote = ""   'static
'                          .BackColor = mNonErrorColor
'                    Else
'                        If strTkThangQuy = "TK_QUY" Then
'                            If CheckNgayInKyKK(CStr(ngayHoaDonXK), TAX_Utilities_v1.ThreeMonths, TAX_Utilities_v1.Year, True) = False Then
'                                 .Col = .ColLetterToNumber("N")
'                                 .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                 .CellNote = vErrorNgayHoaDonXK   'static
'                                 .BackColor = mErrorColor
'                                 iflagNgayHoaDonXK = 1
'                                 'capture the first error cell
'                                 If iFocusFlag = 0 Then
'                                     strFocusSheetName = "'" & .SheetName & "'"
'                                     strFocusCol = str(.Col)
'                                     strFocusRow = str(.Row)
'                                     iFocusFlag = 1
'                                 End If
'                             Else
'                                  .Col = .ColLetterToNumber("N")
'                                  .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                  .CellNote = ""   'static
'                                  .BackColor = mNonErrorColor
'                            End If
'                         Else
'                            If CheckNgayInKyKK(CStr(ngayHoaDonXK), TAX_Utilities_v1.Month, TAX_Utilities_v1.Year, False) = False Then
'                                 .Col = .ColLetterToNumber("N")
'                                 .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                 .CellNote = vErrorNgayHoaDonXK   'static
'                                 .BackColor = mErrorColor
'                                 iflagNgayHoaDonXK = 1
'                                 'capture the first error cell
'                                 If iFocusFlag = 0 Then
'                                     strFocusSheetName = "'" & .SheetName & "'"
'                                     strFocusCol = str(.Col)
'                                     strFocusRow = str(.Row)
'                                     iFocusFlag = 1
'                                 End If
'                             Else
'                                  .Col = .ColLetterToNumber("N")
'                                  .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                  .CellNote = ""   'static
'                                  .BackColor = mNonErrorColor
'                            End If
'                         End If
'                    End If
'
'                    ' Kiem tra ngay chung tu thanh toan
''                    If (Trim(ngayChungTu) = "" Or ngayChungTu = vbNullString) Then
''                          .Col = .ColLetterToNumber("R")
''                          .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                          .CellNote = ""   'static
''                          .BackColor = mNonErrorColor
''                    Else
''                          If DateDiff("D", Date, Format(ngayChungTu, "dd/mm/yyyy")) > 0 Then
''                                 .Col = .ColLetterToNumber("R")
''                                 .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                 .CellNote = vErrorNgayChungTuTT   'static
''                                 .BackColor = mErrorColor
''                                 iflagNgayChungTuTT = 1
''                                 'capture the first error cell
''                                 If iFocusFlag = 0 Then
''                                     strFocusSheetName = "'" & .SheetName & "'"
''                                     strFocusCol = str(.Col)
''                                     strFocusRow = str(.Row)
''                                     iFocusFlag = 1
''                                 End If
''                             Else
''                                  .Col = .ColLetterToNumber("R")
''                                  .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                  .CellNote = ""   'static
''                                  .BackColor = mNonErrorColor
''                            End If
''                    End If
'
'                    'Kiem tra ngay van ban
''                    If (Trim(ngayVanBan) = "" Or ngayVanBan = vbNullString) Then
''                          .Col = .ColLetterToNumber("V")
''                          .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                          .CellNote = ""   'static
''                          .BackColor = mNonErrorColor
''                    Else
''                          If DateDiff("D", Date, Format(ngayVanBan, "dd/mm/yyyy")) > 0 Then
''                                 .Col = .ColLetterToNumber("V")
''                                 .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                 .CellNote = vErrorNgayVBXN   'static
''                                 .BackColor = mErrorColor
''                                 iflagNgayVBXN = 1
''                                 'capture the first error cell
''                                 If iFocusFlag = 0 Then
''                                     strFocusSheetName = "'" & .SheetName & "'"
''                                     strFocusCol = str(.Col)
''                                     strFocusRow = str(.Row)
''                                     iFocusFlag = 1
''                                 End If
''                             Else
''                                  .Col = .ColLetterToNumber("V")
''                                  .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                  .CellNote = ""   'static
''                                  .BackColor = mNonErrorColor
''                            End If
''                    End If
'
'                    'kiem tra ngay hop dong nhap khau
''                    If (Trim(ngayHDNK) = "" Or ngayHDNK = vbNullString) Then
''                          .Col = .ColLetterToNumber("Z")
''                          .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                          .CellNote = ""   'static
''                          .BackColor = mNonErrorColor
''                    Else
''                          If DateDiff("D", Date, Format(ngayVanBan, "dd/mm/yyyy")) > 0 Then
''                                 .Col = .ColLetterToNumber("Z")
''                                 .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                 .CellNote = vErrorNgayHDNK   'static
''                                 .BackColor = mErrorColor
''                                 iflagNgayHDNK = 1
''                                 'capture the first error cell
''                                 If iFocusFlag = 0 Then
''                                     strFocusSheetName = "'" & .SheetName & "'"
''                                     strFocusCol = str(.Col)
''                                     strFocusRow = str(.Row)
''                                     iFocusFlag = 1
''                                 End If
''                             Else
''                                  .Col = .ColLetterToNumber("Z")
''                                  .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                  .CellNote = ""   'static
''                                  .BackColor = mNonErrorColor
''                            End If
''                    End If
'
'                    'Kiem tra ngay hang hoa nhap khau
''                    If (Trim(ngayHangHoaNK) = "" Or ngayHangHoaNK = vbNullString) Then
''                          .Col = .ColLetterToNumber("AD")
''                          .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                          .CellNote = ""   'static
''                          .BackColor = mNonErrorColor
''                    Else
''                        If strTkThangQuy = "TK_QUY" Then
''                            If CheckNgayInKyKK(CStr(ngayHangHoaNK), TAX_Utilities_v1.ThreeMonths, TAX_Utilities_v1.Year, True) = False Then
''                                 .Col = .ColLetterToNumber("AD")
''                                 .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                 .CellNote = vErrorNgayTKNK   'static
''                                 .BackColor = mErrorColor
''                                 iflagNgayTKNK = 1
''                                 'capture the first error cell
''                                 If iFocusFlag = 0 Then
''                                     strFocusSheetName = "'" & .SheetName & "'"
''                                     strFocusCol = str(.Col)
''                                     strFocusRow = str(.Row)
''                                     iFocusFlag = 1
''                                 End If
''                             Else
''                                  .Col = .ColLetterToNumber("AD")
''                                  .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                  .CellNote = ""   'static
''                                  .BackColor = mNonErrorColor
''                            End If
''                         Else
''                            If CheckNgayInKyKK(CStr(ngayHangHoaNK), TAX_Utilities_v1.Month, TAX_Utilities_v1.Year, False) = False Then
''                                 .Col = .ColLetterToNumber("AD")
''                                 .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                 .CellNote = vErrorNgayTKNK   'static
''                                 .BackColor = mErrorColor
''                                 iflagNgayTKNK = 1
''                                 'capture the first error cell
''                                 If iFocusFlag = 0 Then
''                                     strFocusSheetName = "'" & .SheetName & "'"
''                                     strFocusCol = str(.Col)
''                                     strFocusRow = str(.Row)
''                                     iFocusFlag = 1
''                                 End If
''                             Else
''                                  .Col = .ColLetterToNumber("AD")
''                                  .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                                  .CellNote = ""   'static
''                                  .BackColor = mNonErrorColor
''                            End If
''                         End If
''                    End If
'
'              End If
'
'
'             .Col = .ColLetterToNumber("B")
'             .Row = i + 10
'             i = i + 1
'        Loop Until .Text = "aa"
'          .Sheet = .SheetCount
'        .SetText .ColLetterToNumber("B"), 70, IIf(iflagNgayHopDongXK = "1", "0", "1")
'        .SetText .ColLetterToNumber("B"), 71, IIf(iflagNgayDangKyTKXK = 1, "0", "1")
'        .SetText .ColLetterToNumber("B"), 72, IIf(iflagNgayHoaDonXK = 1, "0", "1")
'        .SetText .ColLetterToNumber("B"), 73, IIf(iflagNgayChungTuTT = 1, "0", "1")
'        .SetText .ColLetterToNumber("B"), 74, IIf(iflagNgayVBXN = 1, "0", "1")
'
'        .SetText .ColLetterToNumber("B"), 75, IIf(iflagNgayHDNK = 1, "0", "1")
'        .SetText .ColLetterToNumber("B"), 76, IIf(iflagNgayTKNK = 1, "0", "1")
'
'    End If
    
'    ' Phu luc 01-4A
'    .Sheet = 5
'    blCheck_S5 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'
    .Sheet = sheetPl01_5_GTGT
     blCheck_S7 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    If blCheck_S7 = True Then
    .Col = .ColLetterToNumber("B")
    .Row = 6
    i = 1
    Do
          iWhiteRow = 1
          .GetText .ColLetterToNumber("C"), .Row, vSoChungTu
          .GetText .ColLetterToNumber("D"), .Row, vNgayNop
          .GetText .ColLetterToNumber("E"), .Row, vNoiNop
          .GetText .ColLetterToNumber("G"), .Row, vCoQuanQL
          .GetText .ColLetterToNumber("I"), .Row, vSoTienDN
          If Trim(vSoChungTu) <> "" Or Trim(vNgayNop) <> "" Or Trim(vNoiNop) <> "" Or Trim(vCoQuanQL) <> "" Or Trim(vSoTienDN) <> "0" Then
            iWhiteRow = 0
          End If
          
          
          If (Trim(vSoChungTu) = "" Or vSoChungTu = vbNullString) And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("C")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrorSoChungTu   'static
                .BackColor = mErrorColor
                iflagSoChungTu = 1
                'capture the first error cell
                If iFocusFlag = 0 Then
                    strFocusSheetName = "'" & .SheetName & "'"
                    strFocusCol = str(.Col)
                    strFocusRow = str(.Row)
                    iFocusFlag = 1
                End If
          Else
                .Col = .ColLetterToNumber("C")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""   'static
                .BackColor = mNonErrorColor
          End If
          ' Kiem tra noi nop
           If (Trim(vNoiNop) = "" Or vNoiNop = vbNullString) And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("E")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrorSoChungTu   'static
                .BackColor = mErrorColor
                iflagSoChungTu = 1
                'capture the first error cell
                If iFocusFlag = 0 Then
                    strFocusSheetName = "'" & .SheetName & "'"
                    strFocusCol = str(.Col)
                    strFocusRow = str(.Row)
                    iFocusFlag = 1
                End If
          Else
                .Col = .ColLetterToNumber("E")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""   'static
                .BackColor = mNonErrorColor
          End If
          ' Kiem tra co quan thue
          If (Trim(vCoQuanQL) = "" Or vCoQuanQL = vbNullString) And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("G")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrorSoChungTu   'static
                .BackColor = mErrorColor
                iflagSoChungTu = 1
                'capture the first error cell
                If iFocusFlag = 0 Then
                    strFocusSheetName = "'" & .SheetName & "'"
                    strFocusCol = str(.Col)
                    strFocusRow = str(.Row)
                    iFocusFlag = 1
                End If
          Else
                .Col = .ColLetterToNumber("G")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""   'static
                .BackColor = mNonErrorColor
          End If
          ' Kiem tra so thue
          If (Trim(vSoTienDN) = "0" Or vSoTienDN = vbNullString) And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("I")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrorSoChungTu   'static
                .BackColor = mErrorColor
                iflagSoChungTu = 1
                'capture the first error cell
                If iFocusFlag = 0 Then
                    strFocusSheetName = "'" & .SheetName & "'"
                    strFocusCol = str(.Col)
                    strFocusRow = str(.Row)
                    iFocusFlag = 1
                End If
          Else
                .Col = .ColLetterToNumber("I")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""   'static
                .BackColor = mNonErrorColor
          End If
          
          
          If (Trim(vNgayNop) = "" Or vNgayNop = vbNullString) And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("D")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrorSoChungTu   'static
                .BackColor = mErrorColor
                iflagSoChungTu = 1
                'capture the first error cell
                If iFocusFlag = 0 Then
                    strFocusSheetName = "'" & .SheetName & "'"
                    strFocusCol = str(.Col)
                    strFocusRow = str(.Row)
                    iFocusFlag = 1
                End If
          Else
                If iWhiteRow = 0 Then
                    If DateDiff("D", Date, Format(vNgayNop, "dd/mm/yyyy")) > 0 Then
                           .Col = .ColLetterToNumber("D")
                           .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                           .CellNote = vErrorNgayNop   'static
                           .BackColor = mErrorColor
                           iflagNgayNop = 1
                           'capture the first error cell
                           If iFocusFlag = 0 Then
                               strFocusSheetName = "'" & .SheetName & "'"
                               strFocusCol = str(.Col)
                               strFocusRow = str(.Row)
                               iFocusFlag = 1
                           End If
                       Else
                            .Col = .ColLetterToNumber("D")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = ""   'static
                            .BackColor = mNonErrorColor
                      End If
                Else
                            .Col = .ColLetterToNumber("D")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = ""   'static
                            .BackColor = mNonErrorColor
                End If
          End If
                 
         ' .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
                 
         .Col = .ColLetterToNumber("B")
         .Row = i + 6
         i = i + 1
    Loop Until .Text = "aa"
      .Sheet = .SheetCount
    .SetText .ColLetterToNumber("B"), 57, IIf(iflagSoChungTu = "1", "0", "1")
    .SetText .ColLetterToNumber("B"), 55, IIf(iflagNgayNop = 1, "0", "1")
 End If
    
    .Sheet = sheetPl01_6_GTGT
    iFlagvMst_PL6 = 1
    blCheck_S8 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    For i = 8 To .MaxRows - 13
    If blCheck_S8 Then
        .Col = .ColLetterToNumber("B")
        .Row = rowStartPL01_6
        i = 1
        Do
    
            iWhiteRow = 1
            '.Row = i
            .Col = .ColLetterToNumber("C")
            .GetText .ColLetterToNumber("C"), .Row, vKyHieuHD
            .GetText .ColLetterToNumber("D"), .Row, vMST
            .GetText .ColLetterToNumber("E"), .Row, vSoHD
            .GetText .ColLetterToNumber("F"), .Row, vTennguoimua
            .GetText .ColLetterToNumber("G"), .Row, vDoanhSo
            .GetText .ColLetterToNumber("H"), .Row, vGTGT
            If Trim(vKyHieuHD) <> "" Or Trim(vSoHD) <> "" Or Trim(vMST) <> "" Or Trim(vTennguoimua) <> "" Or Trim(vDoanhSo) <> "0" Or Trim(vGTGT) <> "0" Then iWhiteRow = 0
          ' Check MST
                    If Trim(vMST) <> vbNullString Then
                       vMST = Replace(vMST, " ", "")
                       vMST = Replace(vMST, "-", "")
                        'vMST = Left(vMST, 10)
    
                       strCheck = ""
                       If Len(vMST) = 10 Or Len(vMST) = 13 Then
                            strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                       End If
                       
                       If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                            .Col = .ColLetterToNumber(colD)
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgMST   'static
                            .BackColor = mErrorColor
                            iFlagvMst_PL6 = 0
                            
                            'capture the first error cell
                            If iFocusFlag = 0 Then
                                strFocusSheetName = "'" & .SheetName & "'"
                                strFocusCol = str(.Col)
                                strFocusRow = str(.Row)
                                iFocusFlag = 1
                            End If
                        Else
                            .Col = .ColLetterToNumber(colD)
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    Else
                        .Col = .ColLetterToNumber(colD)
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                    
                 ' kiem tra bat buoc nhap truong ten va dia chi
                  If Trim(vKyHieuHD) = "" And iWhiteRow = 0 Then
                        .Col = .ColLetterToNumber("C")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrorTenCSPT   'static
                        .BackColor = mErrorColor
                        flagTenCSPT = 1
                        'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
                  Else
                         .Col = .ColLetterToNumber("C")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = ""   'static
                        .BackColor = mNonErrorColor
                  End If
                        
                  If Trim(vSoHD) = "" And iWhiteRow = 0 Then
                        .Col = .ColLetterToNumber(colE)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrorDiaChi   'static
                        .BackColor = mErrorColor
                        flagDiaChi = 1
                        'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
                  Else
                         .Col = .ColLetterToNumber(colE)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = ""   'static
                        .BackColor = mNonErrorColor
                  End If
                  
                  'Kiem tra co quan thue quan ly
                  If Trim(vTennguoimua) = "" And iWhiteRow = 0 Then
                        .Col = .ColLetterToNumber(colF)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrorDiaChi   'static
                        .BackColor = mErrorColor
                        flagDiaChi = 1
                        'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
                  Else
                         .Col = .ColLetterToNumber(colF)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = ""   'static
                        .BackColor = mNonErrorColor
                  End If
                  
              .Col = .ColLetterToNumber("B")
             .Row = i + 8
             i = i + 1
        Loop Until .Text = "aa"
                       
                        
    '    Next
   End If
    .Sheet = .SheetCount
    If blCheck_S8 Then
         .SetText .ColLetterToNumber("B"), 48, IIf(iFlagvMst_PL6 = "1", "1", "0")
         .SetText .ColLetterToNumber("B"), 62, IIf(flagTenCSPT = "1", "0", "1")
         .SetText .ColLetterToNumber("B"), 63, IIf(flagDiaChi = "1", "0", "1")
    Else
         .SetText .ColLetterToNumber("B"), 48, "1"
         .SetText .ColLetterToNumber("B"), 62, "1"
         .SetText .ColLetterToNumber("B"), 63, "1"
    End If
  
    
'    ' phu luc 01-7/GTGT
'    .Sheet = sheetPl01_7_GTGT
''         blCheck_S4 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
''         If blCheck_S4 = True Then
''                .Col = .ColLetterToNumber("B")
''                .Row = 8
''                i = 1
''                Do
''                      iWhiteRow = 1
''                      .GetText .ColLetterToNumber("C"), .Row, vLoaiXe
''                      .GetText .ColLetterToNumber("D"), .Row, vDVtinh
''                      .GetText .ColLetterToNumber("F"), .Row, vSoLuongxe
''                      .GetText .ColLetterToNumber("G"), .Row, vGiaxe
''                      If Trim(vLoaiXe) <> "" Or Trim(vDVtinh) <> "" Or Trim(vSoLuongxe) <> "0" Or Trim(vGiaxe) <> "0" Then
''                        iWhiteRow = 0
''                      End If
''                      If (Val(Trim(vTongsoxe)) < Val(Trim(vSoLuongxe))) And (vTongsoxe <> vbNullString) And (vSoLuongxe <> vbNullString) Then
''                            ' Tong so luong xe ban ra
''                            .Col = .ColLetterToNumber("E")
''                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                            .CellNote = vError_pl_3   'static
''                            .BackColor = mErrorColor
''                            ' So luong xe ban cho nguoi tieu dung
''                            .Col = .ColLetterToNumber("F")
''                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                            .CellNote = vError_pl_3   'static
''                            .BackColor = mErrorColor
''
''                            iflag_pl_3 = 1
''                            'capture the first error cell
''                            If iFocusFlag = 0 Then
''                                strFocusSheetName = "'" & .SheetName & "'"
''                                strFocusCol = str(.Col)
''                                strFocusRow = str(.Row)
''                                iFocusFlag = 1
''                            End If
''                      Else
''                            .Col = .ColLetterToNumber("E")
''                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                            .CellNote = ""   'static
''                            .BackColor = mNonErrorColor
''                            ' so luong xe ban tieu dung
''                            .Col = .ColLetterToNumber("F")
''                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                            .CellNote = ""   'static
''                            .BackColor = mNonErrorColor
''                      End If
''                     .Col = .ColLetterToNumber("B")
''                     .Row = i + 8
''                     i = i + 1
''                     If .Text = "aa" Then
''                        i = i + 3
''                     End If
''                Loop Until .Text = "bb"
''                .Sheet = .SheetCount
''                .SetText .ColLetterToNumber("B"), 58, IIf(iflag_pl_3 = "1", "0", "1")
''        End If
'
    
     ' phu luc 01-01TD/GTGT
   .Sheet = sheetPl01_1TD_GTGT
   Dim flagMSTTD As Integer
   
   
    iFlagvMst_PL6 = 0
    flagMSTTD = 0
    blCheck_S8 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    For i = 8 To .MaxRows - 13
    If blCheck_S8 Then
        .Col = .ColLetterToNumber("B")
        .Row = 7
        i = 1
        Do
            iWhiteRow = 1
            '.Row = i
            .Col = .ColLetterToNumber("C")
            .GetText .ColLetterToNumber("C"), .Row, vKyHieuHD
            .GetText .ColLetterToNumber("E"), .Row, vSoHD
            .GetText .ColLetterToNumber("F"), .Row, vDoanhSo
            .GetText .ColLetterToNumber("H"), .Row, vGTGT
            If Trim(vKyHieuHD) <> "" Or Trim(vSoHD) <> "" Or Trim(vDoanhSo) <> "0" Or Trim(vGTGT) <> "0" Then iWhiteRow = 0
          ' Check MST


                 ' kiem tra bat buoc nhap truong ten va dia chi
                  If Trim(vKyHieuHD) = "" And iWhiteRow = 0 Then
                        .Col = .ColLetterToNumber("C")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrorTenCSPT   'static
                        .BackColor = mErrorColor
                        flagTenCSPT = 1
                        'capture the first error cell
                       ' If iFocusFlag = 0 Then
'                            strFocusSheetName = "'" & .SheetName & "'"
'                            strFocusCol = str(.Col)
'                            strFocusRow = str(.Row)
'                            iFocusFlag = 1
'                        End If
                  Else
                         .Col = .ColLetterToNumber("C")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = ""   'static
                        .BackColor = mNonErrorColor
                  End If
                  
                  vMST = vSoHD
                  If Trim(vMST) <> vbNullString Then
                       vMST = Replace(vMST, " ", "")
                       vMST = Replace(vMST, "-", "")
                        'vMST = Left(vMST, 10)
    
                       strCheck = ""
                       If Len(vMST) = 10 Or Len(vMST) = 13 Then
                            strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                       End If
                       
                       If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                            .Col = .ColLetterToNumber("E")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgMST   'static
                            .BackColor = mErrorColor
                            iFlagvMst_PL6 = 1
                            
                          
                        Else
                            .Col = .ColLetterToNumber("E")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    Else
                        .Col = .ColLetterToNumber("E")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If

              .Col = .ColLetterToNumber("B")
             .Row = i + 7
             i = i + 1
            
        Loop Until .Text = "aa"


     End If
     .Sheet = .SheetCount
     If blCheck_S8 Then
       .SetText .ColLetterToNumber("B"), 62, IIf(flagTenCSPT = "1", "0", "1")
       .SetText .ColLetterToNumber("B"), 82, IIf(flagMSTTD = "1", "0", "1")
       
     Else
       .SetText .ColLetterToNumber("B"), 62, "1"
       .SetText .ColLetterToNumber("B"), 82, "1"
     End If
     ' phu luc 01-01TD/GTGT
'
   .Sheet = sheetPl01_2TD_GTGT
   Dim vTenNhaMay As Variant
   Dim vChon As Variant
   Dim vMaSoThue As Variant
   Dim vDimCQT As Variant
   Dim vTyLePB As Variant
   Dim vSoThueGTGTPN As Variant
   Dim flagCQT As Integer
   
    iFlagvMst_PL6 = 0
    blCheck_S8 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
''    For i = 8 To .MaxRows - 13
    If blCheck_S8 Then
    .Col = .ColLetterToNumber("B")
        .Row = 42
        i = 1
        Do

            iWhiteRow = 1
            '.Row = i
            .GetText .ColLetterToNumber("C"), .Row, vTenNhaMay
            .GetText .ColLetterToNumber("G"), .Row, vChon
            .GetText .ColLetterToNumber("H"), .Row, vMaSoThue
            .GetText .ColLetterToNumber("P"), .Row, vDimCQT
            .GetText .ColLetterToNumber("R"), .Row, vTyLePB
            .GetText .ColLetterToNumber("U"), .Row, vSoThueGTGTPN
            
            If Trim(vTenNhaMay) <> "" Or Val(Trim(vChon)) <> "0" Or Trim(vMaSoThue) <> "" Or Trim(vDimCQT) <> "" Or Trim(vTyLePB) <> "0" Or Trim(vSoThueGTGTPN) <> "0" Then iWhiteRow = 0
          ' Check MST


                 ' kiem tra bat buoc nhap truong ten nha may thuy dien
                  If Trim(vTenNhaMay) = "" And Trim(vChon) = "1" And iWhiteRow = 0 Then
                        .Col = .ColLetterToNumber("C")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrorTenCSPT   'static
                        .BackColor = mErrorColor
                        flagTenCSPT = 1
                        'capture the first error cell
                  Else
                         .Col = .ColLetterToNumber("C")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = ""   'static
                        .BackColor = mNonErrorColor
                  End If

                  ' Kiem tra bat buoc nhap CQT
                  If Trim(vDimCQT) = "" And Trim(vChon) = "0" And iWhiteRow = 0 Then
                        .Col = .ColLetterToNumber("P")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrorTenCSPT   'static
                        .BackColor = mErrorColor
                        flagCQT = 1
'                        'capture the first error cell
'                        If iFocusFlag = 0 Then
'                            strFocusSheetName = "'" & .SheetName & "'"
'                            strFocusCol = str(.Col)
'                            strFocusRow = str(.Row)
'                            iFocusFlag = 1
'                        End If
                  Else
                         .Col = .ColLetterToNumber("P")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = ""   'static
                        .BackColor = mNonErrorColor
                  End If
                  
                  
                  If Trim(vMaSoThue) <> vbNullString Then
                       vMaSoThue = Replace(vMaSoThue, " ", "")
                       vMaSoThue = Replace(vMaSoThue, "-", "")
                        'vMST = Left(vMST, 10)
    
                       strCheck = ""
                       If Len(vMaSoThue) = 10 Or Len(vMaSoThue) = 13 Then
                            strCheck = CheckTaxCode(Mid(vMaSoThue, 1, 1), Mid(vMaSoThue, 2, 1), Mid(vMaSoThue, 3, 1), Mid(vMaSoThue, 4, 1), Mid(vMaSoThue, 5, 1), Mid(vMaSoThue, 6, 1), Mid(vMaSoThue, 7, 1), Mid(vMaSoThue, 8, 1), Mid(vMaSoThue, 9, 1), Mid(vMaSoThue, 10, 1), Mid(vMaSoThue, 11, 1), Mid(vMaSoThue, 12, 1), Mid(vMaSoThue, 13, 1))
                       End If
                       
                       If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMaSoThue) Then
                            .Col = .ColLetterToNumber("H")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgMST   'static
                            .BackColor = mErrorColor
                            iFlagvMst_PL6 = 1
                            
                          
                        Else
                            .Col = .ColLetterToNumber("H")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    Else
                        .Col = .ColLetterToNumber("H")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If

              .Col = .ColLetterToNumber("B")
             .Row = .Row + 1
             'i = i + 1
        Loop Until .Text = "aa"


     End If
     .Sheet = .SheetCount
     If blCheck_S8 Then
       .SetText .ColLetterToNumber("B"), 83, IIf(flagTenCSPT = "1", "0", "1")
       .SetText .ColLetterToNumber("B"), 84, IIf(flagCQT = "1", "0", "1")
       .SetText .ColLetterToNumber("B"), 85, IIf(iFlagvMst_PL6 = "1", "0", "1")
     Else
       .SetText .ColLetterToNumber("B"), 83, "1"
       .SetText .ColLetterToNumber("B"), 84, "1"
       .SetText .ColLetterToNumber("B"), 85, "1"
     End If
'
'     If iFocusFlag = 1 Then
'        ' setcell for focus
'
'        strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
'        .Col = .ColLetterToNumber("C")
'        .Row = 62
'        .Formula = strFormula
'
'    End If
     
    .Sheet = iCurrentSheet
    '.ReDraw = True
    '.Visible = True
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
        arrRow = Array(10, 10, 16, 19, 19, 21, 21, 28, 28, 30, 31, 33, 20, 22, 23, 24, 23, 24)
        arrCol = Array("J", "L", "L", "J", "L", "J", "L", "J", "L", "L", "L", "L", "J", "J", "J", "J", "L", "L")
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
    Dim iCol As Long, iRow As Long, i As Long
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
                  
        Set xmlCellNode = Nothing
        
        CheckDynamicError 'Set Exception Error on cells of interface
            
        UpdateSheet ' set gia tri ma hinh xuong XML
            
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
        
    End With
End Sub

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim strId As Variant
Dim iCol As Long, iRow As Long
Dim intIndexCombo As Integer, i As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim itemp As Variant, temp As Variant
Dim cValue As Variant

Dim intIndexCuc As Integer, intIndexChiCuc As Integer
Dim varIdCucThue As Variant
Dim xmlDomData As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode
Dim arrDanhsach() As String
Dim tenCucThue As String
Dim maCucThue As String
Dim tenChiCucThue As String
Dim maChiCucThue As String

Dim varTemp As Variant

Dim loaiCq As String
Dim maLoaiCq As String
With fps
    mCurrentSheet = .ActiveSheet
   If .ActiveSheet = .SheetCount - 1 Then
'   ' lay row cua dong commbo
'           Set xmlCellNode = TAX_Utilities_v1.Data(TAX_Utilities_v1.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell")(TAX_Utilities_v1.Data(TAX_Utilities_v1.NodeValidity.childNodes.length - 1).getElementsByTagName("Cell").length - 15)
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
'
'
'
''        .Col = Col
''        .Row = Row
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
''
''
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
''            .Col = .ColLetterToNumber("BB")
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
''      .Col = .ColLetterToNumber("BE")
''      .Row = Row
''      If .Text = "23" Then
''         .Col = .ColLetterToNumber("BG")
''         .Row = .Row
''        .TypeNumberMin = Val("-99999999999999")
''        .TypeNumberMax = Val("99999999999999")
''      Else
''        .Col = .ColLetterToNumber("BG")
''        .Row = .Row
''        .TypeNumberMin = Val("0")
''        .TypeNumberMax = Val("99999999999999")
''      End If
    ElseIf .ActiveSheet = sheetPl01_1_GTGT Then
        .Col = Col
        .Row = Row
        varTemp = .Text
         If Col = .ColLetterToNumber("C") And .Text <> "" Then
            .EventEnabled(EventChange) = False
            .Col = .ColLetterToNumber(colD)
            .Text = varTemp
            .EventEnabled(EventChange) = True
         End If
    ElseIf .ActiveSheet = sheetPl01_2_GTGT Then
        .Col = Col
        .Row = Row
        varTemp = .Text
         If Col = .ColLetterToNumber("C") And .Text <> "" Then
            .EventEnabled(EventChange) = False
            .Col = .ColLetterToNumber(colD)
            .Text = varTemp
            .EventEnabled(EventChange) = True
         End If
    ElseIf .ActiveSheet = 1 Then
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("D") And Row = 20 Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("P")
            .typeComboboxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
        End If
    End If
    
End With
    
    'CellChange Col, Row
    LockCellBySheet
    TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
End Sub

Private Sub setLast43Target()
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    'Dim strOriginDataFileName As String
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim str43Target As String
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
        ' TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(1), "DataFile") & ".xml"
        strCurrentDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(iDom), "DataFile") & "_" & TAX_Utilities_v1.Month & TAX_Utilities_v1.Year & ".xml"
        '***************************************
        'ThanhDX added
        'Date: 14/04/06
        
        'strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(GetValidityNode("01", Format(Month(MY), "0#"), , Format(Year(MY), "####")).childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        '***************************************
        Else
            MY = dNgayDauQuy
            
            MY = DateAdd("D", -1, MY)
            
            ' TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(1), "DataFile") & ".xml"
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
            Set xmlNode = .nodeFromID("L_47")  'L_47: col & row of 43 target
            str43Target = GetAttribute(xmlNode, "Value")
            Set xmlNode = Nothing
            If Trim(str43Target) <> "" And Trim(str43Target) <> "0" Then
                fps.Sheet = 1
                fps.Col = fps.ColLetterToNumber(S1Txt11Col)
                fps.Row = S1Txt11Row
                If Trim(fps.Text) = "0" Or Trim(fps.Text) = "" Then
                    fps.Text = str43Target
                    mCurrentSheet = 1
                    
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
    SaveErrorLog "cls001", "setLast43Target", Err.number, Err.Description
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

'******************************
'Description: Set active properties in validity node
'Author: ThanhDX
'Date: 24/11/2005

'******************************
Public Sub SetActiveSheet() 'ByRef xmlNodeValid As MSXML.IXMLDOMNode)
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    Dim strValue3 As String
    Dim strValue4 As Variant
    Dim strValue5 As Variant
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "01_1"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 2)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 3)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
                'Set active voi group 1
                fps.Sheet = 2
                fps.GetText fps.ColLetterToNumber("P"), 1, strValue4
                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Or Val(strValue4) <> 0 Then
                            blnNullValue = False
                    End If

                    If blnNullValue Then
                        SetAttribute xmlNodeValid, "Active", "0"
                    End If

             Case "01_2"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(2).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(2).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_v1.Data(2).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(2).getElementsByTagName("Cell").length - 2)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 'Set active voi dong 5
                fps.Sheet = 3
                fps.GetText fps.ColLetterToNumber("P"), 1, strValue5

                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Or Val(strValue5) <> 0 Then
                            blnNullValue = False
                    End If

                    If blnNullValue Then
                        SetAttribute xmlNodeValid, "Active", "0"
                    End If
                    
'            Case "01_3"
'               blnNullValue = True
'                'Check the third sheet
'               For Each xmlCellNode In TAX_Utilities_v1.Data(3).getElementsByTagName("Cell")
'                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
'                    'Check value in E column and F column
''                    If lCol = fps.ColLetterToNumber("E") Or _
''                       lCol = fps.ColLetterToNumber("F") Then
'                        strValue = GetAttribute(xmlCellNode, "Value")
'                        If Not IsNullValue(strValue) Then
'                            blnNullValue = False
'                            Exit For
'                        End If
''                    End If
'                Next
'
'                If blnNullValue Then
'                    SetAttribute xmlNodeValid, "Active", "0"
'                End If
                
'            Case "01_4A"
'                'Check the fourth sheet
'
'                blnNullValue = True
'                For Each xmlCellNode In TAX_Utilities_v1.Data(4).getElementsByTagName("Cell")
'                    If GetAttribute(xmlCellNode, "CellID") = THIRD_SHEET_CELL1 Then _
'                        Exit For
'                Next
'
'                strValue = GetAttribute(xmlCellNode, "Value")
'
'                Set xmlCellNode = TAX_Utilities_v1.Data(4).getElementsByTagName("Cell").Item(0)
'                strValue2 = GetAttribute(xmlCellNode, "Value")
'
'                If Not (IsNullValue(strValue) And IsNullValue(strValue2)) Then
'                    blnNullValue = False
'                End If
'
'                If blnNullValue Then
'                    SetAttribute xmlNodeValid, "Active", "0"
'                End If
'
'             Case "01_4B"
'                'Check the fivth sheet
'                blnNullValue = True
'                For Each xmlCellNode In TAX_Utilities_v1.Data(5).getElementsByTagName("Cell")
'                    If GetAttribute(xmlCellNode, "CellID") = FOURTH_SHEET_CELL1 Then _
'                        Exit For
'                Next
'
'                strValue = GetAttribute(xmlCellNode, "Value")
'                If Not IsNullValue(strValue) Then
'                    blnNullValue = False
'                End If
'
'                If blnNullValue Then
'                    SetAttribute xmlNodeValid, "Active", "0"
'                End If
             
             Case "01_5"
              blnNullValue = True
                'Check the last sheet
               For Each xmlCellNode In TAX_Utilities_v1.Data(sheetPl01_5_GTGT - 1).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value G F column
                    If lCol = fps.ColLetterToNumber("G") Then
                        strValue = GetAttribute(xmlCellNode, "Value")
                        If Not IsNullValue(strValue) Then
                            blnNullValue = False
                            Exit For
                        End If
                    End If
                Next
            
                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If
             Case "01_6"
                'Check the last sheet
               For Each xmlCellNode In TAX_Utilities_v1.Data(7).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value I column
                   ' If lCol = fps.ColLetterToNumber("I") Then
                        strValue = GetAttribute(xmlCellNode, "Value")
                        If Not IsNullValue(strValue) Then
                            blnNullValue = False
                            Exit For
                        End If
                    'End If
                Next
            
                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If
'             Case "01_6B"
'                'Check the last sheet
'               For Each xmlCellNode In TAX_Utilities_v1.Data(8).getElementsByTagName("Cell")
'                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
'                    'Check value I column
'                    If lCol = fps.ColLetterToNumber("C") Then
'                        strValue = GetAttribute(xmlCellNode, "Value")
'                        If Not IsNullValue(strValue) Then
'                            blnNullValue = False
'                            Exit For
'                        End If
'                    End If
'                Next
'
'                If blnNullValue Then
'                    SetAttribute xmlNodeValid, "Active", "0"
'                End If
'             Case "01_7"
'                'Check the last sheet
'               For Each xmlCellNode In TAX_Utilities_v1.Data(8).getElementsByTagName("Cell")
'                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
'                    'Check value I column
'                    If lCol = fps.ColLetterToNumber("C") Then
'                        strValue = GetAttribute(xmlCellNode, "Value")
'                        If Not IsNullValue(strValue) Then
'                            blnNullValue = False
'                            Exit For
'                        End If
'                    End If
'                Next
'
'                If blnNullValue Then
'                    SetAttribute xmlNodeValid, "Active", "0"
'                End If
'             Case "01_1_TD"
'                'Check the last sheet
'               For Each xmlCellNode In TAX_Utilities_v1.Data(10).getElementsByTagName("Cell")
'                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
'                    'Check value I column
'                    If lCol = fps.ColLetterToNumber("C") Then
'                        strValue = GetAttribute(xmlCellNode, "Value")
'                        If Not IsNullValue(strValue) Then
'                            blnNullValue = False
'                            Exit For
'                        End If
'                    End If
'                Next
'
'                If blnNullValue Then
'                    SetAttribute xmlNodeValid, "Active", "0"
'                End If
'           Case "01_2_TD"
'                'Check the last sheet
'               For Each xmlCellNode In TAX_Utilities_v1.Data(11).getElementsByTagName("Cell")
'                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
'                    'Check value I column
'                    If lCol = fps.ColLetterToNumber("C") Then
'                        strValue = GetAttribute(xmlCellNode, "Value")
'                        If Not IsNullValue(strValue) Then
'                            blnNullValue = False
'                            Exit For
'                        End If
'                    End If
'                Next
'
'                If blnNullValue Then
'                    SetAttribute xmlNodeValid, "Active", "0"
'                End If
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

'******************************
'Description: Reset all error satus in sheet contain dynamic error.
'Author: ThanhDX
'Date: 31/12/2005

'******************************
Public Sub ResetErrorCells()
    Dim lRow As Long, lCtrl As Long
    Dim lBColumnColor As Long, lCColumnColor As Long
    Dim lDColumnColor As Long, lEColumnColor As Long
    Dim lFColumnColor As Long
    Dim intCurrentSheet As Integer
    
    With fps
        intCurrentSheet = .Sheet
        
        .Sheet = 2 'Sheet contain dynamic data
        For lCtrl = 1 To .MaxRows
            .Col = .ColLetterToNumber("C")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If

            .Col = .ColLetterToNumber("D")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If

            .Col = .ColLetterToNumber("E")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If

            .Col = .ColLetterToNumber("F")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If
            .Col = .ColLetterToNumber("G")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If
            
            .Col = .ColLetterToNumber("H")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If
            
            .Col = .ColLetterToNumber("I")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If

        Next lCtrl
        
        .Sheet = 3 'Sheet contain dynamic data

        For lCtrl = 1 To .MaxRows
            .Col = .ColLetterToNumber("C")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If

            .Col = .ColLetterToNumber("D")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If

            .Col = .ColLetterToNumber("E")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If

            .Col = .ColLetterToNumber("F")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If
            .Col = .ColLetterToNumber("G")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If
            
            .Col = .ColLetterToNumber("H")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If
            
            .Col = .ColLetterToNumber("I")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If
        Next lCtrl
        
        .Sheet = intCurrentSheet
    End With
End Sub

Private Function GetLast43Previous(ByRef strValue As String) As Boolean
     
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim strDataFileName As String
    Dim str43Target As String
    Dim iDom As Integer
    Dim MY As Date
    
    iDom = 0
    GetLast43Previous = True
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
        strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(GetValidityNode("01", Format(Month(MY), "0#"), , Format(Year(MY), "####")).childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        '***************************************
        Else
            MY = dNgayDauQuy
            
            MY = DateAdd("D", -1, MY)
            
            '***************************************
            'ThanhDX added
            'Date: 14/04/06
            
            'strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
            'strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(GetValidityNode("01", Format(Month(MY), "0#"), , Format(Year(MY), "####")).childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
            strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(iDom), "DataFile") & "_Q0" & DatePart("Q", MY) & DatePart("YYYY", MY) & ".xml"
            '***************************************
        End If
                
        If .Load(strDataFileName) = True Then
            Set xmlNode = .nodeFromID("L_47")  'L_33: col & row of 43 target
            strValue = GetAttribute(xmlNode, "Value")
            Set xmlNode = Nothing
        ElseIf .parseError.reason <> vbNullString Then
            If InStr(1, .parseError.errorCode, "2146697210") <> 0 Then
                'file data of last month does not exist
            Else
                MsgBox .parseError.reason
            End If
            GetLast43Previous = False
        End If
        
    End With
    
    Set xmlDomLastData = Nothing
    Set xmlNode = Nothing
    
    Exit Function
    
ErrorHandle:
    SaveErrorLog "cls001", "GetLast43Previous", Err.number, Err.Description
End Function

'*******************************************************
'Description: ResetData procedure reset specified cells
'Author:ThanhDX
'Date:23/01/2006

'*******************************************************
Public Sub ResetData()
    Dim intCurSheet As Integer, intCtrl As Long
    Dim lCol As Long, lRow As Long
    Dim arrStrCellUnLock() As String
    Dim strCell As String
    
    LockCellBySheet
End Sub

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()
      
'        If fps.ActiveSheet = 6 Then
'            mCurrentSheet = 6
'            fps.Sheet = mCurrentSheet
'                With fps
'                    .Col = fps.ColLetterToNumber("F")
'                    .Row = 2
'                    .Text = TAX_Utilities_v1.Year - 1
'                    .TypeHAlign = TypeHAlignLeft
'                    UpdateCell fps, .Col, .Row, .Text
'                End With
'        End If
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

Private Function GetValueKHBS(StrKytinhthue As String, StrNgayKHBS As String) As Double
    Dim lngIndex As Long
    Dim fso As New FileSystemObject
    Dim fle As File
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlDocument As New MSXML.DOMDocument
    Dim cValue As Long
    Dim StrFileName As String
    For Each fle In fso.GetFolder(GetAbsolutePath(TAX_Utilities_v1.DataFolder)).Files
       If Right$(fle.Name, 4) = ".xml" Then
            If Mid$(fle.Name, 1, Len(fle.Name) - 4) = "KHBS_01_GTGT_" & StrKytinhthue & "_" & StrNgayKHBS Then
                StrFileName = TAX_Utilities_v1.DataFolder & fle.Name
                 xmlDocument.Load StrFileName
                  Set xmlNode = xmlDocument.childNodes(2).childNodes(2).lastChild.lastChild
                        cValue = CDbl(GetAttribute(xmlNode, "Value"))
                        If cValue < 0 Then
                            GetValueKHBS = GetValueKHBS + cValue
                            Exit For
                        End If
                  
            End If
       End If
    Next
      Set xmlNode = Nothing
      Set xmlDocument = Nothing
End Function

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function


Private Function CheckValidNgayHD(vdate As String) As Boolean
  
   Dim dd As String, mm As String, yyyy As String, dDate As Date, dDateNow As Date
   Dim sDateNow As String
   On Error GoTo e
    dd = Right("00" & Left(vdate, InStr(vdate, "/") - 1), 2)
    mm = Right("00" & Mid(vdate, 4, 2), 2)
    yyyy = Right("0000" & vdate, 4)
    sDateNow = "01/" & TAX_Utilities_v1.Month & "/" & TAX_Utilities_v1.Year
    dDateNow = Format(CDate(sDateNow), "dd/mm/yyyy")
    dDate = Format(CDate(vdate), "dd/mm/yyyy")
        If CInt(yyyy) > CInt(TAX_Utilities_v1.Year) Then
            CheckValidNgayHD = False
            Exit Function
        End If
        If DateAdd("m", 5, dDate) < dDateNow Then
        CheckValidNgayHD = False
            Exit Function
        End If
   CheckValidNgayHD = True
   Exit Function
   
e:
    CheckValidNgayHD = False
 End Function


Public Sub UpdateChangeKHBS()
    Dim varTemp As Variant
    Dim var28 As Variant
    Dim var17 As Variant
    Dim var7 As Variant
    Dim var30 As Variant
    Dim var31 As Variant
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
        
        If loaiKHBS = "frmKHBS_BS" Then
            var28 = GetAttribute(TAX_Utilities_v1.DataKHBS.nodeFromID("L_28"), "Value")
            var17 = GetAttribute(TAX_Utilities_v1.DataKHBS.nodeFromID("L_17"), "Value")
            var7 = GetAttribute(TAX_Utilities_v1.DataKHBS.nodeFromID("L_7"), "Value")

            If var28 - var17 - var7 > 0 Then
                SetAttribute TAX_Utilities_v1.DataKHBS.nodeFromID("L_30"), "Value", CStr(var28 - var17 - var7)
                SetAttribute TAX_Utilities_v1.DataKHBS.nodeFromID("L_31"), "Value", 0
                SetAttribute TAX_Utilities_v1.DataKHBS.nodeFromID("L_33"), "Value", 0
                
            Else
                SetAttribute TAX_Utilities_v1.DataKHBS.nodeFromID("L_30"), "Value", 0
                SetAttribute TAX_Utilities_v1.DataKHBS.nodeFromID("L_31"), "Value", CStr(var28 - var17 - var7)
                SetAttribute TAX_Utilities_v1.DataKHBS.nodeFromID("L_33"), "Value", 0
            End If
        End If
        .Sheet = mCurrentSheet
    End With
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

Public Sub FinishImport()
Dim i As Integer, iCol As Long, iRow As Long
Dim varTemp As Variant
With fps
fps.EventEnabled(EventAllEvents) = False
    If .ActiveSheet = 2 Then
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 8
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("K")
'                         .Text = "0"
'                         UpdateCell fps, .Col, .Row, .Text
'                          ' chuyen ky hieu thanh chu hoa
'                          .Col = .ColLetterToNumber("E")
'                          .GetText .Col, .Row, varTemp
'                          .Text = UCase$(CStr(varTemp))
'                          UpdateCell fps, .Col, .Row, .Text
'                          ' end
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 8
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)

                         .Col = .ColLetterToNumber("M")
                         .Text = "0"
                         UpdateCell fps, .Col, .Row, .Text
                         
'                         ' chuyen ky hieu thanh chu hoa
'                         .Col = .ColLetterToNumber("E")
'                          .GetText .Col, .Row, varTemp
'                          .Text = UCase$(CStr(varTemp))
'                          UpdateCell fps, .Col, .Row, .Text
'                         ' end
                         
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("J")
'                         .TypeHAlign = TypeHAlignCenter
'                         .Text = "5%"
'                         UpdateCell fps, .Col, .Row, .Text

'                        ' chuyen ky hieu thanh chu hoa
'                         .Col = .ColLetterToNumber("E")
'                          .GetText .Col, .Row, varTemp
'                          .Text = UCase$(CStr(varTemp))
'                          UpdateCell fps, .Col, .Row, .Text
'                         ' end

                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "cc"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("J")
'                         .TypeHAlign = TypeHAlignCenter
'                         .Text = "10%"
'                         UpdateCell fps, .Col, .Row, .Text
'                          ' chuyen ky hieu thanh chu hoa
'                         .Col = .ColLetterToNumber("E")
'                          .GetText .Col, .Row, varTemp
'                          .Text = UCase$(CStr(varTemp))
'                          UpdateCell fps, .Col, .Row, .Text
'                         ' end

                         .Col = .ColLetterToNumber("B")

                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "dd"
                    i = 1
                    .Row = .Row + 4
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("J")
'                         .TypeHAlign = TypeHAlignCenter
'                         .Text = "10%"
'                         UpdateCell fps, .Col, .Row, .Text

'                         ' chuyen ky hieu thanh chu hoa
'                         .Col = .ColLetterToNumber("E")
'                          .GetText .Col, .Row, varTemp
'                          .Text = UCase$(CStr(varTemp))
'                          UpdateCell fps, .Col, .Row, .Text
'                         ' end

                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "ee"
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol

                    .SetActiveCell iCol, iRow
                    CellChange2
                    CellChange3
            End If
            If .ActiveSheet = 3 Then
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 8
                    Do
                         .Text = str(i)
                         
'                         ' chuyen ky hieu thanh chu hoa
'                         .Col = .ColLetterToNumber("E")
'                          .GetText .Col, .Row, varTemp
'                          .Text = UCase$(CStr(varTemp))
'                          UpdateCell fps, .Col, .Row, .Text
'                         ' end
                         
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 8
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         
'                         ' chuyen ky hieu thanh chu hoa
'                         .Col = .ColLetterToNumber("E")
'                          .GetText .Col, .Row, varTemp
'                          .Text = UCase$(CStr(varTemp))
'                          UpdateCell fps, .Col, .Row, .Text
'                         ' end
                         
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         
'                         ' chuyen ky hieu thanh chu hoa
'                         .Col = .ColLetterToNumber("E")
'                          .GetText .Col, .Row, varTemp
'                          .Text = UCase$(CStr(varTemp))
'                          UpdateCell fps, .Col, .Row, .Text
'                         ' end
                         
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "cc"
                     i = 1
                    .Row = .Row + 4
                    Do
                         .Text = str(i)
                         
'                         ' chuyen ky hieu thanh chu hoa
'                         .Col = .ColLetterToNumber("E")
'                          .GetText .Col, .Row, varTemp
'                          .Text = UCase$(CStr(varTemp))
'                          UpdateCell fps, .Col, .Row, .Text
'                         ' end
                         
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "dd"
                     i = 1
                    .Row = .Row + 4
                    Do
                         .Text = str(i)
                         
'                         ' chuyen ky hieu thanh chu hoa
'                         .Col = .ColLetterToNumber("E")
'                          .GetText .Col, .Row, varTemp
'                          .Text = UCase$(CStr(varTemp))
'                          UpdateCell fps, .Col, .Row, .Text
'                         ' end
                         
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "ee"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
                    CellChange2
                    CellChange3
            End If
            
'            If .ActiveSheet = 4 Then
'                    i = 1
'                    .Col = .ColLetterToNumber("B")
'                    .Row = rowStartPL01_3
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("B")
'                         .Row = i + rowStartPL01_3
'                         i = i + 1
'                    Loop Until .Text = "aa"
'                    .SetActiveCell iCol, iRow
'
'                    .Row = iRow
'                    .Col = iCol
'                    .SetActiveCell iCol, iRow
'            End If

        fps.EventEnabled(EventAllEvents) = True
        End With
    'CellChange iCol, iRow
    'CellChange1 Col, Row
End Sub

Public Function getDieuChinhGiam(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue, strDieuChinhGiam, strDieuChinhTang, strDieuChinhTang4043 As String
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
                .GetText .ColLetterToNumber("L"), 24, varTemp
                strCurrentValue = "CT22_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 27, varTemp
'                strCurrentValue = strCurrentValue & "~CT24_" & Trim(varTemp)
                .GetText .ColLetterToNumber("L"), 28, varTemp
                strCurrentValue = strCurrentValue & "~CT25_" & Trim(varTemp)
                .GetText .ColLetterToNumber("L"), 33, varTemp
                strCurrentValue = strCurrentValue & "~CT31_" & Trim(varTemp)
                .GetText .ColLetterToNumber("L"), 34, varTemp
                strCurrentValue = strCurrentValue & "~CT33_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 43, varTemp
'                strCurrentValue = strCurrentValue & "~CT40b_" & Trim(varTemp)
                .GetText .ColLetterToNumber("L"), 46, varTemp
                strCurrentValue = strCurrentValue & "~CT42_" & Trim(varTemp)
                .GetText .ColLetterToNumber("L"), 44, varTemp
                strCurrentValue = strCurrentValue & "~CT40_" & Trim(varTemp)
                .GetText .ColLetterToNumber("L"), 47, varTemp
                strCurrentValue = strCurrentValue & "~CT43_" & Trim(varTemp)
                ' tinh chenh lech giua 2 lan dieu chinh
                tempArrOld = Split(strOldValue, "~")
                tempArrCurr = Split(strCurrentValue, "~")
                arrMapCT = Split(strMaCT, "~")
                strDieuChinhGiam = ""
                strDieuChinhTang = ""
                strDieuChinhTang4043 = ""
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
                    If (Left(strMaCT, 3) = "40_" Or Val(Left(strMaCT, 2)) = 43) Then
                        'If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) <> 0 Then
                            ' dieu chinh giam
                            strDieuChinhTang4043 = strDieuChinhTang4043 & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                        'End If
                    End If
                    If Left(strMaCT, 3) = "22_" Or Left(strMaCT, 3) = "25_" Or Left(strMaCT, 3) = "42_" Then
                        If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 And Left(strMaCT, 3) <> "40_" And Val(Left(strMaCT, 2)) <> 43 Then
                            ' dieu chinh giam
                            strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                        End If
                        If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 And Left(strMaCT, 3) <> "40_" And Val(Left(strMaCT, 2)) <> 43 Then
                            ' dieu chinh tang
                            strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                        End If
                    Else
                        If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 And Left(strMaCT, 3) <> "40_" And Val(Left(strMaCT, 2)) <> 43 Then
                                ' dieu chinh giam
                                strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                            End If
                            If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 And Left(strMaCT, 3) <> "40_" And Val(Left(strMaCT, 2)) <> 43 Then
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
                If strDieuChinhTang4043 <> "" Then
                    strDieuChinhTang4043 = Mid(strDieuChinhTang4043, 2, Len(strDieuChinhTang4043) - 1)
                End If
                getDieuChinhGiam = strDieuChinhGiam & "###" & strDieuChinhTang & "###" & strDieuChinhTang4043
      End With
End Function
'Public Function getDieuChinhGiam(strOldValue As String) As String
'    Dim varTemp As Variant
'    Dim strCurrentValue, strDieuChinhGiam, strDieuChinhTang As String
'    Dim tempArrOld() As String
'    Dim tempArrCurr() As String
'    Dim tempArrOldValue() As String
'    Dim tempArrCurrValue() As String
'    Dim i, j As Integer
'    Dim strMaCT As String
'    Dim arrMapCT() As String ' Luu gia tri cua dong chi tieu
'    Dim arrMapCTValue() As String ' Luu gia tri cua cac cell chi tieu
'    strMaCT = getMapCT
'    With fps
'                .Sheet = 1
'                .GetText .ColLetterToNumber("L"), 24, varTemp
'                strCurrentValue = "CT22_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 27, varTemp
'                strCurrentValue = strCurrentValue & "~CT24_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 29, varTemp
'                strCurrentValue = strCurrentValue & "~CT25_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 33, varTemp
'                strCurrentValue = strCurrentValue & "~CT31_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 34, varTemp
'                strCurrentValue = strCurrentValue & "~CT33_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 38, varTemp
'                strCurrentValue = strCurrentValue & "~CT37_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 39, varTemp
'                strCurrentValue = strCurrentValue & "~CT38_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 43, varTemp
'                strCurrentValue = strCurrentValue & "~CT40b_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 46, varTemp
'                strCurrentValue = strCurrentValue & "~CT42_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 44, varTemp
'                strCurrentValue = strCurrentValue & "~CT40_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 47, varTemp
'                strCurrentValue = strCurrentValue & "~CT43_" & Trim(varTemp)
'                ' tinh chenh lech giua 2 lan dieu chinh
'                tempArrOld = Split(strOldValue, "~")
'                tempArrCurr = Split(strCurrentValue, "~")
'                arrMapCT = Split(strMaCT, "~")
'                strDieuChinhGiam = ""
'                strDieuChinhTang = ""
'                For i = 0 To UBound(tempArrOld)
'                    tempArrOldValue = Split(tempArrOld(i), "_")
'                    tempArrCurrValue = Split(tempArrCurr(i), "_")
'                    ' lay cac gia tri cua tung cot trong sheet KHBS
'                    For j = 0 To UBound(arrMapCT)
'                        arrMapCTValue = Split(arrMapCT(j), "_")
'                        ' phan tu so 0 la ma chi tieu
'                        If arrMapCTValue(0) = Mid(tempArrOldValue(0), 3) Then
'                            strMaCT = arrMapCT(j)
'                            Exit For
'                        End If
'                    Next j
'
'                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
'                        ' dieu chinh giam
'                        strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
'                    End If
'                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
'                        ' dieu chinh tang
'                        strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
'                    End If
'                Next i
'                If strDieuChinhGiam <> "" Then
'                    strDieuChinhGiam = Mid(strDieuChinhGiam, 2, Len(strDieuChinhGiam) - 1)
'                End If
'                If strDieuChinhTang <> "" Then
'                    strDieuChinhTang = Mid(strDieuChinhTang, 2, Len(strDieuChinhTang) - 1)
'                End If
'                getDieuChinhGiam = strDieuChinhGiam & "###" & strDieuChinhTang
'      End With
'End Function

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
            Set xmlNode = .nodeFromID("L_24")
            result = "CT22_" & Val(GetAttribute(xmlNode, "Value"))
'            Set xmlNode = .nodeFromID("L_27")
'            result = result & "~" & "CT24_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("L_28")
            result = result & "~" & "CT25_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("L_33")
            result = result & "~" & "CT31_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("L_34")
            result = result & "~" & "CT33_" & Val(GetAttribute(xmlNode, "Value"))
'            Set xmlNode = .nodeFromID("L_43")
'            result = result & "~" & "CT40b_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("L_46")
            result = result & "~" & "CT42_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("L_44")
            result = result & "~" & "CT40_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("L_47")
            result = result & "~" & "CT43_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
            

        End If
    End With
    getValueTK = result
    Exit Function
    
ErrorHandle:
    SaveErrorLog "cls_02GTGT", "getValueTK", Err.number, Err.Description
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
    SaveErrorLog "01_GTGT", "getMapCT", Err.number, Err.Description
End Function

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
            If lCol = .ColLetterToNumber("I") And lRow = 23 Then
            ElseIf lCol = .ColLetterToNumber("B") And lRow = 18 Then
            Else
                
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            End If
        Next

        .Sheet = 2
        mCurrentSheet = 2
        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            Next
            
            ' xu ly truong hop datafile ban 320 se set encode cell dau tien = 0 (bo ma hoa don)
            Set xmlCellNodeData4 = TAX_Utilities_v1.Data(1).getElementsByTagName("Sections").Item(0)
            temp1 = GetAttribute(xmlCellNodeData4, "Version")
            If Val(temp1) = 320 Then
                For Each xmlCellNodeData1 In TAX_Utilities_v1.Data(1).getElementsByTagName("Section")
                    If GetAttribute(xmlCellNodeData1, "Dynamic") = "1" Then
                        For Each xmlCellNodeData2 In xmlCellNodeData1.childNodes
                            Set xmlCellNodeData3 = xmlCellNodeData2.firstChild
                            If GetAttribute(xmlCellNodeData3, "FirstCell") <> vbNullString Then
                                SetAttribute xmlCellNodeData3, "Encode", "0"
                            End If
                        Next
                    End If
                Next
            End If
            
            ' end xu ly data data
         End If
        .Sheet = 3
        mCurrentSheet = 3
        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(2).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            Next
            
            
            ' xu ly truong hop datafile ban 320 se set encode cell dau tien = 0 (bo ma hoa don)
            Set xmlCellNodeData4 = TAX_Utilities_v1.Data(2).getElementsByTagName("Sections").Item(0)
            temp1 = GetAttribute(xmlCellNodeData4, "Version")
            If Val(temp1) = 320 Then
                For Each xmlCellNodeData1 In TAX_Utilities_v1.Data(2).getElementsByTagName("Section")
                    If GetAttribute(xmlCellNodeData1, "Dynamic") = "1" Then
                        For Each xmlCellNodeData2 In xmlCellNodeData1.childNodes
                            Set xmlCellNodeData3 = xmlCellNodeData2.firstChild
                            If GetAttribute(xmlCellNodeData3, "FirstCell") <> vbNullString Then
                                SetAttribute xmlCellNodeData3, "Encode", "0"
                            End If
                        Next
                    End If
                Next
            End If
            
            ' end xu ly data data
        End If
        .Sheet = 4
        mCurrentSheet = 4
        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(3).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
        .Sheet = 5
        mCurrentSheet = 5
        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(4).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
'        .Sheet = 6
'        mCurrentSheet = 6
'        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'        If blCheck_S2 = True Then
'            For Each xmlCellNodeData In TAX_Utilities_v1.Data(5).getElementsByTagName("Cell")
'                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'                .GetText lCol, lRow, varTemp
'                UpdateCell fps, lCol, lRow, varTemp
'            Next
'        End If
'        .Sheet = 7
'        mCurrentSheet = 7
'        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'        If blCheck_S2 = True Then
'            For Each xmlCellNodeData In TAX_Utilities_v1.Data(6).getElementsByTagName("Cell")
'                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'                .GetText lCol, lRow, varTemp
'                UpdateCell fps, lCol, lRow, varTemp
'            Next
'        End If
'        .Sheet = 8
'        mCurrentSheet = 8
'        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'        If blCheck_S2 = True Then
'            For Each xmlCellNodeData In TAX_Utilities_v1.Data(7).getElementsByTagName("Cell")
'                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'                .GetText lCol, lRow, varTemp
'                UpdateCell fps, lCol, lRow, varTemp
'            Next
'        End If
        
'        .Sheet = sheetPl01_7_GTGT
'        mCurrentSheet = sheetPl01_7_GTGT
'        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'        If blCheck_S2 = True Then
'            For Each xmlCellNodeData In TAX_Utilities_v1.Data(sheetPl01_7_GTGT - 1).getElementsByTagName("Cell")
'                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'                .GetText lCol, lRow, varTemp
'                UpdateCell fps, lCol, lRow, varTemp
'            Next
'        End If
        
        .Sheet = sheetPl01_1TD_GTGT
        mCurrentSheet = sheetPl01_1TD_GTGT
        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(sheetPl01_1TD_GTGT - 1).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            Next
        End If

        .Sheet = sheetPl01_2TD_GTGT
        mCurrentSheet = sheetPl01_2TD_GTGT
        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(sheetPl01_2TD_GTGT - 1).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                If (lRow >= 42 And lCol = .ColLetterToNumber("G")) Then
                    .GetText lCol, lRow, varTemp
                    If varTemp = "1" Or UCase(CStr(varTemp)) = "X" Then
                        UpdateCell fps, lCol, lRow, "x"
                    Else
                        UpdateCell fps, lCol, lRow, ""
                    End If
                Else
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
                End If
            Next
        End If
        
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


' set cho du ham
' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        getChiTieuTongDC = "0"
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

 
    getValueCTDC = "0"
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01BTNDN", "getValueCTDC", Err.number, Err.Description
End Function


' Ham set lai trang thai cua cac chi tieu khi check chi tieu 21
Sub CellLock(ByVal Col As Long, ByVal Row As Long)

    Dim i As Integer
    Dim a As Double, b As Double, c As Double
    Dim arrRow As Variant, arrCol As Variant
    Dim arrCellTag() As String
    Dim cCol As Long, cRow As Long
    Dim varTemp As Variant
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim iCol  As Long, iRow As Long
    Dim strValue1  As Variant, strValue2 As Variant
    With fps
       .EventEnabled(EventAllEvents) = False
       .Sheet = .ActiveSheet
       Select Case .Sheet
           Case 1 ' To khai GTGT
               mCurrentSheet = .Sheet
               GetCellSpan fps, Col, Row
               If Row = S1ChkRow And Col = .ColLetterToNumber(S1ChkCol) Then
                   arrRow = Array(27, 27, 28, 30, 31, 31, 32, 33, 33, 34, 34, 35, 35)
                   arrCol = Array("J", "L", "L", "J", "J", "L", "J", "J", "L", "J", "L", "J", "L")
                   .Col = .ColLetterToNumber("I")
                   .Row = 23
                   If .Text = "1" Then
                       For i = 0 To UBound(arrRow)
                           .Row = arrRow(i)
                           .Col = .ColLetterToNumber(arrCol(i))
                           .Text = "0"
                           UpdateCell fps, .Col, .Row, .Text
                           .Lock = True
                       Next
                   ElseIf .Text = "0" Then
                       For i = 0 To UBound(arrRow)
                           .Row = arrRow(i)
                           .Col = .ColLetterToNumber(arrCol(i))
                           If (arrCol(i) = "J" Or arrCol(i) = "L") And (arrRow(i) = 31 Or arrRow(i) = 35) Then
                               .Lock = True
                           Else
                               .Lock = False
                           End If
                       Next
                   End If
                   
            End If
            
       End Select
       
       .EventEnabled(EventAllEvents) = True
    End With
End Sub


' Ham set cac cong thuc
Sub CellChange1(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
    Dim varTemp As Variant
    Dim cellTagArr() As String
    Dim ct21 As Variant
    
'    Dim blCheck_S4A As Boolean
'    Dim blCheck_S4B As Boolean
'
'    blCheck_S4A = IIf(TAX_Utilities_v1.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S4B = IIf(TAX_Utilities_v1.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    
    With fps
        
        .Sheet = 1
        ' Chi tieu 21
        .Col = .ColLetterToNumber("I")
        .Row = 23
        ct21 = .Text
        .Sheet = .ActiveSheet
        Select Case .Sheet
            ' Tong hop du lieu tu phu luc 01_1 len to khai
            Case sheetPl01_1_GTGT ' phu luc 01_1
                .Col = .ColLetterToNumber("C")
                .Row = Row
                varTemp = .CellTag
                cellTagArr = Split(varTemp, "~")
                If varTemp <> "" Then
                    If (Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("M")) And cellTagArr(1) = "4" And ct21 = "0" Then
                        ' chi tieu 33
                            .Sheet = 1
                            .Col = .ColLetterToNumber("L")
                            .Row = 34
                            '.Formula = "L53"
                            .Formula = colL & rowHidden53
                            .Sheet = .ActiveSheet
                    End If
                    If Col = .ColLetterToNumber("K") And cellTagArr(1) = "1" And ct21 = "0" Then
                             ' chi tieu 26
                            .Sheet = 1
                            .Col = .ColLetterToNumber("J")
                            .Row = 30
                            '.Formula = "E54"
                            .Formula = colE & rowHidden54
                            .Sheet = .ActiveSheet
                    End If
                    If Col = .ColLetterToNumber("K") And cellTagArr(1) = "2" And ct21 = "0" Then
                             ' chi tieu 29
                            .Sheet = 1
                            .Col = .ColLetterToNumber("J")
                            .Row = 32
                            '.Formula = "H53"
                            .Formula = colH & rowHidden53
                            .Sheet = .ActiveSheet
                    End If
                    If Col = .ColLetterToNumber("K") And cellTagArr(1) = "3" And ct21 = "0" Then
                             ' chi tieu 30
                            .Sheet = 1
                            .Col = .ColLetterToNumber("J")
                            .Row = 33
                            '.Formula = "I53"
                            .Formula = colI & rowHidden53
                            .Sheet = .ActiveSheet
                    End If
                    If Col = .ColLetterToNumber("K") And cellTagArr(1) = "4" And ct21 = "0" Then
                             ' chi tieu 32
                            .Sheet = 1
                            .Col = .ColLetterToNumber("J")
                            .Row = 34
                            '.Formula = "K53"
                            .Formula = colK & rowHidden53
                            .Sheet = .ActiveSheet
                    End If
                    If (Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("M")) And cellTagArr(1) = "3" And ct21 = "0" Then
                             ' chi tieu 31
                            .Sheet = 1
                            .Col = .ColLetterToNumber("L")
                            .Row = 33
                            '.Formula = "J53"
                            .Formula = colJ & rowHidden53
                            .Sheet = .ActiveSheet
                    End If
                    
'                        ' kiem tra neu co phu luc 01_4A thi moi set cong thuc
'                    If blCheck_S4A = True Then
'                        If Col = .ColLetterToNumber("K") And (cellTagArr(1) = "1" Or cellTagArr(1) = "2" Or cellTagArr(1) = "3" Or cellTagArr(1) = "4") Then
'                                .Sheet = 5
'                                ' chi tieu tong doanh so ban ra tren phu luc 1
'                                .Col = .ColLetterToNumber("D")
'                                .Row = 16
'                                .Formula = "C24"
'                                .Sheet = .ActiveSheet
'                        End If
'                        If Col = .ColLetterToNumber("K") And (cellTagArr(1) = "2" Or cellTagArr(1) = "3" Or cellTagArr(1) = "4") Then
'                                ' chi tieu T?ng doanh thu hàng hóa, d?ch v? bán ra ch?u thu? GTGT
'                                .Sheet = 5
'                                .Col = .ColLetterToNumber("D")
'                                .Row = 17
'                                .Formula = "D24"
'                                .Sheet = .ActiveSheet
'                        End If
'                    End If
                End If
            Case sheetPl01_2_GTGT ' phu luc 01_2
                .Col = .ColLetterToNumber("C")
                .Row = Row
                varTemp = .CellTag
                cellTagArr = Split(varTemp, "~")
                If varTemp <> "" Then
'                    If Col = .ColLetterToNumber("K") And (cellTagArr(1) = "1" Or cellTagArr(1) = "2" Or cellTagArr(1) = "3") And ct21 = "0" Then
'                            ' chi tieu 23
'                            .Sheet = 1
'                            .Col = .ColLetterToNumber("J")
'                            .Row = 27
'                            '.Formula = "D53"
'                            .Formula = colD & rowHidden53
'                            .Sheet = .ActiveSheet
'                    End If
'                    If (Col = .ColLetterToNumber("M") Or Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("L")) And (cellTagArr(1) = "1" Or cellTagArr(1) = "2" Or cellTagArr(1) = "3") And ct21 = "0" Then
'                             ' chi tieu 24
'                            .Sheet = 1
'                            .Col = .ColLetterToNumber("L")
'                            .Row = 27
'                            '.Formula = "E53"
'                            .Formula = colE & rowHidden53
'                            .Sheet = .ActiveSheet
'                    End If
'
'                    If (Col = .ColLetterToNumber("M") Or Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("L")) And (cellTagArr(1) = "1") Then
'                             ' chi tieu 25
'                            .Sheet = 1
'                            .Col = .ColLetterToNumber("L")
'                            .Row = 28
'                            '.Formula = "F53"
'                            .Formula = colF & rowHidden53
'                            .Sheet = .ActiveSheet
'                    End If
                    
                    
'                    ' kiem tra neu co phu luc 01_4A thi moi set cong thuc
'                    If blCheck_S4A = True Then
'                        If (Col = .ColLetterToNumber("M") Or Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("L")) And (cellTagArr(1) = "1") Then
'                                ' chi tieu tong dong 1 cot 10
'                                'A1
'                                .Sheet = 5
'                                .Col = .ColLetterToNumber("D")
'                                .Row = 7
'                                .Formula = "C23"
'                                .Sheet = .ActiveSheet
'                        End If
'
'                        If (Col = .ColLetterToNumber("M") Or Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("L")) And (cellTagArr(1) = "2") Then
'                                 ' chi tieu A2-04A
'                                .Sheet = 5
'                                .Col = .ColLetterToNumber("D")
'                                .Row = 9
'                                .Formula = "D23"
'                                .Sheet = .ActiveSheet
'                        End If
'
'                        If (Col = .ColLetterToNumber("M") Or Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("L")) And (cellTagArr(1) = "3") Then
'                                'chi tieu tong dong 3 cot 10
'                                'A3
'                                .Sheet = 5
'                                .Col = .ColLetterToNumber("D")
'                                .Row = 11
'                                .Formula = "E23"
'                                .Sheet = .ActiveSheet
'                        End If
'
'                        If (Col = .ColLetterToNumber("M") Or Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("L")) And (cellTagArr(1) = "3") Then
'                                .Sheet = 5
'                                 ' chitieu B4
'                                .Col = .ColLetterToNumber("D")
'                                .Row = 19
'                                .Formula = "E24"
'                                .Sheet = .ActiveSheet
'                        End If
'                   End If
                End If
'            Case 5 ' phu luc 01_4A
'                        ' bo cac ct tren phu luc 01_4A
'                        reset_01_4A
'                        If Col = .ColLetterToNumber("D") And (Row = 19 Or Row = 16 Or Row = 17) Then
'                                ' chi tieu 25
'                                .Sheet = 1
'                                .Col = .ColLetterToNumber("L")
'                                .Row = 28
'                                '.Formula = "F53"
'                                .Formula = colF & rowHidden53
'                                .Sheet = .ActiveSheet
'                        End If
'            Case 6 ' phu luc 01_4B
'                        If Col = .ColLetterToNumber("E") And Row = 21 Then
'                                ' chi tieu 25
'                                .Sheet = 1
'                                .Col = .ColLetterToNumber("L")
'                                .Row = 28
'                                '.Formula = "F53"
'                                .Formula = colF & rowHidden53
'                                .Sheet = .ActiveSheet
'                        End If
            Case sheetPl01_5_GTGT ' phu luc 01_5
                    If Col = .ColLetterToNumber("I") Then
                            ' chi tieu 39
                            .Sheet = 1
                            .Col = .ColLetterToNumber("L")
                            .Row = 40
                            '.Formula = "D54"
                            .Formula = colD & rowHidden54
                            .Sheet = .ActiveSheet
                    End If
        End Select
    End With
End Sub

' Ham set lai gia tri cua cong thuc khi tai bang ke 01_1
Sub CellChange2()
    Dim varTemp As Variant
    Dim cellTagArr() As String
    With fps
        .Sheet = .ActiveSheet
        Select Case .Sheet
            Case 2 ' phu luc 01_1
        ' chi tieu 33
            .Sheet = 1
            .Col = .ColLetterToNumber("L")
            .Row = 34
            '.Formula = "L53"
            .Formula = colL & rowHidden53
             ' chi tieu 26
            .Sheet = 1
            .Col = .ColLetterToNumber("J")
            .Row = 30
            '.Formula = "E54"
            .Formula = colE & rowHidden54
             ' chi tieu 29
            .Sheet = 1
            .Col = .ColLetterToNumber("J")
            .Row = 32
            '.Formula = "H53"
            .Formula = colH & rowHidden53
             ' chi tieu 30
            .Sheet = 1
            .Col = .ColLetterToNumber("J")
            .Row = 33
            '.Formula = "I53"
            .Formula = colI & rowHidden53
             ' chi tieu 32
            .Sheet = 1
            .Col = .ColLetterToNumber("J")
            .Row = 34
            '.Formula = "K53"
            .Formula = colK & rowHidden53
             ' chi tieu 31
            .Sheet = 1
            .Col = .ColLetterToNumber("L")
            .Row = 33
            '.Formula = "J53"
            .Formula = colJ & rowHidden53
'        Case 3 ' phu luc 01_2
'            ' chi tieu 23
'            .Sheet = 1
'            .Col = .ColLetterToNumber("J")
'            .Row = 27
'            '.Formula = "D53"
'            .Formula = colD & rowHidden53
'             ' chi tieu 24
'            .Sheet = 1
'            .Col = .ColLetterToNumber("L")
'            .Row = 27
'            '.Formula = "E53"
'            .Formula = colE & rowHidden53
'             ' chi tieu 25
'            .Sheet = 1
'            .Col = .ColLetterToNumber("L")
'            .Row = 28
'            '.Formula = "F53"
'            .Formula = colF & rowHidden53
            
        End Select
    End With
End Sub
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
    Dim strValue4 As String
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
                Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell").length - 4)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell")(1)
                   strValue3 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell")(7)
                    strValue4 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell").length > 34 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Or (Not IsNullValue(strValue4)) Then
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
'When cell's data is changed or a control at a cell is handled, This sub is called
'parameter Col: the column of cell that rise event
'parameter Row: the row of cell that rise event
'parameter f: the flag to mark some special option
'**: when this sub is on executing, all events of fpspread will be disabled until the end of the sub
Public Sub LockCellBySheet()
    Dim i As Integer
    Dim a As Double, b As Double, c As Double
    Dim arrRow As Variant, arrCol As Variant
    Dim arrCellTag() As String
    Dim cCol As Long, cRow As Long
    Dim varTemp As Variant
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim iCol  As Long, iRow As Long
    Dim strValue1  As Variant, strValue2 As Variant
With fps
    'Because of unclearly-rised events, this function is here for turning it off
    'at the end of this sub, this event will be turned on
    .EventEnabled(EventAllEvents) = False

    
    ' Kiem tra neu co phu luc moi set cong thuc tong hop so lieu tu tinh
    Dim tempCurrSheet As Integer
    Dim ct21 As Variant
    Dim blCheck_S1 As Boolean
    Dim blCheck_S2 As Boolean
    Dim blCheck_S5 As Boolean
'    Dim blCheck_S4A As Boolean
'    Dim blCheck_S4B As Boolean
    blCheck_S1 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl01_1_GTGT - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl01_2_GTGT - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S5 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl01_5_GTGT - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S4A = IIf(TAX_Utilities_v1.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S4B = IIf(TAX_Utilities_v1.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    tempCurrSheet = mCurrentSheet
    .Sheet = 1
    ' Chi tieu 21
    .Col = .ColLetterToNumber("I")
    .Row = 23
    ct21 = .Text
            If blCheck_S1 = True And ct21 = "0" Then
                .Sheet = 1
               ' chi tieu 26
               .Col = .ColLetterToNumber("J")
               .Row = 30
               .Lock = True
                .Sheet = 1
                ' chi tieu 29
                .Col = .ColLetterToNumber("J")
                .Row = 32
                .Lock = True
                 ' chi tieu 30
                .Col = .ColLetterToNumber("J")
                .Row = 33
                .Lock = True
                ' chi tieu 32
                .Col = .ColLetterToNumber("J")
                .Row = 34
                .Lock = True
                ' chi tieu 31
                .Col = .ColLetterToNumber("L")
                .Row = 33
                .Lock = True
                ' chi tieu 33
                .Col = .ColLetterToNumber("L")
                .Row = 34
                 .Lock = True
            ElseIf (blCheck_S1 = False Or blCheck_S1 = True) And ct21 = "1" Then
                .Sheet = 1
               ' chi tieu 26
               .Col = .ColLetterToNumber("J")
               .Row = 30
               .Lock = True
                .Sheet = 1
                ' chi tieu 29
                .Col = .ColLetterToNumber("J")
                .Row = 32
                .Lock = True
                 ' chi tieu 30
                .Col = .ColLetterToNumber("J")
                .Row = 33
                .Lock = True
                ' chi tieu 32
                .Col = .ColLetterToNumber("J")
                .Row = 34
                .Lock = True
                ' chi tieu 31
                .Col = .ColLetterToNumber("L")
                .Row = 33
                .Lock = True
                ' chi tieu 33
                .Col = .ColLetterToNumber("L")
                .Row = 34
                 .Lock = True
            Else
                .Sheet = 1
               ' chi tieu 26
               .Col = .ColLetterToNumber("J")
               .Row = 30
               .Lock = False
                .Sheet = 1
                ' chi tieu 29
                .Col = .ColLetterToNumber("J")
                .Row = 32
                .Lock = False
                 ' chi tieu 30
                .Col = .ColLetterToNumber("J")
                .Row = 33
                .Lock = False
                ' chi tieu 32
                .Col = .ColLetterToNumber("J")
                .Row = 34
                .Lock = False
                ' chi tieu 31
                .Col = .ColLetterToNumber("L")
                .Row = 33
                .Lock = False
                ' chi tieu 33
                .Col = .ColLetterToNumber("L")
                .Row = 34
                 .Lock = False
            End If
            
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
             If ct21 = "1" Then
                .Sheet = 1
                ' Chi tieu 23
                .Col = .ColLetterToNumber("J")
                .Row = 27
                .Lock = True
                ' Chi tieu 24
                .Col = .ColLetterToNumber("L")
                .Row = 27
                .Lock = True
            Else
                .Sheet = 1
                ' Chi tieu 23
                .Col = .ColLetterToNumber("J")
                .Row = 27
                .Lock = False
                ' Chi tieu 24
                .Col = .ColLetterToNumber("L")
                .Row = 27
                .Lock = False
            End If
            
            If blCheck_S5 = True Then
                .Sheet = 1
                ' Chi tieu 39
                .Col = .ColLetterToNumber("L")
                .Row = 40
                .Lock = True
            Else
                .Sheet = 1
                ' Chi tieu 39
                .Col = .ColLetterToNumber("L")
                .Row = 40
                .Lock = False
            End If
    .EventEnabled(EventAllEvents) = True
End With
End Sub



' Ham set lai gia tri cua cong thuc khi tai bang ke 01_2
Sub CellChange3()
''    Dim blCheck_S4A As Boolean
''    Dim blCheck_S4B As Boolean
''
''    blCheck_S4A = IIf(TAX_Utilities_v1.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
''    blCheck_S4B = IIf(TAX_Utilities_v1.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'
'    ' Tong hop len bang ke 01_4A
'        With fps
'        .Sheet = .ActiveSheet
'        Select Case .Sheet
'            Case 2 ' Phu luc 01_1
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
'
''                If blCheck_S4B = True Then
''                    .Sheet = 6
''                    ' chi tieu B1
''                    .Col = .ColLetterToNumber("E")
''                    .Row = 16
''                    .Formula = "D25"
''
''                    'chi tieu B2
''                    .Col = .ColLetterToNumber("E")
''                    .Row = 17
''                    .Formula = "E26"
''
''                    'chi tieu B4
''                    .Col = .ColLetterToNumber("E")
''                    .Row = 19
''                    .Formula = "F25"
''                End If
'            Case 3 ' phu luc 01_2
''                If blCheck_S4A = True Then
''                    ' chi tieu tong dong 1 cot 10
''                    'A1
''                    .Sheet = 5
''                    .Col = .ColLetterToNumber("D")
''                    .Row = 7
''                    .Formula = "C23"
''                     ' chi tieu A2-04A
''                    .Sheet = 5
''                    .Col = .ColLetterToNumber("D")
''                    .Row = 9
''                    .Formula = "D23"
''                    'chi tieu tong dong 3 cot 10
''                    'A3
''                    .Sheet = 5
''                    .Col = .ColLetterToNumber("D")
''                    .Row = 11
''                    .Formula = "E23"
''                    .Sheet = 5
''                     ' chitieu B4
''                    .Col = .ColLetterToNumber("D")
''                    .Row = 19
''                    .Formula = "E24"
''                End If
'
''                If blCheck_S4B = True Then
''                    .Sheet = 6
''                    ' chi tieu A1
''                    .Col = .ColLetterToNumber("E")
''                    .Row = 7
''                    .Formula = "D24"
''                    ' chi tieu A2
''                    .Col = .ColLetterToNumber("E")
''                    .Row = 9
''                    .Formula = "E24"
''                    ' chi tieu A3
''                    .Col = .ColLetterToNumber("E")
''                    .Row = 11
''                    .Formula = "F24"
''
''                    '
''                End If
''             Case 4
''                If blCheck_S4B = True Then
''                    .Sheet = 6
''                    .Col = .ColLetterToNumber("E")
''                    .Row = 21
''                    .Formula = "D26"
''                End If
'         End Select
'    End With
End Sub


Sub SetValueOCheck()
    With fps
        ' Phu luc 01
        ' chi tieu 33
        .Sheet = 1
        .Col = .ColLetterToNumber("L")
        .Row = 34
        '.Formula = "L53"
        .Formula = colL & rowHidden53
         ' chi tieu 26
        .Sheet = 1
        .Col = .ColLetterToNumber("J")
        .Row = 30
        '.Formula = "E54"
        .Formula = colE & rowHidden54
         ' chi tieu 29
        .Sheet = 1
        .Col = .ColLetterToNumber("J")
        .Row = 32
        '.Formula = "H53"
        .Formula = colH & rowHidden53
         ' chi tieu 30
        .Sheet = 1
        .Col = .ColLetterToNumber("J")
        .Row = 33
        '.Formula = "I53"
        .Formula = colI & rowHidden53
         ' chi tieu 32
        .Sheet = 1
        .Col = .ColLetterToNumber("J")
        .Row = 34
        '.Formula = "K53"
        .Formula = colK & rowHidden53
         ' chi tieu 31
        .Sheet = 1
        .Col = .ColLetterToNumber("L")
        .Row = 33
        '.Formula = "J53"
        .Formula = colJ & rowHidden53
        ' phu luc 01_2
'        ' chi tieu 23
'        .Sheet = 1
'        .Col = .ColLetterToNumber("J")
'        .Row = 27
'        '.Formula = "D53"
'        .Formula = colD & rowHidden53
'         ' chi tieu 24
'        .Sheet = 1
'        .Col = .ColLetterToNumber("L")
'        .Row = 27
'        '.Formula = "E53"
'        .Formula = colE & rowHidden53
'         ' chi tieu 25
'        .Sheet = 1
'        .Col = .ColLetterToNumber("L")
'        .Row = 28
'        '.Formula = "F53"
'        .Formula = colF & rowHidden53
    End With
End Sub
Public Sub update_01_4A()
'    Dim blCheck_S4A As Boolean
'    Dim blCheck_S4B As Boolean
'    Dim blCheck_S01_1 As Boolean
'    Dim blCheck_S01_2 As Boolean
'
'    blCheck_S4A = IIf(TAX_Utilities_v1.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S4B = IIf(TAX_Utilities_v1.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'
'    blCheck_S01_1 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    blCheck_S01_2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'
'    ' Tong hop len bang ke 01_4A
'        With fps
''        .Sheet = .ActiveSheet
''        Select Case .Sheet
''            Case 2 ' Phu luc 01_1
'                If blCheck_S4A = True And blCheck_S01_1 = True Then
'                    .Sheet = 5
'                    ' chi tieu tong doanh so ban ra tren phu luc 1
'                    .Col = .ColLetterToNumber("D")
'                    .Row = 16
'                    .Formula = "C24"
'                    ' chi tieu T?ng doanh thu hàng hóa, d?ch v? bán ra ch?u thu? GTGT
'                    .Sheet = 5
'                    .Col = .ColLetterToNumber("D")
'                    .Row = 17
'                    .Formula = "D24"
'                End If
''            Case 3 ' phu luc 01_2
'                If blCheck_S4A = True And blCheck_S01_2 = True Then
'                    ' chi tieu tong dong 1 cot 10
'                    'A1
'                    .Sheet = 5
'                    .Col = .ColLetterToNumber("D")
'                    .Row = 7
'                    .Formula = "C23"
'                     ' chi tieu A2-04A
'                    .Sheet = 5
'                    .Col = .ColLetterToNumber("D")
'                    .Row = 9
'                    .Formula = "D23"
'                    'chi tieu tong dong 3 cot 10
'                    'A3
'                    .Sheet = 5
'                    .Col = .ColLetterToNumber("D")
'                    .Row = 11
'                    .Formula = "E23"
'                    .Sheet = 5
'                     ' chitieu B4
'                    .Col = .ColLetterToNumber("D")
'                    .Row = 19
'                    .Formula = "E24"
'                End If
'
'                If blCheck_S4A = True Then
'                    ' chi tieu 25
'                    .Sheet = 1
'                    .Col = .ColLetterToNumber("L")
'                    .Row = 28
'                    '.Formula = "F53"
'                    .Formula = colF & rowHidden53
'                End If
''         End Select
'    End With
End Sub

Public Sub reset_01_4A()
'    ' Tong hop len bang ke 01_4A
'        With fps
'        .Sheet = 5
'        ' chi tieu tong doanh so ban ra tren phu luc 1
'        .Col = .ColLetterToNumber("D")
'        .Row = 16
'        .Formula = ""
'        ' chi tieu T?ng doanh thu hàng hóa, d?ch v? bán ra ch?u thu? GTGT
'        .Sheet = 5
'        .Col = .ColLetterToNumber("D")
'        .Row = 17
'        .Formula = ""
'        ' chi tieu tong dong 1 cot 10
'        'A1
'        .Sheet = 5
'        .Col = .ColLetterToNumber("D")
'        .Row = 7
'        .Formula = ""
'         ' chi tieu A2-04A
'        .Sheet = 5
'        .Col = .ColLetterToNumber("D")
'        .Row = 9
'        .Formula = ""
'        'chi tieu tong dong 3 cot 10
'        'A3
'        .Sheet = 5
'        .Col = .ColLetterToNumber("D")
'        .Row = 11
'        .Formula = ""
'        .Sheet = 5
'         ' chitieu B4
'        .Col = .ColLetterToNumber("D")
'        .Row = 19
'        .Formula = ""
'        ' chi tieu 25
'        .Sheet = 1
'        .Col = .ColLetterToNumber("L")
'        .Row = 28
'        .Formula = ""
'    End With
End Sub

' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 22
    denRowDL = 32
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "9_11_2~9_11_3~10_11_4~9_11_5~10_11_6~11_12_7~20_30_8"
    ' end
End Sub

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
                        tenCucThue = tenCucThue + Right$(arrDanhsach(3), Len(arrDanhsach(3)) - 6) + Chr$(9)
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
            .Sheet = sheetPl01_5_GTGT 'set thong tin CQL
            '.EventEnabled(EventAllEvents) = False
            
            .Row = rowStartPL01_5
            .Col = .ColLetterToNumber("F")
            .TypeComboBoxList = tenCucThue
            
            .Row = rowStartPL01_5
            .Col = .ColLetterToNumber("K")
            .TypeComboBoxList = maCucThue
            
            .Row = rowStartPL01_5
            .Col = .ColLetterToNumber("G")
            .TypeComboBoxList = tenChiCucThue
    
            .Row = rowStartPL01_5
            .Col = .ColLetterToNumber("J")
            .TypeComboBoxList = maChiCucThue
            
            ' Phu luc 01_6
            .Sheet = sheetPl01_6_GTGT 'set thong tin CQL
            .Row = rowStartPL01_6
            .Col = .ColLetterToNumber(colE)
            .TypeComboBoxList = tenCucThue
            
            .Row = rowStartPL01_6
            .Col = .ColLetterToNumber(colM)
            .TypeComboBoxList = maCucThue
            
            .Row = rowStartPL01_6
            .Col = .ColLetterToNumber(colF)
            .TypeComboBoxList = tenChiCucThue
    
            .Row = rowStartPL01_6
            .Col = .ColLetterToNumber(colL)
            .TypeComboBoxList = maChiCucThue
            
            'phu luc 01_2TD/GTGT
            .Sheet = sheetPl01_2TD_GTGT 'set thong tin CQL
            .Row = 42
            .Col = .ColLetterToNumber("P")
            .TypeComboBoxList = tenCucThue

            .Row = 42
            .Col = .ColLetterToNumber("AH")
            .TypeComboBoxList = maCucThue
'
'            .Row = 6
'            .Col = .ColLetterToNumber(colF)
'            .TypeComboBoxList = tenChiCucThue
'
'            .Row = 6
'            .Col = .ColLetterToNumber(colJ)
'            .TypeComboBoxList = maChiCucThue
'
'            .Row = 6
'            .Col = .ColLetterToNumber(colG)
'            .TypeComboBoxList = loaiCq
'
'            .Row = 6
'            .Col = .ColLetterToNumber(colL)
'            .TypeComboBoxList = maLoaiCq
            
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

Private Sub setChangeComb(ByVal Col As Long, Row As Long)
    Dim varTemp As Variant
    Dim intIndexCuc As Integer
    Dim intIndexChiCuc As Integer
    Dim varIdCucThue As Variant
    
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim arrDanhsach() As String
    Dim tenChiCucThue As String
    Dim maChiCucThue As String
    
    
    Dim tenCucThue As String
    Dim maCucThue As String
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim intIndexCombo As Integer, i As Integer
    Dim iCol As Long, iRow As Long
  '  Dim checkDL As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        If .ActiveSheet = sheetPl01_5_GTGT Then
            .Sheet = sheetPl01_5_GTGT
            mCurrentSheet = sheetPl01_5_GTGT
            
            .Col = Col
            .Row = Row
            
            If Col = .ColLetterToNumber("F") Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
                intIndexCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("K")
                .Row = Row
                .typeComboboxCurSel = intIndexCuc
                'Nhap lai Combo cot Q
                .Col = .ColLetterToNumber("K")
                UpdateCell fps, .Col, .Row, .Text
                ' Lay Id cua cuc thue vua chon o tren
                fps.GetText .Col, .Row, varIdCucThue
                
                If varIdCucThue <> "" Or varIdCucThue <> vbNullString Then
                    ' Lay thong tin cho chi cuc thue
                    varIdCucThue = Left$(varIdCucThue, 3)
                    
                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
                        For Each xmlNode In xmlNodeListCell
                            If GetAttribute(xmlNode, "Value") <> "" Then
                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                                
                                If arrDanhsach(0) = "0" And arrDanhsach(2) = varIdCucThue Then
                                    tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
                                    maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
                                End If
                            End If
                        Next
                        Set xmlDomData = Nothing
                        Set xmlNodeListCell = Nothing
                        Set xmlNode = Nothing
                    End If
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("G")
                    .TypeComboBoxList = tenChiCucThue
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("J")
                    .TypeComboBoxList = maChiCucThue
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                   
                End If
            End If
                        
            If Col = .ColLetterToNumber("G") Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
                intIndexChiCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("J")
                .Row = Row
                .typeComboboxCurSel = intIndexChiCuc
                'Nhap lai Combo cot Q
                .Col = .ColLetterToNumber("J")
                UpdateCell fps, .Col, .Row, .Text
            End If
        ElseIf .ActiveSheet = sheetPl01_6_GTGT Then
            .Sheet = sheetPl01_6_GTGT
            mCurrentSheet = sheetPl01_6_GTGT
            .Col = Col
            .Row = Row
            
            If Col = .ColLetterToNumber(colE) Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
                intIndexCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber(colM)
                .Row = Row
                .typeComboboxCurSel = intIndexCuc
                'Nhap lai Combo cot Q
                .Col = .ColLetterToNumber(colM)
                UpdateCell fps, .Col, .Row, .Text
                ' Lay Id cua cuc thue vua chon o tren
                fps.GetText .Col, .Row, varIdCucThue
                
                If varIdCucThue <> "" Or varIdCucThue <> vbNullString Then
                    ' Lay thong tin cho chi cuc thue
                    
                                        
                    varIdCucThue = Left$(varIdCucThue, 3)
                    
                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
                        For Each xmlNode In xmlNodeListCell
                            If GetAttribute(xmlNode, "Value") <> "" Then
                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                                
                                If arrDanhsach(0) = "0" And arrDanhsach(2) = varIdCucThue Then
                                    tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
                                    maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
                                End If
                            End If
                        Next
                        Set xmlDomData = Nothing
                        Set xmlNodeListCell = Nothing
                        Set xmlNode = Nothing
                    End If
                    
                    .Row = Row
                    .Col = .ColLetterToNumber(colF)
                    .TypeComboBoxList = tenChiCucThue
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = Row
                    .Col = .ColLetterToNumber(colL)
                    .TypeComboBoxList = maChiCucThue
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                   
                End If
            End If
                        
            If Col = .ColLetterToNumber(colF) Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
                intIndexChiCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber(colL)
                .Row = Row
                .typeComboboxCurSel = intIndexChiCuc
                'Nhap lai Combo cot Q
                .Col = .ColLetterToNumber(colL)
                UpdateCell fps, .Col, .Row, .Text
            End If
        
        ElseIf .ActiveSheet = sheetPl01_2TD_GTGT Then
            .Sheet = sheetPl01_2TD_GTGT
            mCurrentSheet = sheetPl01_2TD_GTGT
            .Col = Col
            .Row = Row
'
'            If Col = .ColLetterToNumber(colD) Then 'Khi chon Combo
'                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
'                intIndexCuc = .TypeComboBoxCurSel
'                .Col = .ColLetterToNumber(colK)
'                .Row = Row
'                .TypeComboBoxCurSel = intIndexCuc
'                'Nhap lai Combo cot Q
'                .Col = .ColLetterToNumber(colK)
'                UpdateCell fps, .Col, .Row, .Text
'                ' Lay Id cua cuc thue vua chon o tren
'                fps.GetText .Col, .Row, varIdCucThue
'
'                If varIdCucThue <> "" Or varIdCucThue <> vbNullString Then
'                    ' Lay thong tin cho chi cuc thue
'
'
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
'                    .Row = Row
'                    .Col = .ColLetterToNumber(colE)
'                    .TypeComboBoxList = tenChiCucThue
'
'                    .Row = Row
'                    .Col = .ColLetterToNumber(colJ)
'                    .TypeComboBoxList = maChiCucThue
'
'                End If
'            End If
'
'            If Col = .ColLetterToNumber(colE) Then 'Khi chon Combo
'                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
'                intIndexChiCuc = .TypeComboBoxCurSel
'                .Col = .ColLetterToNumber(colJ)
'                .Row = Row
'                .TypeComboBoxCurSel = intIndexChiCuc
'                'Nhap lai Combo cot Q
'                .Col = .ColLetterToNumber(colJ)
'                UpdateCell fps, .Col, .Row, .Text
'            End If
'
'
            If Col = .ColLetterToNumber("P") Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot G va gia tri Index cot L
                intIndexChiCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("AH")
                .Row = Row
                .typeComboboxCurSel = intIndexChiCuc
                'Nhap lai Combo cot L
                .Col = .ColLetterToNumber(colL)
                UpdateCell fps, .Col, .Row, .Text
            End If
           
 
            End If
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

' set ma hoa don
Private Sub setMaHoaDon()
    Dim strDataFileName As String
    Dim arrDanhsach() As String
    Dim strComboHien As String
    Dim strCombo As String
    Dim MSTDN As String
    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim strTemp As String
    
    
    strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_loai_HD.xml")
    ' Lay danh muc loai hoa don
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListCell
            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                strCombo = strCombo + CPab(arrDanhsach(0), 10) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(2), 200) + Chr$(9)
                strComboHien = strComboHien + arrDanhsach(1) + Chr$(9)
                strTemp = strTemp & arrDanhsach(0) & "_" & arrDanhsach(1) & "~"
            End If
        Next
    End If
    
    ' set ma loai HD
    arrLoaiHd = Split(Left$(strTemp, Len(strTemp) - 1), "~")
    
    
     With fps
        .Sheet = sheetPl01_1_GTGT
        ' Set gia tri ma HD
        ' Group khong chiu thue
        .Col = .ColLetterToNumber("C")
        .Row = 8
        .TypeComboBoxList = strComboHien
        'Group 0%
        .Col = .ColLetterToNumber("C")
        .Row = 13
        .TypeComboBoxList = strComboHien
        ' Group 5%
        .Col = .ColLetterToNumber("C")
        .Row = 18
        .TypeComboBoxList = strComboHien
        ' Group 10%
        .Col = .ColLetterToNumber("C")
        .Row = 23
        .TypeComboBoxList = strComboHien
        ' Group khong phai tong hop len Tk 01/GTGT
        .Col = .ColLetterToNumber("C")
        .Row = 28
        .TypeComboBoxList = strComboHien
        
        
        .Sheet = sheetPl01_2_GTGT
        ' Set gia tri ma HD
        ' Group khong chiu thue
        .Col = .ColLetterToNumber("C")
        .Row = 8
        .TypeComboBoxList = strComboHien
        'Group 0%
        .Col = .ColLetterToNumber("C")
        .Row = 13
        .TypeComboBoxList = strComboHien
        ' Group 5%
        .Col = .ColLetterToNumber("C")
        .Row = 18
        .TypeComboBoxList = strComboHien
        ' Group 10%
        .Col = .ColLetterToNumber("C")
        .Row = 23
        .TypeComboBoxList = strComboHien
        ' Group khong phai tong hop len Tk 01/GTGT
        .Col = .ColLetterToNumber("C")
        .Row = 28
        .TypeComboBoxList = strComboHien
    End With
End Sub



' ham set thong tin sau khi tai XML
Public Sub finishImportXML()
    Dim i As Integer
    Dim CQT As Variant
    
    Dim maLyDoGiaHan As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
         .Sheet = 1
         .Row = 20
         .GetText .ColLetterToNumber("P"), .Row, maLyDoGiaHan
         If Trim$(maLyDoGiaHan) <> "" Then
               .Row = 18
               .Col = .ColLetterToNumber("B")
               .value = 1
               UpdateCell fps, .Col, .Row, "x"
         Else
                .Row = 18
               .Col = .ColLetterToNumber("B")
               .value = 1
               UpdateCell fps, .Col, .Row, ""
         End If
        
         .Sheet = sheetPl01_2TD_GTGT
         If .SheetVisible = True Then
                .Row = 42
                .Col = .ColLetterToNumber("B")
                i = 1
                Do
                     .GetText .ColLetterToNumber("P"), .Row, CQT
                     If Trim$(CStr(CQT)) = "" Or CQT = vbNullString Then
                        .Text = i
                        .Col = .ColLetterToNumber("G")
                        .value = 1
                        .Col = .ColLetterToNumber("P")
                        .Lock = True
                        i = i + 1
                     End If
                     .Col = .ColLetterToNumber("B")
                    .Row = .Row + 1
                Loop Until .Text = "aa"
                
         End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
