VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01BTNDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_ROW = 30
Const KY_LAP_BO_COL = "E"
Const MA_SO_TEP_ROW = 13
Const MA_SO_TEP_COL = "M"
Const NGAY_NHAN_TO_KHAI_ROW = 32
Const NGAY_NHAN_TO_KHAI_COL = "E"
Const NGUOI_SU_DUNG_ROW = 32
Const NGUOI_SU_DUNG_COL = "V"

Const HEADER_SO_TT_TRONG_TEP_ROW = 39 'Dong 16 da dc Hidden
Const HEADER_SO_TT_TRONG_TEP_COL = "V"

Const HEADER_KY_LAP_BO_ROW = 13
Const HEADER_KY_LAP_BO_COL = "B"

Const PHONG_XU_LY_ROW = 30
Const PHONG_XU_LY_COL = "V"

Const TEN_DL_ROW = 18
Const TEN_DL_COL = "H"

Const DIA_CHI_DL_ROW = 22
Const DIA_CHI_DL_COL = "F"

Const DIEN_THOAI_DL_ROW = 26
Const DIEN_THOAI_DL_COL = "F"

Const FAX_DL_ROW = 26
Const FAX_DL_COL = "N"

Const EMAIL_DL_ROW = 26
Const EMAIL_DL_COL = "V"

Const SO_HOP_DONG_DL_ROW = 28
Const SO_HOP_DONG_DL_COL = "H"

Const NGAY_HOP_DONG_DL_ROW = 28
Const NGAY_HOP_DONG_DL_COL = "N"

' Thong tin nganh nghe KD va to khai PS
Const NGANH_NGHE_KD_ROW = 77
Const NGANH_NGHE_KD_COL = "L"
Const SHEET_PL_01A_2_TD_TNDN = 3

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public strNgayNhanToKhai As String
Public strNguoiSuDung As String
Public strPhongXuLy As String
Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long

Public dNgayTC As Date

Public strTenDL As String
Public strDiaChiDL As String
Public strDienThoaiDL As String
Public strFaxDL As String
Public strEmailDL As String
Public strSoHopDongDL As String
Public strNgayHopDongDL As String

Private Sub DataDM(ByVal Id As String, Optional ByRef TenTN As String)
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
                        If arrDanhsach(1) = Id Then
                            TenTN = arrDanhsach(3)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub

Public Function Prepared1() As Boolean
    Dim xmlDOMdata As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim strDataFileName As String
    Dim xmlNodeListItems As MSXML.IXMLDOMNodeList
    Dim strCboLyDo As String
    Dim strCboLyDoAn As String
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim fldList() As String
    With fps
        .Sheet = 1
        .Col = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_COL)
        .Row = HEADER_SO_TT_TRONG_TEP_ROW
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetNumeric
        .TypeMaxEditLen = 10
        
        'Ma so tep
        .Col = .ColLetterToNumber("M")
        .Row = 13
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetAlphanumeric
        .TypeMaxEditLen = 20
        
        'Ghi chu
        .Col = .ColLetterToNumber("M")
        .Row = 36
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100
        
        'Phong xu ly
        .Col = .ColLetterToNumber("V")
        .Row = 30
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
        
        'Ky lap bo
        SetDateFormat fps, 1, 30, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 32, .ColLetterToNumber("E"), DDMMYYYY
        .Sheet = 1
        .Row = 32
        .Col = .ColLetterToNumber("E")
        .Text = Format(Date, "dd/mm/yyyy")
        
        .TypeHAlign = TypeHAlignLeft
        ' Ngay quet
        SetDateFormat fps, 1, 32, .ColLetterToNumber("M"), DDMMYYYY
        .TypeHAlign = TypeHAlignLeft
        .Sheet = 1
        .Row = 32
        .Col = .ColLetterToNumber("M")
        .Text = Format(Date, "dd/mm/yyyy")
        
        ' set gia tri gia han
        ' Ky dc gia han moi set gia tri ly do gia han
        strDataFileName = GetAbsolutePath("..\InterfaceTemplates\xml\Catalogue_DM_Lydo_GH_TT16.xml")
        If xmlDOMdata.Load(strDataFileName) Then
             Set xmlNodeListItems = xmlDOMdata.getElementsByTagName("Item")
             For Each xmlNode In xmlNodeListItems
                 fldList = Split(GetAttribute(xmlNode, "Value"), "###")
                 strCboLyDo = strCboLyDo + fldList(1) + Chr$(9)
                 strCboLyDoAn = strCboLyDoAn + fldList(0) + Chr$(9)
             Next
         End If
          'set ly do gia han
        .Row = 64
        .Col = .ColLetterToNumber("M")
        .TypeComboBoxList = strCboLyDo
        ' set Ma ly do gia han
        .Col = .ColLetterToNumber("AC")
        .TypeComboBoxList = strCboLyDoAn
    End With
    
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
    Dim i As Integer
    Dim varMaDM As Variant
    Dim tenDM As String
    i = 0
    ReDim Preserve larrPhongXuLy(0)
    larrPhongXuLy(0) = "00"
    ReDim Preserve larrid(0)
    larrid(0) = "00"
    
    With fps
        .Sheet = 1
        .Col = .ColLetterToNumber(PHONG_XU_LY_COL)
        .Row = PHONG_XU_LY_ROW
        If rsPXL.Fields.count > 0 Then
            Do While Not rsPXL.EOF
                .TypeComboBoxIndex = -1
                .TypeComboBoxString = TAX_Utilities_Srv_New.Convert(rsPXL.Fields(1).Value, TCVN, UNICODE)
                i = i + 1
                ReDim Preserve larrPhongXuLy(i)
                larrPhongXuLy(i) = TAX_Utilities_Srv_New.Convert(rsPXL.Fields(1).Value, TCVN, UNICODE)
                ReDim Preserve larrid(i)
                larrid(i) = rsPXL.Fields(0).Value
                rsPXL.MoveNext
            Loop
        Else
            DisplayMessage "0077", msOKOnly
            Exit Function
        End If
        .TypeComboBoxCurSel = 0
        UpdateCell fps, .Col, .Row, .Text
        lSoPhongXL = i
             
        'get ten chi tieu trong danh muc NNKD
        .Sheet = 1
        .Col = .ColLetterToNumber("AC")
        .Row = 64
         .GetText .ColLetterToNumber("AC"), .Row, varMaDM
        If Trim(varMaDM) <> vbNullString Then
            .Col = .ColLetterToNumber("M")
            .Row = 64
            .Text = getTenLyDoGD(varMaDM)
        End If
        
        ' Neu khong co gia han se khong cho chon combobox
         ' Neu check gia han moi cho phep chon GT
        .Col = .ColLetterToNumber("K")
        .Row = 62
        If .Text = "1" Or UCase$(.Text) = "X" Then
            .Col = .ColLetterToNumber("M")
            .Row = 64
            .Lock = False
        Else
            .Col = .ColLetterToNumber("M")
            .Row = 64
            .Lock = True
        End If
        ' End set gia han
        
                .Sheet = 2
        i = 0
        Do
            .GetText .ColLetterToNumber("E"), i + 35, varMaDM
            DataDM varMaDM, tenDM
            If Trim(tenDM) <> vbNullString Then
                .Col = .ColLetterToNumber("G")
                .Row = 35 + i
                .Text = tenDM
            End If
            UpdateCell fps, .Col, .Row, .Text
             i = i + 1
            .Col = .ColLetterToNumber("C")
            .Row = i + 35
        Loop Until .Text = "aa" Or varMaDM = ""
        
                'fill data CQT PL 02 TD TNDN
        .Sheet = SHEET_PL_01A_2_TD_TNDN
        i = 0
        Do
            .GetText .ColLetterToNumber("AE"), i + 18, varMaDM
            DataDM varMaDM, tenDM
            If Trim(tenDM) <> vbNullString Then
                .Col = .ColLetterToNumber("Q")
                .Row = 18 + i
                .Text = tenDM
                UpdateCell fps, .Col, .Row, .Text
            End If
             i = i + 1
            .Col = .ColLetterToNumber("B")
            .Row = i + 18
            tenDM = ""
        Loop Until .Text = "aa"
        
    End With
    
    Prepared2 = True
End Function

Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

Public Function Prepared3() As Boolean
    Dim dYear, dQuy As Integer
    Dim STT As Integer
    
    Dim isNhieuMucThueUD As Variant
    Dim xmlNode As MSXML.IXMLDOMNode
    
    
    With fps
        '.EventEnabled(EventAllEvents) = False
        'Set MaSoTep
        .Sheet = 1
'        .Col = .ColLetterToNumber(MA_SO_TEP_X)
'        .Row = MA_SO_TEP_Y
'        .Text = strMaSoTep
'        UpdateCell fps, .Col, .Row, .Text
        
        'Set NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_COL)
        .Row = NGAY_NHAN_TO_KHAI_ROW
        If strNgayNhanToKhai <> "" Then
            .Text = strNgayNhanToKhai
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        'Set NguoiSuDung
        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_COL)
        .Row = NGUOI_SU_DUNG_ROW
        If strNguoiSuDung <> "" Then
            .Text = strNguoiSuDung
            UpdateCell fps, .Col, .Row, .Text
        End If
        
         'set Phong xu ly
        .Col = .ColLetterToNumber(PHONG_XU_LY_COL)
        .Row = PHONG_XU_LY_ROW
        If strPhongXuLy <> "" Then
            .Text = strPhongXuLy
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        ' Dat lai check E_17 Duoc gian thue theo NQ30/NQ-CP
        ' Thoi gian gian thue trong ca nam 2009 cua to khai TNDN
        dYear = Year(dNgayTC)
        dQuy = DatePart("Q", dNgayTC)
'        'If Val(TAX_Utilities_Srv_New.Year) = 2009 Then
'        If dYear = 2009 Then
'            .Row = 38
'            .Col = .ColLetterToNumber("E")
'            UpdateCell fps, .Col, .Row, .Text
'            .Row = 38
'            .Col = .ColLetterToNumber("F")
'            .Text = "GIA HAN NOP THEO NGHI DINH 30/NQ-CP"
'        ElseIf dYear = 2010 And (dQuy = 1 Or dQuy = 2 Or dQuy = 3) Then
'            .Row = 38
'            .Col = .ColLetterToNumber("E")
'            UpdateCell fps, .Col, .Row, .Text
'            .Row = 38
'            .Col = .ColLetterToNumber("F")
'            .Text = "GIA HAN THEO QUYET DINH 12/2010/QD-TTG"
'        ElseIf (dYear = 2010 And dQuy = 4) Then
'            .Row = 38
'            .Col = .ColLetterToNumber("E")
'            UpdateCell fps, .Col, .Row, .Text
'            .Row = 38
'            .Col = .ColLetterToNumber("F")
'            .Text = "GIA HAN THEO QUYET DINH 21/2011/QD-TTG"
'        ElseIf dYear = 2011 Then
'            .Row = 38
'            .Col = .ColLetterToNumber("E")
'            UpdateCell fps, .Col, .Row, .Text
'            .Row = 38
'            .Col = .ColLetterToNumber("F")
'            .Text = "GIA HAN THEO QUYET DINH 21/2011/QD-TTG; QD 54/2011/QD-TTG"
'        Else
'            .Row = 38
'            .Col = .ColLetterToNumber("E")
'            UpdateCell fps, .Col, .Row, .Text
'            .Row = 38
'            .Col = .ColLetterToNumber("F")
'            .Text = "GIA HAN THEO NQ 02/NQ-CP"
'            ' Neu ngoai hieu luc cua gian thue thi visible cac dieu kien check nay di
''            .Row = 38
''            .Col = .ColLetterToNumber("B")
''            .CellType = CellTypeEdit
''            .Lock = True
''            .Text = ""
''
''            .Row = 38
''            .Col = .ColLetterToNumber("E")
''            .CellType = CellTypeEdit
''            .Lock = True
''            .Text = ""
''            UpdateCell fps, .Col, .Row, ""
''
''            .Row = 38
''            .Col = .ColLetterToNumber("F")
''            .CellType = CellTypeEdit
''            .Lock = True
''            .Text = ""
'        End If
        ' set thong tin dai ly thue
        ' ten dai ly thue
        
                
        .Row = 62
        .Col = .ColLetterToNumber("k")
        UpdateCell fps, .Col, .Row, .Text
        
        .Col = .ColLetterToNumber(TEN_DL_COL)
        .Row = TEN_DL_ROW
        If strTenDL <> "" Then
            .Text = strTenDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' dia chi dai ly thue
        .Col = .ColLetterToNumber(DIA_CHI_DL_COL)
        .Row = DIA_CHI_DL_ROW
        If strDiaChiDL <> "" Then
            .Text = strDiaChiDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' dien thoai dai ly
        .Col = .ColLetterToNumber(DIEN_THOAI_DL_COL)
        .Row = DIEN_THOAI_DL_ROW
        If strDienThoaiDL <> "" Then
            .Text = strDienThoaiDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' fax dai ly thue
        .Col = .ColLetterToNumber(FAX_DL_COL)
        .Row = FAX_DL_ROW
        If strFaxDL <> "" Then
            .Text = strFaxDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' Email dai ly
        .Col = .ColLetterToNumber(EMAIL_DL_COL)
        .Row = EMAIL_DL_ROW
        If strEmailDL <> "" Then
            .Text = strEmailDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' so hop dong
        .Col = .ColLetterToNumber(SO_HOP_DONG_DL_COL)
        .Row = SO_HOP_DONG_DL_ROW
        If strSoHopDongDL <> "" Then
            .Text = strSoHopDongDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' ngay hop dong
        .Col = .ColLetterToNumber(NGAY_HOP_DONG_DL_COL)
        .Row = NGAY_HOP_DONG_DL_ROW
        If strNgayHopDongDL <> "" Then
            .Text = strNgayHopDongDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        ' Doi voi to khai 01B/TNDN khi chon nhieu muc thue uu dai -> set chi tieu 26 ve kieu text
        .GetText .ColLetterToNumber("L"), 51, isNhieuMucThueUD
        If isNhieuMucThueUD = "1" Or UCase$(CStr(isNhieuMucThueUD)) = "X" Then
            .Col = .ColLetterToNumber("U")
            .Row = 51
            .CellType = CellTypeEdit
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeHAlign = TypeHAlignRight
            .TypeVAlign = TypeHAlignCenter
            ' Lay chi tieu 26 trong datafile
            Set xmlNode = TAX_Utilities_Srv_New.Data(0).nodeFromID("U_51")
            .Text = GetAttribute(xmlNode, "Value")
        End If
        
        
        
        .Sheet = 2
        If .SheetVisible = True Then
            STT = 0
            .Row = 35
                   Do
                        .Col = .ColLetterToNumber("C")
                        If (.Row = 35 + STT) Then
                           .Text = str(STT + 1)
                           STT = STT + 1
                        End If
                        .Col = .ColLetterToNumber("C")
                        .Row = .Row + 1
                   Loop Until .Text = "aa"
        End If
        
        '.EventEnabled(EventAllEvents) = True
    End With
    
    Prepared3 = True
End Function

Public Sub Prepared4(ByVal dNgayDauKy As Date)
Dim strKyLapBo As String
Dim blnValid As Boolean

    blnValid = True
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        
        ' Get ky lap bo
        .Col = .ColLetterToNumber(KY_LAP_BO_COL)
        .Row = KY_LAP_BO_ROW
        strKyLapBo = .Text
                        
        ' Get Phong xu ly
        .Col = .ColLetterToNumber(PHONG_XU_LY_COL)
        .Row = PHONG_XU_LY_ROW
        strPhongXuLy = .Text
        
        ' Get NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_COL)
        .Row = NGAY_NHAN_TO_KHAI_ROW
        strNgayNhanToKhai = .Text
        
        'Go to last sheet (header sheet)
        .Sheet = .SheetCount
        
'        If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
'            blnValid = False
'        End If
'
'        .Col = .ColLetterToNumber(HEADER_KY_LAP_BO_COL)
'        .Row = HEADER_KY_LAP_BO_ROW
'
'        If Not blnValid Then
'            .Formula = "0"
'        Else
'            .Formula = "1"
'        End If
        ' Kiem tra neu ky lap bo trang se bao loi
        If Trim(strKyLapBo) = "" Or Trim(strKyLapBo) = "../...." Then
            blnValid = False
            .Col = .ColLetterToNumber("B")
            .Row = 19
            .Formula = IIf(blnValid, "1", "0")
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 19
            .Formula = IIf(blnValid, "1", "0")
            
            If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
                blnValid = False
            Else
                blnValid = True
            End If
            .Col = .ColLetterToNumber(HEADER_KY_LAP_BO_COL)
            .Row = HEADER_KY_LAP_BO_ROW
            
            If Not blnValid Then
                .Formula = "0"
            Else
                .Formula = "1"
            End If
            
            
            ' Kiem tra xem ky ke khai trong cung ky hay khac ky
'            If (DateAdd("M", 1, dNgayDauKy) = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Or (dNgayDauKy = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Then
'                blnValid = True
'            Else
'                blnValid = False
'            End If
'
'            .Col = .ColLetterToNumber("B")
'            .Row = 17
'
'            If Not blnValid Then
'                .Formula = "0"
'            Else
'                .Formula = "1"
'            End If
            ' kiem tra ky lap bo khong duoc lon hon ky hien tai
            If DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) > Date Then
                blnValid = False
            Else
                blnValid = True
            End If
            .Col = .ColLetterToNumber("B")
            .Row = 18
            
            If Not blnValid Then
                .Formula = "0"
            Else
                .Formula = "1"
            End If
        End If

        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Function GetLastMonthOfThreeMonth(strPeriod As String) As String
    Select Case strPeriod
        Case "01"
            GetLastMonthOfThreeMonth = "03"
        Case "02"
            GetLastMonthOfThreeMonth = "06"
        Case "03"
            GetLastMonthOfThreeMonth = "09"
        Case "04"
            GetLastMonthOfThreeMonth = "12"
    End Select
End Function

Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
' set gia lock cell ly do gia han
   fps.EventEnabled(EventAllEvents) = False
    If fps.ActiveSheet = 1 Then
            ' Neu check gia han moi cho phep chon GT
            If Col = fps.ColLetterToNumber("K") And Row = 62 Then
                fps.Col = fps.ColLetterToNumber("K")
                fps.Row = 62
                If fps.Text = "1" Or UCase$(fps.Text) = "X" Then
                    fps.Col = fps.ColLetterToNumber("M")
                    fps.Row = 64
                    fps.Lock = False
                Else
                    fps.Col = fps.ColLetterToNumber("M")
                    fps.Row = 64
                    fps.Lock = True
                    fps.SetText fps.ColLetterToNumber("M"), 64, ""
                    fps.SetText fps.ColLetterToNumber("AC"), 64, "00"
                    UpdateCell fps, fps.ColLetterToNumber("AC"), 64, "00"
                    UpdateCell fps, fps.ColLetterToNumber("M"), 64, ""
                End If
            End If
    End If
    fps.EventEnabled(EventAllEvents) = True
End Sub

Private Sub fps_Click(ByVal Col As Long, ByVal Row As Long)
    With fps
        .Sheet = .ActiveSheet
        .Col = .ActiveCol
        .Row = .ActiveRow
        If Not (.CellType = CellTypeCheckBox Or .CellType = CellTypeButton) Then
            GetCellSpan fps, Col, Row
            fps.SetActiveCell Col, Row
        End If
    End With
End Sub

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    Dim intIndexCombo As Integer
    With fps
       If .ActiveSheet = 1 Then
            .Col = Col
            .Row = Row
             If Col = .ColLetterToNumber("M") And Row = 64 Then
                intIndexCombo = .TypeComboBoxCurSel
                .Col = .ColLetterToNumber("AC")
                .TypeComboBoxCurSel = intIndexCombo
                'Update combo C
                UpdateCell fps, .Col, .Row, .Text
            End If
        End If
        
    End With
End Sub

Private Sub fps_KeyPress(KeyAscii As Integer)
'    If fps.ActiveCol = fps.ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_COL) And fps.ActiveRow = HEADER_SO_TT_TRONG_TEP_ROW Then _
'        CellEditFormatNumber fps, fps.ActiveSheet, fps.ActiveCol, fps.ActiveRow, KeyAscii
    With fps
        If .ActiveSheet = 1 And .ActiveCol = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_COL) _
           And .ActiveRow = HEADER_SO_TT_TRONG_TEP_ROW Then
            If KeyAscii = 46 Or KeyAscii = 45 Then ' KeyAscii of '.' character
                KeyAscii = 0
            End If
        End If
    End With
End Sub

Public Function CheckValidData() As Boolean
    Dim varCheckValue As Variant, varNoteValue As Variant
    Dim varCheckGhValue As Variant, varNoteGhValue As Variant
    Dim varNgayNop As Variant
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 34, varCheckValue
        .GetText .ColLetterToNumber("M"), 34, varNoteValue
        '
        
        .GetText .ColLetterToNumber("K"), 62, varCheckGhValue
        .GetText .ColLetterToNumber("M"), 64, varNoteGhValue
        
        .GetText .ColLetterToNumber("E"), 32, varNgayNop
        If CStr(varCheckValue) = "1" And Trim(CStr(varNoteValue)) = "" Then
            .Sheet = .SheetCount
            .SetText 2, 14, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText 2, 14, "1"
        End If
        '
        If CStr(Trim(varNgayNop)) = "" Or CStr(Trim(varNgayNop)) = "../../...." Then
            .Sheet = .SheetCount
            .SetText 2, 12, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText 2, 12, "1"
        End If
        
        If CStr(varCheckGhValue) = "1" And Trim(CStr(varNoteGhValue)) = "" Then
            .Sheet = .SheetCount
            .SetText 2, 20, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText 2, 20, "1"
        End If
        
    End With
End Function

Public Function GenerateSQL_Header(xmlDOMdata As MSXML.DOMDocument, strSQL_HDR As String, vHdrID As Variant, vKKBS As Byte, ByVal dNgayDauKy As Date) As String
    Dim xmlList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlAttribute As MSXML.IXMLDOMAttribute
    Dim iRowID As Long, strSQL As String, strTempSQL As String
    Dim dDate As Date, strDate() As String
    Dim vTIN, vTEN_DTNT, vDIA_CHI, vLOAI_TKHAI, vNGAY_NOP, vKyLB
    Dim vKYKK, vNGAY_CAP_NHAT, vNGUOI_CAP_NHAT, vCO_LOI_DDANH, v_CO_GIA_HAN
    Dim vSO_HIEU_TEP, vSO_TT_TK, vDA_NHAN, vGHI_CHU_LOI, vCO_GTRINH_02A, VMaLyDoGH
    Dim vCO_GTRINH_02B, vCO_GTRINH_02C
    Dim vPHONG_XU_LY
    Dim i As Long, j As Long
    Dim strMaPhongXuLy As String
    
    Dim dYear, dQuy As Integer
    
    
    Dim vTIN_DL, vSO_HOP_DONG, vNGAY_HOP_DONG, vLAN_BS, vCHINH_THUC, vNGANH_NGHE_KD, vTO_KHAI_LAN_PS, vNGAY_PS As Variant
    
On Error GoTo ErrHandle
    strSQL = strSQL_HDR
    Set xmlList = xmlDOMdata.getElementsByTagName("Cell")
    For Each xmlNode In xmlList
        With xmlNode.Attributes
        
        If Trim(GetAttribute(xmlNode, "MCT")) = vbNullString Then
            Select Case Trim(GetAttribute(xmlNode, "CellID"))
                Case "F_10"
                    vTIN = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "H_8"
                    vTEN_DTNT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "F_12"
                    vDIA_CHI = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_32"
                    vNGAY_NOP = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_30"
                    vKyLB = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "M_32"
                    vNGAY_CAP_NHAT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "V_32"
                    vNGUOI_CAP_NHAT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_36"
                    vCO_LOI_DDANH = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "M_13"
                    ' vSO_HIEU_TEP = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                    ' Tu phien ban 1.3.0 song tep se thuc hien bang tay=> So hieu tep luon bang ""
                    vSO_HIEU_TEP = ""
                Case HEADER_SO_TT_TRONG_TEP_COL & "_" & HEADER_SO_TT_TRONG_TEP_ROW
                    vSO_TT_TK = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "M_36"
                    vGHI_CHU_LOI = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case PHONG_XU_LY_COL & "_" & PHONG_XU_LY_ROW
                    vPHONG_XU_LY = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "K_62"
                    v_CO_GIA_HAN = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                    If UCase(Trim(v_CO_GIA_HAN)) = "X" Or Trim(v_CO_GIA_HAN) = "1" Then
                        v_CO_GIA_HAN = "'Y'"
                    Else
                        v_CO_GIA_HAN = "Null"
                    End If
                Case "AC_64"
                    VMaLyDoGH = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
            End Select
        End If
        End With
    Next
    
    strSQL = strSQL & "'" & vHdrID & "',"
    strSQL = strSQL & vKKBS & ","
    strSQL = strSQL & "'" & vTIN & "',"
    strSQL = strSQL & "'" & vTEN_DTNT & "',"
    strSQL = strSQL & "'" & vDIA_CHI & "',"
    'vLOAI_TKHAI = TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue
    vLOAI_TKHAI = "01B_TNDN13"
    strSQL = strSQL & "'" & vLOAI_TKHAI & "',"
    strSQL = strSQL & "To_date('" & vNGAY_NOP & "','dd/mm/yyyy'),"
    
    'Ky/Quy LB
    'Ngay dau ky lap bo va ngay cuoi ky lap bo
    strDate = Split(vKyLB, "/")
    dDate = DateSerial(Val(strDate(1)), Val(strDate(0)), 1)
    strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    dDate = DateAdd("m", 1, dDate)
    dDate = DateAdd("d", -1, dDate)
    strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    
    'Ky/ Quy KK
   If Trim(TAX_Utilities_Srv_New.ThreeMonths) <> "" Then
        dDate = dNgayDauKy
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
        dDate = DateAdd("m", 3, dDate)
        dDate = DateAdd("d", -1, dDate)
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    End If
    
    strSQL = strSQL & "To_date('" & vNGAY_CAP_NHAT & "','dd/mm/yyyy'),"
    strSQL = strSQL & "'" & vNGUOI_CAP_NHAT & "',"
    strSQL = strSQL & "'" & vCO_LOI_DDANH & "',"
    strSQL = strSQL & "'" & vSO_HIEU_TEP & "',"
    strSQL = strSQL & "'" & vSO_TT_TK & "',"
    
    strSQL = strSQL & "'" & vDA_NHAN & "',"
    ' Doi voi to khai 01B/TNDN, co gia han theo NQ30/NQ-CP thi phai them vao ghi chu loi
    ' cau sau "Gia han theo NQ30/NQ-CP" va ghep dong thoi voi ghi chu loi cach nhau boi dau ";"
    dYear = Year(dNgayDauKy)
    dQuy = DatePart("Q", dNgayDauKy)
    If v_CO_GIA_HAN = "'Y'" Then
        If Trim(vGHI_CHU_LOI) = vbNullString Then
'            If TAX_Utilities_Srv_New.Year = "2009" Then
'                strSQL = strSQL & "'GIA HAN NOP THEO NQ30/NQ-CP',"
'            ElseIf TAX_Utilities_Srv_New.Year = "2010" And (TAX_Utilities_Srv_New.ThreeMonths = 1 Or TAX_Utilities_Srv_New.ThreeMonths = 2 Or TAX_Utilities_Srv_New.ThreeMonths = 3) Then
'                strSQL = strSQL & "'GIA HAN THEO Q§ 12/2010/Q§-TTg',"
'            ElseIf TAX_Utilities_Srv_New.Year = "2011" Or (TAX_Utilities_Srv_New.Year = "2010" And TAX_Utilities_Srv_New.ThreeMonths = 4) Then
'                strSQL = strSQL & "'GIA HAN THEO Q§ 21/2011/Q§-TTg',"
'            Else
'                strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
'            End If
            If dYear = 2009 Then
                strSQL = strSQL & "'GIA HAN NOP THEO NQ30/NQ-CP',"
            ElseIf dYear = 2010 And (dQuy = 1 Or dQuy = 2 Or dQuy = 3) Then
                strSQL = strSQL & "'GIA HAN THEO Q§ 12/2010/Q§-TTg',"
            ElseIf (dYear = 2010 And dQuy = 4) Then
                strSQL = strSQL & "'GIA HAN THEO Q§ 21/2011/Q§-TTg',"
            ElseIf dYear = 2011 Then
                strSQL = strSQL & "'GIA HAN THEO Q§ 21/2011/Q§-TTg;Q§ 54/2011/Q§-TTg',"
            Else
                'strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
'                If VMaLyDoGH = "01" Then
'                    strSQL = strSQL & "'DN VUA VA NHO GIA HAN THEO NQ 02/NQ-CP',"
'                ElseIf VMaLyDoGH = "02" Then
'                    strSQL = strSQL & "'DN NHIEU LD GIA HAN THEO NQ 02/NQ-CP',"
'                ElseIf VMaLyDoGH = "03" Then
'                    strSQL = strSQL & "'DN DAU TU GIA HAN THEO NQ 02/NQ-CP',"
'                ElseIf VMaLyDoGH = "99" Then
'                    strSQL = strSQL & "'LY DO KHAC GIA HAN THEO NQ 02/NQ-CP',"
'                Else
'                    strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
'                End If
                 strSQL = strSQL & "'" & VMaLyDoGH & "',"
            End If

        Else
'            If TAX_Utilities_Srv_New.Year = "2009" Then
'                strSQL = strSQL & "'GIA HAN NOP THEO NQ30/NQ-CP; " & vGHI_CHU_LOI & "',"
'            ElseIf TAX_Utilities_Srv_New.Year = "2010" And (TAX_Utilities_Srv_New.ThreeMonths = 1 Or TAX_Utilities_Srv_New.ThreeMonths = 2 Or TAX_Utilities_Srv_New.ThreeMonths = 3) Then
'                strSQL = strSQL & "'GIA HAN THEO Q§ 12/2010/Q§-TTg; " & vGHI_CHU_LOI & "',"
'            ElseIf TAX_Utilities_Srv_New.Year = "2011" Or (TAX_Utilities_Srv_New.Year = "2010" And TAX_Utilities_Srv_New.ThreeMonths = 4) Then
'                strSQL = strSQL & "'GIA HAN THEO Q§ 21/2011/Q§-TTg; " & vGHI_CHU_LOI & "',"
'            Else
'                strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
'            End If
            If dYear = 2009 Then
                strSQL = strSQL & "'GIA HAN NOP THEO NQ30/NQ-CP; " & vGHI_CHU_LOI & "',"
            ElseIf dYear = 2010 And (dQuy = 1 Or dQuy = 2 Or dQuy = 3) Then
                strSQL = strSQL & "'GIA HAN THEO Q§ 12/2010/Q§-TTg; " & vGHI_CHU_LOI & "',"
            ElseIf (dYear = 2010 And dQuy = 4) Then
                strSQL = strSQL & "'GIA HAN THEO Q§ 21/2011/Q§-TTg; " & vGHI_CHU_LOI & "',"
            ElseIf dYear = 2011 Then
                strSQL = strSQL & "'GIA HAN THEO Q§ 21/2011/Q§-TTg;Q§ 54/2011/Q§-TTg; " & vGHI_CHU_LOI & "',"
            Else
                'strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
'                 If VMaLyDoGH = "01" Then
'                    strSQL = strSQL & "'DN VUA VA NHO GIA HAN THEO NQ02/NQ-CP; " & vGHI_CHU_LOI & "',"
'                ElseIf VMaLyDoGH = "02" Then
'                    strSQL = strSQL & "'DN NHIEU LD GIA HAN THEO NQ02/NQ-CP; " & vGHI_CHU_LOI & "',"
'                ElseIf VMaLyDoGH = "03" Then
'                    strSQL = strSQL & "'DN DAU TU GIA HAN THEO NQ02/NQ-CP; " & vGHI_CHU_LOI & "',"
'                ElseIf VMaLyDoGH = "99" Then
'                    strSQL = strSQL & "'LY DO KHAC GIA HAN THEO NQ02/NQ-CP; " & vGHI_CHU_LOI & "',"
'                Else
'                    strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
'                End If
                 strSQL = strSQL & "'" & VMaLyDoGH & "; " & vGHI_CHU_LOI & "',"
            End If
            
        End If
    Else
        strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
    End If
    
    'Chi co voi TH GTGT
    strSQL = strSQL & "'','','','','','',null,null,"
    
    With fps
        For i = 1 To lSoPhongXL
            If vPHONG_XU_LY = TAX_Utilities_Srv_New.Convert(larrPhongXuLy(i), UNICODE, TCVN) Then
                strMaPhongXuLy = larrid(i)
                Exit For
            End If
        Next
        
        .Sheet = 1
       ' Ma dai ly thue
        .GetText .ColLetterToNumber("F"), 20, vTIN_DL
        .GetText .ColLetterToNumber(SO_HOP_DONG_DL_COL), SO_HOP_DONG_DL_ROW, vSO_HOP_DONG
        .GetText .ColLetterToNumber(NGAY_HOP_DONG_DL_COL), NGAY_HOP_DONG_DL_ROW, vNGAY_HOP_DONG
        .GetText .ColLetterToNumber("I"), 77, vLAN_BS
        .GetText .ColLetterToNumber("C"), 77, vCHINH_THUC
        If Trim(vCHINH_THUC) = "1" Then
            vLAN_BS = "0"
        End If
        
        .GetText .ColLetterToNumber("K"), 62, v_CO_GIA_HAN
                            If UCase(Trim(v_CO_GIA_HAN)) = "X" Or Trim(v_CO_GIA_HAN) = "1" Then
                        v_CO_GIA_HAN = "'Y'"
                    Else
                        v_CO_GIA_HAN = "Null"
                    End If
        
        If Len(Trim(vTIN_DL)) = 13 Then
            vTIN_DL = Left(vTIN_DL, 10) & "-" & Right(vTIN_DL, 3)
        End If

        
        .GetText .ColLetterToNumber(NGANH_NGHE_KD_COL), NGANH_NGHE_KD_ROW, vNGANH_NGHE_KD
        vNGANH_NGHE_KD = "'" & vNGANH_NGHE_KD & "'"
        vTO_KHAI_LAN_PS = "null"
        vNGAY_PS = "null"
        
        If Trim(vNGAY_HOP_DONG) = "" Or vNGAY_HOP_DONG = vbNullString Then
            vNGAY_HOP_DONG = "null"
        Else
            vNGAY_HOP_DONG = "'" & vNGAY_HOP_DONG & "'"
            vNGAY_HOP_DONG = "to_date(" & vNGAY_HOP_DONG & ",'dd/mm/yyyy')"
        End If
    End With
    strSQL = strSQL & "'" & strMaPhongXuLy & "', null,"
    'strSQL = strSQL & v_CO_GIA_HAN & ")"
    strSQL = strSQL & v_CO_GIA_HAN & ","
    strSQL = strSQL & "'" & Trim(vTIN_DL) & "','" & Trim(vSO_HOP_DONG) & "'," & vNGAY_HOP_DONG & "," & Trim(vLAN_BS) & "," & vNGANH_NGHE_KD & "," & vTO_KHAI_LAN_PS & "," & vNGAY_PS & ")"
    GenerateSQL_Header = strSQL
    
    Exit Function
ErrHandle:
    SaveErrorLog "cls_01BTNDN", "GenerateSQL_Header", Err.Number, Err.Description
End Function

Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
    KiemTraKhoaSo = True
    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Srv_New.Data(0).nodeFromID("E_30"), "Value"), "dd/mm/yyyy")

    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
  
End Function

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    With fps
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("E") And Row = 32 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    Else
                        .SetActiveCell Col, Row
                    End If
'                Else
'                 .SetActiveCell Col, Row
                End If
               .Col = Col
               .Row = Row
               UpdateCell fps, .Col, .Row, .Text
            End If
            'ky lap bo
            
            If Col = .ColLetterToNumber("E") And Row = 30 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../...." Then
                    If Format_mmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_mmyyyy(CStr(varTemp))
                    Else
                        .SetActiveCell Col, Row
                    End If
'                Else
'                 .SetActiveCell Col, Row
                End If
                .Col = Col
                .Row = Row
               UpdateCell fps, .Col, .Row, .Text
            End If
        End If
    End With
End Sub

'ham lay ten DM NNKD
Private Function getTenLyDoGD(maGH As Variant) As String
    Dim maDM1 As String
    Dim strDuongDanFile As String
    Dim xmlDOMdata      As New MSXML.DOMDocument
    Dim xmlNodelist     As MSXML.IXMLDOMNodeList
    Dim xmlNode         As MSXML.IXMLDOMNode
    Dim arrStrName()    As String
    
    getTenLyDoGD = ""
    maDM1 = CStr(Trim(maGH))
    'strDuongDanFile = GetCatalogueFileName
    'lay file chua tenDM
    strDuongDanFile = "..\InterfaceTemplates\xml\Catalogue_DM_Lydo_GH_TT16.xml"
    strDuongDanFile = GetAbsolutePath(strDuongDanFile)
    xmlDOMdata.Load (strDuongDanFile)
    
    Set xmlNodelist = xmlDOMdata.getElementsByTagName("Item")
        For Each xmlNode In xmlNodelist
            If Left(GetAttribute(xmlNode, "Value"), 2) = maDM1 Then
                arrStrName = Split(GetAttribute(xmlNode, "Value"), "###")
                getTenLyDoGD = arrStrName(1)
            End If
        Next
End Function

