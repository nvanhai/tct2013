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
Option Explicit
Private clsConn As New TAX_Utilities_Svr_New.clsADO

Const KY_LAP_BO_Y = 30
Const KY_LAP_BO_X = "E"
Const NGAY_NHAN_TO_KHAI_Y = 32
Const NGAY_NHAN_TO_KHAI_X = "E"
Const NGUOI_SU_DUNG_Y = 32
Const NGUOI_SU_DUNG_X = "V"
Const PHONG_XU_LY_Y = 30
Const PHONG_XU_LY_X = "V"
Const PHONG_QUAN_LY_X = 34
Const PHONQ_QUAN_LY_Y = "V"
Const NGAY_NOP_Y = 32
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 32
Const NGAY_QUET_X = "M"
'Const TEN_BPQL_X = "V"
'Const TEN_BPQL_Y = 26
'Const MA_BPQL_X = "V"
'Const MA_BPQL_Y = 29
Private rs As ADODB.Recordset
' thong tin din vi
Const MA_SO_THUE_X = "F"
Const MA_SO_THUE_Y = 10
Const TEN_GOI_X = "H"
Const TEN_GOI_Y = 8
Const DIA_CHI_X = "F"
Const DIA_CHI_Y = 12
Const QUAN_HUYEN_X = "F"
Const QUAN_HUYEN_Y = 14
Const TINH_TPHO_X = "Q"
Const TINH_TPHO_Y = 14
Const DIEN_THOAI_X = "F"
Const DIEN_THOAI_Y = 16
Const FAX_X = "N"
Const FAX_Y = 16
Const MAIL_X = "V"
Const MAIL_Y = 16
'thong tin Dai ly thue
Const MA_SO_THUE_DL_X = "F"
Const MA_SO_THUE_DL_Y = 20
Const TEN_GOI_DL_X = "H"
Const TEN_GOI_DL_Y = 18
Const DIA_CHI_DL_X = "F"
Const DIA_CHI_DL_Y = 22
Const QUAN_HUYEN_DL_X = "F"
Const QUAN_HUYEN_DL_Y = 24
Const TINH_TPHO_DL_X = "Q"
Const TINH_TPHO_DL_Y = 24
Const DIEN_THOAI_DL_X = "F"
Const DIEN_THOAI_DL_Y = 26
Const FAX_DL_X = "N"
Const FAX_DL_Y = 26
Const MAIL_DL_X = "V"
Const MAIL_DL_Y = 26
Const SO_HDDL_DL_X = "H"
Const SO_HDDL_DL_Y = 28
Const NGAY_HDDL_DL_X = "N"
Const NGAY_HDDL_DL_Y = 28

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
Public isExistFile As Boolean
'Public strMaSoTep As String
Public strNgayNhanToKhai As String
Public strNguoiSuDung As String
Public strPhongXuLy As String
Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long
Private LAN_QUET As Variant
Private DAGHI2 As Variant
Private LOAI_TO_KHAI As String

Private isTkQuy As Boolean

'Longvh
Public strMST As String
Public strTenGoi As String
Public strDchi As String
Public strNganh As String
Public strMaBPQL As String
Public strDThoai As String
Public strFax As String
Public strMail As String
Public strQHuyen As String
Public strTPho As String
' thong tin dai ly
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
Public spathVat As String
Public sSaiCT11 As String
Public sKyKeKhai As String
Public dNgayNopDB As Variant
Public HANNOP As String
Public TTHTK As String
Public strTenBpql As String
Private tempKyLb As Variant
Private tempNgayNop As Variant
Public hdrId As String
Public strMaCQT As String
Public strMaPQL As String
Public strTenPQL As String
Public strSoTTTKhai As String
Public isTKTonTai As Boolean
Public checkDLTTonTai As Boolean

Public Function Prepared1() As Boolean
    Dim ThangKK          As Integer
    Dim xmlDOMdata       As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim strDataFileName  As String
    Dim xmlNodeListItems As MSXML.IXMLDOMNodeList
    Dim strCboLyDo       As String
    Dim strCboLyDoAn     As String
    Dim xmlNode          As MSXML.IXMLDOMNode
    Dim fldList()        As String

    With fps
        .Sheet = 1
        '        .Col = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X)
        '        .Row = HEADER_SO_TT_TRONG_TEP_Y
        '        .CellType = CellTypeEdit
        '        .TypeEditCharSet = TypeEditCharSetNumeric
        '        .TypeMaxEditLen = 10

        'Ma so tep
        '        .Col = .ColLetterToNumber("M")
        '        .Row = 10
        '        .CellType = CellTypeEdit
        '        .TypeEditCharSet = TypeEditCharSetAlphanumeric
        '        .TypeMaxEditLen = 20
        
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
        
        'Phong quan ly
        .Col = .ColLetterToNumber("V")
        .Row = 34
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
        
        ' Ky lap bo
        SetDateFormat fps, 1, 30, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 32, .ColLetterToNumber("E"), DDMMYYYY
        .Row = 32
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 32, .ColLetterToNumber("M"), DDMMYYYY
        .Row = 32
        .Col = .ColLetterToNumber("M")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
        ' set gia tri gia han
        ' Ky dc gia han moi set gia tri ly do gia han
        strDataFileName = GetAbsolutePath("..\InterfaceTemplates\Catalogue_DM_Lydo_GH_TT16.xml")

        If xmlDOMdata.Load(strDataFileName) Then
            Set xmlNodeListItems = xmlDOMdata.getElementsByTagName("Item")

            For Each xmlNode In xmlNodeListItems
                fldList = Split(GetAttribute(xmlNode, "Value"), "###")
                strCboLyDo = strCboLyDo + fldList(1) + Chr$(9)
                strCboLyDoAn = strCboLyDoAn + fldList(0) + Chr$(9)
            Next

        End If

        'set ly do gia han
        .Row = 41
        .Col = .ColLetterToNumber("E")
        .TypeComboBoxList = strCboLyDo
        ' set Ma ly do gia han
        .Col = .ColLetterToNumber("AB")
        .TypeComboBoxList = strCboLyDoAn
        
    End With
    
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
    Dim i        As Integer
    Dim varMaDM  As Variant
    Dim LoaiKyKK As Variant
    Dim TenDM    As String
    i = 0
    ReDim Preserve larrPhongXuLy(0)
    larrPhongXuLy(0) = "00"
    ReDim Preserve larrid(0)
    larrid(0) = "00"
    
    With fps
        .Sheet = 1
        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y
        PHONG_XU_LY_X1 = PHONG_XU_LY_X
        PHONG_XU_LY_Y1 = PHONG_XU_LY_Y

        If rsPXL.Fields.count > 0 Then

            Do While Not rsPXL.EOF
                'dhdang sua combox ngay xu ly
                'ngay 05-08
                .TypeComboBoxIndex = -1
                .TypeComboBoxString = TAX_Utilities_Svr_New.Convert(rsPXL.Fields(1).Value, TCVN, UNICODE) + "     {" + TAX_Utilities_Svr_New.Convert(Trim(rsPXL.Fields(0).Value), TCVN, UNICODE) + "}"
                .TextTip = TextTipOff
                i = i + 1
                ReDim Preserve larrPhongXuLy(i)
                larrPhongXuLy(i) = TAX_Utilities_Svr_New.Convert(rsPXL.Fields(1).Value, TCVN, UNICODE)
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
        
        ' Loai ky ke khai
        .GetText .ColLetterToNumber("AB"), 42, LoaiKyKK

        If LoaiKyKK = "0" Then
            LOAI_TO_KHAI = "01/GTGT"
            isTkQuy = False
        Else
            LOAI_TO_KHAI = "01Q/GTGT"
            isTkQuy = True
        End If
        
        'get ten chi tieu trong danh muc NNKD
        .Sheet = 1
        .Col = .ColLetterToNumber("AB")
        .Row = 41
        .GetText .ColLetterToNumber("AB"), .Row, varMaDM

        If Trim(varMaDM) <> vbNullString Then
            .Col = .ColLetterToNumber("E")
            .Row = 41
            .Text = getTenLyDoGD(varMaDM)
        End If
        
        ' Neu khong co gia han se khong cho chon combobox
        ' Neu check gia han moi cho phep chon GT
        .Col = .ColLetterToNumber("E")
        .Row = 38

        If .Text = "1" Or UCase$(.Text) = "X" Then
            .Col = .ColLetterToNumber("E")
            .Row = 41
            .Lock = False
        Else
            .Col = .ColLetterToNumber("E")
            .Row = 41
            .Lock = True
        End If

        ' End set gia han
        
        ' Phu luc 05
        .Sheet = 7
        .Col = .ColLetterToNumber("F")
        i = 0

        Do
            .GetText .ColLetterToNumber("F"), i + 22, varMaDM
            GetTenCQT varMaDM, TenDM

            If Trim(TenDM) <> vbNullString Then
                .Col = .ColLetterToNumber("G")
                .Row = 22 + i
                .Text = TenDM
            End If

            UpdateCell fps, .Col, .Row, .Text
            i = i + 1
            .Col = .ColLetterToNumber("B")
            .Row = i + 22
        Loop Until .Text = "aa" Or varMaDM = ""
        
        ' Phu luc 06
        
        .Sheet = 8
        .Col = .ColLetterToNumber("E")
        i = 0

        Do
            .GetText .ColLetterToNumber("E"), 8 + i, varMaDM
            GetTenCQT varMaDM, TenDM

            If Trim(TenDM) <> vbNullString Then
                .Col = .ColLetterToNumber("F")
                .Row = 8 + i
                .Text = TenDM
            End If

            UpdateCell fps, .Col, .Row, .Text
            
            i = i + 1
            .Col = .ColLetterToNumber("B")
            .Row = i + 8
        Loop Until .Text = "aa" Or varMaDM = ""

    End With
    
    Prepared2 = True
End Function
Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

Public Function Prepared3() As Boolean
    With fps
        .Sheet = 1
        
        'Set NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
        .Row = NGAY_NHAN_TO_KHAI_Y

        If strNgayNhanToKhai <> "" Then
            .Text = strNgayNhanToKhai
            UpdateCell fps, .Col, .Row, .Text
        End If

        '**************************************************************

        '        set Phong xu ly
        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y

        If strPhongXuLy <> "" Then
            .Text = strPhongXuLy
            UpdateCell fps, .Col, .Row, .Text
        End If

        .Col = .ColLetterToNumber(MA_SO_THUE_X)
        .Row = MA_SO_THUE_Y

        If strMST <> "" Then
            .Text = strMST
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(KY_LAP_BO_X)
        .Row = KY_LAP_BO_Y
 
        If Len(Month(Now())) = 1 Then
            .Text = "0" & Month(Now()) & "/" & Year(Now())
            sKyKeKhai = "0" & Month(Now()) & Year(Now())
        Else
            .Text = Month(Now()) & "/" & Year(Now())
            sKyKeKhai = Month(Now()) & Year(Now())
        End If

        UpdateCell fps, .Col, .Row, .Text

        .Col = .ColLetterToNumber(NGAY_NOP_X)
        .Row = NGAY_NOP_Y
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        UpdateCell fps, .Col, .Row, .Text
        .Lock = False

        .Col = .ColLetterToNumber(NGAY_QUET_X)
        .Row = NGAY_QUET_Y
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        UpdateCell fps, .Col, .Row, .Text

        .Col = .ColLetterToNumber(TEN_GOI_X)
        .Row = TEN_GOI_Y

        If strTenGoi <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenGoi, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If

        .Col = .ColLetterToNumber(DIA_CHI_X)
        .Row = DIA_CHI_Y

        If strDchi <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strDchi, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If

        .Col = .ColLetterToNumber(DIEN_THOAI_X)
        .Row = DIEN_THOAI_Y

        If strDThoai <> "" Then
            .Text = strDThoai
            UpdateCell fps, .Col, .Row, .Text
        End If

        .Col = .ColLetterToNumber(FAX_X)
        .Row = FAX_Y

        If strFax <> "" Then
            .Text = strFax
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        ' Lay thong tin dai ly
        
        .Col = .ColLetterToNumber(MA_SO_THUE_DL_X)
        .Row = MA_SO_THUE_DL_Y

        If strMST_DLT <> "" Then
            .Text = strMST_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(TEN_GOI_DL_X)
        .Row = TEN_GOI_DL_Y

        If strTen_DLT <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strTen_DLT, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(DIA_CHI_DL_X)
        .Row = DIA_CHI_DL_Y

        If strDchi_DLT <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strDchi_DLT, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(QUAN_HUYEN_DL_X)
        .Row = QUAN_HUYEN_DL_Y

        If strQHuyen_DLT <> "" Then
            .Text = strQHuyen_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(TINH_TPHO_DL_X)
        .Row = TINH_TPHO_DL_Y

        If strTTPho_DLT <> "" Then
            .Text = strTTPho_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(DIEN_THOAI_DL_X)
        .Row = DIEN_THOAI_DL_Y

        If strDthoai_DLT <> "" Then
            .Text = strDthoai_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(FAX_DL_X)
        .Row = FAX_DL_Y

        If strFax_DLT <> "" Then
            .Text = strFax_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(MAIL_DL_X)
        .Row = MAIL_DL_Y

        If strMail_DLT <> "" Then
            .Text = strMail_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(NGAY_HDDL_DL_X)
        .Row = NGAY_HDDL_DL_Y

        If strNgayHD_DLT <> "" Then
            .Text = strNgayHD_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(SO_HDDL_DL_X)
        .Row = SO_HDDL_DL_Y

        If strSoHD_DLT <> "" Then
            .Text = strSoHD_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If

        'set check loi dinh danh
        If checkDLTTonTai = False Then
            .Col = .ColLetterToNumber("E")
            .Row = 36
            .Value = 1
        End If
        
        'so thu tu to khai
        .Col = .ColLetterToNumber("M")
        .Row = 30
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        
        ''        .Col = .ColLetterToNumber(MA_BPQL_X)
        ''        .Row = MA_BPQL_Y
        ''        If strMaBPQL <> "" Then
        ''            .Text = strMaBPQL
        ''            UpdateCell fps, .Col, .Row, .Text
        ''        End If
        ''
        ''        .Col = .ColLetterToNumber(TEN_BPQL_X)
        ''        .Row = TEN_BPQL_Y
        ''        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        
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
        .Col = .ColLetterToNumber(KY_LAP_BO_X)
        .Row = KY_LAP_BO_Y
        strKyLapBo = .Text
        'Hlnam sua tao bang de ghi du lieu de ghi vao VAT theo Ky Lap Bo (Tren to khai)
        'Ngay 14/07/2010
        'Begin
'        If Len(Month(strKyLapBo)) = 1 Then
'            sKyKeKhai = "0" & Month(strKyLapBo) & Year(strKyLapBo)
'        Else
'            sKyKeKhai = Month(strKyLapBo) & Year(strKyLapBo)
'        End If
        'End
        
        ' Get Phong xu ly
        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y
        strPhongXuLy = .Text
        
        ' Get NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
        .Row = NGAY_NHAN_TO_KHAI_Y
        strNgayNhanToKhai = .Text
        
        'Go to last sheet (header sheet)
        .Sheet = .SheetCount
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
            If Col = fps.ColLetterToNumber("E") And Row = 38 Then
                fps.Col = fps.ColLetterToNumber("E")
                fps.Row = 38
                If fps.Text = "1" Or UCase$(fps.Text) = "X" Then
                    fps.Col = fps.ColLetterToNumber("E")
                    fps.Row = 41
                    fps.Lock = False
                Else
                    fps.Col = fps.ColLetterToNumber("E")
                    fps.Row = 41
                    fps.Lock = True
                    fps.SetText fps.ColLetterToNumber("E"), 41, ""
                    fps.SetText fps.ColLetterToNumber("AB"), 41, "03"
                    UpdateCell fps, fps.ColLetterToNumber("AB"), 41, "03"
                    UpdateCell fps, fps.ColLetterToNumber("E"), 41, ""
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

'Private Sub fps_KeyPress(KeyAscii As Integer)
'    With fps
'        If .ActiveSheet = 1 And .ActiveCol = .ColLetterToNumber("M") _
'           And .ActiveRow = 14 And KeyAscii = 13 Then
'                .SetActiveCell .ColLetterToNumber("E"), 14
'        End If
'    End With
'End Sub

Public Function CheckValidData() As Boolean
    Dim varCheckValue As Variant, varNoteValue As Variant
    
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 36, varCheckValue
        .GetText .ColLetterToNumber("M"), 36, varNoteValue
        If CStr(varCheckValue) = "1" And Trim(CStr(varNoteValue)) = "" Then
            .Sheet = .SheetCount
            .SetText 2, 14, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText 2, 14, "1"
        End If

    End With
End Function

Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
    KiemTraKhoaSo = True
    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("E_30"), "Value"), "dd/mm/yyyy")
    
    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function

Public Function InsertDTL_PL1() As String
    Dim sSQL          As String
    Dim sSQLCol       As String
    Dim sSQLVal       As String
    Dim bln           As Boolean
    Dim madtnt        As Variant
    Dim ngnop         As Variant
    Dim KYLBO         As Variant
    Dim kykkhai       As Variant
    Dim KYTTBS        As Variant
    Dim matkhai       As Variant
    Dim LOAIHD        As Variant
    Dim KYHIEU        As Variant
    Dim SOHDON        As Variant
    Dim NGAYPH        As Variant
    Dim HANG          As Variant
    Dim NGMUA         As Variant
    Dim MST           As Variant
    Dim DSOCHUA       As Variant
    Dim THUESUAT      As Variant
    Dim THUEGTGT      As Variant
    Dim mact          As Variant
    Dim STT           As Variant
    Dim STT2          As Variant
    Dim STTIN         As Variant
    Dim GhiChu        As Variant
    Dim MAPL          As Variant
    Dim i             As Integer, iCol As Long, iRow As Long
    Dim xmlNode       As MSXML.IXMLDOMNode
    Dim strFileNameP1 As String
    Dim fso           As New FileSystemObject
    Dim DANHAN        As Variant
    Dim MAUHD         As Variant
    'kiem tra ton tai tep *.dbf chua
    '    strFileNameP1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P1" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileNameP1) = False Then
    '    Dim strTepmauP1 As String
    '        strTepmauP1 = spathVat & "\tepmau\P1.DBF"
    '        fso.CopyFile strTepmauP1, strFileNameP1, False
    '    End If
   
    sSQLCol = "MATKHAI, KYLBO, KYKKHAI, MADTNT, TTHTK, NGNOP, LOAIHD, KYHIEU, SOHDON, NGAYPH, HANG, NGMUA," & "MST, DSOCHUA, THUESUAT, THUEGTGT, MACT, STT, STT2, STTIN, GHICHU, MAPL,LAN_QUET, DANHAN, MAUHD"
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
                
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        .Sheet = 2
            
        'insert 'Hang hoa khong chiu thue'
            
        LOAIHD = "''"
        KYHIEU = "''"
        SOHDON = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0097" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'1001'"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        GhiChu = "''"
        MAPL = "'PL01GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = 25

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("H"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("K"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If

            THUESUAT = 0
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 1
            STT2 = i
            STTIN = "'I'"
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL01GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                 
            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"
             
        'insert HHDV chiu thue 0%
        LOAIHD = "''"
        SOHDON = "''"
        KYHIEU = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0098" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'1002'"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        GhiChu = "''"
        MAPL = "'PL01GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = .Row + 4

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("H"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("K"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If

            THUESUAT = 0
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If
            mact = "''"
            STT = 2
            STT2 = i
            STTIN = "'II'"
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL01GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                           
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                
            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "bb"
             
        'insert HHDV chiu thue 5%
        LOAIHD = "''"
        SOHDON = "''"
        KYHIEU = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0099" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'1003'"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        GhiChu = "''"
        MAPL = "'PL01GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = .Row + 4

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("H"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("K"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If

            THUESUAT = 0
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If
            mact = "''"
            STT = 3
            STT2 = i
            STTIN = "'III'"
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL01GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                           
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "cc"
             
        'insert HHDV chiu thue 10%
        LOAIHD = "''"
        SOHDON = "''"
        KYHIEU = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0100" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'1004'"
        STT = 4
        STT2 = 0
        STTIN = "'IV'"
        GhiChu = "''"
        MAPL = "'PL01GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = .Row + 4

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("H"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("K"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If

            THUESUAT = 0
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If
            
            mact = "''"
            STT = 4
            STT2 = i
            STTIN = "'IV'"
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL01GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                           
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "dd"
            
        'insert HHDV không phai tong hop trên to khai 01/GTGT:
        LOAIHD = "''"
        SOHDON = "''"
        KYHIEU = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0138" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'1005'"
        STT = 5
        STT2 = 0
        STTIN = "'V'"
        GhiChu = "''"
        MAPL = "'PL01GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = .Row + 4

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("H"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("K"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If

            THUESUAT = 0
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If
            mact = "''"
            STT = 5
            STT2 = i
            STTIN = "'V'"
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL01GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                           
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "ee"

    End With
   
    clsDAO.Disconnect
    InsertDTL_PL1 = vbNullString
   
End Function

Public Function InsertDTL_PL2() As String
    Dim sSQL          As String
    Dim sSQLCol       As String
    Dim sSQLVal       As String
    Dim bln           As Boolean
    Dim madtnt        As Variant
    Dim ngnop         As Variant
    Dim KYLBO         As Variant
    Dim kykkhai       As Variant
    Dim KYTTBS        As Variant
    Dim matkhai       As Variant
    Dim LOAIHD        As Variant
    Dim KYHIEU        As Variant
    Dim SOHDON        As Variant
    Dim NGAYPH        As Variant
    Dim HANG          As Variant
    Dim NGMUA         As Variant
    Dim MST           As Variant
    Dim DSOCHUA       As Variant
    Dim THUESUAT      As Variant
    Dim THUEGTGT      As Variant
    Dim mact          As Variant
    Dim STT           As Variant
    Dim STT2          As Variant
    Dim STTIN         As Variant
    Dim GhiChu        As Variant
    Dim MAPL          As Variant
    Dim i             As Integer, iCol As Long, iRow As Long
    Dim xmlNode       As MSXML.IXMLDOMNode
    Dim DANHAN        As Variant
    Dim strFileNameP1 As String
    Dim fso           As New FileSystemObject
    Dim MAUHD         As Variant
    'kiem tra ton tai tep *.dbf chua
    '    strFileNameP1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P1" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileNameP1) = False Then
    '    Dim strTepmauP1 As String
    '        strTepmauP1 = spathVat & "\tepmau\P1.DBF"
    '        fso.CopyFile strTepmauP1, strFileNameP1, False
    '    End If
   
    sSQLCol = "MATKHAI, KYLBO, KYKKHAI, MADTNT, TTHTK, NGNOP, LOAIHD, KYHIEU, SOHDON, NGAYPH, HANG, NGMUA," & "MST, DSOCHUA, THUESUAT, THUEGTGT, MACT, STT, STT2, STTIN, GHICHU, MAPL , LAN_QUET, DANHAN, MAUHD"
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
                
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        .Sheet = 3
            
        'insert HHDV dung rieng SXKD chiu thue GTGT
            
        LOAIHD = "''"
        KYHIEU = "''"
        SOHDON = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0101" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2001'"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        GhiChu = "''"
        MAPL = "'PL02GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = 25

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("H"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("K"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            'THUESUAT = 0
                
            .GetText .ColLetterToNumber("M"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 1
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"
             
        'insert HHDV dung rieng SXKD khong chiu thue GTGT
        LOAIHD = "''"
        SOHDON = "''"
        KYHIEU = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0102" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2002'"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        GhiChu = "''"
        MAPL = "'PL02GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = .Row + 4

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("H"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("K"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            'THUESUAT = 0
                
            .GetText .ColLetterToNumber("M"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 2
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                           
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "bb"
             
        'insert HHDV dung chung SXKD chiu, khong chiu thue GTGT
        LOAIHD = "''"
        SOHDON = "''"
        KYHIEU = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0103" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2003'"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        GhiChu = "''"
        MAPL = "'PL02GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = .Row + 4

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("H"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("K"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            'THUESUAT = 0
                
            .GetText .ColLetterToNumber("M"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If
            mact = "''"
            STT = 3
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                           
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "cc"
             
        'insert Hàng HDV dùng cho du án dauu tu du dieu kien khau tru thue

        LOAIHD = "''"
        SOHDON = "''"
        KYHIEU = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0139" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2004'"
        STT = 4
        STT2 = 0
        STTIN = "'IV'"
        GhiChu = "''"
        MAPL = "'PL02GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = .Row + 4

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("H"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("K"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            'THUESUAT = 0
                
            .GetText .ColLetterToNumber("M"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 3
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                           
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "dd"
             
        'insert HHDV khong phai tong hop tren TK 01/GTGT

        LOAIHD = "''"
        SOHDON = "''"
        KYHIEU = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0138" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2005'"
        STT = 5
        STT2 = 0
        STTIN = "'V'"
        GhiChu = "''"
        MAPL = "'PL02GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = .Row + 4

        Do
            LOAIHD = "''"
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("H"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("K"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            'THUESUAT = 0
                
            .GetText .ColLetterToNumber("M"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 5
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                           
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "ee"
             
    End With
   
    clsDAO.Disconnect
    InsertDTL_PL2 = vbNullString
   
End Function

Public Function InsertDTL_PL017() As String
    Dim sSQL        As String
    Dim sSQLCol     As String
    Dim sSQLVal     As String
    Dim bln         As Boolean
   
    Dim madtnt      As Variant
    Dim ngnop       As Variant
    Dim KYLBO       As Variant
    Dim kykkhai     As Variant
    Dim matkhai     As Variant
    Dim TTHKT       As Variant
    Dim LOAIXE      As Variant
    Dim DVTinh      As Variant
    Dim TONGBANRA   As Variant
    Dim BANNTDDL    As Variant
    Dim GIAHD       As Variant
    Dim GhiChu      As Variant
    Dim mact        As Variant
    Dim STT         As Variant
    Dim STT2        As Variant
    Dim STTIN       As Variant
    Dim MAPL        As Variant
    Dim i           As Integer
    Dim DANHAN      As Variant
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim strFileName As String
    Dim fso         As New FileSystemObject
    
    'kiem tra ton tai tep *.dbf chua
    '    strFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P3" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileName) = False Then
    '    Dim strTepmau As String
    '        strTepmau = spathVat & "\tepmau\P3.DBF"
    '        fso.CopyFile strTepmau, strFileName, False
    '    End If
   
    sSQLCol = "MATKHAI , KYLBO , KYKKHAI , MADTNT , TTHTK  , NGNOP,LOAIXE ,DVTINH,TONGBANRA,BANNTDDL,GIAHD,GHICHU,MACT,STT,STT2,STTIN,MAPL,DANHAN,LAN_QUET"
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        mact = "'1001'"
        STT = 1
        STT2 = 0
        STTIN = "'A'"
        MAPL = "'PL07GTKT'"
        DANHAN = "''"
        DVTinh = "''"
        TONGBANRA = "0"
        BANNTDDL = "0"
        GIAHD = "0"
        GhiChu = "''"
        '        TTHTK = "'NULL'"
        
        'insert chi tieu tinh

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0167" Then
                LOAIXE = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIXE & ","
        sSQLVal = sSQLVal & DVTinh & "," & TONGBANRA & "," & BANNTDDL & "," & GIAHD & "," & GhiChu & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & MAPL & "," & DANHAN & "," & LAN_QUET
             
        sSQL = "INSERT INTO TMP_P017" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        .Sheet = 9
        i = 0
        .Col = .ColLetterToNumber("B")
        .Row = 9
        STT2 = 1
        mact = "''"
        Do
            STTIN = "'" & CStr(STT2) & "'"
            'get loai xe
            .Col = .ColLetterToNumber("C")
            LOAIXE = .Text

            If Trim(LOAIXE) = vbNullString Then
                LOAIXE = "''"
            Else
                LOAIXE = "'" & LOAIXE & "'"
            End If

            'get Don vi tinh
            .Col = .ColLetterToNumber("D")
            DVTinh = .Text

            If Trim(DVTinh) = vbNullString Then
                DVTinh = "''"
            Else
                DVTinh = "'" & Trim(TAX_Utilities_Svr_New.Convert(CStr(DVTinh), UNICODE, TCVN)) & "'"
            End If

            'get Tong so luong xe ban ra
            .Col = .ColLetterToNumber("E")
            TONGBANRA = .Text

            If TONGBANRA = vbNullString Then
                TONGBANRA = "0"
            Else
                TONGBANRA = Replace(TONGBANRA, ".", "")
            End If

            'get so luong ban cho nguoi tieu dung
            .Col = .ColLetterToNumber("F")
            BANNTDDL = .Text

            If BANNTDDL = vbNullString Then
                BANNTDDL = "0"
            Else
                BANNTDDL = Replace(BANNTDDL, ".", "")
            End If

            'get gia hoa don
            .Col = .ColLetterToNumber("G")
            GIAHD = .Text

            If GIAHD = vbNullString Then
                GIAHD = "0"
            Else
                GIAHD = Replace(GIAHD, ".", "")
            End If

            ' get ghi chu
            .Col = .ColLetterToNumber("H")
            GhiChu = .Text

            If GhiChu = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & GhiChu & "'"
            End If

            ' sSQLCol = "MATKHAI , KYLBO , KYKKHAI , MADTNT , TTHKT , NGNOP,LOAIXE ,DVTINH,TONGBANRA,BANNTDDL,GIAHD,GHICHU,MACT,STT,STT2,STTIN,MAPL"

            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIXE & ","
            sSQLVal = sSQLVal & DVTinh & "," & TONGBANRA & "," & BANNTDDL & "," & GIAHD & "," & GhiChu & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & MAPL & "," & DANHAN & "," & LAN_QUET
             
            sSQL = "INSERT INTO TMP_P017" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
            STT2 = STT2 + 1
        Loop Until .Text = "aa"
             
        ' xu ly chuyen muc 2
        .Row = .Row + 4
        mact = "'1002'"
        STT = 2
        STT2 = 0
        STTIN = "'B'"
        MAPL = "'PL07GTKT'"
        DVTinh = "''"
        TONGBANRA = "0"
        BANNTDDL = "0"
        GIAHD = "0"
        GhiChu = "''"
        'insert chi tieu tinh
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0168" Then
                LOAIXE = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIXE & ","
        sSQLVal = sSQLVal & DVTinh & "," & TONGBANRA & "," & BANNTDDL & "," & GIAHD & "," & GhiChu & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & MAPL & "," & DANHAN & "," & LAN_QUET
             
        sSQL = "INSERT INTO TMP_P017" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        .Sheet = 9
        .Col = .ColLetterToNumber("B")
        STT2 = 1
        mact = "''"
        Do
            STTIN = "'" & CStr(STT2) & "'"
            'get loai xe
            .Col = .ColLetterToNumber("C")
            LOAIXE = .Text

            If Trim(LOAIXE) = vbNullString Then
                LOAIXE = "''"
            Else
                LOAIXE = "'" & LOAIXE & "'"
            End If

            'get Don vi tinh
            .Col = .ColLetterToNumber("D")
            DVTinh = .Text

            If Trim(DVTinh) = vbNullString Then
                DVTinh = "''"
            Else
                DVTinh = "'" & Trim(TAX_Utilities_Svr_New.Convert(CStr(DVTinh), UNICODE, TCVN)) & "'"
            End If

            'get Tong so luong xe ban ra
            .Col = .ColLetterToNumber("E")
            TONGBANRA = .Text

            If TONGBANRA = vbNullString Then
                TONGBANRA = "0"
            Else
                TONGBANRA = Replace(TONGBANRA, ".", "")
            End If

            'get so luong ban cho nguoi tieu dung
            .Col = .ColLetterToNumber("F")
            BANNTDDL = .Text

            If BANNTDDL = vbNullString Then
                BANNTDDL = "0"
            Else
                BANNTDDL = Replace(BANNTDDL, ".", "")
            End If

            'get gia hoa don
            .Col = .ColLetterToNumber("G")
            GIAHD = .Text

            If GIAHD = vbNullString Then
                GIAHD = "0"
            Else
                GIAHD = Replace(GIAHD, ".", "")
            End If

            ' get ghi chu
            .Col = .ColLetterToNumber("H")
            GhiChu = .Text

            If GhiChu = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & GhiChu & "'"
            End If
                
            ' sSQLCol = "MATKHAI , KYLBO , KYKKHAI , MADTNT , TTHKT , NGNOP,LOAIXE ,DVTINH,TONGBANRA,BANNTDDL,GIAHD,GHICHU,MACT,STT,STT2,STTIN,MAPL"

            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIXE & ","
            sSQLVal = sSQLVal & DVTinh & "," & TONGBANRA & "," & BANNTDDL & "," & GIAHD & "," & GhiChu & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & MAPL & "," & DANHAN & "," & LAN_QUET
             
            sSQL = "INSERT INTO TMP_P017" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
            STT2 = STT2 + 1
        Loop Until .Text = "bb"

    End With
   
    clsDAO.Disconnect
    InsertDTL_PL017 = vbNullString
   
End Function

Public Function InsertDTL_PL4A() As String
    Dim sSQL    As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim bln     As Boolean
   
    Dim madtnt  As Variant
    Dim ngnop   As Variant
    Dim KYLBO   As Variant
    Dim kykkhai As Variant
    Dim matkhai As Variant
  
    Dim MATHUE  As Variant
    Dim MaMuc   As Variant
    Dim maTM    As Variant
    Dim MAPP    As Variant
    '  Dim TTHTK As Variant
  
    Dim CTTN    As Variant
    Dim TIEN    As Variant
    Dim DANHAN  As Variant
    Dim SOTIEN2 As Variant
    Dim tt      As Variant
    sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN,MATHUE,MAMUC,MATM,MAPP,LAN_QUET,TTHTK,SOTIEN2,TT"
  
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If

    MATHUE = "''"
    MaMuc = "''"
    maTM = "''"
    MAPP = "''"

    ' TTHTK = "''"
    '--------------------
  
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'001'"
        .Sheet = 5
        .Col = .ColLetterToNumber("D")
        .Row = 4
        TIEN = .Text

        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
        End If

        SOTIEN2 = TIEN
        tt = "'01'"
        '  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P4" & sKyKeKhai & " ( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
    End With
  
    '--------------------
  
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'003'"
        .Sheet = 5
        .Col = .ColLetterToNumber("D")
        .Row = 6
        TIEN = .Text

        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
        End If

        SOTIEN2 = TIEN
        tt = "'02'"

        '  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P4" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
    End With
  
    '--------------------
  
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'004'"
        .Sheet = 5
        .Col = .ColLetterToNumber("D")
        .Row = 8
        TIEN = .Text

        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
        End If

        SOTIEN2 = TIEN
        tt = "'03'"

        '  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P4" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
    End With
  
    '--------------------
  
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'005'"
        .Sheet = 5
        .Col = .ColLetterToNumber("D")
        .Row = 10
        TIEN = .Text

        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
        End If

        SOTIEN2 = TIEN
        tt = "'04'"

        '  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P4" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
    End With
  
    '--------------------
  
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'007'"
        .Sheet = 5
        .Col = .ColLetterToNumber("D")
        .Row = 15
        TIEN = .Text

        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
        End If

        SOTIEN2 = TIEN
        tt = "'05'"

        '  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P4" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
    End With
  
    '--------------------
  
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'008'"
        .Sheet = 5
        .Col = .ColLetterToNumber("D")
        .Row = 16
        TIEN = .Text

        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
        End If

        SOTIEN2 = TIEN
        tt = "'06'"

        '  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P4" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
    End With
  
    '-----------------------------------

    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'009'"
        .Sheet = 5
        .Col = .ColLetterToNumber("D")
        .Row = 17
        TIEN = .Text

        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
            TIEN = Round(Replace(TIEN, ",", "."))
        End If

        SOTIEN2 = TIEN
        tt = "'07'"
        '  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P4" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
    End With
  
    '--------------------
  
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'010'"
        .Sheet = 5
        .Col = .ColLetterToNumber("D")
        .Row = 18
        TIEN = .Text

        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
        End If

        SOTIEN2 = TIEN
        tt = "'08'"
        '  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P4" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
    End With
  
    '--------------------
  
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'011'"
        .Sheet = 5
        .Col = .ColLetterToNumber("D")
        .Row = 19
        TIEN = .Text

        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
        End If

        SOTIEN2 = TIEN
        tt = "'09'"
        '  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P4" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
    End With
  
    clsDAO.Disconnect
   
    InsertDTL_PL4A = vbNullString
  
End Function

Public Function InsertDTL_PL4B() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim bln  As Boolean
   
  Dim madtnt As Variant
  Dim ngnop As Variant
  Dim KYLBO As Variant
  Dim kykkhai As Variant
  Dim matkhai As Variant
  
  Dim MATHUE As Variant
  Dim MaMuc As Variant
  Dim maTM As Variant
  
  Dim CTTN As Variant
  Dim TIEN As Variant
  Dim DANHAN As Variant
      Dim SOTIEN2 As Variant
    Dim tt      As Variant
  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN,MATHUE,MAMUC,MATM,LAN_QUET,TTHTK,SOTIEN2,TT"
  
  If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
  End If
  MATHUE = "''"
  MaMuc = "''"
  maTM = "''"
 ' TTHTK = "''"
'--------------------
  
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
                
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'001'"
        .Sheet = 6
        .Col = .ColLetterToNumber("E")
        .Row = 7
        TIEN = .Text
        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
            TIEN = Val(CStr(TIEN))
        End If
        SOTIEN2 = TIEN
        tt = "'01'"
'  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P5" & sKyKeKhai & " ( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
  End With
  
'--------------------
  
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
                
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'003'"
        .Sheet = 6
        .Col = .ColLetterToNumber("E")
        .Row = 9
        TIEN = .Text
        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
            TIEN = Val(CStr(TIEN))
        End If
                SOTIEN2 = TIEN
        tt = "'02'"
'  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P5" & sKyKeKhai & " ( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
  End With
  
  '--------------------
  
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
                
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'004'"
        .Sheet = 6
        .Col = .ColLetterToNumber("E")
        .Row = 11
        TIEN = .Text
        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
            TIEN = Val(CStr(TIEN))
        End If
                SOTIEN2 = TIEN
        tt = "'03'"
'  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P5" & sKyKeKhai & " ( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
  End With
  
  '--------------------
  
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
                
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'005'"
        .Sheet = 6
        .Col = .ColLetterToNumber("E")
        .Row = 13
        TIEN = .Text
        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
            TIEN = Val(CStr(TIEN))
        End If
                SOTIEN2 = TIEN
        tt = "'04'"
'  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P5" & sKyKeKhai & " ( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
  End With
   
   '--------------------
  
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
                
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'007'"
        .Sheet = 6
        .Col = .ColLetterToNumber("E")
        .Row = 18
        TIEN = .Text
        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
            TIEN = Val(CStr(TIEN))
        End If
                SOTIEN2 = TIEN
        tt = "'05'"
'  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P5" & sKyKeKhai & " ( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
  End With
   '--------------------
  
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
                
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'008'"
        .Sheet = 6
        .Col = .ColLetterToNumber("E")
        .Row = 19
        TIEN = .Text
        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
            TIEN = Val(CStr(TIEN))
        End If
                SOTIEN2 = TIEN
        tt = "'06'"
'  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P5" & sKyKeKhai & " ( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
  End With
  
  '--------------------
  
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
                
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'009'"
        .Sheet = 6
        .Col = .ColLetterToNumber("E")
        .Row = 20
        TIEN = .Text
        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
            TIEN = Val(CStr(TIEN))
        End If
                SOTIEN2 = TIEN
        tt = "'07'"
'  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P5" & sKyKeKhai & " ( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
  End With
  
  
  '--------------------
  
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
                
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'010'"
        .Sheet = 6
        .Col = .ColLetterToNumber("E")
        .Row = 21
        TIEN = .Text
        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
            TIEN = Val(CStr(TIEN))
        End If
                SOTIEN2 = TIEN
        tt = "'08'"
'  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P5" & sKyKeKhai & " ( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
  End With
  
  
  '--------------------
  
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
                
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'011'"
        .Sheet = 6
        .Col = .ColLetterToNumber("E")
        .Row = 22
        TIEN = .Text
        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
            TIEN = Val(CStr(TIEN))
        End If
                SOTIEN2 = TIEN
        tt = "'09'"
'  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P5" & sKyKeKhai & " ( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
  End With
  
  '--------------------
  
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
                
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'012'"
        .Sheet = 6
        .Col = .ColLetterToNumber("E")
        .Row = 23
        TIEN = .Text
        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
            TIEN = Val(CStr(TIEN))
        End If
                SOTIEN2 = TIEN
        tt = "'10'"
'  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P5" & sKyKeKhai & " ( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
  End With
  '--------------------
  
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
                
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        CTTN = "'013'"
        .Sheet = 6
        .Col = .ColLetterToNumber("E")
        .Row = 24
        TIEN = .Text
        If TIEN = vbNullString Then
            TIEN = "0"
        Else
            TIEN = Replace(TIEN, ".", "")
            TIEN = Val(CStr(TIEN))
        End If
                SOTIEN2 = TIEN
        tt = "'11'"
'  sSQLCol = "MADTNT,KYLBO,NGNOP,MATKHAI,KYKKHAI,CTTN,TIEN,DANHAN"
        sSQLVal = madtnt & "," & KYLBO & "," & ngnop & "," & matkhai & "," & kykkhai & "," & CTTN & "," & TIEN & "," & DANHAN & "," & MATHUE & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & TTHTK & "," & SOTIEN2 & "," & tt
        
        sSQL = "INSERT INTO TMP_P5" & sKyKeKhai & " ( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
  End With
        clsDAO.Disconnect
        InsertDTL_PL4B = vbNullString
End Function

Public Function InsertDTL_PL5() As String
    Dim sSQL      As String
    Dim sSQLCol   As String
    Dim sSQLVal   As String
    Dim bln       As Boolean
   
    Dim madtnt    As Variant
    Dim ngnop     As Variant
    Dim KYLBO     As Variant
  Dim kykkhai As Variant
  Dim matkhai As Variant
    Dim SOCT      As Variant
    Dim NGNOPTHUE As Variant
    Dim NOINOP    As Variant
    Dim CQUANQL   As Variant
    Dim THUEDN    As Variant
    Dim DANHAN    As Variant
   
    sSQLCol = "MADTNT, NGNOP, TTHTK, KYLBO, KYKKHAI, MATKHAI, SOCT, NGNOPTHUE, NOINOP, CQUANQL, THUEDN, LAN_QUET, DANHAN"
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        .Sheet = 7
        .Col = .ColLetterToNumber("B")
        .Row = 22

        Do
                                       
            .Col = .ColLetterToNumber("D")
            NGNOPTHUE = .Text

            If Trim(NGNOPTHUE) = vbNullString Then
                NGNOPTHUE = "CTOD('')"
            Else
                NGNOPTHUE = ToDate(Trim(NGNOPTHUE), DDMMYYYY)
                NGNOPTHUE = "CTOD('" & Format(NGNOPTHUE, "mm/dd/yyyy") & "')"
            End If
                                                    
            .Col = .ColLetterToNumber("F")
            CQUANQL = .Text

            If Trim(CQUANQL) = vbNullString Then
                CQUANQL = "''"
            Else
                CQUANQL = "'" & TAX_Utilities_Svr_New.Convert(CStr(CQUANQL), UNICODE, TCVN) & "'"
            End If
                       
            .Col = .ColLetterToNumber("H")
            THUEDN = .Value

            If Trim(THUEDN) = vbNullString Then
                THUEDN = "0"
            End If
                        
            .Col = .ColLetterToNumber("E")
            NOINOP = .Text
            If Trim(NOINOP) = vbNullString Then
                NOINOP = "''"
            Else
                NOINOP = "'" & TAX_Utilities_Svr_New.Convert(CStr(NOINOP), UNICODE, TCVN) & "'"
            End If
                        
            .Col = .ColLetterToNumber("C")
            SOCT = .Text
            If Trim(SOCT) = vbNullString Then
                SOCT = "''"
            Else
                SOCT = "'" & TAX_Utilities_Svr_New.Convert(CStr(SOCT), UNICODE, TCVN) & "'"
            End If
                        
            sSQLVal = madtnt & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & kykkhai & "," & matkhai & "," & SOCT & "," & NGNOPTHUE & "," & NOINOP & "," & CQUANQL & "," & THUEDN & "," & LAN_QUET & "," & DANHAN
                       
            sSQL = "INSERT INTO TMP_P6" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)

            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
             
    End With

    clsDAO.Disconnect
   
    InsertDTL_PL5 = vbNullString
   
End Function

Public Function InsertDTL_PL3() As String
    Dim sSQL    As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim bln     As Boolean
   
    Dim STT     As Variant
    Dim SOCT    As Variant
    Dim Cot2    As Variant
    Dim Cot3    As Variant
    Dim cot4    As Variant
    Dim cot5    As Variant
    Dim Cot6    As Variant
    Dim Cot7    As Variant
    Dim Cot8    As Variant
    Dim Cot9    As Variant
    Dim Cot10   As Variant
    Dim Cot11   As Variant
    Dim Cot12   As Variant
    Dim Cot13   As Variant
    Dim Cot14   As Variant
    Dim Cot15   As Variant
    Dim Cot16   As Variant
    Dim Cot17   As Variant
    Dim Cot18   As Variant
    Dim Cot19   As Variant
    Dim Cot20   As Variant
    Dim Cot21   As Variant
    Dim Cot22   As Variant
    Dim Cot23   As Variant
    Dim Cot24   As Variant
    Dim Cot25   As Variant
    Dim Cot26   As Variant
    Dim Cot27   As Variant
    Dim Cot28   As Variant
    Dim Cot29   As Variant
    Dim Cot30   As Variant
    Dim Cot31   As Variant
    Dim Cot32   As Variant
    Dim Cot33   As Variant
    
    Dim madtnt  As Variant
    Dim ngnop   As Variant
    Dim KYLBO   As Variant
    Dim kykkhai As Variant
    Dim matkhai As Variant
    Dim DANHAN  As Variant
  
    sSQLCol = "MADTNT,NGNOP,TTHTK,KYLBO,KYKKHAI,MATKHAI,SOCT,STT,Cot2,Cot3,Cot4,Cot5,Cot6,Cot7,Cot8,Cot9,Cot10,Cot11,Cot12,Cot13,Cot14,Cot15,Cot16,Cot17,Cot18,Cot19,Cot20,Cot21,Cot22,Cot23,Cot24,Cot25,Cot26,Cot27,Cot28,Cot29,Cot30,Cot31,Cot32,Cot33,LAN_QUET,DANHAN"

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        
        .Sheet = 4
        .Row = 10
        .Col = .ColLetterToNumber("B")
        SOCT = "''"

        Do
            .GetText .ColLetterToNumber("B"), .Row, STT

            If Trim(STT) = vbNullString Then
                STT = "''"
            Else
                STT = "'" & STT & "'"
            End If

            .GetText .ColLetterToNumber("C"), .Row, Cot2

            If Trim(Cot2) = vbNullString Then
                Cot2 = "''"
            Else
                Cot2 = "'" & Cot2 & "'"
            End If

            .GetText .ColLetterToNumber("D"), .Row, Cot3

            If Trim(Cot3) = vbNullString Then
                Cot3 = "CTOD('')"
            Else
                Cot3 = ToDate(Trim(Cot3), DDMMYYYY)
                Cot3 = "CTOD('" & Format(Cot3, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("E"), .Row, cot4

            If Trim(cot4) = vbNullString Then
                cot4 = "''"
            Else
                cot4 = "'" & cot4 & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, cot5

            If Trim(cot5) = vbNullString Then
                cot5 = "0"
            End If

            .GetText .ColLetterToNumber("G"), .Row, Cot6

            If Trim(Cot6) = vbNullString Then
                Cot6 = "''"
            Else
                Cot6 = "'" & Cot6 & "'"
            End If

            .GetText .ColLetterToNumber("H"), .Row, Cot7

            If Trim(Cot7) = vbNullString Then
                Cot7 = "CTOD('')"
            Else
                Cot7 = ToDate(Trim(Cot7), DDMMYYYY)
                Cot7 = "CTOD('" & Format(Cot7, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("I"), .Row, Cot8

            If Trim(Cot8) = vbNullString Then
                Cot8 = "''"
            Else
                Cot8 = "'" & Cot8 & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, Cot9

            If Trim(Cot9) = vbNullString Then
                Cot9 = "CTOD('')"
            Else
                Cot9 = ToDate(Trim(Cot9), DDMMYYYY)
                Cot9 = "CTOD('" & Format(Cot9, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("K"), .Row, Cot10

            If Trim(Cot10) = vbNullString Then
                Cot10 = "''"
            Else
                Cot10 = "'" & Cot10 & "'"
            End If

            .GetText .ColLetterToNumber("L"), .Row, Cot11

            If Trim(Cot11) = vbNullString Then
                Cot11 = "0"
            End If

            .GetText .ColLetterToNumber("M"), .Row, Cot12

            If Trim(Cot12) = vbNullString Then
                Cot12 = "''"
            Else
                Cot12 = "'" & Cot12 & "'"
            End If

            .GetText .ColLetterToNumber("N"), .Row, Cot13

            If Trim(Cot13) = vbNullString Then
                Cot13 = "CTOD('')"
            Else
                Cot13 = ToDate(Trim(Cot13), DDMMYYYY)
                Cot13 = "CTOD('" & Format(Cot13, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("O"), .Row, Cot14

            If Trim(Cot14) = vbNullString Then
                Cot14 = "''"
            Else
                Cot14 = "'" & Cot14 & "'"
            End If

            .GetText .ColLetterToNumber("P"), .Row, Cot15

            If Trim(Cot15) = vbNullString Then
                Cot15 = "0"
            End If

            .GetText .ColLetterToNumber("Q"), .Row, Cot16

            If Trim(Cot16) = vbNullString Then
                Cot16 = "''"
            Else
                Cot16 = "'" & Cot16 & "'"
            End If

            .GetText .ColLetterToNumber("R"), .Row, Cot17

            If Trim(Cot17) = vbNullString Then
                Cot17 = "CTOD('')"
            Else
                Cot17 = ToDate(Trim(Cot17), DDMMYYYY)
                Cot17 = "CTOD('" & Format(Cot17, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("S"), .Row, Cot18

            If Trim(Cot18) = vbNullString Then
                Cot18 = "''"
            Else
                Cot18 = "'" & Cot18 & "'"
            End If

            .GetText .ColLetterToNumber("T"), .Row, Cot19

            If Trim(Cot19) = vbNullString Then
                Cot19 = "0"
            End If

            .GetText .ColLetterToNumber("U"), .Row, Cot20

            If Trim(Cot20) = vbNullString Then
                Cot20 = "''"
            Else
                Cot20 = "'" & Cot20 & "'"
            End If

            .GetText .ColLetterToNumber("V"), .Row, Cot21

            If Trim(Cot21) = vbNullString Then
                Cot21 = "CTOD('')"
            Else
                Cot21 = ToDate(Trim(Cot21), DDMMYYYY)
                Cot21 = "CTOD('" & Format(Cot21, "mm/dd/yyyy") & "')"
            End If
            .GetText .ColLetterToNumber("W"), .Row, Cot22

            If Trim(Cot22) = vbNullString Then
                Cot22 = "''"
            Else
                Cot22 = "'" & Cot22 & "'"
            End If

            .GetText .ColLetterToNumber("X"), .Row, Cot23

            If Trim(Cot23) = vbNullString Then
                Cot23 = "0"
            End If

            .GetText .ColLetterToNumber("Y"), .Row, Cot24

            If Trim(Cot24) = vbNullString Then
                Cot24 = "''"
            Else
                Cot24 = "'" & Cot24 & "'"
            End If

            .GetText .ColLetterToNumber("Z"), .Row, Cot25

            If Trim(Cot25) = vbNullString Then
                Cot25 = "CTOD('')"
            Else
                Cot25 = ToDate(Trim(Cot25), DDMMYYYY)
                Cot25 = "CTOD('" & Format(Cot25, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("AA"), .Row, Cot26

            If Trim(Cot26) = vbNullString Then
                Cot26 = "''"
            Else
                Cot26 = "'" & Cot26 & "'"
            End If

            .GetText .ColLetterToNumber("AB"), .Row, Cot27

            If Trim(Cot27) = vbNullString Then
                Cot27 = "0"
            End If

            .GetText .ColLetterToNumber("AC"), .Row, Cot28

            If Trim(Cot28) = vbNullString Then
                Cot28 = "''"
            Else
                Cot28 = "'" & Cot28 & "'"
            End If

            .GetText .ColLetterToNumber("AD"), .Row, Cot29

            If Trim(Cot29) = vbNullString Then
                Cot29 = "CTOD('')"
            Else
                Cot29 = ToDate(Trim(Cot29), DDMMYYYY)
                Cot29 = "CTOD('" & Format(Cot29, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("AE"), .Row, Cot30

             If Trim(Cot30) = vbNullString Then
                Cot30 = "''"
            Else
                Cot30 = "'" & Cot30 & "'"
            End If

            .GetText .ColLetterToNumber("AF"), .Row, Cot31

            If Trim(Cot31) = vbNullString Then
                Cot31 = "0"
            End If

            .GetText .ColLetterToNumber("AG"), .Row, Cot32

            If Trim(Cot32) = vbNullString Then
                Cot32 = "''"
            Else
                Cot32 = "'" & Cot32 & "'"
            End If

            .GetText .ColLetterToNumber("AH"), .Row, Cot33

            If Trim(Cot33) = vbNullString Then
                Cot33 = "''"
            Else
                Cot33 = "'" & Cot33 & "'"
            End If

            sSQLVal = madtnt & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & kykkhai & "," & matkhai & "," & SOCT & "," & STT & "," & Cot2 & "," & Cot3 & "," & cot4 & "," & cot5 & "," & Cot6 & "," & Cot7 & "," & Cot8 & "," & Cot9 & "," & Cot10 & "," & Cot11 & "," & Cot12 & "," & Cot13 & "," & Cot14 & "," & Cot15 & "," & Cot16 & "," & Cot17 & "," & Cot18 & "," & Cot19 & "," & Cot20 & "," & Cot21 & "," & Cot22 & "," & Cot23 & "," & Cot24 & "," & Cot25 & "," & Cot26 & "," & Cot27 & "," & Cot28 & "," & Cot29 & "," & Cot30 & "," & Cot31 & "," & Cot32 & "," & Cot33 & "," & LAN_QUET & "," & DANHAN
          
            sSQL = "INSERT INTO TMP_PL03" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)

            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"

    End With

    clsDAO.Disconnect
    InsertDTL_PL3 = vbNullString
           
End Function

Public Function InsertDTL_PL016() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim bln  As Boolean
   
  Dim madtnt As Variant
  Dim ngnop As Variant
  Dim KYLBO As Variant
  Dim kykkhai As Variant
  Dim matkhai As Variant
  
  Dim TENCSSX As Variant
  Dim DCHI As Variant
  Dim MADTNT1 As Variant
  Dim THUE5 As Variant
  Dim THUE10 As Variant
  Dim TONGDT As Variant
  Dim THUEPN1 As Variant
  Dim THUEPN2 As Variant
  Dim DANHAN As Variant
  Dim CQUANQL As Variant
  sSQLCol = "MADTNT,NGNOP,TTHTK,KYLBO,KYKKHAI,MATKHAI,TENCSSX,MADTNT1,THUE5,THUE10,TONGDT,THUEPN1,THUEPN2,DANHAN,LAN_QUET,CQQL,DCHI"

   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        
        .Sheet = 8
        .Row = 8
        .Col = .ColLetterToNumber("B")
        Do
            'get TEN CSSXPT
            .GetText .ColLetterToNumber("C"), .Row, TENCSSX
            If Trim(TENCSSX) = vbNullString Then
                TENCSSX = "''"
            Else
                TENCSSX = "'" & Trim(TAX_Utilities_Svr_New.Convert(CStr(TENCSSX), UNICODE, TCVN)) & "'"
            End If
            
'            'get DIACHI
'            .GetText .ColLetterToNumber("D"), .Row, DCHI
'            If Trim(DCHI) = vbNullString Then
'                DCHI = "''"
'            Else
'                DCHI = "'" & Trim(TAX_Utilities_Svr_New.Convert(CStr(DCHI), UNICODE, TCVN)) & "'"
'            End If
            
            'get MADTNT 1
            .GetText .ColLetterToNumber("D"), .Row, MADTNT1
            If Trim(MADTNT1) = vbNullString Then
                MADTNT1 = "''"
            Else
                MADTNT1 = "'" & MADTNT1 & "'"
            End If
                        'get CQT
            .GetText .ColLetterToNumber("E"), .Row, CQUANQL
            If Trim(CQUANQL) = vbNullString Then
                CQUANQL = "''"
            Else
                CQUANQL = "'" & CQUANQL & "'"
            End If
            'get HHCT 5%
            .GetText .ColLetterToNumber("G"), .Row, THUE5
            If Trim(THUE5) = vbNullString Then
                THUE5 = "0"
            Else
                THUE5 = Replace(THUE5, ".", "")
            End If
            'get HHCT 10%
            .GetText .ColLetterToNumber("H"), .Row, THUE10
            If Trim(THUE10) = vbNullString Then
                THUE10 = "0"
            Else
                THUE10 = Replace(THUE10, ".", "")
            End If
            'get TongDT
            .GetText .ColLetterToNumber("I"), .Row, TONGDT
            If Trim(TONGDT) = vbNullString Then
                TONGDT = "0"
            Else
                TONGDT = Replace(TONGDT, ".", "")
            End If
            'get THUEPN1
            .GetText .ColLetterToNumber("J"), .Row, THUEPN1
            If Trim(THUEPN1) = vbNullString Then
                THUEPN1 = "0"
            Else
                THUEPN1 = Replace(THUEPN1, ".", "")
            End If
            'get THUEPN2
            .GetText .ColLetterToNumber("K"), .Row, THUEPN2
            If Trim(THUEPN2) = vbNullString Then
                THUEPN2 = "0"
            Else
                THUEPN2 = Replace(THUEPN2, ".", "")
            End If
        ' sSQLCol = "MADTNT,NGNOP,TTHTK,KYLBO,KYKKHAI,MATKHAI,TENCSSX,DCHI,MADTNT1,THUE5,THUE10,TONGDT,THUEPN1,THUEPN2"
        
            sSQLVal = madtnt & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & kykkhai & "," & matkhai & "," & TENCSSX & "," & MADTNT1 & "," & THUE5 & "," & THUE10 & "," & TONGDT & "," & THUEPN1 & "," & THUEPN2 & "," & DANHAN & "," & LAN_QUET & "," & CQUANQL & ",''"
          
            sSQL = "INSERT INTO TMP_P016" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)

           .Col = .ColLetterToNumber("B")
           .Row = .Row + 1
        Loop Until .Text = "aa"
    End With
           clsDAO.Disconnect
           InsertDTL_PL016 = vbNullString
           
End Function

Public Function InsertDTL_PL016B() As String
    Dim sSQL     As String
    Dim sSQLCol  As String
    Dim sSQLVal  As String
    Dim bln      As Boolean
   
    Dim madtnt   As Variant
    Dim ngnop    As Variant
    Dim KYLBO    As Variant
    Dim kykkhai  As Variant
    Dim matkhai  As Variant
    Dim CHINHANH As Variant
    Dim CT07     As Variant
    Dim MADTNT1  As Variant
    Dim STT      As Variant
    Dim TONGDT   As Variant
    Dim THUEPN1  As Variant
    Dim THUEPN2  As Variant
    Dim DANHAN   As Variant
    Dim CQUANQL  As Variant
    sSQLCol = "MADTNT,NGNOP,TTHTK,KYLBO,KYKKHAI,MATKHAI,CT07,STT,TENCN,CQQL,MADTNT01,DOANHTHU,THUEPN1,THUEPN2,DANHAN,LANQUET"

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If

    ' xet ky ke khai
    If isTkQuy = True Then
        kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
    Else
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
    End If
        
    matkhai = "'" & LOAI_TO_KHAI & "'"

    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
        
        .GetText .ColLetterToNumber("G"), 13, CT07

        If Trim(CT07) = vbNullString Then
            CT07 = "0"
        Else
            CT07 = Replace(CT07, ".", "")
        End If

        DANHAN = "''"
        
        .Sheet = 9
        .Row = 7
        .Col = .ColLetterToNumber("B")

        Do
            'get STT
            .GetText .ColLetterToNumber("B"), .Row, STT

            If Trim(STT) = vbNullString Then
                STT = "''"
            Else
                STT = "'" & Trim(TAX_Utilities_Svr_New.Convert(CStr(STT), UNICODE, TCVN)) & "'"
            End If

            'Get ten chi nhanh
            .GetText .ColLetterToNumber("C"), .Row, CHINHANH

            If Trim(CHINHANH) = vbNullString Then
                CHINHANH = "''"
            Else
                CHINHANH = "'" & CHINHANH & "'"
            End If

            'get CQT
            .GetText .ColLetterToNumber("D"), .Row, CQUANQL

            If Trim(CQUANQL) = vbNullString Then
                CQUANQL = "''"
            Else
                CQUANQL = "'" & CQUANQL & "'"
            End If

            'get dtnt
            .GetText .ColLetterToNumber("F"), .Row, MADTNT1

            If Trim(MADTNT1) = vbNullString Then
                MADTNT1 = "''"
            Else
                MADTNT1 = "'" & MADTNT1 & "'"
            End If

            'get TongDT
            .GetText .ColLetterToNumber("G"), .Row, TONGDT

            If Trim(TONGDT) = vbNullString Then
                TONGDT = "0"
            Else
                TONGDT = Replace(TONGDT, ".", "")
            End If

            'get THUEPN1
            .GetText .ColLetterToNumber("H"), .Row, THUEPN1

            If Trim(THUEPN1) = vbNullString Then
                THUEPN1 = "0"
            Else
                THUEPN1 = Replace(THUEPN1, ".", "")
            End If

            'get THUEPN2
            .GetText .ColLetterToNumber("I"), .Row, THUEPN2

            If Trim(THUEPN2) = vbNullString Then
                THUEPN2 = "0"
            Else
                THUEPN2 = Replace(THUEPN2, ".", "")
            End If

            ' sSQLCol = "MADTNT,NGNOP,TTHTK,KYLBO,KYKKHAI,MATKHAI,TENCSSX,DCHI,MADTNT1,THUE5,THUE10,TONGDT,THUEPN1,THUEPN2"
            sSQLVal = madtnt & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & kykkhai & "," & matkhai & "," & CT07 & "," & STT & "," & CHINHANH & "," & CQUANQL & "," & MADTNT1 & "," & TONGDT & "," & THUEPN1 & "," & THUEPN2 & "," & DANHAN & "," & LAN_QUET
          
            sSQL = "INSERT INTO tmp_p016B" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)

            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"

    End With

    clsDAO.Disconnect
    InsertDTL_PL016B = vbNullString
           
End Function

'Public Function InsertDTL_PL017() As String
'    Dim sSQL    As String
'    Dim sSQLCol As String
'    Dim sSQLVal As String
'    Dim bln     As Boolean
'
'    Dim madtnt  As Variant
'    Dim ngnop   As Variant
'    Dim KYLBO   As Variant
'    Dim CT07    As Variant
'    Dim kykkhai As Variant
'    Dim MATKHAI As Variant
'    Dim MADTNT1 As Variant
'    Dim THUE5   As Variant
'    Dim THUE10  As Variant
'    Dim TONGDT  As Variant
'    Dim THUEPN1 As Variant
'    Dim THUEPN2 As Variant
'    Dim DANHAN  As Variant
'    Dim CQUANQL As Variant
'
'    sSQLCol = "MADTNT,NGNOP,TTHTK,KYLBO,KYKKHAI,MATKHAI,CT07,CQQL,MADTNT01,THUE5,THUE10,DOANHTHU,THUEPN1,THUEPN2,DANHAN,LAN_QUET"
'
'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'    End If
'
'    ' xet ky ke khai
'    If isTkQuy = True Then
'        kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
'    Else
'        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'    End If
'
'    MATKHAI = "'" & LOAI_TO_KHAI & "'"
'
'    With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("F"), 10, madtnt
'
'        If Trim(madtnt) = vbNullString Then
'            madtnt = "''"
'        Else
'            madtnt = "'" & madtnt & "'"
'        End If
'
'        .GetText .ColLetterToNumber("E"), 32, ngnop
'
'        'NGNOP = Date
'        If Trim(ngnop) = vbNullString Then
'            ngnop = "CTOD('')"
'        Else
'            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
'            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
'        End If
'
'        .GetText .ColLetterToNumber("E"), 30, KYLBO
'
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "CTOD('')"
'        Else
'
'            If Len(Trim(KYLBO)) = 6 Then
'                KYLBO = "'0" & KYLBO & "'"
'            Else
'                KYLBO = "'" & KYLBO & "'"
'            End If
'
'        End If
'
'        .GetText .ColLetterToNumber("E"), 14, CT07
'
'        DANHAN = "''"
'
'        .Sheet = 8
'        .Row = 8
'        .Col = .ColLetterToNumber("B")
'
'        Do
'            'get TEN DIA PHUONG
'            .GetText .ColLetterToNumber("C"), .Row, CQUANQL
'
'            If Trim(CQUANQL) = vbNullString Then
'                CQUANQL = "''"
'            Else
'                CQUANQL = "'" & Trim(TAX_Utilities_Svr_New.Convert(CStr(CQUANQL), UNICODE, TCVN)) & "'"
'            End If
'
'            'get MADTNT 1
'            .GetText .ColLetterToNumber("D"), .Row, MADTNT1
'
'            If Trim(MADTNT1) = vbNullString Then
'                MADTNT1 = "''"
'            Else
'                MADTNT1 = "'" & MADTNT1 & "'"
'            End If
'
'            'get HHCT 5%
'            .GetText .ColLetterToNumber("E"), .Row, THUE5
'
'            If Trim(THUE5) = vbNullString Then
'                THUE5 = "0"
'            Else
'                THUE5 = Replace(THUE5, ".", "")
'            End If
'
'            'get HHCT 10%
'            .GetText .ColLetterToNumber("F"), .Row, THUE10
'
'            If Trim(THUE10) = vbNullString Then
'                THUE10 = "0"
'            Else
'                THUE10 = Replace(THUE10, ".", "")
'            End If
'
'            'get TongDT
'            .GetText .ColLetterToNumber("G"), .Row, TONGDT
'
'            If Trim(TONGDT) = vbNullString Then
'                TONGDT = "0"
'            Else
'                TONGDT = Replace(TONGDT, ".", "")
'            End If
'
'            'get THUEPN1
'            .GetText .ColLetterToNumber("H"), .Row, THUEPN1
'
'            If Trim(THUEPN1) = vbNullString Then
'                THUEPN1 = "0"
'            Else
'                THUEPN1 = Replace(THUEPN1, ".", "")
'            End If
'
'            'get THUEPN2
'            .GetText .ColLetterToNumber("I"), .Row, THUEPN2
'
'            If Trim(THUEPN2) = vbNullString Then
'                THUEPN2 = "0"
'            Else
'                THUEPN2 = Replace(THUEPN2, ".", "")
'            End If
'
'            ' sSQLCol = "MADTNT,NGNOP,TTHTK,KYLBO,KYKKHAI,MATKHAI,TENCSSX,DCHI,MADTNT1,THUE5,THUE10,TONGDT,THUEPN1,THUEPN2"
'            If CT07 = vbNullString Then
'                CT07 = "0"
'            End If
'            sSQLVal = madtnt & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & kykkhai & "," & MATKHAI & "," & CT07 & "," & CQUANQL & "," & MADTNT1 & "," & THUE5 & "," & THUE10 & "," & TONGDT & "," & THUEPN1 & "," & THUEPN2 & "," & DANHAN & "," & LAN_QUET
'
'            sSQL = "INSERT INTO TMP_P017" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'            bln = clsDAO.ExecuteDLL(sSQL)
'
'            .Col = .ColLetterToNumber("B")
'            .Row = .Row + 1
'        Loop Until .Text = "aa"
'
'    End With
'
'    clsDAO.Disconnect
'    InsertDTL_PL017 = vbNullString
'
'End Function

Public Function InsertDTL_KHBS() As String
    Dim sSQL        As String
    Dim sSQLCol     As String
    Dim sSQLVal     As String
    Dim bln         As Boolean
    Dim madtnt      As Variant
    Dim ngnop       As Variant
    Dim KYLBO       As Variant
    Dim kykkhai     As Variant
    Dim matkhai     As Variant
    Dim MATHUE      As Variant
    Dim MaMuc       As Variant
    Dim maTM        As Variant
    Dim MAPP        As Variant
    Dim mact        As Variant
    Dim MACT2       As Variant
    Dim MACT3       As Variant
    Dim STT         As Variant
    Dim STT2        As Variant
    Dim STTIN       As Variant
    Dim SOKK        As Variant
    Dim SODC        As Variant
    Dim SOCL        As Variant
    Dim i           As Integer, iCol As Long, iRow As Long
    Dim xmlNode     As MSXML.IXMLDOMNode
    Dim DANHAN      As Variant
    Dim strFileName As String
    Dim fso         As New FileSystemObject
    
    'them bien de luu So Tien Nop Cham va so Ngay Nop Cham theo TT28
    Dim SOTIENNC    As Variant
    Dim SONGAYNC    As Variant
    'kiem tra ton tai tep *.dbf chua
    '    strFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "BS" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileName) = False Then
    '    Dim strTepmau As String
    '        strTepmau = spathVat & "\tepmau\BS.DBF"
    '        fso.CopyFile strTepmau, strFileName, False
    '    End If
   
    sSQLCol = "MADTNT, KYKKHAI, MATKHAI, MATHUE, MAMUC, MATM, MAPP, NGNOP, TTHTK, KYLBO , MACT, MACT2, MACT3, STT, STT2, STTIN," & " SOKK, SODC, SOCL,SOTIENNC,  SONGAYNC, LAN_QUET, DANHAN"
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        MATHUE = "'01'"
        MAPP = "'1'"
          
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
        
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                
        ' Insert so dieu chinh tang
        
        mact = "'1000'"
        MACT2 = "''"
        MACT3 = "''"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        SOKK = "0"
        SODC = "0"
        SOCL = "0"
        SOTIENNC = "0"
        SONGAYNC = "0"
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                       
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        .Sheet = .SheetCount - 1
        
        'get so Tien nop cham
        .GetText .ColLetterToNumber("BE"), .MaxRows - 27, SOTIENNC

        If SOTIENNC = vbNullString Then
            SOTIENNC = "0"
        Else
            SOTIENNC = Replace(SOTIENNC, ".", "")
        End If

        'get so ngay nop cham
        .GetText .ColLetterToNumber("BE"), .MaxRows - 28, SONGAYNC

        If SONGAYNC = vbNullString Then
            SONGAYNC = "0"
        Else
            SONGAYNC = Replace(SONGAYNC, ".", "")
        End If
        
        i = 1
        .Row = 9

        Do
            mact = "''"
        
            .GetText .ColLetterToNumber("BE"), .Row, MACT2

            If Trim(MACT2) = vbNullString Then
                MACT2 = "''"
            ElseIf MACT2 = "11" Then
                MACT2 = "'005'"
            ElseIf MACT2 = "23" Then
                MACT2 = "'040'"
            ElseIf MACT2 = "31" Then
                MACT2 = "'052'"
            ElseIf MACT2 = "33" Then
                MACT2 = "'054'"
            ElseIf MACT2 = "42" Then
                MACT2 = "'080'"
            ElseIf MACT2 = "22" Then
                MACT2 = "'001'"
            ElseIf MACT2 = "25" Then
                MACT2 = "'005'"
            ElseIf MACT2 = "31" Then
                MACT2 = "'010'"
            ElseIf MACT2 = "33" Then
                MACT2 = "'011'"
            ElseIf MACT2 = "42" Then
                MACT2 = "'023'"
            ElseIf MACT2 = "40" Then
                MACT2 = "'021'"
            ElseIf MACT2 = "43" Then
                MACT2 = "'024'"
            Else
                MACT2 = "'000'"
            End If
                
            MACT3 = "'2'"
            STT = 1
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("BF"), .Row, SOKK

            If Trim(SOKK) = vbNullString Then
                SOKK = "0"
            Else
                SOKK = SOKK
            End If
                
            .GetText .ColLetterToNumber("BG"), .Row, SODC

            If Trim(SODC) = vbNullString Then
                SODC = "0"
            Else
                SODC = SODC
            End If
                
            .GetText .ColLetterToNumber("BH"), .Row, SOCL

            If Trim(SOCL) = vbNullString Then
                SOCL = "0"
            Else
                SOCL = SOCL
            End If
            
            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                           
            sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                
            If MACT2 <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
        
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"

        ' Insert so dieu chinh giam

        mact = "'2000'"
        MACT2 = "''"
        MACT3 = "''"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        SOKK = "0"
        SODC = "0"
        SOCL = "0"
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                       
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

        i = 1
        .Row = .Row + 3

        Do
            mact = "''"
            .GetText .ColLetterToNumber("BE"), .Row, MACT2

            If Trim(MACT2) = vbNullString Then
                MACT2 = "''"
            ElseIf MACT2 = "11" Then
                MACT2 = "'005'"
            ElseIf MACT2 = "23" Then
                MACT2 = "'040'"
            ElseIf MACT2 = "31" Then
                MACT2 = "'052'"
            ElseIf MACT2 = "33" Then
                MACT2 = "'054'"
            ElseIf MACT2 = "42" Then
                MACT2 = "'080'"
            ElseIf MACT2 = "22" Then
                MACT2 = "'001'"
            ElseIf MACT2 = "25" Then
                MACT2 = "'005'"
            ElseIf MACT2 = "31" Then
                MACT2 = "'010'"
            ElseIf MACT2 = "33" Then
                MACT2 = "'011'"
            ElseIf MACT2 = "42" Then
                MACT2 = "'023'"
            ElseIf MACT2 = "40" Then
                MACT2 = "'021'"
            ElseIf MACT2 = "43" Then
                MACT2 = "'024'"
            Else
                MACT2 = "'000'"
            End If
                
            MACT3 = "'2'"
            STT = 2
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("BF"), .Row, SOKK

            If Trim(SOKK) = vbNullString Then
                SOKK = "0"
            Else
                SOKK = SOKK
            End If
                
            .GetText .ColLetterToNumber("BG"), .Row, SODC

            If Trim(SODC) = vbNullString Then
                SODC = "0"
            Else
                SODC = SODC
            End If
                
            .GetText .ColLetterToNumber("BH"), .Row, SOCL

            If Trim(SOCL) = vbNullString Then
                SOCL = "0"
            Else
                SOCL = SOCL
            End If
            
            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                           
            sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If MACT2 <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
        
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "bb"
        
        ' Insert Tong hop dieu chinh

        mact = "'3000'"
        MACT2 = "''"
        MACT3 = "''"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        SOKK = "0"
        SODC = "0"
        SOCL = "0"
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                       
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        .Row = .Row + 3
        mact = "''"
        .GetText .ColLetterToNumber("BE"), .Row, MACT2

        If Trim(MACT2) = vbNullString Then
            MACT2 = "''"
        ElseIf MACT2 = "11" Then
            MACT2 = "'005'"
        ElseIf MACT2 = "23" Then
            MACT2 = "'040'"
        ElseIf MACT2 = "31" Then
            MACT2 = "'052'"
        ElseIf MACT2 = "33" Then
            MACT2 = "'054'"
        ElseIf MACT2 = "42" Then
            MACT2 = "'080'"
        ElseIf MACT2 = "22" Then
            MACT2 = "'001'"
        ElseIf MACT2 = "25" Then
            MACT2 = "'005'"
        ElseIf MACT2 = "31" Then
            MACT2 = "'010'"
        ElseIf MACT2 = "33" Then
            MACT2 = "'011'"
        ElseIf MACT2 = "42" Then
            MACT2 = "'023'"
        ElseIf MACT2 = "40" Then
            MACT2 = "'021'"
        ElseIf MACT2 = "43" Then
            MACT2 = "'024'"
        Else
            MACT2 = "'000'"
        End If
                
        MACT3 = "'2'"
        STT = 3
        STT2 = 1
        STTIN = "'1'"
                
        .GetText .ColLetterToNumber("BF"), .Row, SOKK

        If Trim(SOKK) = vbNullString Then
            SOKK = "0"
        Else
            SOKK = SOKK
        End If
                
        .GetText .ColLetterToNumber("BG"), .Row, SODC

        If Trim(SODC) = vbNullString Then
            SODC = "0"
        Else
            SODC = SODC
        End If
                
        .GetText .ColLetterToNumber("BH"), .Row, SOCL

        If Trim(SOCL) = vbNullString Then
            SOCL = "0"
        Else
            SOCL = SOCL
        End If
            
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                           
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        If MACT2 <> "''" Then
            bln = clsDAO.ExecuteDLL(sSQL)
        End If
                
        .Row = .Row + 1
        mact = "''"
        .GetText .ColLetterToNumber("BE"), .Row, MACT2

        If Trim(MACT2) = vbNullString Then
            MACT2 = "''"
        ElseIf MACT2 = "11" Then
            MACT2 = "'005'"
        ElseIf MACT2 = "23" Then
            MACT2 = "'040'"
        ElseIf MACT2 = "31" Then
            MACT2 = "'052'"
        ElseIf MACT2 = "33" Then
            MACT2 = "'054'"
        ElseIf MACT2 = "42" Then
            MACT2 = "'080'"
        ElseIf MACT2 = "22" Then
            MACT2 = "'001'"
        ElseIf MACT2 = "25" Then
            MACT2 = "'005'"
        ElseIf MACT2 = "31" Then
            MACT2 = "'010'"
        ElseIf MACT2 = "33" Then
            MACT2 = "'011'"
        ElseIf MACT2 = "42" Then
            MACT2 = "'023'"
        ElseIf MACT2 = "40" Then
            MACT2 = "'021'"
        ElseIf MACT2 = "43" Then
            MACT2 = "'024'"
        Else
            MACT2 = "'000'"
        End If
                
        MACT3 = "'2'"
        STT = 3
        STT2 = 2
        STTIN = "'2'"
                
        .GetText .ColLetterToNumber("BF"), .Row, SOKK

        If Trim(SOKK) = vbNullString Then
            SOKK = "0"
        Else
            SOKK = SOKK
        End If
                
        .GetText .ColLetterToNumber("BG"), .Row, SODC

        If Trim(SODC) = vbNullString Then
            SODC = "0"
        Else
            SODC = SODC
        End If
                
        .GetText .ColLetterToNumber("BH"), .Row, SOCL

        If Trim(SOCL) = vbNullString Then
            SOCL = "0"
        Else
            SOCL = SOCL
        End If
            
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                           
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        If MACT2 <> "''" Then
            bln = clsDAO.ExecuteDLL(sSQL)
        End If

    End With
   
    clsDAO.Disconnect
    InsertDTL_KHBS = vbNullString
   
End Function

Public Function InsertDTL() As String
    Dim sSQL      As String
    Dim sSQLCol   As String
    Dim sSQLVal   As String
   
    Dim madtnt    As Variant
    Dim MATHUE    As Variant
    Dim maTM      As Variant
    Dim MAPP      As Variant
    Dim KYLBO     As Variant
    Dim LoaiTien  As Variant
    Dim ngnop     As Variant
    Dim CT_KT     As Variant
    Dim DSO       As Variant
    Dim THUEGTGT  As Variant
    Dim DSO2      As Variant
    Dim THUEGTGT2 As Variant
    Dim tt        As Variant
    Dim MaMuc     As Variant
    Dim CT30      As Variant
    Dim CT31      As Variant
    Dim CT32      As Variant
    Dim CT33      As Variant
    Dim SCT31     As Boolean
    Dim SCT33     As Boolean
    Dim matkhai   As Variant
    Dim kykkhai   As Variant
    Dim DANHAN    As Variant
 
    SCT31 = True
    SCT33 = True
   
    Dim bln As Boolean

    If clsDAO.Connected = False Then
        '        Me.MousePointer = vbHourglass
        '        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        '        frmSystem.MousePointer = vbDefault
        '        Me.MousePointer = vbDefault
    End If
    
    sSQLCol = "MADTNT , MATHUE, MATM, MAPP, KYLBO, LOAITIEN, NGNOP, "
    sSQLCol = sSQLCol & "CT_KT , DSO, THUEGTGT, DSO2, THUEGTGT2, TT, MAMUC, TTHTK, MATKHAI, KYKKHAI, LAN_QUET, DANHAN"
   
    '-------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"

        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'001'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 46, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = "0"
        THUEGTGT2 = THUEGTGT
        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
        
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'002'"
        DSO = "0"

        THUEGTGT = "0"
        
        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'003'"
        DSO = "0"
        THUEGTGT = "0"
        
        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
    '----------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'004'"
        
        .GetText .ColLetterToNumber("Q"), 49, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If

        .GetText .ColLetterToNumber("W"), 49, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
      
    '-------------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'005'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 50, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
  
    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'006'"
        DSO = "0"
        THUEGTGT = "0"
        
        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'007'"
        THUEGTGT = "0"
        .GetText .ColLetterToNumber("Q"), 52, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'008'"
        
        .GetText .ColLetterToNumber("Q"), 53, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If

        .GetText .ColLetterToNumber("W"), 53, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'009'"
        THUEGTGT = "0"
        .GetText .ColLetterToNumber("Q"), 54, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'010'"
        
        .GetText .ColLetterToNumber("Q"), 55, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If

        .GetText .ColLetterToNumber("W"), 55, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'011'"
        
        .GetText .ColLetterToNumber("Q"), 56, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If

        .GetText .ColLetterToNumber("W"), 56, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'012'"
        
        .GetText .ColLetterToNumber("Q"), 57, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If

        .GetText .ColLetterToNumber("W"), 57, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'013'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 58, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '---------------------
   
    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'014'"
        DSO = "0"
        THUEGTGT = "0"
        
        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'015'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 60, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'016'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 61, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'017'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 62, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '----------------------

    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'018'"
        DSO = "0"
        THUEGTGT = "0"
        
        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
            
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '-----------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'019'"
        
        .GetText .ColLetterToNumber("W"), 64, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'020'"
        
        .GetText .ColLetterToNumber("W"), 65, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '-----------------------
    '    With fps
    '        .Sheet = 1
    '        .GetText .ColLetterToNumber("F"), 10, madtnt
    '
    '        If Trim(madtnt) = vbNullString Then
    '            madtnt = "''"
    '        Else
    '            madtnt = "'" & madtnt & "'"
    '        End If
    '
    '        MATHUE = "'01'"
    '        MAPP = "'1'"
    '        .GetText .ColLetterToNumber("E"), 30, KYLBO
    '
    '        If Trim(KYLBO) = vbNullString Then
    '            KYLBO = "''"
    '        Else
    '
    '            If Len(Trim(KYLBO)) = 6 Then
    '                KYLBO = "'0" & KYLBO & "'"
    '            Else
    '                KYLBO = "'" & KYLBO & "'"
    '            End If
    '        End If
    '
    '        LOAITIEN = "''"
    '        .GetText .ColLetterToNumber("E"), 32, ngnop
    '
    '        'NGNOP = Date
    '        If Trim(ngnop) = vbNullString Then
    '            ngnop = "CTOD('')"
    '        Else
    '            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
    '            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
    '        End If
    '
    '        'Chi tieu
    '        CT_KT = "'021'"
    '
    '        .GetText .ColLetterToNumber("W"), 66, THUEGTGT
    '
    '        If Trim(THUEGTGT) = vbNullString Then
    '            THUEGTGT = "0"
    '        End If
    '
    '        DSO2 = DSO
    '        THUEGTGT2 = THUEGTGT
    '        tt = "''"
    '
    '        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
    '            MATM = "'01'"
    '            MaMuc = "'014'"
    '        Else
    '            MaMuc = "'1700'"
    '            MATM = "'1701'"
    '        End If
    '
    '
    '        MATKHAI = "'" & LOAI_TO_KHAI & "'"
    '        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
    '        ' xet ky ke khai
    '        If isTkQuy = True Then
    '            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
    '        Else
    '            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
    '        End If
    '   End With
    '
    '    sSQLVal = madtnt & "," & MATHUE & "," & MATM & "," & MAPP & "," & KYLBO & "," & LOAITIEN & "," & ngnop & ","
    '    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    '    sSQLVal = sSQLVal & TTHTK & "," & MATKHAI & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
    '
    '   sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & _
    '     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    '
    '    bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'021'"
        
        .GetText .ColLetterToNumber("W"), 67, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'022'"
        
        .GetText .ColLetterToNumber("W"), 68, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'023'"
        
        .GetText .ColLetterToNumber("W"), 70, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------
    '    With fps
    '        .Sheet = 1
    '        .GetText .ColLetterToNumber("F"), 10, madtnt
    '
    '        If Trim(madtnt) = vbNullString Then
    '            madtnt = "''"
    '        Else
    '            madtnt = "'" & madtnt & "'"
    '        End If
    '
    '        MATHUE = "'01'"
    '        MAPP = "'1'"
    '        .GetText .ColLetterToNumber("E"), 30, KYLBO
    '
    '        If Trim(KYLBO) = vbNullString Then
    '            KYLBO = "''"
    '        Else
    '
    '            If Len(Trim(KYLBO)) = 6 Then
    '                KYLBO = "'0" & KYLBO & "'"
    '            Else
    '                KYLBO = "'" & KYLBO & "'"
    '            End If
    '        End If
    '
    '        LOAITIEN = "''"
    '        .GetText .ColLetterToNumber("E"), 32, ngnop
    '
    '        'NGNOP = Date
    '        If Trim(ngnop) = vbNullString Then
    '            ngnop = "CTOD('')"
    '        Else
    '            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
    '            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
    '        End If
    '
    '        'Chi tieu
    '        CT_KT = "'025'"
    '
    '        .GetText .ColLetterToNumber("W"), 71, THUEGTGT
    '
    '        If Trim(THUEGTGT) = vbNullString Then
    '            THUEGTGT = "0"
    '        End If
    '
    '        DSO2 = DSO
    '        THUEGTGT2 = THUEGTGT
    '        tt = "''"
    '
    '        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
    '            MATM = "'01'"
    '            MaMuc = "'014'"
    '        Else
    '            MaMuc = "'1700'"
    '            MATM = "'1701'"
    '        End If
    '
    '    End With
    '
    '    sSQLVal = madtnt & "," & MATHUE & "," & MATM & "," & MAPP & "," & KYLBO & "," & LOAITIEN & "," & ngnop & ","
    '    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    '    sSQLVal = sSQLVal & TTHTK & "," & MATKHAI & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
    '
    '    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    '
    '    bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CT_KT = "'024'"
        
        .GetText .ColLetterToNumber("W"), 72, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"
      
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CT_KT & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & ","
    sSQLVal = sSQLVal & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_KT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '----------------------
    '----------------------
    '----------------------
    '----------------------
    clsDAO.Disconnect
   
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL1
    End If
   
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2
    End If
   
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(3).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL3
    End If
   
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL4A
    End If
   
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL4B
    End If
   
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(6).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL5
    End If
   
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(7).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL016
    End If
   
    '    If TAX_Utilities_Svr_New.NodeValidity.childNodes(8).Attributes.getNamedItem("Active").nodeValue = 1 Then
    '        InsertDTL = InsertDTL_PL016B
    '    End If
   
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(8).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL017
    End If
   
    InsertDTL = vbNullString
End Function

Public Function InsertHDR() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    Dim kykkhai    As Variant
    Dim matkhai    As Variant
    Dim TENNV      As Variant
    Dim CHUNGCHI   As Variant
    Dim madtnt     As Variant
    Dim MATHUE     As Variant
    Dim maTM       As Variant
    Dim MAPP       As Variant
    Dim KYLBO      As Variant
    Dim ngnop      As Variant
    Dim LoaiTien   As Variant
    Dim MABPQL     As Variant
    Dim SHTEP      As Variant
    Dim NgNhap     As Variant
    Dim ThueTKy    As Variant
    Dim GTTinhThue As Variant
    Dim DaGhi      As Variant
    Dim KTTRUOC    As Variant
    Dim DoanhSo    As Variant
    Dim MAPP2      As Variant
    Dim MaMuc      As Variant
    Dim MaVach     As Variant
    Dim CT27       As Variant
    Dim ct34       As Variant
    Dim CT36       As Variant
    Dim HANNOP2    As Variant
    Dim ThueTKy2   As Variant
    Dim TENDA      As Variant
    Dim NGANH      As Variant
    Dim DAGHI1     As Variant
    Dim AUTO_CAL   As Variant
    Dim DANHAN     As Variant
    'MADLT
    Dim madlt      As Variant
    Dim LanBS      As Variant
    Dim MANNKD     As Variant
    Dim THUECL     As Variant
    Dim DSoCL      As Variant
    Dim GTTTCL     As Variant
    Dim NgayPs     As Variant
    Dim GiaHan     As Variant
    Dim MALD       As Variant
    Dim LoaiKyKK   As Variant
    
    sSQLCol = "MADTNT, MATHUE, MATM, MAPP, KYLBO, NGNOP, "
    sSQLCol = sSQLCol & "LOAITIEN, MABPQL, SHTEP, TTHTK, KYKKHAI, "
    sSQLCol = sSQLCol & "NGNHAP, THUETKY, GTTINHTHUE, DAGHI, "
    sSQLCol = sSQLCol & "KTTRUOC, DOANHSO, MAPP2, MAMUC, MAVACH,MATKHAI,HANNOP, THUETKY2,HANNOP2,TENDA,NGANH,DAGHI1,DAGHI2,AUTO_CAL,LAN_QUET, DANHAN,MADLT ,LANBS,MANNKD,THUECL,DSOCL,GTTTCL,NGAYPS,GIAHAN,MALD,TENNV,CHUNGCHI"
    
    With fps
        .Sheet = 1
        'lay ten nhan vien dai ly thue, chung chi hanh nghe
        .GetText .ColLetterToNumber("F"), 77, TENNV
        'nshung - convert font
        If Trim(TENNV) <> vbNullString Then
            TENNV = TAX_Utilities_Svr_New.Convert(Trim(TENNV), UNICODE, TCVN)
        End If
        .GetText .ColLetterToNumber("F"), 79, CHUNGCHI
        'nshung - convert font
        If Trim(CHUNGCHI) <> vbNullString Then
            CHUNGCHI = TAX_Utilities_Svr_New.Convert(Trim(CHUNGCHI), UNICODE, TCVN)
        End If
        
        'set DAGHI2 theo checkbox loi dinh danh
        .Col = .ColLetterToNumber("E")
        .Row = 36

        If .Value = 1 Then
            DAGHI2 = ".T."
        Else
            DAGHI2 = ".F."
        End If
        
        'set cac chi so
        DSoCL = "0"
        GTTTCL = "0"
        NgayPs = "CTOD('')"
        
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'1'"
        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        LoaiTien = "''"
        '        .GetText .ColLetterToNumber(MA_BPQL_X), MA_BPQL_Y, MABPQL
        
        MABPQL = strMaBPQL

        If Trim(MABPQL) = vbNullString Then
            MABPQL = "''"
        Else
            MABPQL = "'" & MABPQL & "'"
        End If
        
        '       .GetText .ColLetterToNumber("M"), 10, SHTEP
        If Trim(SHTEP) = vbNullString Then
            SHTEP = "''"
        Else
            SHTEP = "'" & TAX_Utilities_Svr_New.Convert(Trim(SHTEP), UNICODE, TCVN) & "'"
        End If

        ' xet ky ke khai
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

        '.GetText .ColLetterToNumber("M"), 12, NGNHAP
        NgNhap = Date

        If Trim(NgNhap) = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If
        
        .GetText .ColLetterToNumber("W"), 67, ThueTKy

        If Trim(ThueTKy) = "0" Then
            ThueTKy = "0"
            .GetText .ColLetterToNumber("W"), 72, ThueTKy

            If Trim(ThueTKy) = "0" Then
                ThueTKy = "0"
            Else
                ThueTKy = -ThueTKy
            End If
        End If
        
        'get MaNNKD
        .GetText .ColLetterToNumber("T"), 82, MANNKD

        If Trim(MANNKD) = vbNullString Then
            MANNKD = "''"
        Else
            MANNKD = "'" & TAX_Utilities_Svr_New.Convert(Trim(MANNKD), UNICODE, TCVN) & "'"
        End If
        
        .GetText .ColLetterToNumber("Q"), 53, CT27

        If Trim(CT27) = vbNullString Then
            CT27 = 0
        End If

        .GetText .ColLetterToNumber("Q"), 58, ct34

        If Trim(ct34) = vbNullString Then
            ct34 = 0
        End If

        '       .GetText .ColLetterToNumber("Q"), 59, CT36
        If Trim(CT36) = vbNullString Then
            CT36 = 0
        End If
        
        GTTinhThue = CT27 + ct34 - CT36

        If GTTinhThue = 0 Then
            GTTinhThue = "0"
        Else
            GTTinhThue = Trim(GTTinhThue)
        End If

        DaGhi = ".F."
        .GetText .ColLetterToNumber("W"), 46, KTTRUOC

        If Trim(KTTRUOC) = vbNullString Then
            KTTRUOC = "0"
        End If

        'DOANHSO = "0"
        DoanhSo = CT27 + ct34 - CT36

        If DoanhSo = 0 Then

            DoanhSo = "0"
        Else

            DoanhSo = Trim(DoanhSo)
        End If

        MAPP2 = "'1'"
        
        '       lay MADLT
        .GetText .ColLetterToNumber("F"), 20, madlt

        If Trim(madlt) = vbNullString Then
            madlt = "''"
        Else
            madlt = "'" & madlt & "'"
        End If

        '       lay so lan bs
        .GetText .ColLetterToNumber("T"), 6, LanBS

        If Trim(LanBS) = vbNullString Then
            LanBS = "0"
        End If

        ' Lay loai ky ke khai(thang/quy)
        .GetText .ColLetterToNumber("AB"), 42, LoaiKyKK

        If LoaiKyKK = "0" Then
            LoaiKyKK = ".T."
        Else
            LoaiKyKK = ".F."
        End If

        ' Lay so lan quet
        '       .GetText .ColLetterToNumber("M"), 30, LAN_QUET
        If Trim(LAN_QUET) = vbNullString Then
            LAN_QUET = "1"
        End If

        DANHAN = "''"
        
        ' sua MAMUC & MATM tu nam 2009
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        MaVach = ".T."
        matkhai = "'" & LOAI_TO_KHAI & "'"

        If Trim(HANNOP) = vbNullString Then
            HANNOP = "CTOD('')"
        Else
            HANNOP = ToDate(Trim(HANNOP), DDMMYYYY)
            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
        End If
        
        ThueTKy2 = "0"
        HANNOP2 = HANNOP
        TENDA = "''"
        DAGHI1 = ".F."
        AUTO_CAL = ".F."
        THUECL = "0"
        NGANH = TinhMaNganh
        
        '       luu check gia han va ly do
        'get check gia han
        .GetText .ColLetterToNumber("E"), 38, GiaHan
        'get ly do gia han
        .GetText .ColLetterToNumber("AB"), 41, MALD

        If Trim(GiaHan) = "1" Or UCase(Trim(GiaHan)) = "X" Then
            GiaHan = ".T."
            '            If Trim(MALD) = "01" Then
            '                MALD = "'01'"
            '            ElseIf Trim(MALD) = "02" Then
            '                MALD = "'02'"
            '            Else
            '                MALD = "Null"
            '            End If
            MALD = "'" & TAX_Utilities_Svr_New.Convert(Trim(MALD), UNICODE, TCVN) & "'"

            'set lai han nop
            If TAX_Utilities_Svr_New.Year = "2012" Then
                If TAX_Utilities_Svr_New.Month = "04" Then
                    HANNOP = "CTOD('" & Format("20/11/2012", "mm/dd/yyyy") & "')"
                ElseIf TAX_Utilities_Svr_New.Month = "05" Then
                    HANNOP = "CTOD('" & Format("20/12/2012", "mm/dd/yyyy") & "')"
                ElseIf TAX_Utilities_Svr_New.Month = "06" Then
                    HANNOP = "CTOD('" & Format("21/1/2013", "mm/dd/yyyy") & "')"
                End If

                HANNOP2 = HANNOP
            ElseIf TAX_Utilities_Svr_New.Year = "2013" Then

                If TAX_Utilities_Svr_New.Month = "01" Then
                    HANNOP = "CTOD('" & Format("20/08/2013", "mm/dd/yyyy") & "')"
                ElseIf TAX_Utilities_Svr_New.Month = "02" Then
                    HANNOP = "CTOD('" & Format("20/09/2013", "mm/dd/yyyy") & "')"
                ElseIf TAX_Utilities_Svr_New.Month = "03" Then
                    HANNOP = "CTOD('" & Format("21/10/2013", "mm/dd/yyyy") & "')"
                End If

                HANNOP2 = HANNOP
            End If

        Else
            GiaHan = ".F."
            MALD = "Null"
        End If
        
    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & ngnop & ", "
    sSQLVal = sSQLVal & LoaiTien & "," & MABPQL & "," & SHTEP & "," & TTHTK & "," & kykkhai & ", "
    sSQLVal = sSQLVal & NgNhap & "," & ThueTKy & "," & GTTinhThue & "," & DaGhi & ", "
    sSQLVal = sSQLVal & KTTRUOC & "," & DoanhSo & "," & MAPP2 & "," & MaMuc & "," & MaVach & ", "
    sSQLVal = sSQLVal & matkhai & "," & HANNOP & "," & ThueTKy2 & "," & HANNOP2 & "," & TENDA & ",'" & NGANH & "'," & DAGHI1 & "," & DAGHI2 & "," & AUTO_CAL & "," & LAN_QUET & "," & DANHAN & "," & madlt & "," & LanBS & "," & MANNKD & "," & THUECL & "," & DSoCL & "," & GTTTCL & "," & NgayPs & "," & GiaHan & "," & MALD & ",'" & TENNV & "','" & CHUNGCHI & "'"
       
    sSQL = "INSERT INTO TMP_TK" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    InsertHDR = sSQL
    
End Function

Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    
    
        strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
        strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_KT" & sKyKeKhai & ".DBF"
        
        Dim strFileNameBS As String
        strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"
        'file HDR
        If fso.FileExists(strFileNameHDR) = False Then
            Dim strTepmauHDR As String
            strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_TK.DBF"
            fso.CopyFile strTepmauHDR, strFileNameHDR, False
        End If
        'file dtl
        If fso.FileExists(strFileNameDTL) = False Then
            Dim strTepmauDTL As String
            strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_KT.DBF"
            fso.CopyFile strTepmauDTL, strFileNameDTL, False
        End If
        'file phu luc
        Dim strFileNamePL1 As String
        strFileNamePL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_P1" & sKyKeKhai & ".DBF"
        If fso.FileExists(strFileNamePL1) = False Then
            Dim strTepmauPL1 As String
            strTepmauPL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_P1.DBF"
            fso.CopyFile strTepmauPL1, strFileNamePL1, False
        End If
        'file phu luc
        Dim strFileNamePL3 As String
        strFileNamePL3 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_PL03" & sKyKeKhai & ".DBF"
        If fso.FileExists(strFileNamePL3) = False Then
            Dim strTepmauPL3 As String
            strTepmauPL3 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_PL03.DBF"
            fso.CopyFile strTepmauPL3, strFileNamePL3, False
        End If
        'file phu luc
        Dim strFileNamePL4A As String
        strFileNamePL4A = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_P4" & sKyKeKhai & ".DBF"
        If fso.FileExists(strFileNamePL4A) = False Then
            Dim strTepmauPL4A As String
            strTepmauPL4A = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_P4.DBF"
            fso.CopyFile strTepmauPL4A, strFileNamePL4A, False
        End If
        'file phu luc
        Dim strFileNamePL4B As String
        strFileNamePL4B = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_P5" & sKyKeKhai & ".DBF"
        If fso.FileExists(strFileNamePL4B) = False Then
            Dim strTepmauPL4B As String
            strTepmauPL4B = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_P5.DBF"
            fso.CopyFile strTepmauPL4B, strFileNamePL4B, False
        End If
        'file phu luc
        
        Dim strFileNamePL6 As String
        strFileNamePL6 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_P6" & sKyKeKhai & ".DBF"
        If fso.FileExists(strFileNamePL6) = False Then
            Dim strTepmauPL6 As String
            strTepmauPL6 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_P6.DBF"
            fso.CopyFile strTepmauPL6, strFileNamePL6, False
        End If
        
        'file phu luc
        
        Dim strFileNamePL016 As String
        strFileNamePL016 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_P016" & sKyKeKhai & ".DBF"
        If fso.FileExists(strFileNamePL016) = False Then
            Dim strTepmauPL016 As String
            strTepmauPL016 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_P016.DBF"
            fso.CopyFile strTepmauPL016, strFileNamePL016, False
        End If
                'file phu luc
        
'        Dim strFileNamePL016B As String
'        strFileNamePL016B = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tmp_p016B" & sKyKeKhai & ".DBF"
'        If fso.FileExists(strFileNamePL016B) = False Then
'            Dim strTepmauPL016B As String
'            strTepmauPL016B = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tepmau\tmp_p016B.DBF"
'            fso.CopyFile strTepmauPL016B, strFileNamePL016B, False
'        End If
                'file phu luc
        
        Dim strFileNamePL017 As String
        strFileNamePL017 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_P017" & sKyKeKhai & ".DBF"
        If fso.FileExists(strFileNamePL017) = False Then
            Dim strTepmauPL017 As String
            strTepmauPL017 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tepmau\TMP_P017.DBF"
            fso.CopyFile strTepmauPL017, strFileNamePL017, False
        End If
        'file bo xung
        If fso.FileExists(strFileNameBS) = False Then
            Dim strTepmauBS As String
            strTepmauBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_BS.DBF"
            fso.CopyFile strTepmauBS, strFileNameBS, False
        End If

        isExistFile = True
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   Dim vKyKK As String
   If isTkQuy = True Then
        vKyKK = Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year
   Else
        vKyKK = TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
   End If
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyKeKhai & _
        " WHERE MATHUE = '01' AND  MAPP = '1' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & vKyKK & "'" & _
        " AND TTHTK = '1' AND MATKHAI = '" & LOAI_TO_KHAI & "'"
    
   Set rs = clsDAO.Execute(sSQL)
   If Not rs Is Nothing Then
        dNgayNopDB = rs.Fields("NGNOP")
        If Trim(dNgayNopDB) = vbNullString Then
            dNgayNopDB = "CTOD('')"
        Else
            dNgayNopDB = "CTOD('" & Format(dNgayNopDB, "mm/dd/yyyy") & "')"
        End If
        TKTT = True
   Else
        TKTT = False
   End If
        LAN_QUET = TinhSoLanQuet + 1
      
   clsDAO.Disconnect
    
End Function


Public Function UpdateTHUETKY2() As Boolean
    Dim sSQL As String
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   Dim vKyKK As String
   If isTkQuy = True Then
        vKyKK = Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year
   Else
        vKyKK = TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
   End If
   
   sSQL = "UPDATE TMP_TK" & sKyKeKhai & _
        " SET THUETKY2=1" & _
        " WHERE MATHUE = '01' AND MAPP = '1' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & vKyKK & "'" & _
        " AND MATKHAI = '" & LOAI_TO_KHAI & "'"
    
   clsDAO.ExecuteDLL sSQL
   
   UpdateTHUETKY2 = True
   
   clsDAO.Disconnect

End Function

Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim KYLBO As Variant
    Dim rs As ADODB.Recordset
    
    If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
   Dim vKyKK As String
   If isTkQuy = True Then
        vKyKK = Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year
   Else
        vKyKK = TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
   End If
   
   sSQL = "DELETE FROM TMP_TK" & sKyKeKhai & _
        " WHERE MATHUE = '01' AND MAPP = '1' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & vKyKK & "'" & _
        " AND MATKHAI = '" & LOAI_TO_KHAI & "'"
    
   clsDAO.ExecuteDLL sSQL
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 10, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If
   End With
   sSQL = "DELETE FROM KT" & sKyKeKhai & _
        " WHERE MATHUE = '01' AND MAPP = '1' AND MADTNT = " & "'" & strMST & "' AND NGNOP = " & dNgayNopDB & _
        " AND MATKHAI='" & LOAI_TO_KHAI & "' AND KYKKHAI ='" & vKyKK & "'"
    
   clsDAO.ExecuteDLL sSQL
   
   'Xoa phu luc PL3
   sSQL = "DELETE FROM TMP_PL03" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "' AND NGNOP = " & dNgayNopDB & _
        " AND MATKHAI='" & LOAI_TO_KHAI & "' AND KYKKHAI ='" & vKyKK & "'"
   
   clsDAO.ExecuteDLL sSQL
   
   ' Xoa phu luc PL5
   sSQL = "DELETE FROM TMP_P6" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & _
        " AND MATKHAI='" & LOAI_TO_KHAI & "' AND KYKKHAI ='" & vKyKK & "'"
   clsDAO.ExecuteDLL sSQL
   
   
   XoaTKTT = True
   
   clsDAO.Disconnect

End Function

Public Function TKRB() As Boolean

    Dim sSQL As String
    Dim THUEGTGT_KT As Double
    Dim rs As ADODB.Recordset
    Dim sKyTruoc As String
    Dim sKyKKTruoc As String
    Dim fso As New FileSystemObject
    Dim THUEGTGT As Variant
    Dim dNgayNopDBKT As Variant
    
    If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
   If CInt(Month(Now())) = 1 Then
        sKyTruoc = "12" & Trim(str(CInt(Year(Now())) - 1))
   Else
        sKyTruoc = IIf(Len(Trim(str(CInt(Month(Now())) - 1))) = 1, "0" & Trim(str(CInt(Month(Now())) - 1)), Trim(str(CInt(Month(Now())) - 1))) & Year(Now())
   End If
   
   If CInt(TAX_Utilities_Svr_New.Month) = 1 Then
        sKyKKTruoc = "12" & TAX_Utilities_Svr_New.Year
   Else
        sKyKKTruoc = IIf(Len(Trim(str(CInt(TAX_Utilities_Svr_New.Month) - 1))) = 1, "0" & Trim(str(CInt(TAX_Utilities_Svr_New.Month) - 1)), Trim(str(CInt(TAX_Utilities_Svr_New.Month) - 1))) & TAX_Utilities_Svr_New.Year
   End If
   
   If Not fso.FileExists(spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyTruoc & ".dbf") Then
        TKRB = True
        clsDAO.Disconnect
        Exit Function
   End If
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyTruoc & _
        " WHERE MATHUE = '01' AND MAPP = '1' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & Left(sKyKKTruoc, 2) & "/" & Right(sKyKKTruoc, 4) & "'"
    
   Set rs = clsDAO.Execute(sSQL)
   If Not rs Is Nothing Then
        dNgayNopDBKT = rs.Fields("NGNOP")
        If Trim(dNgayNopDBKT) = vbNullString Then
            dNgayNopDBKT = "CTOD('')"
        Else
            dNgayNopDBKT = "CTOD('" & Format(dNgayNopDBKT, "mm/dd/yyyy") & "')"
        End If
   Else
        TKRB = True
        clsDAO.Disconnect
        Exit Function
   End If
   If Not fso.FileExists(spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "KT" & sKyTruoc & ".dbf") Then
        TKRB = True
        clsDAO.Disconnect
        Exit Function
   End If
   
   sSQL = "SELECT thuegtgt FROM KT" & sKyTruoc & _
        " WHERE CT_KT = '085' AND MADTNT = " & "'" & strMST & "' AND NGNOP = " & dNgayNopDBKT
    
   Set rs = clsDAO.Execute(sSQL)
   If Not rs Is Nothing Then
        THUEGTGT_KT = rs.Fields("thuegtgt")
        If THUEGTGT_KT = 0 Then
            TKRB = True
            clsDAO.Disconnect
            Exit Function
        Else
            With fps
                .Sheet = 1
                .GetText .ColLetterToNumber("W"), 39, THUEGTGT
                If Trim(THUEGTGT) = vbNullString Then
                    THUEGTGT = "0"
                End If
                If THUEGTGT_KT = THUEGTGT Then
                     TKRB = True
                     clsDAO.Disconnect
                     Exit Function
                Else
                      TKRB = False
                      clsDAO.Disconnect
                      Exit Function
                End If
            End With
        End If
   Else
        TKRB = True
        clsDAO.Disconnect
        Exit Function
   End If
   
   TKRB = True
   clsDAO.Disconnect

End Function

Public Function TKTTNGNOP() As Boolean
   Dim ngnop As Variant
   Dim NGNOPDB As Variant
   Dim sSQL As String
   Dim THUEGTGT_KT As Double
   Dim rs As ADODB.Recordset
   Dim i As Integer
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 12, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = ""
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
   End With
   
   If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
   Dim vKyKK As String
   If isTkQuy = True Then
        vKyKK = Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year
   Else
        vKyKK = TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
   End If
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyKeKhai & _
        " WHERE MATHUE = '01' AND MAPP = '1' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI <> '" & vKyKK & "'" & _
        " AND MATKHAI='" & LOAI_TO_KHAI & "'"
   Set rs = clsDAO.Execute(sSQL)
   If Not rs Is Nothing Then
        For i = 1 To rs.RecordCount
            If Not rs.EOF Then
                NGNOPDB = rs.Fields("NGNOP")
                If Trim(NGNOPDB) = vbNullString Then
                    NGNOPDB = ""
                Else
                    NGNOPDB = "CTOD('" & Format(NGNOPDB, "mm/dd/yyyy") & "')"
                End If
                If Trim(NGNOPDB) <> vbNullString And Trim(NGNOPDB) = Trim(ngnop) Then
                    clsDAO.Disconnect
                    TKTTNGNOP = False
                    Exit Function
                Else
                    TKTTNGNOP = True
                End If
            End If
            rs.MoveNext
        Next
   Else
        clsDAO.Disconnect
        TKTTNGNOP = True
   End If

End Function

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    Dim intIndexCombo As Integer
    With fps
       If .ActiveSheet = 1 Then
            .Col = Col
            .Row = Row
             If Col = .ColLetterToNumber("E") And Row = 41 Then
                intIndexCombo = .TypeComboBoxCurSel
                .Col = .ColLetterToNumber("AB")
                .TypeComboBoxCurSel = intIndexCombo
                'Update combo C
                UpdateCell fps, .Col, .Row, .Text
            End If
        End If
        
    End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
 Dim varTemp As Variant
 Dim vNgayQyet As Variant
 With fps
    If .ActiveSheet = 1 Then
        If Col = .ColLetterToNumber("E") And Row = 32 Then
            .GetText Col, Row, varTemp
            .GetText .ColLetterToNumber("M"), Row, vNgayQyet
            If varTemp <> "" And varTemp <> "../../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                     checkNgayNop varTemp, vNgayQyet
                Else
                    .SetActiveCell Col, Row
                End If
            Else
             .SetActiveCell Col, Row
            End If
'           UpdateCell fps, .Col, .Row, .Text
        End If

        If Col = .ColLetterToNumber("E") And Row = 30 Then
            .GetText Col, Row, varTemp
            If varTemp <> "" And varTemp <> "../...." Then
                If Format_mmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_mmyyyy(CStr(varTemp))
                Else
                    .SetActiveCell Col, Row
                End If
            Else
             .SetActiveCell Col, Row
            End If
           UpdateCell fps, .Col, .Row, .Text
        End If
    End If
End With
End Sub
Public Function TinhSoLanQuet() As Integer
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim LAN_QUET1 As Variant
    
    Dim vKyKK As String
    
    'kiem tra ton tai tep *.dbf chua
    'Tinh so lan quyet de ghi vao DB trung gian
    'dhdang sua
    'ngay 24/09/2010
'-----------------------------
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        End If
       
            If isTkQuy = True Then
                vKyKK = Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year
            Else
                vKyKK = TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
            End If
       
           sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM  TMP_TK" & sKyKeKhai & _
                " WHERE MATHUE = '01' AND MADTNT = " & "'" & strMST & "'" & _
                " AND KYKKHAI = '" & vKyKK & "'" & _
                " AND MATKHAI = '" & LOAI_TO_KHAI & "'"
                              
             Set rs = clsDAO.Execute(sSQL)
            If Not rs Is Nothing Then
                 LAN_QUET1 = rs.Fields("LAN_QUET")
                 TinhSoLanQuet = LAN_QUET1
            Else
                 TinhSoLanQuet = 0
            End If
            clsDAO.Disconnect
    Else
    TinhSoLanQuet = 0
    End If
End Function
'dhdang tinh ma nganh
'ngay 25/10/10
Public Function TinhMaNganh() As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim NGANH As Variant
    Dim TEMP As New clsADO

        If TEMP.Connected = False Then
            TEMP.CreateConnectionString spathVat & "\dtnt\"
            TEMP.Connect
         End If
         sSQL = "SELECT nganh FROM dtnt2 WHERE MADTNT =  '" & strMST & "'"
        
        Set rs = TEMP.Execute(sSQL)
        If Not rs Is Nothing Then
            NGANH = rs.Fields("nganh")
               
            TinhMaNganh = NGANH
        Else
            TinhMaNganh = 0
        End If
        TEMP.Disconnect
End Function

Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim vKeKhai As Variant
    Dim vKeKhai_temp As Variant
    Dim vKyLB As Variant
    'them bien de kiem tra su ton tai cua bang dl
    Dim strTestFileName As String
    Dim fso As New FileSystemObject
    
    Dim vKyKK As String
    
    Dim flag As Boolean
    'kiem tra da co to khai chinh thuc chua
    flag = True
    vKyLB = DateSerial(Int(Year(Now())), Int(Month(Now())), 1)
    vKeKhai = DateSerial(Int(Val(TAX_Utilities_Svr_New.Year)), Int(Val(TAX_Utilities_Svr_New.Month)), 1)
    vKeKhai_temp = sKyKeKhai
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   Do While flag
        strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & vKeKhai_temp & ".DBF"
        'kiem tra neu ton tai bang thi moi thuc hien truy van
        If isTkQuy = True Then
            vKyKK = Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year
        Else
            vKyKK = TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
        End If
        
        If fso.FileExists(strTestFileName) Then
             sSQL = "SELECT * FROM TMP_TK" & vKeKhai_temp & _
             " WHERE MATHUE = '01' AND  MADTNT = " & "'" & strMST & "'" & _
             " AND KYKKHAI = '" & vKyKK & "'" & _
             " AND (TTHTK = '1' or TTHTK = '3' or TTHTK = '4') AND MATKHAI = '" & LOAI_TO_KHAI & "'"
        
             Set rs = clsDAO.Execute(sSQL)
        End If
        If rs Is Nothing Then
             isToKhaiChinhThuc = False
        Else
             isToKhaiChinhThuc = True
             Exit Do
        End If
        
        If vKeKhai < vKyLB Then
            vKyLB = DateAdd("M", -1, vKyLB)
            ' ghep ten bang de kiem tra du lieu
            If Len(Month(vKyLB)) = 1 Then
                vKeKhai_temp = "0" & Month(vKyLB) & Year(vKyLB)
            Else
                vKeKhai_temp = Month(vKyLB) & Year(vKyLB)
            End If
            flag = True
        Else
            flag = False
        End If
   Loop
   clsDAO.Disconnect
End Function


Private Function checkNgayNop(ByVal vNgayNop As Variant, ByVal vNgayQuet As Variant)
    Dim varNgayNop As Variant
    Dim varNgayQuet As Variant
    On Error GoTo e
    
    varNgayNop = DateSerial(Int(Right(CStr(vNgayNop), 4)), Int(Mid(CStr(vNgayNop), 4, 2)), Int(Left(CStr(vNgayNop), 2)))
    varNgayQuet = DateSerial(Int(Right(CStr(vNgayQuet), 4)), Int(Mid(CStr(vNgayQuet), 4, 2)), Int(Left(CStr(vNgayQuet), 2)))

    If Year(vNgayNop) > Year(vNgayQuet) Then
        GoTo e
    ElseIf Month(vNgayNop) > Month(vNgayQuet) Then
        GoTo e
    ElseIf Day(vNgayNop) > Day(vNgayQuet) Then
        GoTo e
    End If
    Exit Function
e:
    DisplayMessage "0144", msOKOnly, miCriticalError

End Function

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
    strDuongDanFile = "..\InterfaceTemplates\Catalogue_DM_Lydo_GH_TT16.xml"
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
