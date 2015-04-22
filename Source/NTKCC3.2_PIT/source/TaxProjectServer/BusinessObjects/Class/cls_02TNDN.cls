VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_02TNDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_Y = 42
Const KY_LAP_BO_X = "E"
Const NGAY_NHAN_TO_KHAI_Y = 25
Const NGAY_NHAN_TO_KHAI_X = "R"
Const NGUOI_SU_DUNG_Y = 37
Const NGUOI_SU_DUNG_X = "AF"
Const PHONG_XU_LY_Y = 42
Const PHONG_XU_LY_X = "V"
Const NGAY_NOP_Y = 44
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 44
Const NGAY_QUET_X = "M"
'Const TEN_BPQL_X = "V"
'Const TEN_BPQL_Y = 34
Const MA_BPQL_X = "V"
Const MA_BPQL_Y = 46

Const STT_X = "M"
Const STT_Y = 42

Private rs As ADODB.Recordset

' ----- thong tin din vi  ------
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

'---  thong tin Dai ly thue ----
Const MA_SO_THUE_DL_X = "F"
Const MA_SO_THUE_DL_Y = 32
Const TEN_GOI_DL_X = "H"
Const TEN_GOI_DL_Y = 118
Const DIA_CHI_DL_X = "F"
Const DIA_CHI_DL_Y = 34
Const QUAN_HUYEN_DL_X = "F"
Const QUAN_HUYEN_DL_Y = 36
Const TINH_TPHO_DL_X = "Q"
Const TINH_TPHO_DL_Y = 36
Const DIEN_THOAI_DL_X = "F"
Const DIEN_THOAI_DL_Y = 38
Const FAX_DL_X = "N"
Const FAX_DL_Y = 38
Const MAIL_DL_X = "V"
Const MAIL_DL_Y = 38
Const SO_HDDL_DL_X = "H"
Const SO_HDDL_DL_Y = 40
Const NGAY_HDDL_DL_X = "N"
Const NGAY_HDDL_DL_Y = 40

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
Public TTHTK As String
Private LOAI_TO_KHAI As String


'end

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
Public isExistFile As Boolean
Public Lan_quet As Variant


'Public strMaSoTep As String
Public strNgayNhanToKhai As String
Public strNguoiSuDung As String
Public strPhongXuLy As String
Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long

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

Public spathVat As String
Public sSaiCT11 As String
Public sKyKeKhai As String
Public dNgayNopDB As Variant
Public HANNOP As String
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
Private TKTHANG As Variant
Private NGAY_PS As Variant

Public Function Prepared1() As Boolean
   With fps
        .Sheet = 1
        'Ghi chu
        .Col = .ColLetterToNumber("M")
        .Row = 48
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100

        ' Ky lap bo
        SetDateFormat fps, 1, 42, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft

        'Phong xu ly
        .Col = .ColLetterToNumber("V")
        .Row = 42
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60

        ' Ngay nop
        SetDateFormat fps, 1, 44, .ColLetterToNumber("E"), DDMMYYYY
        .Sheet = 1
        .Row = 44
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft

        ' Ngay quet
        SetDateFormat fps, 1, 44, .ColLetterToNumber("M"), DDMMYYYY
        .Sheet = 1
        .Row = 44
        .Col = .ColLetterToNumber("M")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
    End With
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
    Dim i As Integer
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
    End With
    Prepared2 = True
End Function
Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String) 'ByRef strPMaSoTep, ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
' If fps.ActiveCol = fps.ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) And fps.ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then _
'        CellEditFormatNumber fps, fps.ActiveSheet, fps.ActiveCol, fps.ActiveRow, KeyAscii
'    With fps
'        If .ActiveSheet = 1 And .ActiveCol = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) _
'           And .ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then
'            If KeyAscii = 46 Or KeyAscii = 45 Then ' KeyAscii of '.' character
'                KeyAscii = 0
'            End If
'        End If
'    End With
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

Public Function Prepared3() As Boolean

    With fps

        .Sheet = 1

'        'set check loi dinh danh
'        If checkDLTTonTai = False Then
'            .Col = .ColLetterToNumber("E")
'            .Row = 48
'            .Value = 1
'        End If

        'Set NgayNhanToKhai
        '        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
        '        .Row = NGAY_NHAN_TO_KHAI_Y
        '        If strNgayNhanToKhai <> "" Then
        '            .Text = strNgayNhanToKhai
        '            UpdateCell fps, .Col, .Row, .Text
        '        End If

        .Col = .ColLetterToNumber(MA_BPQL_X)
        .Row = MA_BPQL_Y

        If strMaBPQL <> "" Then
            .Text = strMaBPQL
            UpdateCell fps, .Col, .Row, .Text
        End If

        '        .Col = .ColLetterToNumber(TEN_BPQL_X)
        '        .Row = TEN_BPQL_Y
        '        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))

        'set Phong xu ly
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

        ' them thong tin dai ly thue
        'ten
        .Col = .ColLetterToNumber(TEN_GOI_DL_X)
        .Row = TEN_GOI_DL_Y

        If strTen_DLT <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strTen_DLT, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If

        'mst
        .Col = .ColLetterToNumber(MA_SO_THUE_DL_X)
        .Row = MA_SO_THUE_DL_Y

        If strMST_DLT <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strMST_DLT, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If

        'dia chi
        .Col = .ColLetterToNumber(DIA_CHI_DL_X)
        .Row = DIA_CHI_DL_Y

        If strDchi_DLT <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strDchi_DLT, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If

        ' quan huyen
        .Col = .ColLetterToNumber(QUAN_HUYEN_DL_X)
        .Row = QUAN_HUYEN_DL_Y

        If strQHuyen_DLT <> "" Then
            .Text = strQHuyen_DLT
        End If

        ' thanh pho
        .Col = .ColLetterToNumber(TINH_TPHO_DL_X)
        .Row = TINH_TPHO_DL_Y

        If strTTPho_DLT <> "" Then
            .Text = strTTPho_DLT
        End If

        'dien thoai
        .Col = .ColLetterToNumber(DIEN_THOAI_DL_X)
        .Row = DIEN_THOAI_DL_Y

        If strDthoai_DLT <> vbNullString Then
            .Text = strDthoai_DLT
        End If

        'fax
        .Col = .ColLetterToNumber(FAX_DL_X)
        .Row = FAX_DL_Y

        If strFax_DLT <> "" Then
            .Text = strFax_DLT
        End If

        ' mail
        .Col = .ColLetterToNumber(MAIL_DL_X)
        .Row = MAIL_DL_Y

        If strMail_DLT <> "" Then
            .Text = strMail_DLT
        End If

        ' so hop dong
        .Col = .ColLetterToNumber(SO_HDDL_DL_X)
        .Row = SO_HDDL_DL_Y

        If strSoHD_DLT <> "" Then
            .Text = strSoHD_DLT
        End If

        ' ngay hop dong
        .Col = .ColLetterToNumber(NGAY_HDDL_DL_X)
        .Row = NGAY_HDDL_DL_Y

        If strNgayHD_DLT <> "" Then
            .Text = strNgayHD_DLT
        End If

'        If Len(Month(Now())) = 1 Then
'            sKyKeKhai = "0" & Month(Now()) & Year(Now())
'        Else
'            sKyKeKhai = Month(Now()) & Year(Now())
'        End If

        'to khai thang
        Dim tokhai As Variant
        .GetText .ColLetterToNumber("Q"), 49, tokhai

        If tokhai = "1" Then
            TKTHANG = False
        Else
            TKTHANG = True
        End If

        ' so thu tu to khai
        .Col = .ColLetterToNumber(STT_X)
        .Row = STT_Y
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))

        ' ngay phat sinh
        .GetText .ColLetterToNumber("M"), 50, NGAY_PS

        If Trim(NGAY_PS) = vbNullString Then
            LOAI_TO_KHAI = "02/TNDN"
            NGAY_PS = "CTOD('')"
        Else
            TAX_Utilities_Svr_New.Month = Mid$(NGAY_PS, 4, 2)
            TAX_Utilities_Svr_New.Year = Right$(NGAY_PS, 4)
            TAX_Utilities_Svr_New.ThreeMonths = ""
            NGAY_PS = ToDate(Trim(NGAY_PS), DDMMYYYY)
            NGAY_PS = "CTOD('" & Format(NGAY_PS, "mm/dd/yyyy") & "')"
            LOAI_TO_KHAI = "02PS/TNDN"
        End If
        
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

        ' Get MaSoTep
'        .Col = .ColLetterToNumber(MA_SO_TEP_X)
'        .Row = MA_SO_TEP_Y
'        strMaSoTep = .Text

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

'        If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
'            blnValid = False
'        End If
'
'        .Col = .ColLetterToNumber(HEADER_KY_LAP_BO_X)
'        .Row = HEADER_KY_LAP_BO_Y
'
'        If Not blnValid Then
'            .Formula = "0"
'        Else
'            .Formula = "1"
'        End If
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

Private Sub fps_KeyPress(KeyAscii As Integer)
'    If fps.ActiveCol = fps.ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) And fps.ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then _
'        CellEditFormatNumber fps, fps.ActiveSheet, fps.ActiveCol, fps.ActiveRow, KeyAscii
'    With fps
'        If .ActiveSheet = 1 And .ActiveCol = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) _
'           And .ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then
'            If KeyAscii = 46 Or KeyAscii = 45 Then ' KeyAscii of '.' character
'                KeyAscii = 0
'            End If
'        End If
'    End With
End Sub

Public Function CheckValidData() As Boolean
    Dim NGAYNOP As Variant, ngayHienTai As Variant
    
    CheckValidData = True
    With fps
        .Sheet = 1
        
        ' check ngay nop khong dc lon hon ngay hien tai
        .GetText .ColLetterToNumber("E"), 44, NGAYNOP
        ngayHienTai = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        ngayHienTai = DateSerial(Int(Mid$(ngayHienTai, 7, 4)), Int(Mid$(ngayHienTai, 4, 2)), Int(Mid$(ngayHienTai, 1, 2)))
        NGAYNOP = DateSerial(Int(Mid$(NGAYNOP, 7, 4)), Int(Mid$(NGAYNOP, 4, 2)), Int(Mid$(NGAYNOP, 1, 2)))
        If NGAYNOP > ngayHienTai Then
            CheckValidData = False
        End If

    End With
End Function



Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
    KiemTraKhoaSo = True
    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("E_10"), "Value"), "dd/mm/yyyy")
    
    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function

Public Function InsertDTL() As String
    Dim sSQL      As String
    Dim sSQLCol   As String
    Dim sSQLVal   As String
    Dim bln       As Boolean

    Dim madtnt    As Variant
    Dim MaMuc     As Variant
    Dim maTM      As Variant
    Dim ngnop     As Variant
    Dim MAHIEU    As Variant
    Dim giatri    As Variant
    Dim TTHTK     As Variant
    Dim QUI   As Variant ' Thay the truong kykkhai = QUI theo TT151
    Dim matkhai   As Variant
    Dim giatri2   As Variant
    Dim GIATRICL  As Variant ' Bo sung them theo TT151
    Dim tt        As Variant
    Dim KD_MADTNT As Variant
    Dim danhan    As Variant

    sSQLCol = "MADTNT,MAMUC,MATM,NGNOP,TTHTK,MAHIEU,GIATRI,QUI,MATKHAI,GIATRI2,TT,KD_MADTNT,DANHAN,LAN_QUET,GIATRICL"

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If

    With fps

        .Sheet = 1
        '------------------ Thong tin chung -------------------
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MaMuc = "'1050'"
        maTM = "'1054'"
        'ma so thue don vi van chuyen
        .GetText .ColLetterToNumber("R"), 22, KD_MADTNT

        If Trim(KD_MADTNT) = vbNullString Then
            KD_MADTNT = "''"
        Else
            KD_MADTNT = "'" & KD_MADTNT & "'"
        End If

        'trang thai to khai
        Dim TTHTK1 As Variant

        .GetText .ColLetterToNumber("M"), 49, TTHTK1

        If Trim(TTHTK1) = vbNullString Then
            TTHTK = "'" & 2 & "'"
        Else
            TTHTK = "'" & 1 & "'"
        End If

        ' ngay nop
        .GetText .ColLetterToNumber("E"), 44, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'ma to khai
        matkhai = "'" & LOAI_TO_KHAI & "'"
        danhan = "''"
        'ky ke khai thay the bang QUI theo TT151
        If NGAY_PS <> "CTOD('')" Then
            QUI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            QUI = "'" & TAX_Utilities_Svr_New.ThreeMonths & "/" & TAX_Utilities_Svr_New.Year & "'"
            QUI = "'" & Right$(QUI, 7)
        End If
        tt = "' '"
        giatri2 = "0"
        GIATRICL = "0" ' Bo sung theo TT151
        
        '--- End thong tin chung -------------------------------
        .Row = 55
        Dim i As Integer
        i = 0
        Dim arrMaHieu() As Variant
           
        arrMaHieu() = Array("002", "003", "004", "005", "006", "007", "008", "009", "010", "011", "012")

        Do
            If .Row <> 64 Then
                If .Row <> 65 Then
                    MAHIEU = "'" & arrMaHieu(i) & "'"
        
                    .GetText .ColLetterToNumber("V"), .Row, giatri
                    
                    If Trim(giatri) = vbNullString Then
                        giatri = "0"
                    End If
        
                    sSQLVal = madtnt & "," & MaMuc & "," & maTM & "," & ngnop & "," & TTHTK & "," & MAHIEU & "," & giatri & ", "
                    sSQLVal = sSQLVal & QUI & "," & matkhai & "," & giatri2 & "," & tt & "," & KD_MADTNT & "," & danhan & "," & Lan_quet & "," & GIATRICL
        
                    sSQL = "INSERT INTO TMP_SDD" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                    bln = clsDAO.ExecuteDLL(sSQL)
                
                i = i + 1
                End If
            End If
            .Row = .Row + 1
        Loop Until .Row = 68

    End With

    clsDAO.Disconnect
     
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL1
    End If
    
    InsertDTL = vbNullString
End Function

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
    Dim SONGAYNC    As Variant
    Dim SOTIENNC    As Variant
    Dim danhan      As Variant
    
    Dim i           As Integer, iCol As Long, iRow As Long
    Dim xmlNode     As MSXML.IXMLDOMNode

    Dim strFileName As String
    Dim fso         As New FileSystemObject

    'kiem tra ton tai tep *.dbf chua
    strFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"

    If fso.FileExists(strFileName) = False Then
        Dim strTepmau As String
        strTepmau = spathVat & "\tepmau\TMP_BS.DBF"
        fso.CopyFile strTepmau, strFileName, False
    End If

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

        If NGAY_PS <> "CTOD('')" Then
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.ThreeMonths & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

        MaMuc = "'1050'"
        MATHUE = "' '"
        MAPP = "'2'"
        matkhai = "'" & LOAI_TO_KHAI & "'"
        maTM = "'1054'"
        
        danhan = "''"
        
        .GetText .ColLetterToNumber("E"), 44, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 42, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
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
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & Lan_quet & "," & danhan
                       
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        .Sheet = .SheetCount - 1
        
        'get so tien nop cham va so ngay nop cham
        .GetText .ColLetterToNumber("BE"), .MaxRows - 27, SOTIENNC

        If SOTIENNC = vbNullString Then
            SOTIENNC = "0"
        Else
            SOTIENNC = Replace(SOTIENNC, ".", "")
        End If

        .GetText .ColLetterToNumber("BE"), .MaxRows - 28, SONGAYNC

        If SONGAYNC = vbNullString Then
            SONGAYNC = "0"
        Else
            SONGAYNC = Replace(SONGAYNC, ".", "")
        End If

        'end
        
        i = 1
        .Row = 9

        Do
            mact = "''"
            .GetText .ColLetterToNumber("BE"), .Row, MACT2

            If Trim(MACT2) = vbNullString Then
                MACT2 = "''"
            ElseIf MACT2 = "14" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "16" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "23" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "24" Then
                MACT2 = "'050'"
            ElseIf MACT2 = "21" Then
                MACT2 = "'039'"
            ElseIf MACT2 = "31" Then
                MACT2 = "'050'"
            ElseIf MACT2 = "30" Then
                MACT2 = "'031'"
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
            
            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & Lan_quet & "," & danhan
                           
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
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & Lan_quet & "," & danhan
                       
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

        i = 1
        .Row = .Row + 3

        Do
            mact = "''"
            .GetText .ColLetterToNumber("BE"), .Row, MACT2

            If Trim(MACT2) = vbNullString Then
                MACT2 = "''"
            ElseIf MACT2 = "14" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "16" Then
                MACT2 = "'041'"
                '                ElseIf MACT2 = "23" Then
                '                    MACT2 = "'041'"
            ElseIf MACT2 = "24" Then
                MACT2 = "'050'"
            ElseIf MACT2 = "21" Then
                MACT2 = "'039'"
            ElseIf MACT2 = "31" Then
                MACT2 = "'050'"
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
            
            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & Lan_quet & "," & danhan
                           
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
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & Lan_quet & "," & danhan
                       
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        .Row = .Row + 3
        mact = "''"
        .GetText .ColLetterToNumber("BE"), .Row, MACT2

        If Trim(MACT2) = vbNullString Then
            MACT2 = "''"
        ElseIf MACT2 = "14" Then
            MACT2 = "'041'"
        ElseIf MACT2 = "16" Then
            MACT2 = "'041'"
            '                ElseIf MACT2 = "23" Then
            '                    MACT2 = "'041'"
        ElseIf MACT2 = "24" Then
            MACT2 = "'050'"
        ElseIf MACT2 = "21" Then
            MACT2 = "'039'"
        ElseIf MACT2 = "31" Then
            MACT2 = "'050'"
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
            
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & Lan_quet & "," & danhan
                           
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        If MACT2 <> "''" Then
            bln = clsDAO.ExecuteDLL(sSQL)
        End If

    End With
   
    clsDAO.Disconnect
    
    InsertHDR_KHBS
    
    InsertDTL_KHBS = vbNullString
End Function

Public Function InsertHDR() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    Dim bln        As Boolean
    Dim TENNV      As Variant
    Dim CHUNGCHI   As Variant

    Dim madtnt     As Variant
    Dim MaMuc      As Variant
    Dim maTM       As Variant
    Dim KD_MADTNT  As Variant
    Dim TEN_TC     As Variant
    Dim DC_TC      As Variant
    Dim QD_TC      As Variant
    Dim HOPDONG    As Variant
    Dim ngnop      As Variant
    Dim TTHTK      As Variant
    Dim QUI    As Variant ' Thay the cho KYKKHAI theo TT151
    Dim matkhai    As Variant
    Dim HANNOP     As Variant
    Dim HANNOP2    As Variant
    Dim ThueTKy    As Variant
    Dim ThueTKy2   As Variant
    Dim NGAYHD     As Variant
    Dim NGAYXN     As Variant
    Dim GhiChu     As Variant
    Dim SHTEP      As Variant
    Dim MaVach     As Variant
    Dim AUTO_CAL   As Variant
    Dim DaGhi      As Variant
    Dim KYLBO      As Variant
    Dim danhan     As Variant
    Dim THUECL     As Variant
    Dim madlt      As Variant
    Dim lanbs      As Variant
    Dim DoanhSo    As Variant
    Dim GTTinhThue As Variant
    Dim GTTTCL     As Variant
    Dim DSoCL      As Variant
    Dim DAGHI2     As Variant
    Dim GIA_HAN    As Variant
    
    'Bo sung them theo TT151: bo truong ngay_ps, thay truong kykkhai = qui
    'Bo sung them 3 truong duoi
    Dim dn_nho    As Variant
    Dim dn_kk    As Variant
    Dim yn_da30    As Variant
 
'    sSQLCol = "MADTNT,MAMUC,MATM,KD_MADTNT,TEN_TC,DC_TC,QD_TC,HOPDONG,NGNOP,TTHTK,"
'    sSQLCol = sSQLCol & "KYKKHAI,MATKHAI,HANNOP,HANNOP2,THUETKY,THUETKY2,NGAYHD,NGAYXN,GHICHU,SHTEP,MAVACH,AUTO_CAL,DAGHI,KYLBO,THUECL,MADLT,LANBS,LAN_QUET,DoanhSo,GTTinhThue,GTTTCL,DSOCL,DaGhi2,DANHAN,NGAY_PS,TENNV,CHUNGCHI,GIAHAN"

    sSQLCol = "MADTNT,MAMUC,MATM,KD_MADTNT,TEN_TC,DC_TC,QD_TC,HOPDONG,NGNOP,TTHTK,"
    sSQLCol = sSQLCol & "QUI,MATKHAI,HANNOP,HANNOP2,THUETKY,THUETKY2,NGAYHD,NGAYXN,GHICHU,SHTEP,MAVACH,AUTO_CAL,DAGHI,KYLBO,THUECL,MADLT,LANBS,LAN_QUET,DoanhSo,GTTinhThue,GTTTCL,DSOCL,DaGhi2,DANHAN,NGAY_PS,TENNV,CHUNGCHI,GIAHAN,dn_nho,dn_kk,yn_da30"

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
       
    With fps
        .Sheet = 1
        'lay ten nhan vien dai ly thue, chung chi hanh nghe
        .GetText .ColLetterToNumber("G"), 78, TENNV
        If Trim(TENNV) = vbNullString Then
            TENNV = "''"
        Else
            TENNV = "'" & TAX_Utilities_Svr_New.Convert(Trim(TENNV), UNICODE, TCVN) & "'"
        End If
        
        .GetText .ColLetterToNumber("G"), 80, CHUNGCHI
        If Trim(CHUNGCHI) = vbNullString Then
            CHUNGCHI = "''"
        Else
            CHUNGCHI = "'" & TAX_Utilities_Svr_New.Convert(Trim(CHUNGCHI), UNICODE, TCVN) & "'"
        End If
        

        ' ma doi tuong not thue
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        ' loi dinh danh
        Dim loiDD As Variant
        .GetText .ColLetterToNumber("E"), 48, loiDD

        If loiDD = "0" Then
            DAGHI2 = ".F."
        Else
            DAGHI2 = ".T."
        End If

        DaGhi = ".F."

        'ma dai ly thue
        .GetText .ColLetterToNumber("F"), 32, madlt

        If Trim(madlt) = vbNullString Then
            madlt = "''"
        Else
            madlt = "'" & madlt & "'"
        End If
        
        MaMuc = "'1050'"
        maTM = "'1054'"
        
        'ma so thue don vi van chuyen
        .GetText .ColLetterToNumber("R"), 22, KD_MADTNT

        If Trim(KD_MADTNT) = vbNullString Then
            KD_MADTNT = "''"
        Else
            KD_MADTNT = "'" & KD_MADTNT & "'"
        End If

        'Gia han
        .GetText .ColLetterToNumber("E"), 47, GIA_HAN

        If GIA_HAN = "1" Then
            GIA_HAN = ".T."
        Else
            GIA_HAN = ".F."
        End If

        ' ten don vi van chuyen
        .GetText .ColLetterToNumber("K"), 20, TEN_TC

        If Trim(TEN_TC) = vbNullString Then
            TEN_TC = "''"
        Else
            TEN_TC = "'" & TAX_Utilities_Svr_New.Convert(Trim(TEN_TC), UNICODE, TCVN) & "'"
        End If

        'dia chi don vi van chuyen
        .GetText .ColLetterToNumber("E"), 24, DC_TC

        If Trim(DC_TC) = vbNullString Then
            DC_TC = "''"
        Else
            DC_TC = "'" & TAX_Utilities_Svr_New.Convert(Trim(DC_TC), UNICODE, TCVN) & "'"
        End If

        ' doanh so
        DoanhSo = "0"
        'so quyet dinh
        QD_TC = "''"
        'hop dong
        HOPDONG = "''"
        .GetText .ColLetterToNumber("K"), 26, HOPDONG

        If Trim(HOPDONG) = vbNullString Then
            HOPDONG = "''"
        Else
            HOPDONG = "'" & HOPDONG & "'"
        End If
        
        'trang thai to khai
        Dim TTHTK1 As Variant

        .GetText .ColLetterToNumber("M"), 49, TTHTK1

        If Trim(TTHTK1) = vbNullString Then
            TTHTK = "'" & 2 & "'"
        Else
            TTHTK = "'" & 1 & "'"
        End If
  
        ' ngay nop
        Dim NgNop2 As Variant
        .GetText .ColLetterToNumber("E"), 44, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            NgNop2 = ngnop
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
        'ngay hop dong
        .GetText .ColLetterToNumber("N"), 26, NGAYHD

        If Trim(NGAYHD) = vbNullString Then
            NGAYHD = "CTOD('')"
        Else
            NGAYHD = ToDate(Trim(NGAYHD), DDMMYYYY)
            NGAYHD = "CTOD('" & Format(NGAYHD, "mm/dd/yyyy") & "')"
        End If

        ' ngay xac nhan
        
        .GetText .ColLetterToNumber("Q"), 28, NGAYXN

        If Trim(NGAYXN) = vbNullString Then
            NGAYXN = "CTOD('')"
        Else
            NGAYXN = ToDate(Trim(NGAYXN), DDMMYYYY)
            NGAYXN = "CTOD('" & Format(NGAYXN, "mm/dd/yyyy") & "')"
        End If
        
        AUTO_CAL = ".F."
       
        'ma to khai
        matkhai = "'" & LOAI_TO_KHAI & "'"

        'Thay the kykkhai = QUI
        If NGAY_PS <> "CTOD('')" Then
            QUI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            QUI = "'" & TAX_Utilities_Svr_New.ThreeMonths & "/" & TAX_Utilities_Svr_New.Year & "'"
            QUI = Right$(QUI, 7)
            QUI = "'" & QUI
        End If
        
        ' ---------- han nop -----------------------
        ' HanNop = "CTOD('')"
        If NGAY_PS <> "CTOD('')" Then
            HANNOP = DateAdd("d", 10, NgNop2)
            HANNOP = ToString(HANNOP, "DD/MM/YYYY")
        ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2009" Then

            If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
                HANNOP = "01/02/2010"
            ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
                HANNOP = "04/05/2010"
            ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
                HANNOP = "30/07/2010"
            ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
                HANNOP = "01/11/2010"
            End If

        ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" Then

            If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
                HANNOP = "30/07/2010"
            ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
                HANNOP = "01/11/2010"
            ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
                HANNOP = "08/02/2011"
            ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
                'HANNOP = "03/05/2011"
                HANNOP = "30/01/2012"
            End If

        ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2011" Then

            If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
                HANNOP = "02/05/2012"
            ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
                HANNOP = "30/07/2012"
            ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
                HANNOP = "30/10/2012"
            ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
                HANNOP = "01/04/2013"
            End If

        ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2013" Then

            If GIA_HAN = ".T." Then
                If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
                    HANNOP = "30/10/2013"
                ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
                    HANNOP = "30/10/2013"
                ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
                    HANNOP = "31/01/2014"
                ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
                    HANNOP = "01/04/2015"
                End If

            Else

                If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
                    HANNOP = "30/04/2013"
                ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
                    HANNOP = "30/07/2013"
                ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
                    HANNOP = "30/10/2013"
                ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
                    HANNOP = "30/01/2014"
                End If
            End If
        End If

        If Trim(HANNOP) = vbNullString Then
            HANNOP = "CTOD('')"
        Else
            HANNOP = DateSerial(Int(Mid$(HANNOP, 7, 4)), Int(Mid$(HANNOP, 4, 2)), Int(Mid$(HANNOP, 1, 2)))

            If TTHTK = "'2'" And DateDiff("d", HANNOP, NgNop2) > 0 And (Trim(TAX_Utilities_Svr_New.Year) = "2011" Or (Trim(TAX_Utilities_Svr_New.Year) = "2010" And Val(TAX_Utilities_Svr_New.ThreeMonths) = 4)) Then
                HANNOP = NgNop2
            End If

            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
        End If

        HANNOP2 = HANNOP
        
        MaVach = ".T."
        
        ' ky lap bo
        .GetText .ColLetterToNumber("E"), 42, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If
        
        's0 hieu tep
        SHTEP = "''"
        
        'ghi chu
        .GetText .ColLetterToNumber("M"), 48, GhiChu

        If Trim(GhiChu) = vbNullString Then
            GhiChu = "''"
        Else
            GhiChu = "'" & TAX_Utilities_Svr_New.Convert(Trim(GhiChu), UNICODE, TCVN) & "'"
        End If
        
        ' lan bo sung
        .GetText .ColLetterToNumber("O"), 49, lanbs

        If Trim(lanbs) = vbNullString Then
            lanbs = "0"
        End If
        
        'Thue trong ky
        Dim ThueTKy11 As Variant
        Dim ThueTKy12 As Variant
        
        .GetText .ColLetterToNumber("V"), 69, ThueTKy11
        .GetText .ColLetterToNumber("V"), 75, ThueTKy12
        
        ThueTKy = ThueTKy11 + ThueTKy12
        ThueTKy2 = ThueTKy
        THUECL = "0"
        GTTinhThue = "0"
        GTTTCL = "0"
        DSoCL = "0"
        danhan = "''"
        
        'Bo sung theo TT151
        .GetText .ColLetterToNumber("O"), 49, dn_nho
        If Trim(dn_nho) = vbNullString Then
            dn_nho = ".F."
        Else
            dn_nho = ".T."
        End If
        
        .GetText .ColLetterToNumber("O"), 49, dn_kk
        If Trim(dn_kk) = vbNullString Then
            dn_kk = ".F."
        Else
            dn_kk = ".T."
        End If
        
        yn_da30 = "''"

                
        sSQLVal = madtnt & "," & MaMuc & "," & maTM & "," & KD_MADTNT & "," & TEN_TC & "," & DC_TC & "," & QD_TC & ", "
        sSQLVal = sSQLVal & HOPDONG & "," & ngnop & "," & TTHTK & "," & QUI & "," & matkhai & "," & HANNOP & "," & HANNOP2 & ", "
        sSQLVal = sSQLVal & ThueTKy & "," & ThueTKy2 & "," & NGAYHD & "," & NGAYXN & "," & GhiChu & ", "
        sSQLVal = sSQLVal & SHTEP & "," & MaVach & "," & AUTO_CAL & "," & DaGhi & ", "
        sSQLVal = sSQLVal & KYLBO & "," & THUECL & "," & madlt & "," & lanbs & "," & Lan_quet & "," & DoanhSo & "," & GTTinhThue & "," & GTTTCL & "," & DSoCL & "," & DAGHI2 & "," & danhan & "," & NGAY_PS & "," & TENNV & "," & CHUNGCHI & "," & GIA_HAN & "," & dn_nho & "," & dn_kk & "," & yn_da30

        sSQL = "INSERT INTO TMP_SMD" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
       
    End With

    clsDAO.Disconnect
    
    InsertHDR = vbNullString
    
End Function

Public Function InsertHDR_KHBS() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    Dim bln        As Boolean

    Dim madtnt As Variant
    Dim kykkhai As Variant
    Dim matkhai As Variant
    Dim MATHUE As Variant
    Dim MaMuc As Variant
    Dim maTM As Variant
    Dim MAPP As Variant
    Dim ngnop As Variant
    'Dim TTHTK As Variant
    Dim KYLBO As Variant
    Dim songaycn As Variant
    Dim sotiencn As Variant
    Dim tienhoan As Variant
    Dim lenhhoan As Variant
    Dim ngayhoan As Variant
    Dim HANNOP As Variant
    Dim cqt As Variant
    Dim songay As Variant
    Dim SOTIEN As Variant
    Dim lydokhac As Variant
    Dim danhan As Variant

    With fps
        .Sheet = 1

        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        kykkhai = "'" & TAX_Utilities_Svr_New.Year & "'"
        MaMuc = "'1050'"
        maTM = "'1054'"
        matkhai = "'02/TNDN'"



        MATHUE = "'08'"
        MAPP = "'1'"



        .GetText .ColLetterToNumber("E"), 44, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 42, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If

        'Lay sheet KHBS tren to khai
        .Sheet = .SheetCount - 1

        .GetText .ColLetterToNumber("BE"), .MaxRows - 28, songaycn
        If Trim(songaycn) = vbNullString Then
            songaycn = "0"
        End If

        .GetText .ColLetterToNumber("BE"), .MaxRows - 27, sotiencn
        If Trim(sotiencn) = vbNullString Then
            sotiencn = "0"
        End If

        .GetText .ColLetterToNumber("BE"), .MaxRows - 21, tienhoan
        If Trim(tienhoan) = vbNullString Then
            tienhoan = "0"
        End If

        .GetText .ColLetterToNumber("BE"), .MaxRows - 19, lenhhoan
        If Trim(lenhhoan) = vbNullString Then
            lenhhoan = "''"
        Else
            lenhhoan = "'" & TAX_Utilities_Svr_New.Convert(CStr(lenhhoan), UNICODE, TCVN) & "'"
        End If

        .GetText .ColLetterToNumber("E"), .MaxRows - 17, ngayhoan

        If Trim(ngayhoan) = vbNullString Then
            ngayhoan = "CTOD('')"
        Else
            ngayhoan = ToDate(Trim(ngayhoan), DDMMYYYY)
            ngayhoan = "CTOD('" & Format(ngayhoan, "mm/dd/yyyy") & "')"
        End If

        HANNOP = "CTOD('')"

        .GetText .ColLetterToNumber("BI"), .MaxRows - 13, cqt
        If Trim(cqt) = vbNullString Then
            cqt = "''"
        Else
            cqt = "'" & cqt & "'"
        End If

        .GetText .ColLetterToNumber("BE"), .MaxRows - 11, songay
        If Trim(songay) = vbNullString Then
            songay = "0"
        End If

        .GetText .ColLetterToNumber("BE"), .MaxRows - 9, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        .GetText .ColLetterToNumber("BD"), .MaxRows - 5, lydokhac
        If Trim(lydokhac) = vbNullString Then
            lydokhac = "''"
        Else
            lydokhac = "'" & TAX_Utilities_Svr_New.Convert(CStr(lydokhac), UNICODE, TCVN) & "'"
        End If

        danhan = "''"

        If clsDAO.Connected = False Then
            clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
            clsDAO.Connect
        End If

        sSQLCol = "madtnt,kykkhai,matkhai,MATHUE,MaMuc,maTM,MAPP,ngnop,TTHTK,KYLBO,songaycn,sotiencn,tienhoan," & _
            "lenhhoan,ngayhoan,HANNOP,cqt,songay,SOTIEN,lydokhac,DANHAN,LAN_QUET"

        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & songaycn & "," & sotiencn & "," & tienhoan & "," & _
            lenhhoan & "," & ngayhoan & "," & HANNOP & "," & cqt & "," & songay & "," & SOTIEN & "," & lydokhac & "," & danhan & "," & Lan_quet

        sSQL = "INSERT INTO TMP_ND" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        bln = clsDAO.ExecuteDLL(sSQL)

    End With

    InsertHDR_KHBS = vbNullString
    
    clsDAO.Disconnect
    
End Function

Public Function TKTT() As Boolean
    Dim sSQL           As String
    Dim rs             As ADODB.Recordset
    Dim fso            As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim strFileNameBS  As String
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_SMD" & Right(sKyKeKhai, 4) & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_SDD" & Right(sKyKeKhai, 4) & ".DBF"
    strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"
    Dim strFileNameBSHDR As String
    strFileNameBSHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_ND" & sKyKeKhai & ".DBF"
    
    'file HDR
    If fso.FileExists(strFileNameHDR) = False Then
        Dim strTepmauHDR As String
        strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_SMD.DBF"
        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    End If

    'file dtl
    If fso.FileExists(strFileNameDTL) = False Then
        Dim strTepmauDTL As String
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_SDD.DBF"
        fso.CopyFile strTepmauDTL, strFileNameDTL, False
    End If
    
    'file bo xung
    If fso.FileExists(strFileNameBS) = False Then
        Dim strTepmauBS As String
        strTepmauBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_BS.DBF"
        fso.CopyFile strTepmauBS, strFileNameBS, False
    End If
    
    If fso.FileExists(strFileNameBSHDR) = False Then
        Dim strTepmauBSHDR As String
        strTepmauBSHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_ND.DBF"
        fso.CopyFile strTepmauBSHDR, strFileNameBSHDR, False
    End If

    'Cac file fu luc
    Dim strFileNamePL1 As String
    Dim strTepmauPL1   As String
    'Them bang phu luc PL03-1A/TNDN
    strFileNamePL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_P2DN" & Right(sKyKeKhai, 4) & ".DBF"

    If fso.FileExists(strFileNamePL1) = False Then
        strTepmauPL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_P2DN.DBF"
        fso.CopyFile strTepmauPL1, strFileNamePL1, False
    End If
    
    isExistFile = True

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    If TKTHANG = True Then
        sSQL = "SELECT MADTNT, NGNOP FROM TMP_SMD" & Right(sKyKeKhai, 4) & " WHERE MADTNT = " & "'" & strMST & "' AND QUI = '" & Right$(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='" & LOAI_TO_KHAI & "' AND NGAY_PS = CTOD('')"
    
    Else
        sSQL = "SELECT MADTNT, NGNOP FROM TMP_SMD" & Right(sKyKeKhai, 4) & " WHERE MADTNT = " & "'" & strMST & "' AND NGAY_PS = " & NGAY_PS & " AND MATKHAI='" & LOAI_TO_KHAI & "'"
    End If

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

    Lan_quet = TinhSoLanQuet + 1
    clsDAO.Disconnect

End Function

Public Function UpdateTHUETKY2() As Boolean
'    Dim sSQL As String
'   If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'   End If
'
'   sSQL = "UPDATE TMP_QT6" & sKyKeKhai & _
'        " SET DATHAY=.T." & _
'        " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT = " & TAX_Utilities_Svr_New.Year
'
'   clsDAO.ExecuteDLL sSQL
'
   UpdateTHUETKY2 = True
'
'   clsDAO.Disconnect

End Function


Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim ngnop As Variant
    Dim rs As ADODB.Recordset
    
    If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
     
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 12, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
   End With
   sSQL = "DELETE FROM TMP_QT6D" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
    
   clsDAO.ExecuteDLL sSQL
   
   ' xoa hrd
   sSQL = "DELETE FROM TMP_QT6" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT=" & TAX_Utilities_Svr_New.Year
    
   clsDAO.ExecuteDLL sSQL
   
'   With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("E"), 12, NGNOP
'        'NGNOP = Date
'        If Trim(NGNOP) = vbNullString Then
'            NGNOP = "CTOD('')"
'        Else
'            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
'            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
'        End If
'   End With
'   sSQL = "DELETE FROM Q6D" & sKyKeKhai & _
'        " WHERE MADTNT = " & "'" & strMST & "' AND NGNOP = " & dNgayNopDB
'
'   clsDAO.ExecuteDLL sSQL
   '------------
   'Xoa phu luc
   'Phu luc PL1A
   sSQL = "DELETE FROM TMP_Q31A" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   
   'Phu luc PL1B
   sSQL = "DELETE FROM TMP_Q31B" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   
   ' phuc luc PL1C
   sSQL = "DELETE FROM TMP_Q31C" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' xoa phu luc PL3
   sSQL = "DELETE FROM TMP_Q33" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   
   ' xoa phuc luc PL4A
    sSQL = "DELETE FROM TMP_Q34A" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' xoa phu luc PL4B
   sSQL = "DELETE FROM TMP_Q34B" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   'xoa phu luc PL4C
   sSQL = "DELETE FROM TMP_Q34C" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' xoa phu luc PL4D
   sSQL = "DELETE FROM TMP_Q34D" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' xoa phu luc PL4DD
   sSQL = "DELETE FROM TMP_Q34Z" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' Xoa phu luc PL4E
   sSQL = "DELETE FROM TMP_Q34E" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' Xoa phu luc PL4G
   sSQL = "DELETE FROM TMP_Q34G" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' Xoa phu luc PL4H
   sSQL = "DELETE FROM TMP_Q34H" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' Xoa phu luc PL5
   sSQL = "DELETE FROM TMP_Q35" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   'Phu luc PL2A
   sSQL = "DELETE FROM TMP_Q321" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   'Xoa du lieu bang Q322
   sSQL = "DELETE FROM TMP_Q322" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL

   XoaTKTT = True
   clsDAO.Disconnect
End Function

Public Function TKRB() As Boolean

   TKRB = True

End Function

Public Function InsertDTL_PL1() As String
    Dim sSQL      As String
    Dim sSQLCol   As String
    Dim sSQLVal   As String
    Dim bln       As Boolean

    Dim madtnt    As Variant
    Dim MST     As Variant
    Dim diaChi      As Variant
    Dim ngnop     As Variant
    Dim TENBN    As Variant
    Dim HOPDONG    As Variant
    Dim TTHTK     As Variant
    
    Dim QUI   As Variant ' KyKKhai thay bang QUI theo TT151
    Dim matkhai   As Variant
    Dim tt        As Variant
    Dim danhan    As Variant
    Dim KYLBO As Variant

    sSQLCol = "MATKHAI,MADTNT,KYLBO,QUI,NGNOP,TTHTK,STT,MST,DCHI,TENBN,HOPDONG,DANHAN,LAN_QUET"

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If

    With fps

        .Sheet = 1
        '------------------ Thong tin chung -------------------
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        ' ky lap bo
        .GetText .ColLetterToNumber("E"), 42, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If
        'trang thai to khai
        Dim TTHTK1 As Variant

        .GetText .ColLetterToNumber("M"), 49, TTHTK1

        If Trim(TTHTK1) = vbNullString Then
            TTHTK = "'" & 2 & "'"
        Else
            TTHTK = "'" & 1 & "'"
        End If

        ' ngay nop
        .GetText .ColLetterToNumber("E"), 44, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'ma to khai
        matkhai = "'" & LOAI_TO_KHAI & "'"
        danhan = "''"
        'ky ke khai thay bang QUI theo TT151
        If NGAY_PS <> "CTOD('')" Then
            QUI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            QUI = "'" & TAX_Utilities_Svr_New.ThreeMonths & "/" & TAX_Utilities_Svr_New.Year & "'"
            QUI = "'" & Right$(QUI, 7)
        End If
        tt = "' '"
        
        
        .Sheet = 2
        .Row = 7
        Do
        
            .GetText .ColLetterToNumber("B"), .Row, tt
            .GetText .ColLetterToNumber("C"), .Row, TENBN
            .GetText .ColLetterToNumber("D"), .Row, MST
            .GetText .ColLetterToNumber("E"), .Row, diaChi
            .GetText .ColLetterToNumber("F"), .Row, HOPDONG
            
            tt = "'" & tt & "'"
            TENBN = "'" & TENBN & "'"
            MST = "'" & MST & "'"
            diaChi = "'" & diaChi & "'"
            HOPDONG = "'" & HOPDONG & "'"
            
            sSQLVal = matkhai & "," & madtnt & "," & KYLBO & "," & QUI & "," & ngnop & "," & TTHTK _
            & "," & tt & "," & MST & "," & diaChi & "," & TENBN & "," & HOPDONG & "," & danhan & "," & Lan_quet
            
            sSQL = "INSERT INTO TMP_P2DN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
            
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
    End With
   
    InsertDTL_PL1 = vbNullString
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
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_QT6" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT <> " & TAX_Utilities_Svr_New.Year
    
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

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
 Dim varTemp As Variant
 With fps
    If .ActiveSheet = 1 Then
        If Col = .ColLetterToNumber("E") And Row = 12 Then
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
           UpdateCell fps, .Col, .Row, .Text
        End If
        If Col = .ColLetterToNumber("E") And Row = 10 Then
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
    Dim sSQL           As String
    Dim rs             As ADODB.Recordset
    Dim fso            As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim LAN_QUET1      As Variant
    
    'kiem tra ton tai tep *.dbf chua
    'Tinh so lan quyet de ghi vao DB trung gian
    'dhdang sua
    'ngay 24/09/2010
    '-----------------------------
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_SMD" & Right(sKyKeKhai, 4) & ".DBF"

    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
            clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
            clsDAO.Connect
        End If

'        If Trim(NGAY_PS) = vbNullString Then
'            LOAI_TO_KHAI = "02/TNDN"
'            NGAY_PS = "CTOD('')"
'        Else
'            TAX_Utilities_Svr_New.Month = Mid$(NGAY_PS, 4, 2)
'            TAX_Utilities_Svr_New.Year = Right$(NGAY_PS, 4)
'            TAX_Utilities_Svr_New.ThreeMonths = ""
'            NGAY_PS = ToDate(Trim(NGAY_PS), DDMMYYYY)
'            NGAY_PS = "CTOD('" & Format(NGAY_PS, "mm/dd/yyyy") & "')"
'            LOAI_TO_KHAI = "02PS/TNDN"
'        End If
        
        If TKTHANG = True Then
            sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_SMD" & Right(sKyKeKhai, 4) & " WHERE MADTNT = " & "'" & strMST & "' AND QUI = '" & Right$(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='" & LOAI_TO_KHAI & "' AND NGAY_PS = CTOD('')"
        Else
            ' ngay phat sinh
            fps.Sheet = 1
            fps.GetText fps.ColLetterToNumber("M"), 50, NGAY_PS

            If Trim(NGAY_PS) = vbNullString Then
                NGAY_PS = "CTOD('')"
            Else
                NGAY_PS = ToDate(Trim(NGAY_PS), DDMMYYYY)
                NGAY_PS = "CTOD('" & Format(NGAY_PS, "mm/dd/yyyy") & "')"
            End If

            sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_SMD" & Right(sKyKeKhai, 4) & " WHERE MADTNT = " & "'" & strMST & "' AND MATKHAI='" & LOAI_TO_KHAI & "' AND NGAY_PS = " & NGAY_PS

        End If

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

Private Function GetMATM() As Object
    Dim rsReturn1 As New ADODB.Recordset
    Dim strSQL1 As String
    Dim clsDAO1 As New TAX_Utilities_Svr_New.clsADO
    'connect to database
    If clsDAO1.Connected = False Then
        clsDAO1.CreateConnectionString spathVat & "\DB_HT\"
        clsDAO1.Connect
    End If
   
    strSQL1 = "SELECT matm, tengoi FROM dmmt2 WHERE mamuc ='1050' order by matm"
    Set rsReturn1 = clsDAO1.Execute(strSQL1)

    Set GetMATM = rsReturn1
    
    Set rsReturn1 = Nothing
    
    'Connect DB success
    'blnSuccess = True
    clsDAO1.Disconnect
End Function

Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL            As String
    Dim rs              As ADODB.Recordset
    Dim vKeKhai         As Variant
    Dim vKeKhai_temp    As Variant
    Dim vKyLB           As Variant
    Dim strTestFileName As String
    Dim fso             As New FileSystemObject
    Dim matkhai         As Variant
     
    Dim flag            As Boolean
    'kiem tra da co to khai chinh thuc chua
    flag = True
    vKyLB = DateSerial(Int(Year(Now())), Int(Month(Now())), 1)
    vKeKhai = DateSerial(Int(Val(TAX_Utilities_Svr_New.Year)), Int(Val(TAX_Utilities_Svr_New.Month)), 1)
    vKeKhai_temp = Right(sKyKeKhai, 4)

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If

    Dim kykkhai As Variant

    If NGAY_PS <> "CTOD('')" Then
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
    Else
        kykkhai = "'" & TAX_Utilities_Svr_New.ThreeMonths & "/" & TAX_Utilities_Svr_New.Year & "'"
        kykkhai = "'" & Right$(kykkhai, 7)
    End If

    Do While flag
   
        strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_SM" & vKeKhai_temp & ".DBF"

        'kiem tra neu co db thi moi thuc hien truy van
        If fso.FileExists(strTestFileName) Then
            sSQL = "SELECT MADTNT, NGNOP FROM TMP_SMD" & vKeKhai_temp & " WHERE MADTNT = " & "'" & strMST & "' AND QUI = " & kykkhai & " AND MATKHAI='" & LOAI_TO_KHAI & "' AND NGAY_PS = CTOD('')"
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
