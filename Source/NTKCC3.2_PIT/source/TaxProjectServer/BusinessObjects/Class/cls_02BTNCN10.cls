VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_02BTNCN10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_Y = 22
Const KY_LAP_BO_X = "E"
Const NGAY_NHAN_TO_KHAI_Y = 24
Const NGAY_NHAN_TO_KHAI_X = "E"
Const NGUOI_SU_DUNG_Y = 24
Const NGUOI_SU_DUNG_X = "V"
Const HEADER_KY_LAP_BO_Y = 13
Const HEADER_KY_LAP_BO_X = "B"
Const MA_BPQL_X = "V"
Const MA_BPQL_Y = 16 'Dong 16 da dc Hidden
Const PHONG_XU_LY_Y = 22
Const PHONG_XU_LY_X = "V"

'Longvh
Const MA_SO_THUE_X = "G"
Const MA_SO_THUE_Y = 8
Const NGAY_NOP_Y = 24
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 24
Const NGAY_QUET_X = "M"
Const TEN_BPQL_X = "V"
Const TEN_BPQL_Y = 26
Private rs As ADODB.Recordset
' thong tin din vi
Const TEN_GOI_X = "H"
Const TEN_GOI_Y = 7
Const DIA_CHI_X = "G"
Const DIA_CHI_Y = 9
Const QUAN_HUYEN_X = "F"
Const QUAN_HUYEN_Y = 11
Const TINH_TPHO_X = "T"
Const TINH_TPHO_Y = 11
Const DIEN_THOAI_X = "E"
Const DIEN_THOAI_Y = 12
Const FAX_X = "M"
Const FAX_Y = 12
Private tempKyLb As Variant
Private tempNgayNop As Variant
Const MAIL_X = "V"
Const MAIL_Y = 12
'thong tin Dai ly thue
Const MA_SO_THUE_DL_X = "G"
Const MA_SO_THUE_DL_Y = 15
Const TEN_GOI_DL_X = "H"
Const TEN_GOI_DL_Y = 14
Const DIA_CHI_DL_X = "G"
Const DIA_CHI_DL_Y = 16
Const QUAN_HUYEN_DL_X = "F"
Const QUAN_HUYEN_DL_Y = 17
Const TINH_TPHO_DL_X = "T"
Const TINH_TPHO_DL_Y = 17
Const DIEN_THOAI_DL_X = "E"
Const DIEN_THOAI_DL_Y = 18
Const FAX_DL_X = "M"
Const FAX_DL_Y = 18
Const MAIL_DL_X = "V"
Const MAIL_DL_Y = 18
Const SO_HDDL_DL_X = "I"
Const SO_HDDL_DL_Y = 19
Const NGAY_HDDL_DL_X = "S"
Const NGAY_HDDL_DL_Y = 19
'end
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
Public isExistFile As Boolean
Public LAN_QUET As Variant
Const STT_X = "M"
Const STT_Y = 22
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
Public spathVat As String
Public sSaiCT11 As String
Public sKyKeKhai As String
Public dNgayNopDB As Variant
Public HANNOP As String
Public tthtk As String
Public strTenBpql As String
Public hdrId As Integer
Public strMaCQT As String
Public strMaPQL As String
Public strTenPQL As String
Public strSoTTTKhai As String
Public isTKTonTai As Boolean
Public checkDLTTonTai As Boolean


Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
                
        'Ghi chu
        .Col = .ColLetterToNumber("M")
        .Row = 28
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100
        
        'Phong xu ly
        .Col = .ColLetterToNumber("V")
        .Row = 22
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
        
        'Phong quan ly
        .Col = .ColLetterToNumber("V")
        .Row = 26
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
        
        ' Ky lap bo
        SetDateFormat fps, 1, 22, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 24, .ColLetterToNumber("E"), DDMMYYYY
        .Row = 24
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 24, .ColLetterToNumber("M"), DDMMYYYY
        .Row = 24
        .Col = .ColLetterToNumber("M")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
    End With
    ' Lay ve gia tri cua HDR ID khi quet to khai
    hdrId = GetHdrId(spathVat)
    ' end

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

Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

Public Function Prepared3() As Boolean
    With fps
        .Sheet = 1
          'set check loi dinh danh
            If checkDLTTonTai = False Then
                .Col = .ColLetterToNumber("E")
                .Row = 26
                .Value = 1
            End If
        'Set NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
        .Row = NGAY_NHAN_TO_KHAI_Y
        If strNgayNhanToKhai <> "" Then
            .Text = strNgayNhanToKhai
            UpdateCell fps, .Col, .Row, .Text
        End If
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
        UpdateCell fps, .Col, .Row, .Value
       
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
        ' thong tin dai ly thue
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
'        ' quan huyen
'        .Col = .ColLetterToNumber(QUAN_HUYEN_DL_X)
'        .Row = QUAN_HUYEN_DL_Y
'        If strQHuyen_DLT <> "" Then
'            .Text = strQHuyen_DLT
'        End If
'        ' thanh pho
'        .Col = .ColLetterToNumber(TINH_TPHO_DL_X)
'        .Row = TINH_TPHO_DL_Y
'        If strTTPho_DLT <> "" Then
'            .Text = strTTPho_DLT
'        End If
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
        'so thu tu to khai
        .Col = .ColLetterToNumber(STT_X)
        .Row = STT_Y
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        
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
        
'        .Col = .ColLetterToNumber(MA_BPQL_X)
'        .Row = MA_BPQL_Y
'        If strMaBPQL <> "" Then
'            .Text = strMaBPQL
'            UpdateCell fps, .Col, .Row, .Text
'        End If
        .Col = .ColLetterToNumber(TEN_BPQL_X)
        .Row = TEN_BPQL_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        '.EventEnabled(EventAllEvents) = True
        
        'neu ghi vao PIT thi lay kylbo va ngaynop luu trong cookies
        If TAX_Utilities_Svr_New.isCheckPIT Then
                    
            ' Set lai ky lap bo
            ' Lay ky lap bo duoc cua to khai quet truoc do, duoc ghi trong cookie de hien thi cho lan quet to khai tncn lan sau
            tempKyLb = GetSetting("HTKK", "TNCN", "tempKyLb")
            If (tempKyLb <> vbNullString) Then
                .Col = .ColLetterToNumber(KY_LAP_BO_X)
                .Row = KY_LAP_BO_Y
                .Text = tempKyLb
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            ' Set lai ngay nop
            ' Lay ngay nop duoc cua to khai quet truoc do, duoc ghi trong cookie de hien thi cho lan quet to khai tncn lan sau
            tempNgayNop = GetSetting("HTKK", "TNCN", "tempNgayNop")
            If (tempNgayNop <> vbNullString) Then
                .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
                .Row = NGAY_NHAN_TO_KHAI_Y
                .Text = tempNgayNop
                .TypeHAlign = TypeHAlignLeft
            End If
    
        End If
        
               
        ' Dat lai check E_36 Duoc gian thue theo TT27/TT-BTC
        ' Thoi gian gian thue tu thang 1 den thang 6 nam 2009 cua to khai thang
'        If Val(TAX_Utilities_Svr_New.Year) = 2009 And Val(TAX_Utilities_Svr_New.Month) <= 6 Then
'            ' Neu trong hieu luc cua gian thue thi default la checked
'            .Row = 36
'            .Col = .ColLetterToNumber("E")
'            .Text = "1"
'            UpdateCell fps, .Col, .Row, "x"
'        Else
'            ' Neu ngoai hieu luc cua gian thue thi visible cac dieu kien check nay di
'            .Row = 36
'            .Col = .ColLetterToNumber("B")
'            .CellType = CellTypeEdit
'            .Lock = True
'            .Text = ""
'
'            .Row = 36
'            .Col = .ColLetterToNumber("E")
'            .CellType = CellTypeEdit
'            .Lock = True
'            .Text = ""
'            UpdateCell fps, .Col, .Row, ""
'        End If
        
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


Public Function CheckValidData() As Boolean

    Dim NGAYNOP As Variant, ngayHienTai As Variant
    CheckValidData = True
    With fps
        .Sheet = 1
        
        ' check ngay nop khong dc lon hon ngay hien tai
        .GetText .ColLetterToNumber("E"), 24, NGAYNOP
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
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("E_22"), "Value"), "dd/mm/yyyy")

    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function

Public Function InsertDTL() As String
    Dim sSQL    As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim matkhai As Variant
    Dim madtnt  As Variant
    Dim kylbo   As Variant
    Dim kykkhai As Variant
    Dim ngnop   As Variant
    Dim CTTN    As Variant
    Dim giatri  As Variant
    Dim giatri2 As Variant
    Dim tt      As Variant
    Dim sSQL1   As String
    Dim hdr     As TNCN_DTL
    Dim DANHAN  As Variant
    Dim arr()   As String
    Dim bln     As Boolean
    Dim cnt     As Integer
    Dim TTHTK1  As Variant
    Dim i       As Integer
    Dim pit     As Boolean
    
    If clsDAO.Connected = False Then
        '        Me.MousePointer = vbHourglass
        '        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        '        frmSystem.MousePointer = vbDefault
        '        Me.MousePointer = vbDefault
    End If
  
    '-------------
    With fps
   
        .Sheet = 1
        matkhai = "'02B/TNCN11'"
        DANHAN = "'T'"   ' To khai khong chuyen sang VAT
        .GetText .ColLetterToNumber("G"), 8, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        'ky lap bo
        .GetText .ColLetterToNumber("E"), 22, kylbo

        If Trim(kylbo) = vbNullString Then
            kylbo = "''"
        Else

            If Len(Trim(kylbo)) = 6 Then
                kylbo = "'0" & kylbo & "'"
            Else
                kylbo = "'" & kylbo & "'"
            End If
        End If

        'ngay nop
        .GetText .ColLetterToNumber("E"), 24, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        kykkhai = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        giatri2 = "0"
        tt = "''"
        'trang thai to khai
        .GetText .ColLetterToNumber("L"), 5, TTHTK1

        If Trim(TTHTK1) = vbNullString Then
            tthtk = "'" & 2 & "'"
        End If
       
        '----------------------------------------------------------
        sSQLCol = "MATKHAI, MADTNT, KYLBO, KYKKHAI, TTHTK, NGNOP, CTTN, GIATRI, GIATRI2,TT, LAN_QUET, DANHAN"
            
        '-------  Insert bang TMP_T2, tmp_tncn_dtl theo chi tieu CTTN   -----------
        Dim cttnArr()   As Variant 'mang cua cac Cell chi tieu
        Dim cellArr()   As Variant 'mang gia tri cua cac Cell chi tieu
        Dim kyhieuArr() As Variant ' mang ky hieu theo chi tieu
    
        cttnArr() = Array("21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37")
        kyhieuArr() = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17")
        cellArr() = Array(39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 52, 53, 54, 55, 56)

        If TAX_Utilities_Svr_New.isCheckPIT = True Then
            pit = True
        End If

        For i = 0 To UBound(cttnArr)
            
            CTTN = "'0" & cttnArr(i) & "'"
                
            .GetText .ColLetterToNumber("U"), cellArr(i), giatri
            giatri2 = giatri
                
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
            
            arr = Split(GetMaCqt, "~")
            hdr.matkhai = "'02B_TNCN13'"
            hdr.madtnt = madtnt
            hdr.kylbo = kylbo
            hdr.kykkhai = kykkhai
            hdr.tthtk = tthtk
            hdr.ngnop = ngnop
            hdr.CTTN = CTTN
            hdr.giatri = giatri
                            
            hdr.LAN_QUET = LAN_QUET
            hdr.ky_hieu = "'" & kyhieuArr(i) & "'"
            hdr.MA_CQT = arr(0)
            hdr.DANHAN = 0
            hdr.Id = GetDtlId(spathVat)
            hdr.Hdr_id = hdrId

            bln = InsertDTL_By_CTTN(hdr, pit)

        Next
        
    End With

    ' Ghi du lieu doi voi truong hop khong trien khai PIT
    If TAX_Utilities_Svr_New.isCheckPIT = False Then
        sSQLCol = "MATKHAI, MADTNT, KYLBO, KYKKHAI, TTHTK, NGNOP, CTTN, GIATRI, GIATRI2,TT, LAN_QUET, DANHAN "
         
        With fps
        
            .Sheet = 1
            matkhai = "'02Q/KK-TNCN'"
            ' ma doi tuong nop thue
            .GetText .ColLetterToNumber("G"), 8, madtnt

            If Trim(madtnt) = vbNullString Then
                madtnt = "''"
            Else
                madtnt = "'" & madtnt & "'"
            End If

            'ky lap bo
            .GetText .ColLetterToNumber("E"), 22, kylbo

            If Trim(kylbo) = vbNullString Then
                kylbo = "''"
            Else

                If Len(Trim(kylbo)) = 6 Then
                    kylbo = "'" & kylbo & "'"
                Else
                    kylbo = "'" & kylbo & "'"
                End If
            End If

            ' ky ke khai
            kykkhai = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
            'ngay nop
            .GetText .ColLetterToNumber("E"), 24, ngnop

            If Trim(ngnop) = vbNullString Then
                ngnop = "CTOD('')"
            Else
                ngnop = ToDate(Trim(ngnop), DDMMYYYY)
                ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
            End If

            giatri2 = "0"
            tt = "''"
            DANHAN = "''"
            'trang thai to khai
            .GetText .ColLetterToNumber("L"), 5, TTHTK1

            If Trim(TTHTK1) = vbNullString Then
                tthtk = "'" & 2 & "'"
            End If

            'Chi tieu "'21'"
            CTTN = "'001'"
            .GetText .ColLetterToNumber("U"), 41, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'22'"
            CTTN = "'002'"
            .GetText .ColLetterToNumber("U"), 42, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'23'"
            CTTN = "'003'"
            .GetText .ColLetterToNumber("U"), 43, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'24'"
            CTTN = "'004'"
            .GetText .ColLetterToNumber("U"), 44, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'25'"
            CTTN = "'005'"
            .GetText .ColLetterToNumber("U"), 45, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'26'"
            CTTN = "'006'"
            .GetText .ColLetterToNumber("U"), 46, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'27'"
            CTTN = "'007'"
            .GetText .ColLetterToNumber("U"), 47, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'28'"
            CTTN = "'008'"
            .GetText .ColLetterToNumber("U"), 48, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'29'"
            CTTN = "'009'"
            .GetText .ColLetterToNumber("U"), 49, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'30'"
            CTTN = "'010'"
            .GetText .ColLetterToNumber("U"), 50, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'31'"
            CTTN = "'011'"
            .GetText .ColLetterToNumber("U"), 51, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'32'"
            CTTN = "'012'"
            .GetText .ColLetterToNumber("U"), 52, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'33'"
            CTTN = "'013'"
            .GetText .ColLetterToNumber("U"), 53, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'34'"
            CTTN = "'014'"
            .GetText .ColLetterToNumber("U"), 54, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
            'Chi tieu "'35'"
            CTTN = "'015'"
            .GetText .ColLetterToNumber("U"), 55, giatri

            If Trim(giatri) = vbNullString Then
                giatri = "0"
            End If

            giatri2 = giatri
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & kykkhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DANHAN
             
            sSQL = "INSERT INTO TMP_TQ2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
             
        End With

    End If
   
    clsDAO.Disconnect
   
    InsertDTL = vbNullString
End Function
'--------------------------------------------------------------------------------
' Project    :       Tax_Business_Svr_New
' Procedure  :       InsertDTL_By_CTTN
' Description:       [Insert bang TMP_T2, tmp_tncn_dtl theo chi tieu CTTN]
' Machine    :       KHAIHOAN
' Date-Time  :       7/22/2011-20:01:28
' Parameters :       pit : kiem tra truong PIT
'--------------------------------------------------------------------------------
Public Function InsertDTL_By_CTTN(hdr As TNCN_DTL, pit As Boolean) As Boolean
  
    Dim bln     As Boolean
    Dim sSQL1   As String

    If pit = True Then
        sSQL1 = InsertDTL_TGTC(hdr)
        bln = clsDAO.ExecuteDLL(sSQL1)
    
    End If

    InsertDTL_By_CTTN = bln
End Function

Public Function InsertDTL_KHBS() As String
'    Dim sSQL As String
'    Dim sSQLCol As String
'    Dim sSQLVal As String
'    Dim bln  As Boolean
'    Dim MADTNT As Variant
'    Dim NGNOP As Variant
'    Dim KYLBO As Variant
'    Dim KYKKHAI As Variant
'    Dim MATKHAI As Variant
'    Dim MATHUE As Variant
'    Dim MAMUC As Variant
'    Dim MATM As Variant
'    Dim MAPP As Variant
'    Dim MACT As Variant
'    Dim MACT2 As Variant
'    Dim MACT3 As Variant
'    Dim STT As Variant
'    Dim STT2 As Variant
'    Dim STTIN As Variant
'    Dim SOKK As Variant
'    Dim SODC As Variant
'    Dim SOCL As Variant
'    Dim i As Integer, iCol As Long, iRow As Long
'    Dim xmlNode As MSXML.IXMLDOMNode
'
'    Dim strFileName As String
'    Dim fso As New FileSystemObject
'
'   'kiem tra ton tai tep *.dbf chua
'    strFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "BS" & sKyKeKhai & ".DBF"
'    If fso.FileExists(strFileName) = False Then
'    Dim strTepmau As String
'        strTepmau = spathVat & "\tepmau\BS.DBF"
'        fso.CopyFile strTepmau, strFileName, False
'    End If
'
'
'   sSQLCol = "MADTNT, KYKKHAI, MATKHAI, MATHUE, MAMUC, MATM, MAPP, NGNOP, TTHTK, KYLBO , MACT, MACT2, MACT3, STT, STT2, STTIN," & _
'             " SOKK, SODC, SOCL"
'
'   If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'   End If
'
'
'   With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("G"), 4, MADTNT
'        If Trim(MADTNT) = vbNullString Then
'            MADTNT = "''"
'        Else
'            MADTNT = "'" & MADTNT & "'"
'        End If
'
'        KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'        MATKHAI = "'03/TNDN'"
'        MATHUE = "'01'"
'        MAMUC = "'014'"
'        MATM = "'01'"
'        MAPP = "'2'"
'
'        .GetText .ColLetterToNumber("E"), 12, NGNOP
'        'NGNOP = Date
'        If Trim(NGNOP) = vbNullString Then
'            NGNOP = "CTOD('')"
'        Else
'            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
'            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
'        End If
'
'        .GetText .ColLetterToNumber("E"), 10, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "CTOD('')"
'        Else
'            KYLBO = "'" & KYLBO & "'"
'        End If
'
'' Insert so dieu chinh tang
'
'        MACT = "'1000'"
'        MACT2 = "''"
'        MACT3 = "''"
'        STT = 1
'        STT2 = 0
'        STTIN = "'I'"
'        SOKK = "0"
'        SODC = "0"
'        SOCL = "0"
'
'            sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                      NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                      SOKK & "," & SODC & "," & SOCL
'
'
'            sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'            bln = clsDAO.ExecuteDLL(sSQL)
'
'        .Sheet = .SheetCount - 1
'        i = 1
'        .Row = 24
'        Do
'                MACT = "''"
'                .GetText .ColLetterToNumber("E"), .Row, MACT2
'                If Trim(MACT2) = vbNullString Then
'                    MACT2 = "''"
'                ElseIf MACT2 = "20" Then
'                    MACT2 = "'062'"
'                ElseIf MACT2 = "21" Then
'                    MACT2 = "'076'"
'                End If
'
'                MACT3 = "'2'"
'                STT = 1
'                STT2 = i
'                STTIN = "'" & CStr(i) & "'"
'
'                .GetText .ColLetterToNumber("F"), .Row, SOKK
'                If Trim(SOKK) = vbNullString Then
'                    SOKK = "0"
'                Else
'                    SOKK = SOKK
'                End If
'
'                .GetText .ColLetterToNumber("G"), .Row, SODC
'                If Trim(SODC) = vbNullString Then
'                    SODC = "0"
'                Else
'                    SODC = SODC
'                End If
'
'                .GetText .ColLetterToNumber("H"), .Row, SOCL
'                If Trim(SOCL) = vbNullString Then
'                    SOCL = "0"
'                Else
'                    SOCL = SOCL
'                End If
'
'                sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                          NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                          SOKK & "," & SODC & "," & SOCL
'
'                sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'                If MACT2 <> "''" Then
'                    bln = clsDAO.ExecuteDLL(sSQL)
'                End If
'
'            .Col = .ColLetterToNumber("B")
'            .Row = 1 + .Row
'            i = i + 1
'        Loop Until .Text = "aa"
'' Insert so dieu chinh giam
'
'        MACT = "'2000'"
'        MACT2 = "''"
'        MACT3 = "''"
'        STT = 2
'        STT2 = 0
'        STTIN = "'II'"
'        SOKK = "0"
'        SODC = "0"
'        SOCL = "0"
'
'            sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                      NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                      SOKK & "," & SODC & "," & SOCL
'
'
'            sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'            bln = clsDAO.ExecuteDLL(sSQL)
'
'        i = 1
'        .Row = .Row + 3
'         Do
'                MACT = "''"
'                .GetText .ColLetterToNumber("E"), .Row, MACT2
'                If Trim(MACT2) = vbNullString Then
'                    MACT2 = "''"
'                ElseIf MACT2 = "20" Then
'                    MACT2 = "'062'"
'                ElseIf MACT2 = "21" Then
'                    MACT2 = "'076'"
'                End If
'
'                MACT3 = "'2'"
'                STT = 2
'                STT2 = i
'                STTIN = "'" & CStr(i) & "'"
'
'                .GetText .ColLetterToNumber("F"), .Row, SOKK
'                If Trim(SOKK) = vbNullString Then
'                    SOKK = "0"
'                Else
'                    SOKK = SOKK
'                End If
'
'                .GetText .ColLetterToNumber("G"), .Row, SODC
'                If Trim(SODC) = vbNullString Then
'                    SODC = "0"
'                Else
'                    SODC = SODC
'                End If
'
'                .GetText .ColLetterToNumber("H"), .Row, SOCL
'                If Trim(SOCL) = vbNullString Then
'                    SOCL = "0"
'                Else
'                    SOCL = SOCL
'                End If
'
'                sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                          NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                          SOKK & "," & SODC & "," & SOCL
'
'                sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'                If MACT2 <> "''" Then
'                    bln = clsDAO.ExecuteDLL(sSQL)
'                End If
'
'            .Col = .ColLetterToNumber("B")
'            .Row = 1 + .Row
'            i = i + 1
'        Loop Until .Text = "bb"
'
'' Insert Tong hop dieu chinh
'
'        MACT = "'3000'"
'        MACT2 = "''"
'        MACT3 = "''"
'        STT = 3
'        STT2 = 0
'        STTIN = "'III'"
'        SOKK = "0"
'        SODC = "0"
'        SOCL = "0"
'
'            sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                      NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                      SOKK & "," & SODC & "," & SOCL
'
'            sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'            bln = clsDAO.ExecuteDLL(sSQL)
'
'   End With
'
'   clsDAO.Disconnect
   InsertDTL_KHBS = vbNullString
   
End Function

Public Function InsertHDR() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    Dim matkhai    As Variant
    Dim kylbo      As Variant
    Dim ngnop      As Variant
    Dim NgNhap     As Variant
    Dim kykkhai    As Variant
    Dim madtnt     As Variant
    Dim MADLT      As Variant
    Dim DaGhi      As Variant
    Dim LoaiTien   As Variant
    Dim MaVach     As Variant
    Dim SHTEP      As Variant
    Dim HANNOP2    As Variant
    Dim ThueTKy    As Variant
    Dim ThueTKy2   As Variant
    Dim MaMuc      As Variant
    Dim maTM       As Variant
    'Dim NGNHAP As Variant
    Dim AUTO_CAL   As Variant
    Dim MATHUE     As Variant
    Dim MAPP       As Variant
   
    Dim AUTOCAL    As Variant
    Dim CHKGIAHAN  As Variant
    Dim GTTinhThue As Variant
    Dim LanBS      As Variant
    Dim THUECL     As Variant
    Dim GTTTCL     As Variant
    Dim solanBS    As Variant
    Dim hdr        As TNCN_HDR
    Dim bln        As Boolean
    Dim sSQLTemp   As String
    Dim DANHAN     As Variant
    Dim arr()      As String
    Dim PHONG_XL   As Variant
    Dim TTHTK1     As Variant
    Dim LANDAU     As Variant
        
        
    Dim GHICHU As Variant
    'pit
       
    Dim SOHDDL     As Variant
    Dim NGAYHDDL   As Variant
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
'    sSQLCol = "MATKHAI, KYLBO, NGNOP, KYKKHAI, MADTNT, MADLT, TTHTK, DAGHI, LOAITIEN, MAVACH, HANNOP, HANNOP2, THUETKY, THUETKY2, MAMUC, MATM, NGNHAP, AUTO_CAL, MATHUE, MAPP, SHTEP, GTTINHTHUE, LANBS, THUECL, GTTTCL, LAN_QUET"
         
    With fps
        .Sheet = 1
        'ma doi tuong nop thue
        .GetText .ColLetterToNumber("G"), 8, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
      
        'ky lap bo
        .GetText .ColLetterToNumber("E"), 22, kylbo
        'NGNOP = Date
        'Luu thong tin cua ky lap bo cua to khai quet hien tai, de hien thi cho viec quet to khai lan sau
        SaveSetting "HTKK", "TNCN", "tempKyLb", kylbo

        If Trim(kylbo) = vbNullString Then
            kylbo = "''"
        Else

            If Len(Trim(kylbo)) = 6 Then
                kylbo = "'0" & kylbo & "'"
            Else
                kylbo = "'" & kylbo & "'"
            End If
        End If

        ' ngay nop
        .GetText .ColLetterToNumber("E"), 24, ngnop
        'Luu thong tin cua ngay nop cua to khai quet hien tai, de hien thi cho viec quet to khai lan sau
        SaveSetting "HTKK", "TNCN", "tempNgayNop", ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
        'ma dai ly thue
        .GetText .ColLetterToNumber("G"), 15, MADLT

        If Trim(MADLT) = vbNullString Then
            MADLT = "''"
        Else
            MADLT = "'" & MADLT & "'"
        End If

        ' ngay nhap
        .GetText .ColLetterToNumber("M"), 24, NgNhap

        If Trim(NgNhap) = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = ToDate(Trim(NgNhap), DDMMYYYY)
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If
        
        ' Lay thong tin phong xu ly
        
        .GetText .ColLetterToNumber(PHONG_XU_LY_X1), PHONG_XU_LY_Y1, PHONG_XL
        PHONG_XL = Mid(PHONG_XL, InStr(1, PHONG_XL, "{") + 1, InStr(1, PHONG_XL, "}") - InStr(1, PHONG_XL, "{") - 1)
        
        PHONG_XL = "'" & PHONG_XL & "'"
        
        kykkhai = "'" & TAX_Utilities_Svr_New.ThreeMonths & "/" & TAX_Utilities_Svr_New.Year & "'"
        
        '        .GetText .ColLetterToNumber(MA_BPQL_X), MA_BPQL_Y, MABPQL
        '
        '        If Trim(MABPQL) = vbNullString Then
        '            MABPQL = "''"
        '        Else
        '            MABPQL = "'" & MABPQL & "'"
        '        End If
        
        .GetText .ColLetterToNumber("L"), 5, TTHTK1

        If Trim(TTHTK1) = vbNullString Then
            tthtk = "'" & 2 & "'"
        End If

        ' so lan bo sung
        If tthtk = "'" & 2 & "'" Then
            .GetText .ColLetterToNumber("T"), 5, solanBS
            LanBS = solanBS
        Else
            LanBS = "0"
        End If
        
        'Check gia han nop
        .GetText .ColLetterToNumber("E"), 36, CHKGIAHAN

        If Trim(CHKGIAHAN) = "1" Then
            HANNOP = "CTOD('" & Format("30/07/2009", "mm/dd/yyyy") & "')"
        Else

            If Trim(HANNOP) = vbNullString Then
                HANNOP = "CTOD('')"
            Else
                HANNOP = ToDate(Trim(HANNOP), DDMMYYYY)
                HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
            End If
        End If
        
        HANNOP2 = HANNOP
        DANHAN = "'T'"   ' To khai khong chuyen sang VAT
        matkhai = "'02B/TNCN11'"
        DaGhi = ".F."
        LoaiTien = "''"
         
        MaVach = ".T."
         
        MaMuc = "'1000'"
        maTM = "'1001'"
         
        MATHUE = "'09'"
        MAPP = "'02'"
        SHTEP = "''"
        AUTOCAL = ".T."
        GTTinhThue = "0"
        ' thue ky
'        .GetText .ColLetterToNumber("U"), 37, ThueTKy
        .GetText .ColLetterToNumber("U"), 53, ThueTKy
        If Trim(ThueTKy) = vbNullString Then
            ThueTKy = "0"
        End If

        ThueTKy2 = ThueTKy
        
        .GetText .ColLetterToNumber("G"), 15, MADLT
        .GetText .ColLetterToNumber("I"), 19, SOHDDL
        .GetText .ColLetterToNumber("S"), 19, NGAYHDDL
        .GetText .ColLetterToNumber("C"), 61, LANDAU

        If LANDAU = 1 Then
            LanBS = "0"
        Else
            .GetText .ColLetterToNumber("I"), 61, LanBS
        End If
        
    End With

'    sSQLVal = MATKHAI & "," & KYLBO & "," & NGNOP & "," & KYKKHAI & "," & MADTNT & "," & MADLT & ","
'    sSQLVal = sSQLVal & TTHTK & "," & DAGHI & "," & LOAITIEN & "," & MAVACH & "," & HANNOP & "," & HANNOP2 & ","
'    sSQLVal = sSQLVal & THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & "," & NGNHAP & "," & AUTO_CAL & ","
'    sSQLVal = sSQLVal & MATHUE & "," & MAPP & "," & SHTEP & "," & GTTINHTHUE & "," & LANBS & "," & THUECL & "," & GTTTCL & "," & LAN_QUET
'
'    sSQL = "INSERT INTO tmp_cntk" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
    ' Ghi du lieu vao bang trung gian de day len TC
    ' check PIT
    ' ghi giu lieu vao TG
    If TAX_Utilities_Svr_New.isCheckPIT = True Then
        arr = Split(GetMaCqt, "~")
        'hdrId = GetHdrId
        hdr.Tin = madtnt
        hdr.co_loi_dda = "''"
        
        hdr.DIA_CHI = arr(2)
        hdr.ghi_chu_loi = "''"
        hdr.khoa_so = "''"
        hdr.loai_tkhai = "'02B_TNCN13'"
        hdr.kieu_ky = "Q"

        If tthtk = "1" Then
            hdr.kkbs = 0    ' To khai chinh thuc
        ElseIf tthtk = "2" Then
            hdr.kkbs = 1    ' To khai thay the
        ElseIf tthtk = "3" Then
            hdr.kkbs = 2    ' To khai bo sung
        Else
            hdr.kkbs = 1    ' To khai nop cham
        End If

        hdr.kykkhai = kykkhai
        hdr.kylbo = kylbo
        hdr.MA_CQT = arr(0)
        hdr.ngay_cap_nhat = NgNhap
        hdr.Ngay_nop = ngnop
        'hdr.nguoi_cn = "'ADMIN'"
        hdr.nguoi_cn = "'" & Trim(strNguoiSuDung) & "'"
        hdr.Phong_xly = PHONG_XL
        hdr.so_hieu_tep = SHTEP
        hdr.So_tt_tk = LAN_QUET
        hdr.ten_dtnt = arr(1)
        hdr.DA_NHAN = 0
        hdr.tthtk = tthtk
        hdr.Id = hdrId
        hdr.thueondinh = "''"
  
        If Trim(NGAYHDDL) = vbNullString Then
            NGAYHDDL = "CTOD('')"
        Else
            NGAYHDDL = ToDate(Trim(NGAYHDDL), DDMMYYYY)
            NGAYHDDL = "CTOD('" & Format(NGAYHDDL, "mm/dd/yyyy") & "')"
        End If

        hdr.Ma_dl_thue = "'" & MADLT & "'"
        hdr.So_hd_dl = "'" & SOHDDL & "'"
        hdr.Ngay_hd_dl = NGAYHDDL
        hdr.lan_bs = LanBS
        
        sSQLTemp = InsertHDR_TGTC(hdr)
        bln = clsDAO.ExecuteDLL(sSQLTemp)
        clsDAO.Disconnect
    Else
          sSQLCol = "MATKHAI, KYLBO, NGNOP, KYKKHAI, MADTNT, MADLT, TTHTK, DAGHI, LOAITIEN, MAVACH, HANNOP, HANNOP2, THUETKY, THUETKY2, MAMUC, MATM, NGNHAP, AUTO_CAL, MATHUE, MAPP, SHTEP, GTTINHTHUE, LANBS, THUECL, GTTTCL, LAN_QUET, DANHAN, GHICHU"
         
            With fps
                .Sheet = 1
                matkhai = "'02Q/KK-TNCN'"
                 
                'ky lap bo
                .GetText .ColLetterToNumber("E"), 22, kylbo
                If Trim(kylbo) = vbNullString Then
                    kylbo = "''"
                Else
                    If Len(Trim(kylbo)) = 6 Then
                        kylbo = "'" & kylbo & "'"
                    Else
                        kylbo = "'" & kylbo & "'"
                    End If
                End If
                'ghi chu
                .GetText .ColLetterToNumber("E"), 28, GHICHU
                If Trim(GHICHU) = vbNullString Then
                    GHICHU = "''"
                Else
                    GHICHU = "'" & GHICHU & "'"
                End If
                ' ngay nop
                .GetText .ColLetterToNumber("E"), 24, ngnop
                If Trim(ngnop) = vbNullString Then
                    ngnop = "CTOD('')"
                Else
                    ngnop = ToDate(Trim(ngnop), DDMMYYYY)
                    ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
                End If
                'ky ke khai
                kykkhai = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
                'ma doi tuong nop thue
                .GetText .ColLetterToNumber("G"), 8, madtnt
                If Trim(madtnt) = vbNullString Then
                    madtnt = "''"
                Else
                    madtnt = "'" & madtnt & "'"
                End If
               'ma dai ly thue
                .GetText .ColLetterToNumber("G"), 15, MADLT
                If Trim(MADLT) = vbNullString Then
                    MADLT = "''"
                Else
                    MADLT = "'" & MADLT & "'"
                End If
                 ' ngay nhap
                .GetText .ColLetterToNumber("M"), 24, NgNhap
                If Trim(NgNhap) = vbNullString Then
                    NgNhap = "CTOD('')"
                Else
                    NgNhap = ToDate(Trim(NgNhap), DDMMYYYY)
                    NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
                End If
                 'trang thai to khai
                .GetText .ColLetterToNumber("L"), 5, TTHTK1
                If Trim(TTHTK1) = vbNullString Then
                    tthtk = "'" & 2 & "'"
                End If
                 ' so lan bo sung
                 If tthtk = "'" & 2 & "'" Then
                 .GetText .ColLetterToNumber("T"), 5, solanBS
                    LanBS = solanBS
                 Else
                    LanBS = "0"
                 End If
                ' thue ky
'                .GetText .ColLetterToNumber("U"), 37, ThueTKy
                .GetText .ColLetterToNumber("U"), 52, ThueTKy
                 If Trim(ThueTKy) = vbNullString Then
                        ThueTKy = "0"
                 End If
                 ThueTKy2 = ThueTKy
                 DaGhi = ".F."
                 LoaiTien = "''"
                 MaVach = ".T."
                 If Trim(HANNOP) = vbNullString Then
                    HANNOP = "CTOD('')"
                 End If
                 HANNOP2 = HANNOP
                 MaMuc = "'1000'"
                 maTM = "'1001'"
                 AUTO_CAL = ".T."
                 MATHUE = "'09'"
                 MAPP = "'02'"
                 SHTEP = "''"
                 GTTinhThue = "0"
                 THUECL = "0"
                 GTTTCL = "0"
                 DANHAN = "''"
            End With
                sSQLVal = matkhai & "," & kylbo & "," & ngnop & "," & kykkhai & "," & madtnt & "," & MADLT & ","
                sSQLVal = sSQLVal & tthtk & "," & DaGhi & "," & LoaiTien & "," & MaVach & "," & HANNOP & "," & HANNOP2 & ","
                sSQLVal = sSQLVal & ThueTKy & "," & ThueTKy2 & "," & MaMuc & "," & maTM & "," & NgNhap & "," & AUTO_CAL & ","
                sSQLVal = sSQLVal & MATHUE & "," & MAPP & "," & SHTEP & "," & GTTinhThue & "," & LanBS & "," & THUECL & "," & GTTTCL & "," & LAN_QUET & "," & DANHAN & "," & GHICHU
                
                 sSQL = "INSERT INTO tmp_cntk" & Right(sKyKeKhai, 4) & _
                 "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        
    End If

    ' end PIT
    
    InsertHDR = sSQL
    
End Function

Public Function TKTT() As Boolean
    Dim sSQL            As String
    Dim rs              As ADODB.Recordset
    Dim fso             As New FileSystemObject
    Dim strFileNameHDR  As String
    Dim strFileNameHDR1 As String
    Dim strFileNameDTL  As String
    Dim strFileNameDTL1 As String
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_CNTK" & Right(sKyKeKhai, 4) & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TQ2" & Right(sKyKeKhai, 4) & ".DBF"
    strFileNameHDR1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tmp_tncn_hdr.DBF"
    strFileNameDTL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tmp_tncn_dtl.DBF"

    'Dim strFileNameBS As String
    '    strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & Right(sKyKeKhai, 4) & ".DBF"
    'file HDR
    If fso.FileExists(strFileNameHDR) = False Then
        Dim strTepmauHDR As String
        strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tepmau\TMP_CNTK.DBF"
        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    End If

    'file dtl
    If fso.FileExists(strFileNameDTL) = False Then
        Dim strTepmauDTL As String
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tepmau\TMP_TQ2.DBF"
        fso.CopyFile strTepmauDTL, strFileNameDTL, False
    End If

    'file HDR1
    If fso.FileExists(strFileNameHDR1) = False Then
        Dim strTepmauHDR1 As String
        strTepmauHDR1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tepmau\tmp_tncn_hdr.DBF"
        fso.CopyFile strTepmauHDR1, strFileNameHDR1, False
    End If

    'file dtl1
    If fso.FileExists(strFileNameDTL1) = False Then
        Dim strTepmauDTL1 As String
        strTepmauDTL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tepmau\tmp_tncn_dtl.DBF"
        fso.CopyFile strTepmauDTL1, strFileNameDTL1, False
    End If

    'file bo xung
    '    If fso.FileExists(strFileNameBS) = False Then
    '        Dim strTepmauBS As String
    '        strTepmauBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_BS.DBF"
    '        fso.CopyFile strTepmauBS, strFileNameBS, False
    '    End If

    isExistFile = True
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    'Truong hop pit
    If TAX_Utilities_Svr_New.isCheckPIT = True Then
        sSQL = "SELECT * FROM tmp_tncn_hdr WHERE Tin = '" & strMST & "' AND KYKKHAI = '" & TAX_Utilities_Svr_New.ThreeMonths & "/" & TAX_Utilities_Svr_New.Year & "'" & "  AND Loai_tkhai='02B_TNCN13'"
        Set rs = clsDAO.Execute(sSQL)
        If Not rs Is Nothing Then
             dNgayNopDB = rs.Fields("Ngay_nop")

            If Trim(dNgayNopDB) = vbNullString Then
                dNgayNopDB = "CTOD('')"
            Else
                dNgayNopDB = "CTOD('" & Format(dNgayNopDB, "mm/dd/yyyy") & "')"
            End If
            TKTT = True
        Else
            TKTT = False
        End If
    
    ' khong phai PIT
    Else
        sSQL = "SELECT MADTNT, NGNOP FROM TMP_CNTK" & Right(sKyKeKhai, 4) & " WHERE MADTNT = " & "'" & strMST & "'" & " AND KYKKHAI = '" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='02Q/KK-TNCN'"
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
    End If
   
    LAN_QUET = TinhSoLanQuet + 1
    '   hdrId = GetHdrId(spathVat)
    clsDAO.Disconnect

End Function

Public Function UpdateTHUETKY2() As Boolean
   Dim sSQL As String

   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   sSQL = "UPDATE TMP_CNTK" & Right(sKyKeKhai, 4) & _
        " SET THUETKY2 = 1 " & _
        " WHERE MATHUE = '09' and MAPP = '02' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='02Q/KK-TNCN'"

   clsDAO.ExecuteDLL sSQL

   UpdateTHUETKY2 = True

   clsDAO.Disconnect
End Function

Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL As String

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If

        ' Truong hop Pit
    If TAX_Utilities_Svr_New.isCheckPIT = True Then
        sSQL = "SELECT * FROM tmp_tncn_hdr WHERE Tin = '" & strMST & "' AND KYKKHAI = '" & TAX_Utilities_Svr_New.ThreeMonths & "/" & TAX_Utilities_Svr_New.Year & "'" & " AND Lan_bs = 0 AND Loai_tkhai='02B_TNCN13'"
        Set rs = clsDAO.Execute(sSQL)
        If rs Is Nothing Then
            isToKhaiChinhThuc = False
        Else
            isToKhaiChinhThuc = True
        End If
    Else
        ' Truong hop khong phai Pit
        sSQL = "SELECT * from tmp_cntk" & Right(sKyKeKhai, 4) & " WHERE TTHTK = '1' AND MADTNT = " & "'" & strMST & "'" & " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='02Q/KK-TNCN'"
   
        Set rs = clsDAO.Execute(sSQL)

        If rs Is Nothing Then
            isToKhaiChinhThuc = False
        Else
            isToKhaiChinhThuc = True
        End If
    End If
   
    clsDAO.Disconnect
End Function

Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim kylbo As Variant
    Dim rs As ADODB.Recordset
    
    If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
   sSQL = "DELETE FROM tmp_cntk" & Right(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='02Q/KK-TNCN'"
        
    
   clsDAO.ExecuteDLL sSQL
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 22, kylbo
        If Trim(kylbo) = vbNullString Then
            kylbo = "''"
        Else
            If Len(Trim(kylbo)) = 6 Then
                kylbo = "'0" & kylbo & "'"
            Else
                kylbo = "'" & kylbo & "'"
            End If
        End If
   End With
   sSQL = "DELETE FROM TMP_TQ2" & Right(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "' AND KYKKHAI = '" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "' " & _
        " AND MATKHAI='02Q/KK-TNCN'"
    
   clsDAO.ExecuteDLL sSQL
   
   XoaTKTT = True
   
   clsDAO.Disconnect

End Function

Public Function TKRB() As Boolean

'    Dim sSQL As String
'    Dim THUEGTGT_KT As Double
'    Dim rs As ADODB.Recordset
'    Dim sKyTruoc As String
'    Dim sKyKKTruoc As String
'    Dim fso As New FileSystemObject
'    Dim THUEGTGT As Variant
'    Dim dNgayNopDBKT As Variant
'
'    If clsDAO.Connected = False Then
''        Me.MousePointer = vbHourglass
''        frmSystem.MousePointer = vbHourglass
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
''        frmSystem.MousePointer = vbDefault
''        Me.MousePointer = vbDefault
'   End If
'
'   If CInt(Month(Now())) = 1 Then
'        sKyTruoc = "12" & Trim(Str(CInt(Year(Now())) - 1))
'   Else
'        sKyTruoc = IIf(Len(Trim(Str(CInt(Month(Now())) - 1))) = 1, "0" & Trim(Str(CInt(Month(Now())) - 1)), Trim(Str(CInt(Month(Now())) - 1))) & Year(Now())
'   End If
'
'   If CInt(TAX_Utilities_Svr_New.Month) = 1 Then
'        sKyKKTruoc = "12" & TAX_Utilities_Svr_New.Year
'   Else
'        sKyKKTruoc = IIf(Len(Trim(Str(CInt(TAX_Utilities_Svr_New.Month) - 1))) = 1, "0" & Trim(Str(CInt(TAX_Utilities_Svr_New.Month) - 1)), Trim(Str(CInt(TAX_Utilities_Svr_New.Month) - 1))) & TAX_Utilities_Svr_New.Year
'   End If
'
'   If Not fso.FileExists(spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TK" & sKyTruoc & ".dbf") Then
'        TKRB = True
'        clsDAO.Disconnect
'        Exit Function
'   End If
'
'   sSQL = "SELECT MADTNT, NGNOP FROM TK" & sKyTruoc & _
'        " WHERE MATHUE = '01' AND MATM = '01' AND MAPP = '5' AND MADTNT = " & "'" & strMST & "'" & _
'        " AND KYKKHAI = '" & Left(sKyKKTruoc, 2) & "/" & Right(sKyKKTruoc, 4) & "'"
'
'   Set rs = clsDAO.Execute(sSQL)
'   If Not rs Is Nothing Then
'        dNgayNopDBKT = rs.Fields("NGNOP")
'        If Trim(dNgayNopDBKT) = vbNullString Then
'            dNgayNopDBKT = "CTOD('')"
'        Else
'            dNgayNopDBKT = "CTOD('" & Format(dNgayNopDBKT, "mm/dd/yyyy") & "')"
'        End If
'   Else
'        TKRB = True
'        clsDAO.Disconnect
'        Exit Function
'   End If
'   If Not fso.FileExists(spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "KT" & sKyTruoc & ".dbf") Then
'        TKRB = True
'        clsDAO.Disconnect
'        Exit Function
'   End If
'
'   sSQL = "SELECT thuegtgt FROM KT" & sKyTruoc & _
'        " WHERE CT_KT = '085' AND MADTNT = " & "'" & strMST & "' AND NGNOP = " & dNgayNopDBKT
'
'   Set rs = clsDAO.Execute(sSQL)
'   If Not rs Is Nothing Then
'        THUEGTGT_KT = rs.Fields("thuegtgt")
'        If THUEGTGT_KT = 0 Then
'            TKRB = True
'            clsDAO.Disconnect
'            Exit Function
'        Else
'            With fps
'                .Sheet = 1
'                .GetText .ColLetterToNumber("W"), 39, THUEGTGT
'                If Trim(THUEGTGT) = vbNullString Then
'                    THUEGTGT = "0"
'                End If
'                If THUEGTGT_KT = THUEGTGT Then
'                     TKRB = True
'                     clsDAO.Disconnect
'                     Exit Function
'                Else
'                      TKRB = False
'                      clsDAO.Disconnect
'                      Exit Function
'                End If
'            End With
'        End If
'   Else
'        TKRB = True
'        clsDAO.Disconnect
'        Exit Function
'   End If
'
   TKRB = True
'   clsDAO.Disconnect

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
        .GetText .ColLetterToNumber("E"), 24, ngnop
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
   
   sSQL = "SELECT MADTNT, NGNOP FROM tmp_cntk" & Right(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI <> '" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        
    
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
        If Col = .ColLetterToNumber("E") And Row = 24 Then
            .Sheet = 1
            .GetText Col, Row, varTemp
            If varTemp <> "" And varTemp <> "../../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                Else
                    .SetActiveCell Col, Row
                End If
            Else
            .Col = .ColLetterToNumber("E")
            .Row = 24
             .SetText Col, Row, ""
            End If
           UpdateCell fps, .Col, .Row, .Text
        End If
        If Col = .ColLetterToNumber("E") And Row = 22 Then
            .Sheet = 1
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
            .Col = .ColLetterToNumber("E")
            .Row = 22
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
    
    'kiem tra ton tai tep *.dbf chua
    'Tinh so lan quyet de ghi vao DB trung gian
    'dhdang sua
    'ngay 24/09/2010
'-----------------------------
    If TAX_Utilities_Svr_New.isCheckPIT = False Then
        strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_CNTK" & Right(sKyKeKhai, 4) & ".DBF"
        If fso.FileExists(strFileNameHDR) = True Then
            If clsDAO.Connected = False Then
            clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
            clsDAO.Connect
            End If
            
           sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_CNTK" & Right(sKyKeKhai, 4) & _
                    " WHERE MADTNT = " & "'" & strMST & "'" & _
                    " AND KYKKHAI = '" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='02Q/KK-TNCN' AND (TTHTK = '1' OR TTHTK = '3' OR TTHTK = '4' or TTHTK = '2')"
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
    Else
        Dim kykk_tu_ngay As Variant
        Dim kykk_den_ngay As Variant
        Dim tKyKekhai As Variant
        Dim strDate As Variant
        Dim dDate As Variant
        
        strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TNCN_HDR" & ".DBF"
        tKyKekhai = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        tKyKekhai = Replace(tKyKekhai, "'", "")
        strDate = Split(tKyKekhai, "/")
        dDate = GetNgayDauQuy(Int(strDate(0)), Int(strDate(1)))
        kykk_tu_ngay = dDate
        kykk_den_ngay = DateAdd("m", 3, kykk_tu_ngay)
        kykk_den_ngay = DateAdd("d", -1, kykk_den_ngay)
        If Trim(kykk_tu_ngay) = vbNullString Then
                kykk_tu_ngay = "CTOD('')"
        Else
                kykk_tu_ngay = "CTOD('" & Format(kykk_tu_ngay, "mm/dd/yyyy") & "')"
        End If
        If Trim(kykk_den_ngay) = vbNullString Then
                kykk_den_ngay = "CTOD('')"
        Else
                kykk_den_ngay = "CTOD('" & Format(kykk_den_ngay, "mm/dd/yyyy") & "')"
        End If
  
        If fso.FileExists(strFileNameHDR) = True Then
            If clsDAO.Connected = False Then
                clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
                clsDAO.Connect
            End If
             sSQL = "SELECT max(so_tt_tk) as LAN_QUET FROM TMP_TNCN_HDR" & _
                    " WHERE TIN = " & "'" & strMST & "'" & _
                    " And loai_tkhai = '02B_TNCN13' " & _
                    " And kykk_tu_ng = " & kykk_tu_ngay & _
                    " And kykk_den_n =" & kykk_den_ngay
                Set rs = clsDAO.Execute(sSQL)
                If Not rs Is Nothing Then
                     LAN_QUET1 = Val(rs.Fields("LAN_QUET"))
    
                     TinhSoLanQuet = LAN_QUET1
                Else
                     TinhSoLanQuet = 0
                End If
                clsDAO.Disconnect
        Else
        TinhSoLanQuet = 0
        End If
    End If
End Function

'' Ham de ghi du lieu vao bang trung gian de gui to khai len TC
'Public Function InsertHDR_TGTC(ByRef hdr As TNCN_HDR) As String
'    Dim sSQLCol As String
'    Dim sSQLVal As String
'    Dim sSQL As String
'
'    Dim kykk_tu_ngay As Variant
'    Dim kykk_den_ngay As Variant
'    Dim kyLB_tu_ngay As Variant
'    Dim kyLB_den_ngay As Variant
'    Dim tKyKekhai As Variant
'    Dim tKyLB As Variant
'    Dim dDate As Date, strDate() As String
'     'Ky/ Quy KK
'    If hdr.kieu_ky = "M" Then
'        'Ngay dau ky ke khai va ngay cuoi ky ke khai
'        tKyKekhai = hdr.KYKKHAI
'        tKyKekhai = Replace(tKyKekhai, "'", "")
'        strDate = Split(tKyKekhai, "/")
'        dDate = DateSerial(Int(strDate(1)), Int(strDate(0)), 1)
'        kykk_tu_ngay = dDate
'        kykk_den_ngay = DateAdd("m", 1, kykk_tu_ngay)
'        kykk_den_ngay = DateAdd("d", -1, kykk_den_ngay)
'
'        If Trim(kykk_den_ngay) = vbNullString Then
'                kykk_den_ngay = "CTOD('')"
'        Else
'                kykk_den_ngay = "CTOD('" & Format(kykk_den_ngay, "mm/dd/yyyy") & "')"
'        End If
'
'        If Trim(kykk_tu_ngay) = vbNullString Then
'                kykk_tu_ngay = "CTOD('')"
'        Else
'                kykk_tu_ngay = "CTOD('" & Format(kykk_tu_ngay, "mm/dd/yyyy") & "')"
'        End If
'    ElseIf hdr.kieu_ky = "Q" Then
'        tKyKekhai = hdr.KYKKHAI
'        tKyKekhai = Replace(tKyKekhai, "'", "")
'        strDate = Split(tKyKekhai, "/")
'        dDate = GetNgayDauQuy(Int(strDate(0)), Int(strDate(1)))
'        kykk_tu_ngay = dDate
'        kykk_den_ngay = DateAdd("m", 3, kykk_tu_ngay)
'        kykk_den_ngay = DateAdd("d", -1, kykk_den_ngay)
'        If Trim(kykk_tu_ngay) = vbNullString Then
'                kykk_tu_ngay = "CTOD('')"
'        Else
'                kykk_tu_ngay = "CTOD('" & Format(kykk_tu_ngay, "mm/dd/yyyy") & "')"
'        End If
'        If Trim(kykk_den_ngay) = vbNullString Then
'                kykk_den_ngay = "CTOD('')"
'        Else
'                kykk_den_ngay = "CTOD('" & Format(kykk_den_ngay, "mm/dd/yyyy") & "')"
'        End If
'    Else
'        kykk_tu_ngay = "CTOD('')"
'        kykk_den_ngay = "CTOD('')"
'    End If
'
'    'Ky lb
'    If hdr.kieu_ky = "M" Or hdr.kieu_ky = "Q" Then
'        'Ngay dau ky lap bo va ngay cuoi ky lap bo
'        tKyLB = hdr.KYLBO
'        tKyLB = Replace(tKyLB, "'", "")
'        strDate = Split(tKyLB, "/")
'        dDate = DateSerial(Int(strDate(1)), Int(strDate(0)), 1)
'        kyLB_tu_ngay = dDate
'        kyLB_den_ngay = DateAdd("m", 1, kyLB_tu_ngay)
'        kyLB_den_ngay = DateAdd("d", -1, kyLB_den_ngay)
'        If Trim(kyLB_den_ngay) = vbNullString Then
'                kyLB_den_ngay = "CTOD('')"
'        Else
'                kyLB_den_ngay = "CTOD('" & Format(kyLB_den_ngay, "mm/dd/yyyy") & "')"
'        End If
'
'        If Trim(kyLB_tu_ngay) = vbNullString Then
'                kyLB_tu_ngay = "CTOD('')"
'        Else
'                kyLB_tu_ngay = "CTOD('" & Format(kyLB_tu_ngay, "mm/dd/yyyy") & "')"
'        End If
'    ElseIf hdr.kieu_ky = "Q" Then
'
'    Else
'        kyLB_den_ngay = "CTOD('')"
'        kyLB_tu_ngay = "CTOD('')"
'    End If
'
'
'
'    sSQLCol = "INSERT INTO tmp_tncn_hdr (tin, ten_dtnt, dia_chi, loai_tkhai, ngay_nop, kylb_tu_ng, kylb_den_n, kykk_tu_ng, kykk_den_n, ngay_cap_n,"
'    sSQLCol = sSQLCol + " nguoi_cn, co_loi_dda, so_hieu_te, so_tt_tk, da_nhan, ghi_chu_lo, khoa_so, phong_xly, kkbs, tthtk, kylbo, kykkhai, ma_cqt) "
'    sSQLCol = sSQLCol + " values ("
'
'    sSQLVal = hdr.tin & "," & hdr.ten_dtnt & "," & hdr.DIA_CHI & "," & hdr.loai_tkhai & "," & hdr.ngay_nop & "," & kyLB_tu_ngay & "," & _
'    kyLB_den_ngay & "," & kykk_tu_ngay & "," & kykk_den_ngay & "," & hdr.ngay_cap_nhat & "," & hdr.nguoi_cn & "," & hdr.co_loi_dda & "," & _
'    hdr.so_hieu_tep & "," & hdr.so_tt_tk & "," & hdr.DA_NHAN & "," & hdr.ghi_chu_loi & ", " & hdr.khoa_so & ", " & hdr.phong_xly & ", " & hdr.kkbs & ", " & hdr.TTHTK & "," & hdr.KYLBO & "," & hdr.KYKKHAI & "," & hdr.ma_cqt
'
'    sSQL = sSQLCol & sSQLVal & " )"
'
'   InsertHDR_TGTC = sSQL
'End Function
'' end
'' Ham de ghi du lieu vao bang trung gian de gui to khai len TC
'Public Function InsertDTL_TGTC(ByRef dtl As TNCN_DTL) As String
'    Dim sSQLCol As String
'    Dim sSQLVal As String
'    Dim sSQL As String
'
'    sSQLCol = "INSERT INTO tmp_tncn_dtl (matkhai, madtnt, kylbo, kykkhai, tthtk, ngnop, cttn, giatri, danhan, lan_quet, ky_hieu, ma_cqt) "
'    sSQLCol = sSQLCol + " values ("
'
'    sSQLVal = dtl.MATKHAI & "," & dtl.MADTNT & "," & dtl.KYLBO & "," & dtl.KYKKHAI & "," & dtl.TTHTK & "," & dtl.NGNOP & "," & _
'    dtl.CTTN & "," & dtl.GIATRI & "," & dtl.DANHAN & "," & dtl.LAN_QUET & "," & dtl.ky_hieu & "," & dtl.ma_cqt
'
'    sSQL = sSQLCol & sSQLVal & " )"
'
'   InsertDTL_TGTC = sSQL
'End Function
'' end
'' tinh ngay dau cua Quy
'Public Function GetNgayDauQuy(q As Integer, Y As Integer) As Date
'    Dim intYear As Integer, intDay As Integer, intMonth As Integer
'    intDay = 1
'    intYear = Y
'    Select Case q
'        Case 1
'            intMonth = 1
'        Case 2
'            intMonth = 4
'        Case 3
'            intMonth = 7
'        Case 4
'            intMonth = 10
'    End Select
'    GetNgayDauQuy = DateSerial(intYear, intMonth, intDay)
'End Function
'Get ve ma CQT
Public Function GetMaCqt() As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim maCQT As Variant
    Dim TEMP As String
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO

        If clsConn.Connected = False Then
            clsConn.CreateConnectionString spathVat & "\dtnt\"
            clsConn.Connect
         End If
         sSQL = "SELECT madbhc,tengoi,dchi FROM dtnt2 WHERE MADTNT =  '" & strMST & "'"
        
        Set rs = clsConn.Execute(sSQL)
        If Not rs Is Nothing Then
            maCQT = rs.Fields("madbhc")
            TEMP = "'" & Mid$(Trim(maCQT), 1, 5) & "'"
            TEMP = TEMP & "~'" & Trim(rs.Fields("tengoi")) & "'"
            TEMP = TEMP & "~'" & Trim(rs.Fields("dchi")) & "'"
        Else
            TEMP = "''~''~''"
        End If
        clsConn.Disconnect
        GetMaCqt = TEMP
End Function
