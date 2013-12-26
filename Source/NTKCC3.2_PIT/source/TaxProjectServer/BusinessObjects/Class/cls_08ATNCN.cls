VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_08ATNCN"
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
Public tthtk As String
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
'dntai kieu ke khai la tu thang, den thang thi kieuKeKhai ="T", ke khai quy thi ="Q", bien de luu gia tri ke khai tu thang
Public kieuKeKhai As String
Public vkKhaiTuThang As String


Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
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
        .CellType = CellTypeEdit
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
        
        ' Ngay quet
        SetDateFormat fps, 1, 28, .ColLetterToNumber("N"), DDMMYYYY
        
        ' Lay ve gia tri cua HDR ID khi quet to khai
        hdrId = GetHdrId(spathVat)
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

Public Function Prepared3() As Boolean
    Dim vTuThang As Variant, vDenThang As Variant

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

        'set phong quan ly
        .Col = .ColLetterToNumber(PHONQ_QUAN_LY_Y)
        .Row = PHONG_QUAN_LY_X
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))



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
        
        'get gia tri tu thang, den thang de set lai hinh thuc ke khai tren header
        .GetText .ColLetterToNumber("K"), 3, vTuThang
        .GetText .ColLetterToNumber("P"), 3, vDenThang

        If Trim(vTuThang) <> vbNullString Or Trim(vDenThang) <> vbNullString Then
            .SetText .ColLetterToNumber("E"), 50, "1"
            kieuKeKhai = "T"
            vkKhaiTuThang = CStr(Trim(vTuThang))
        Else
            .SetText .ColLetterToNumber("E"), 50, "2"
            kieuKeKhai = "Q"
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

Public Function CheckValidData() As Boolean
    Dim NGAYNOP As Variant, ngayHienTai As Variant
    CheckValidData = True
    With fps
        .Sheet = 1
        
        ' check ngay nop khong dc lon hon ngay hien tai
        .GetText .ColLetterToNumber("E"), 32, NGAYNOP
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
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("E_30"), "Value"), "dd/mm/yyyy")
    
    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function

Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

Public Function InsertDTL_KHBS() As String

End Function

Public Function InsertDTL() As String
    Dim sSQL    As String
    Dim sSQLCol As String
    Dim sSQLVal As String
   
    Dim matkhai As Variant
    Dim madtnt  As Variant
    Dim kylbo   As Variant
    Dim KyKKhai As Variant
    Dim ngnop   As Variant
    Dim CTTN    As Variant
    Dim giatri  As Variant
    Dim tt      As Variant
    Dim giatri2 As Variant
    Dim MaMuc   As Variant
    Dim maTM    As Variant
    Dim DA_NHAN As Variant
    Dim bln     As Boolean
    Dim vTuThang As Variant, vDenThang As Variant
'bien dung cho PIT
    Dim cttnArr()   As Variant 'mang cua cac Cell chi tieu
    Dim cellArr()   As Variant 'mang gia tri cua cac Cell chi tieu
    Dim kyhieuArr() As Variant ' mang ky hieu theo chi tieu
    Dim hdr As TNCN_DTL
    Dim i As Integer, rowDynamic As Integer, rowID As Integer
    Dim arr()   As String
   
    bln = False
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
    sSQLCol = "MATKHAI,MADTNT,KYLBO,KYKKHAI,TTHTK,NGNOP,CTTN,GIATRI,TT,GIATRI2,MAMUC,MATM,DANHAN,LAN_QUET"
        
    With fps
            .Sheet = 1
            matkhai = "'08A/KK-TNCN'"
            'get MADTNT
            .GetText .ColLetterToNumber("F"), 10, madtnt
    
            If madtnt = vbNullString Then
                madtnt = "''"
            Else
                madtnt = "'" & madtnt & "'"
            End If
        
            'get KYLBO
            .GetText .ColLetterToNumber("E"), 30, kylbo
    
            If kylbo = vbNullString Then
                kylbo = "''"
            Else
                kylbo = "'" & kylbo & "'"
            End If
        
            ' set KYKKHAI
            KyKKhai = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        
            ' get ngaynop
            .GetText .ColLetterToNumber("E"), 32, ngnop
    
            If ngnop = vbNullString Then
                ngnop = "CTOD('')"
            Else
                ngnop = ToDate(Trim(ngnop), DDMMYYYY)
                ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
            End If
        
            DA_NHAN = "''"

            'get tu thang , den thang
            .GetText .ColLetterToNumber("K"), 3, vTuThang
            .GetText .ColLetterToNumber("P"), 3, vDenThang
            If Trim(vTuThang) <> vbNullString And Trim(vDenThang) <> vbNullString Then
                MaMuc = "'1000'"
                maTM = "'1004'"
            Else
                MaMuc = "'1000'"
                maTM = "'1003'"
            End If

        If TAX_Utilities_Svr_New.isCheckPIT Then
            If Trim(vTuThang) <> vbNullString And Trim(vDenThang) <> vbNullString Then
                KyKKhai = "'" & vTuThang & "'"
            End If
            cttnArr() = Array("21", "22", "23", "24")
            kyhieuArr() = Array("21", "22", "23", "24")
            cellArr() = Array(44, 45, 46, 47)
            
            For i = 0 To UBound(cttnArr)
            
            CTTN = "'0" & cttnArr(i) & "'"
                
            .GetText .ColLetterToNumber("T"), cellArr(i), giatri
            giatri2 = giatri
                
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & KyKKhai & "," & tthtk & "," & ngnop & ","
            sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DA_NHAN

            
                arr = Split(GetMaCqt, "~")
                hdr.matkhai = "'08A_TNCN11'"
                hdr.madtnt = madtnt
                hdr.kylbo = kylbo
                hdr.KyKKhai = KyKKhai
                hdr.tthtk = tthtk
                hdr.ngnop = ngnop
                hdr.CTTN = CTTN
                'tren to nay co ghi du lieu kieu chuoi nen truong gia tri da doi thanh kieu du lieu text
                hdr.giatri = "'" & giatri & "'"
                            
                hdr.LAN_QUET = LAN_QUET
                hdr.ky_hieu = "'" & kyhieuArr(i) & "'"
                hdr.MA_CQT = arr(0)
                hdr.DANHAN = 0
                hdr.Id = GetDtlId(spathVat)
                hdr.Hdr_id = hdrId
          

                sSQL = InsertDTL_TGTC08(hdr, rowID)
                bln = clsDAO.ExecuteDLL(sSQL)

            Next

            'dntai 04/02/2012 them phan ghi dong dong
            rowDynamic = 0
            .Row = 61
            Do
                cttnArr() = Array("26", "27", "28", "29", "30", "31", "32", "33", "34", "35")
                kyhieuArr() = Array("26", "27", "28", "29", "30", "31", "32", "33", "34", "35")
                cellArr() = Array(.ColLetterToNumber("C"), .ColLetterToNumber("H"), .ColLetterToNumber("L"), .ColLetterToNumber("N"), .ColLetterToNumber("Q"), .ColLetterToNumber("U"), .ColLetterToNumber("X"), .ColLetterToNumber("AB"), .ColLetterToNumber("AE"), .ColLetterToNumber("AG"))
                For i = 0 To UBound(cttnArr)
                
                    CTTN = "'0" & cttnArr(i) & "'"
                        
                    .GetText cellArr(i), .Row, giatri
                    giatri2 = giatri
                        
                    sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & KyKKhai & "," & tthtk & "," & ngnop & ","
                    sSQLVal = sSQLVal & CTTN & "," & giatri & "," & giatri2 & "," & tt & "," & LAN_QUET & "," & DA_NHAN
        
                    
                    arr = Split(GetMaCqt, "~")
                    hdr.matkhai = "'08A_TNCN11'"
                    hdr.madtnt = madtnt
                    hdr.kylbo = kylbo
                    hdr.KyKKhai = KyKKhai
                    hdr.tthtk = tthtk
                    hdr.ngnop = ngnop
                    hdr.CTTN = CTTN
                    hdr.giatri = "'" & giatri & "'"
                                
                    hdr.LAN_QUET = LAN_QUET
                    hdr.ky_hieu = "'" & kyhieuArr(i) & "'"
                    hdr.MA_CQT = arr(0)
                    hdr.DANHAN = 0
                    hdr.Id = GetDtlId(spathVat)
                    hdr.Hdr_id = hdrId
                    rowID = rowDynamic + 1
    
                    sSQL = InsertDTL_TGTC08(hdr, rowID)
                    bln = clsDAO.ExecuteDLL(sSQL)
    
                Next
                rowDynamic = rowDynamic + 1
                .Row = 61 + rowDynamic
                .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
            
        Else
        '---------------------------------------------------------------------
        
            'set theo CT21
            
            CTTN = "'001'"
            tt = "'1'"
            
            'get GT
            .GetText .ColLetterToNumber("T"), 44, giatri
            
            If giatri = vbNullString Then
                giatri = "0"
            End If
            
            giatri2 = giatri
            
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & KyKKhai & "," & tthtk & "," & ngnop & "," & CTTN
            sSQLVal = sSQLVal & "," & giatri & "," & tt & "," & giatri2 & "," & MaMuc & "," & maTM & "," & DA_NHAN & "," & LAN_QUET
        
            sSQL = "INSERT INTO TMP_C08" & Right$(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            
            bln = clsDAO.ExecuteDLL(sSQL)
            
    '---------------------------------------------------------------------
        
            'set theo CT22
            
            CTTN = "'002'"
            tt = "'2'"
            
            'get GT
            .GetText .ColLetterToNumber("T"), 45, giatri
            
            If giatri = vbNullString Then
                giatri = "0"
            End If
            
            giatri2 = giatri
            
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & KyKKhai & "," & tthtk & "," & ngnop & "," & CTTN
            sSQLVal = sSQLVal & "," & giatri & "," & tt & "," & giatri2 & "," & MaMuc & "," & maTM & "," & DA_NHAN & "," & LAN_QUET
        
            sSQL = "INSERT INTO TMP_C08" & Right$(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            
            bln = clsDAO.ExecuteDLL(sSQL)
            
    '---------------------------------------------------------------------
        
            'set theo CT23
            
            CTTN = "'003'"
            tt = "'3'"
            
            'get GT
            .GetText .ColLetterToNumber("T"), 46, giatri
            
            If giatri = vbNullString Then
                giatri = "0"
            End If
            
            giatri2 = giatri
            
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & KyKKhai & "," & tthtk & "," & ngnop & "," & CTTN
            sSQLVal = sSQLVal & "," & giatri & "," & tt & "," & giatri2 & "," & MaMuc & "," & maTM & "," & DA_NHAN & "," & LAN_QUET
        
            sSQL = "INSERT INTO TMP_C08" & Right$(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            
            bln = clsDAO.ExecuteDLL(sSQL)
            
    '---------------------------------------------------------------------
        
            'set CT24
            
            CTTN = "'004'"
            tt = "'4'"
            
            'get GT
            .GetText .ColLetterToNumber("T"), 47, giatri
            
            If giatri = vbNullString Then
                giatri = "0"
            End If
            
            giatri2 = giatri
            
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & KyKKhai & "," & tthtk & "," & ngnop & "," & CTTN
            sSQLVal = sSQLVal & "," & giatri & "," & tt & "," & giatri2 & "," & MaMuc & "," & maTM & "," & DA_NHAN & "," & LAN_QUET
        
            sSQL = "INSERT INTO TMP_C08" & Right$(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            
            bln = clsDAO.ExecuteDLL(sSQL)

    '---------------------------------------------------------------------


        End If
                  
     End With
        InsertDTL = vbNullString
        clsDAO.Disconnect
End Function

Public Function InsertHDR() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String

    Dim matkhai    As Variant
    Dim kylbo      As Variant
    Dim ngnop      As Variant
    Dim KyKKhai    As Variant
    Dim madtnt     As Variant
    Dim MADLT      As Variant
    Dim DaGhi      As Variant
    Dim LoaiTien   As Variant
    Dim MaVach     As Variant
    Dim HANNOP2    As Variant
    Dim ThueTKy    As Variant
    Dim ThueTKy2   As Variant
    Dim MaMuc      As Variant
    Dim maTM       As Variant
    Dim NgNhap     As Variant
    Dim AUTO_CAL   As Variant
    Dim MATHUE     As Variant
    Dim MAPP       As Variant
    Dim SHTEP      As Variant
    Dim GTTinhThue As Variant
    Dim THUECL     As Variant
    Dim GTTTCL     As Variant
    Dim DANHAN     As Variant
    Dim GHICHU     As Variant
    Dim LanBS      As Variant
    Dim bln        As Boolean
    Dim vTuThang As Variant, vDenThang As Variant
    'bien dung cho PIT
    Dim hdr      As TNCN_HDR
    Dim arr()    As String
    Dim NGAYHDDLT  As Variant
    Dim SOHDDL    As Variant
    Dim PXULY      As Variant
    Dim kKEKHAITUNGAY As Variant
    Dim kKEKHAIDENNGAY As Variant
    Dim KIEUQT As Variant
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    sSQLCol = "MATKHAI,KYLBO,NGNOP,KYKKHAI,MADTNT,MADLT,TTHTK,DAGHI,LOAITIEN,MAVACH,HANNOP,HANNOP2,THUETKY,THUETKY2,"
    sSQLCol = sSQLCol & "MAMUC,MATM,NGNHAP,AUTO_CAL,MATHUE,MAPP,SHTEP,GTTINHTHUE,LANBS,THUECL,GTTTCL,LAN_QUET,DANHAN,GHICHU"
      
    bln = False

    With fps
        .Sheet = 1
        matkhai = "'08A/KK-TNCN'"
        'neu quyet toan tung nam KIEUQT = 2, dau nam KQT= 1
        KIEUQT = ""
        'get KYLBO
        .GetText .ColLetterToNumber("E"), 30, kylbo

        If kylbo = vbNullString Then
            kylbo = "''"
        Else
            kylbo = "'" & kylbo & "'"
        End If
        
        ' get ngaynop
        .GetText .ColLetterToNumber("E"), 32, ngnop

        If ngnop = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
   
        ' set KYKKHAI
        KyKKhai = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        
        'get MADTNT
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If madtnt = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        'get DLT
        .GetText .ColLetterToNumber("F"), 20, MADLT

        If MADLT = vbNullString Then
            MADLT = "''"
        Else
            MADLT = "'" & MADLT & "'"
        End If
        
        'set TTHTK
        If tthtk = vbNullString Then
            tthtk = "''"
        End If
        
        'set DAGHI
        DaGhi = ".T."
        LoaiTien = "''"
        MaVach = ".T."
        
        'set HANNOP
        If Trim(HANNOP) = vbNullString Then
            HANNOP = "CTOD('')"
        Else
            HANNOP = ToDate(Trim(HANNOP), DDMMYYYY)
            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
        End If
        
        HANNOP2 = HANNOP
        ThueTKy = 0
        'set THUETKY
        .GetText .ColLetterToNumber("AB"), 49, ThueTKy

        If ThueTKy = vbNullString Then
            ThueTKy = "0"
        End If

        ThueTKy2 = "0"
        'get tu thang , den thang
        .GetText .ColLetterToNumber("K"), 3, vTuThang
        .GetText .ColLetterToNumber("P"), 3, vDenThang
        If Trim(vTuThang) <> vbNullString And Trim(vDenThang) <> vbNullString Then
            MaMuc = "'1000'"
            maTM = "'1004'"
        Else
            MaMuc = "'1000'"
            maTM = "'1003'"
        End If

        
        'set NGNHAP
        .GetText .ColLetterToNumber("M"), 32, NgNhap

        If Trim(NgNhap) = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = ToDate(Trim(NgNhap), DDMMYYYY)
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If
        
        AUTO_CAL = ".T."
        MATHUE = "''"
        MAPP = "'08'"
        SHTEP = "''"
        
        'set GTTINHTHUE
        .GetText .ColLetterToNumber("T"), 46, GTTinhThue

        If Trim(GTTinhThue) = vbNullString Then
            GTTinhThue = "0"
        End If

        'set LANBS
        .GetText .ColLetterToNumber("T"), 7, LanBS

        If Trim(LanBS) = vbNullString Then
            LanBS = "0"
        End If
        
        THUECL = "0"
        GTTTCL = "0"
        
        If Trim(LAN_QUET) = vbNullString Then
            LAN_QUET = "0"
        End If
        
        DANHAN = "''"
        'get ghichu
        .GetText .ColLetterToNumber("M"), 36, GHICHU
        If GHICHU = vbNullString Then
            GHICHU = "''"
        Else
            GHICHU = "'" & TAX_Utilities_Svr_New.Convert(CStr(GHICHU), UNICODE, TCVN) & "'"
        End If
        
        'get phong xu ly
        .GetText .ColLetterToNumber(PHONG_XU_LY_X1), PHONG_XU_LY_Y1, PXULY
        PXULY = Mid(PXULY, InStr(1, PXULY, "{") + 1, InStr(1, PXULY, "}") - InStr(1, PXULY, "{") - 1)
        
        PXULY = "'" & PXULY & "'"
        
        'get value Tu Ngay, den Ngay
        .GetText .ColLetterToNumber("K"), 3, kKEKHAITUNGAY
        .GetText .ColLetterToNumber("P"), 3, kKEKHAIDENNGAY
        If Trim(kKEKHAIDENNGAY) = vbNullString And Trim(kKEKHAITUNGAY) = vbNullString Then
            kKEKHAITUNGAY = vbNullString
            kKEKHAIDENNGAY = vbNullString
        End If
        
        If TAX_Utilities_Svr_New.isCheckPIT Then
            'lay kieu quyet toan
            .GetText .ColLetterToNumber("R"), 4, KIEUQT
            If Trim(KIEUQT) = vbNullString Then
                .GetText .ColLetterToNumber("R"), 5, KIEUQT
                KIEUQT = "01"
            Else
                KIEUQT = "02"
            End If
            
            ' Ghi du lieu vao bang trung gian de day len TC
            arr = Split(GetMaCqt, "~")
        '     hdrId = GetHdrId
            hdr.Tin = madtnt
            hdr.co_loi_dda = "''"
            hdr.DIA_CHI = arr(2)
            hdr.ghi_chu_loi = "''"
            hdr.khoa_so = "''"
            hdr.loai_tkhai = "'08A_TNCN11'"
            If kKEKHAITUNGAY = vbNullString And kKEKHAIDENNGAY = vbNullString Then
                kKEKHAITUNGAY = ""
                kKEKHAIDENNGAY = ""
                hdr.kieu_ky = "Q"
            Else
                'set lai ky ke khai
                KyKKhai = "'" & kKEKHAITUNGAY & "'"
                kKEKHAIDENNGAY = "01/" & kKEKHAIDENNGAY
                kKEKHAITUNGAY = "01/" & kKEKHAITUNGAY
                kKEKHAIDENNGAY = ToDate(Trim(kKEKHAIDENNGAY), DDMMYYYY)
                kKEKHAIDENNGAY = "CTOD('" & kKEKHAIDENNGAY & "')"
                kKEKHAITUNGAY = ToDate(Trim(kKEKHAITUNGAY), DDMMYYYY)
                kKEKHAITUNGAY = "CTOD('" & kKEKHAITUNGAY & "')"
                hdr.kieu_ky = "M"
            End If
        
    
            If tthtk = "1" Then
                hdr.kkbs = 0    ' To khai chinh thuc
            ElseIf tthtk = "2" Then
                hdr.kkbs = 1    ' To khai thay the
            ElseIf tthtk = "3" Then
                hdr.kkbs = 2    ' To khai bo sung
            Else
                hdr.kkbs = 1    ' To khai nop cham
            End If
            
            hdr.KyKKhai = KyKKhai
            hdr.kylbo = kylbo
            hdr.MA_CQT = arr(0)
            hdr.ngay_cap_nhat = NgNhap
            hdr.Ngay_nop = ngnop
            hdr.nguoi_cn = "'" & Trim(strNguoiSuDung) & "'"
            hdr.Phong_xly = PXULY
            hdr.so_hieu_tep = "''"
            hdr.So_tt_tk = LAN_QUET
            hdr.ten_dtnt = arr(1)
            hdr.tthtk = tthtk
            hdr.DA_NHAN = 0
            hdr.Id = hdrId
            hdr.thueondinh = "''"
            
            'set ngay hop dong dai ly
            NGAYHDDLT = strNgayHD_DLT
            If Trim(NGAYHDDLT) = vbNullString Then
                NGAYHDDLT = "CTOD('')"
            Else
                NGAYHDDLT = ToDate(Trim(NGAYHDDLT), DDMMYYYY)
                NGAYHDDLT = "CTOD('" & Format(NGAYHDDLT, "mm/dd/yyyy") & "')"
            End If
            
            'set So Hop dong Dl
            
            hdr.Ma_dl_thue = MADLT
            hdr.So_hd_dl = "'" & Trim(CStr(SOHDDL)) & "'"
            hdr.Ngay_hd_dl = NGAYHDDLT
            hdr.lan_bs = LanBS

            sSQL = InsertHDR_TGTC_08(hdr, kKEKHAITUNGAY, kKEKHAIDENNGAY, KIEUQT)
            bln = clsDAO.ExecuteDLL(sSQL)
        Else
            sSQLVal = matkhai & "," & kylbo & "," & ngnop & "," & KyKKhai & "," & madtnt & "," & MADLT & "," & tthtk & "," & DaGhi
            sSQLVal = sSQLVal & "," & LoaiTien & "," & MaVach & "," & HANNOP & "," & HANNOP2 & "," & ThueTKy & "," & ThueTKy2
            sSQLVal = sSQLVal & "," & MaMuc & "," & maTM & "," & NgNhap & "," & AUTO_CAL & "," & MATHUE & "," & MAPP & "," & SHTEP
            sSQLVal = sSQLVal & "," & GTTinhThue & "," & LanBS & "," & THUECL & "," & GTTTCL & "," & LAN_QUET & "," & DANHAN & "," & GHICHU
            
            sSQL = "INSERT INTO TMP_CNTK" & Right$(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            'InsertHDR = sSQL
            bln = clsDAO.ExecuteDLL(sSQL)
        End If
                
    End With

    clsDAO.Disconnect
    
End Function

Public Function TKTT() As Boolean
    Dim sSQL           As String
    Dim rs             As ADODB.Recordset
    Dim fso            As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim strFileNameHDR1 As String
    Dim strFileNameDTL1 As String
    Dim vTuThang As Variant, vDenThang As Variant, KyKKhai As Variant
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_CNTK" & Right$(sKyKeKhai, 4) & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_C08" & Right$(sKyKeKhai, 4) & ".DBF"
    strFileNameHDR1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tmp_tncn_hdr.DBF"
    'bang tmp_tncn_dtl_plus dung de ghi to 08A_TNCN, co truong gia tri chuyen thanh truong sau
    strFileNameDTL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tmp_tncn_dtl_plus.DBF"

    '    Dim strFileNameBS As String
    '    strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"
    'file HDR
    If fso.FileExists(strFileNameHDR) = False Then
        Dim strTepmauHDR As String
        strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_CNTK.DBF"
        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    End If

    'file dtl
    If fso.FileExists(strFileNameDTL) = False Then
        Dim strTepmauDTL As String
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_C08.DBF"
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
        strTepmauDTL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tepmau\tmp_tncn_dtl_plus.DBF"
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

    'set kykekhai
    KyKKhai = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"

    'get Tu thang, Den thang
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("K"), 3, vTuThang
        .GetText .ColLetterToNumber("P"), 3, vDenThang
    End With

    
    If Trim(vTuThang) <> vbNullString And Trim(vDenThang) <> vbNullString Then
        KyKKhai = "'" & Trim(vTuThang) & "'"
    End If


    If TAX_Utilities_Svr_New.isCheckPIT Then
        ' Truong hop PIT
        
        If Trim(vTuThang) <> vbNullString And Trim(vDenThang) <> vbNullString Then
             vDenThang = "01/" & Trim(vDenThang)
             vDenThang = ToDate(Trim(vDenThang), DDMMYYYY)
             vDenThang = "CTOD('" & Format(vDenThang, "mm/dd/yyyy") & "')"
             sSQL = "SELECT * FROM tmp_tncn_hdr WHERE Tin = '" & strMST & "' AND KYKKHAI = " & KyKKhai & " AND Lan_bs = 0 AND Loai_tkhai='08A_TNCN11' AND den_ngay = " & vDenThang
        Else
             sSQL = "SELECT * FROM tmp_tncn_hdr WHERE Tin = '" & strMST & "' AND KYKKHAI = " & KyKKhai & " AND Lan_bs = 0 AND Loai_tkhai='08A_TNCN11'"
        End If
        
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
    Else
        sSQL = "SELECT MADTNT, NGNOP FROM TMP_CNTK" & Right$(sKyKeKhai, 4) & " WHERE MADTNT = " & "'" & strMST & "'" & " AND MATKHAI='08A/KK-TNCN'"
    
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
    clsDAO.Disconnect
End Function

Public Function UpdateTHUETKY2() As Boolean
'    Dim sSQL As String
'
'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'   End If
'
'   sSQL = "UPDATE TMP_TK" & sKyKeKhai & _
'        " SET THUETKY2=1" & _
'        " WHERE MATHUE = '01' AND MAPP = '1' AND MADTNT = " & "'" & strMST & "'" & _
'        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'
'   clsDAO.ExecuteDLL sSQL
'
'   UpdateTHUETKY2 = True
'
'   clsDAO.Disconnect

End Function

Public Function XoaTKTT() As Boolean

End Function

Public Function TKRB() As Boolean
    TKRB = True
End Function

Public Function TKTTNGNOP() As Boolean

End Function

Public Function TinhSoLanQuet() As Integer
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim LAN_QUET1 As Variant
    Dim vTungay As Variant, vDenngay As Variant
    
    'kiem tra ton tai tep *.dbf chua
    'Tinh so lan quyet de ghi vao DB trung gian
    'dhdang sua
    'ngay 24/09/2010
    '-----------------------------
    
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("K"), 3, vTungay
        .GetText .ColLetterToNumber("P"), 3, vDenngay
        If Trim(vTungay) = vbNullString And Trim(vDenngay) = vbNullString Then
            vTungay = vbNullString
            vDenngay = vbNullString
        End If
    End With
    
    If TAX_Utilities_Svr_New.isCheckPIT Then
        Dim kykk_tu_ngay As Variant
        Dim kykk_den_ngay As Variant
        Dim tKyKekhai As Variant
        Dim strDate As Variant
        Dim dDate As Variant
        
        strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TNCN_HDR" & ".DBF"
        tKyKekhai = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        tKyKekhai = Replace(tKyKekhai, "'", "")
        If vTungay = vbNullString And vDenngay = vbNullString Then
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
        Else
           vDenngay = "01/" & vDenngay
           vDenngay = ToDate(Trim(vDenngay), DDMMYYYY)
           kykk_den_ngay = "CTOD('" & vDenngay & "')"
           vTungay = "01/" & vTungay
           vTungay = ToDate(Trim(vTungay), DDMMYYYY)
           kykk_tu_ngay = "CTOD('" & vTungay & "')"
        End If
  
        If fso.FileExists(strFileNameHDR) = True Then
            If clsDAO.Connected = False Then
                clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
                clsDAO.Connect
            End If
             sSQL = "SELECT max(so_tt_tk) as LAN_QUET FROM TMP_TNCN_HDR" & _
                    " WHERE TIN = " & "'" & strMST & "'" & _
                    " And loai_tkhai = '08A_TNCN11' " & _
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
 
    Else
        strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_CNTK" & Right$(sKyKeKhai, 4) & ".DBF"
        If fso.FileExists(strFileNameHDR) = True Then
            If clsDAO.Connected = False Then
            clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
            clsDAO.Connect
            End If
             sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_CNTK" & Right$(sKyKeKhai, 4) & _
                    " WHERE MADTNT = " & "'" & strMST & "'" & _
                    " AND MATKHAI='08A/KK-TNCN'" & " AND(TTHTK = '1' OR TTHTK = '3' OR TTHTK = '4')"
            
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
    Dim vTuThang As Variant, vDenThang As Variant, KyKKhai As Variant

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
       
    'set kykekhai
    KyKKhai = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"

    With fps
         'get Tu thang, Den thang
         .Sheet = 1
        .GetText .ColLetterToNumber("K"), 3, vTuThang
        .GetText .ColLetterToNumber("P"), 3, vDenThang
    End With

    
    If Trim(vTuThang) <> vbNullString And Trim(vDenThang) <> vbNullString Then
        KyKKhai = "'" & Trim(vTuThang) & "'"
    End If
       
   If TAX_Utilities_Svr_New.isCheckPIT Then
   
        If Trim(vTuThang) <> vbNullString And Trim(vDenThang) <> vbNullString Then
             vDenThang = "01/" & Trim(vDenThang)
             vDenThang = ToDate(Trim(vDenThang), DDMMYYYY)
             vDenThang = "CTOD('" & Format(vDenThang, "mm/dd/yyyy") & "')"
             sSQL = "SELECT * FROM tmp_tncn_hdr WHERE Tin = '" & strMST & "' AND KYKKHAI = " & KyKKhai & " AND Lan_bs = 0 AND Loai_tkhai='08A_TNCN11' AND den_ngay = " & vDenThang
        Else
             sSQL = "SELECT * FROM tmp_tncn_hdr WHERE Tin = '" & strMST & "' AND KYKKHAI = " & KyKKhai & " AND Lan_bs = 0 AND Loai_tkhai='08A_TNCN11'"
        End If
        
        Set rs = clsDAO.Execute(sSQL)
        If rs Is Nothing Then
            isToKhaiChinhThuc = False
        Else
            isToKhaiChinhThuc = True
        End If
 
   Else
        sSQL = "SELECT TMP_CNTK" & Right$(sKyKeKhai, 4) & _
        " WHERE ( TTHTK = '1' OR TTHTK = '3' OR TTHTK = '4' ) AND MADTNT = " & "'" & strMST & "'" & _
        " AND MATKHAI='08A/KK-TNCN'"
   
        Set rs = clsDAO.Execute(sSQL)
        If rs Is Nothing Then
            isToKhaiChinhThuc = False
        Else
            isToKhaiChinhThuc = True
        End If
    
   End If
          
   clsDAO.Disconnect

End Function



Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    With fps
        If .ActiveSheet = 1 Then
            If (Col = .ColLetterToNumber("E") Or Col = .ColLetterToNumber("F")) And Row = 32 Then
                .GetText .ColLetterToNumber("E"), Row, varTemp

                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                        
                        'check ngay nop khong lon hon ngay hien tai
'                          varTemp = ToDate(CStr(varTemp), DDMMYYYY)
'                          If DateDiff("D", Format(varTemp, "mm/dd/yyyy"), Format(Date, "mm/dd/yyyy")) < 0 Then
'                             DisplayMessage "0159", msOKOnly, miCriticalError
'                            .SetActiveCell Col, Row
'                          End If
                    Else
                        .SetActiveCell Col, Row
                    End If

                Else
                    .SetText Col, Row, ""
                End If

                UpdateCell fps, .Col, .Row, .Text
           
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

