VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_05GTGT"
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

Const NGAY_NOP_Y = 24
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 24
Const NGAY_QUET_X = "M"
Const TEN_BPQL_X = "V"
Const TEN_BPQL_Y = 26
'Const MA_BPQL_X = "V"
'Const MA_BPQL_Y = 26
Private rs As ADODB.Recordset
' thong tin din vi
Const MA_SO_THUE_X = "G"
Const MA_SO_THUE_Y = 8
Const TEN_GOI_X = "G"
Const TEN_GOI_Y = 7
Const DIA_CHI_X = "G"
Const DIA_CHI_Y = 9
Const QUAN_HUYEN_X = "F"
Const QUAN_HUYEN_Y = 11
Const TINH_TPHO_X = "T"
Const TINH_TPHO_Y = 11
'Const TEN_BPQL_X = "V"
'Const TEN_BPQL_Y = 12
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
Const TEN_GOI_DL_X = "G"
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
'Public strMaSoTep As String
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
Public TTHTK As String
Public strTenBpql As String
'Private tempKyLb As Variant
'Private tempNgayNop As Variant
Public hdrId As String
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
        .Col = .ColLetterToNumber("D")
        .Row = 28
        .CellType = CellTypeCheckBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100
        
        'Phong xu ly
        .Col = .ColLetterToNumber("V")
        .Row = 22
        .CellType = CellTypeComboBox
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
        '.Text = Day(Date) & "/" & Month(Date) & "/" & Year(Date)
        .TypeHAlign = TypeHAlignLeft
        
        'Phong quan ly
        .Col = .ColLetterToNumber("V")
        .Row = 26
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
        
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
Dim So_tt As Variant
    With fps
        .Sheet = 1
        
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
        
        .Col = .ColLetterToNumber(MA_BPQL_X)
        .Row = MA_BPQL_Y
        If strMaBPQL <> "" Then
            .Text = strMaBPQL
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        'st phong quan ly
        .Col = .ColLetterToNumber(TEN_BPQL_X)
        .Row = TEN_BPQL_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        UpdateCell fps, .Col, .Row, .Text
        
        '.EventEnabled(EventAllEvents) = True
        
        
        ' Set lai ky lap bo
        ' Lay ky lap bo duoc cua to khai quet truoc do, duoc ghi trong cookie de hien thi cho lan quet to khai tncn lan sau
'        tempKyLb = GetSetting("HTKK", "TNCN", "tempKyLb")
'        If (tempKyLb <> vbNullString) Then
'            .Col = .ColLetterToNumber(KY_LAP_BO_X)
'            .Row = KY_LAP_BO_Y
'            .Text = tempKyLb
'            UpdateCell fps, .Col, .Row, .Text
'        End If
        
        ' Set lai ngay nop
        ' Lay ngay nop duoc cua to khai quet truoc do, duoc ghi trong cookie de hien thi cho lan quet to khai tncn lan sau
'        tempNgayNop = GetSetting("HTKK", "TNCN", "tempNgayNop")
'        If (tempNgayNop <> vbNullString) Then
'            .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
'            .Row = NGAY_NHAN_TO_KHAI_Y
'            .Text = tempNgayNop
'            .TypeHAlign = TypeHAlignLeft
'        End If
          'set check loi dinh danh
        If checkDLTTonTai = False Then
            .Col = .ColLetterToNumber("D")
            .Row = 28
            .Value = 1
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
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim MATKHAI As Variant
   Dim MADTNT As Variant
   Dim MATHUE As Variant
   Dim MAPP As Variant
   Dim MAMUC As Variant
   Dim MATM As Variant
   Dim CTIEUTNDN As Variant
   Dim KYLBO As Variant
   Dim KYKKHAI As Variant
   Dim NGNOP As Variant
   Dim SOTIEN As Variant
   Dim TT As Variant
   Dim SOTIEN2 As Variant
   Dim LOAITIEN As Variant
   Dim NGANH As Variant
   Dim THUECL As Variant
   Dim DANHAN As Variant
   Dim bln As Variant
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
    sSQLCol = "MATKHAI, MADTNT, MATHUE, MAPP, MAMUC, MATM, CTIEUTNDN, TTHTK, KYLBO, KYKKHAI, NGNOP, SOTIEN,"
    sSQLCol = sSQLCol + "TT, SOTIEN2, LOAITIEN, NGANH, THUECL, DANHAN, LAN_QUET"
    
    With fps
   
        .Sheet = 1
        'ma to khai
        MATKHAI = "'05/GTGT'"
        'ma doi tuong nop thue
        .GetText .ColLetterToNumber("G"), 8, MADTNT
        If Trim(MADTNT) = vbNullString Then
            MADTNT = "''"
        Else
            MADTNT = "'" & MADTNT & "'"
        End If
        'ma thue
        MATHUE = "'01'"
        MAPP = "'6'"
        'ky lap bo
        .GetText .ColLetterToNumber("E"), 22, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
        ' ku ke khai
        KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        'ngay nop
        .GetText .ColLetterToNumber("E"), 24, NGNOP
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
        End If
        'trang thai
        TT = "''"
        'ma muc
        If Mid(Trim(KYLBO), 5, 4) < 2009 Then
            MAMUC = "'014'"
        Else
            MAMUC = "'1700'"
        End If
        'ma tieu muc
        If Mid(Trim(KYLBO), 5, 4) < 2009 Then
            MATM = "'01'"
        Else
            MATM = "'1701'"
        End If
        'tinh ma nganh
        NGANH = TinhMaNganh
        If Trim(CStr(NGANH)) = vbNullString Then
            NGANH = "''"
        Else
            NGANH = "'" & NGANH & "'"
        End If
        THUECL = "0"
        DANHAN = "''"
        LOAITIEN = "''"
        
        'chi tieu [21]
        CTIEUTNDN = "'030'"
        .GetText .ColLetterToNumber("K"), 43, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = SOTIEN
        
        sSQLVal = MATKHAI & "," & MADTNT & "," & MATHUE & "," & MAPP & "," & MAMUC & "," & MATM & "," & _
        CTIEUTNDN & "," & TTHTK & "," & KYLBO & "," & KYKKHAI & "," & NGNOP & "," & SOTIEN & "," & _
        TT & "," & SOTIEN2 & "," & LOAITIEN & "," & NGANH & "," & THUECL & "," & DANHAN & "," & LAN_QUET
        sSQL = "INSERT INTO TMP_DC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        'chi tieu [22]
        CTIEUTNDN = "'031'"
        .GetText .ColLetterToNumber("S"), 43, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = SOTIEN
        
        sSQLVal = MATKHAI & "," & MADTNT & "," & MATHUE & "," & MAPP & "," & MAMUC & "," & MATM & "," & _
        CTIEUTNDN & "," & TTHTK & "," & KYLBO & "," & KYKKHAI & "," & NGNOP & "," & SOTIEN & "," & _
        TT & "," & SOTIEN2 & "," & LOAITIEN & "," & NGANH & "," & THUECL & "," & DANHAN & "," & LAN_QUET
        sSQL = "INSERT INTO TMP_DC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        'chi tieu [23]
        CTIEUTNDN = "'032'"
        .GetText .ColLetterToNumber("K"), 44, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = SOTIEN
        
        sSQLVal = MATKHAI & "," & MADTNT & "," & MATHUE & "," & MAPP & "," & MAMUC & "," & MATM & "," & _
        CTIEUTNDN & "," & TTHTK & "," & KYLBO & "," & KYKKHAI & "," & NGNOP & "," & SOTIEN & "," & _
        TT & "," & SOTIEN2 & "," & LOAITIEN & "," & NGANH & "," & THUECL & "," & DANHAN & "," & LAN_QUET
        sSQL = "INSERT INTO TMP_DC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        'chi tieu [24]
        CTIEUTNDN = "'033'"
        .GetText .ColLetterToNumber("S"), 44, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = SOTIEN
        
        sSQLVal = MATKHAI & "," & MADTNT & "," & MATHUE & "," & MAPP & "," & MAMUC & "," & MATM & "," & _
        CTIEUTNDN & "," & TTHTK & "," & KYLBO & "," & KYKKHAI & "," & NGNOP & "," & SOTIEN & "," & _
        TT & "," & SOTIEN2 & "," & LOAITIEN & "," & NGANH & "," & THUECL & "," & DANHAN & "," & LAN_QUET
        sSQL = "INSERT INTO TMP_DC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        'chi tieu [25]
        CTIEUTNDN = "'034'"
        .GetText .ColLetterToNumber("K"), 45, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = SOTIEN
        
        sSQLVal = MATKHAI & "," & MADTNT & "," & MATHUE & "," & MAPP & "," & MAMUC & "," & MATM & "," & _
        CTIEUTNDN & "," & TTHTK & "," & KYLBO & "," & KYKKHAI & "," & NGNOP & "," & SOTIEN & "," & _
        TT & "," & SOTIEN2 & "," & LOAITIEN & "," & NGANH & "," & THUECL & "," & DANHAN & "," & LAN_QUET
        sSQL = "INSERT INTO TMP_DC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        'chi tieu [26]
        CTIEUTNDN = "'035'"
        .GetText .ColLetterToNumber("S"), 45, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = SOTIEN
        
        sSQLVal = MATKHAI & "," & MADTNT & "," & MATHUE & "," & MAPP & "," & MAMUC & "," & MATM & "," & _
        CTIEUTNDN & "," & TTHTK & "," & KYLBO & "," & KYKKHAI & "," & NGNOP & "," & SOTIEN & "," & _
        TT & "," & SOTIEN2 & "," & LOAITIEN & "," & NGANH & "," & THUECL & "," & DANHAN & "," & LAN_QUET
        sSQL = "INSERT INTO TMP_DC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'chi tieu [26]
        CTIEUTNDN = "'036'"
        .GetText .ColLetterToNumber("J"), 48, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = SOTIEN
        
        sSQLVal = MATKHAI & "," & MADTNT & "," & MATHUE & "," & MAPP & "," & MAMUC & "," & MATM & "," & _
        CTIEUTNDN & "," & TTHTK & "," & KYLBO & "," & KYKKHAI & "," & NGNOP & "," & SOTIEN & "," & _
        TT & "," & SOTIEN2 & "," & LOAITIEN & "," & NGANH & "," & THUECL & "," & DANHAN & "," & LAN_QUET
        sSQL = "INSERT INTO TMP_DC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
    End With
   
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
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim bln  As Boolean
    Dim MADTNT As Variant
    Dim NGNOP As Variant
    Dim KYLBO As Variant
    Dim KYKKHAI As Variant
    Dim MATKHAI As Variant
    Dim MATHUE As Variant
    Dim MAMUC As Variant
    Dim MATM As Variant
    Dim MAPP As Variant
    Dim mact As Variant
    Dim MACT2 As Variant
    Dim MACT3 As Variant
    Dim STT As Variant
    Dim STT2 As Variant
    Dim STTIN As Variant
    Dim SOKK As Variant
    Dim SODC As Variant
    Dim SOCL As Variant
    Dim SONGAYNC As Variant
    Dim SOTIENNC As Variant
    Dim i As Integer, iCol As Long, iRow As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim strFileName As String
    Dim fso As New FileSystemObject
    Dim DANHAN As Variant
    Dim arrMTM() As String, arrMTMSave() As String
    Dim arrVitri() As Integer
    Dim idx, idx1, countMTM As Integer
    Dim vTemp As String, vVitriTrung As String
    Dim ROWNUM As Variant
   'kiem tra ton tai tep *.dbf chua
    strFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileName) = False Then
    Dim strTepmau As String
        strTepmau = spathVat & "\tepmau\TMP_BS.DBF"
        fso.CopyFile strTepmau, strFileName, False
    End If

   sSQLCol = "MADTNT, KYKKHAI, MATKHAI, MATHUE, MAMUC, MATM, MAPP, NGNOP, TTHTK, KYLBO , MACT, MACT2, MACT3, STT, STT2, STTIN," & _
             " SOKK, SODC, SOCL, SONGAYNC, SOTIENNC, LAN_QUET, DANHAN"

   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If


   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("G"), 8, MADTNT
        If Trim(MADTNT) = vbNullString Then
            MADTNT = "''"
        Else
            MADTNT = "'" & MADTNT & "'"
        End If

        KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        MATKHAI = "'05/GTGT'"
        MATHUE = "'01'"
        'ma muc
        If Mid(Trim(KYLBO), 5, 4) < 2009 Then
            MAMUC = "'014'"
        Else
            MAMUC = "'1700'"
        End If
        
        MAPP = "'6'"
        'ma tieu muc
        If Mid(Trim(KYLBO), 5, 4) < 2009 Then
            MATM = "'01'"
        Else
            MATM = "'1701'"
        End If

        .GetText .ColLetterToNumber("E"), 24, NGNOP
        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 22, KYLBO
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
        SONGAYNC = "0"
        SOTIENNC = "0"
        DANHAN = "''"
            sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
                      NGNOP & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
                      SOKK & "," & SODC & "," & SOCL & "," & SONGAYNC & "," & SOTIENNC & "," & LAN_QUET & "," & DANHAN


            sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)

        .Sheet = .SheetCount - 1
        ' so ngay nop cham
        .GetText .ColLetterToNumber("E"), 20, SONGAYNC
        If Trim(SONGAYNC) = vbNullString Then
            SONGAYNC = "0"
        Else
            SONGAYNC = SONGAYNC
        End If
        ' so tien nop cham
        .GetText .ColLetterToNumber("G"), 20, SOTIENNC
        If Trim(SOTIENNC) = vbNullString Then
            SOTIENNC = "0"
        Else
            SOTIENNC = SOTIENNC
        End If
        i = 1
        .Row = 24
        Do
'                .GetText .ColLetterToNumber("D"), .Row, MACT
'                If Trim(MACT) = vbNullString Then
                mact = "''"
'                Else
'                    MACT = "'" & MACT & "'"
'                End If
                .GetText .ColLetterToNumber("E"), .Row, MACT2
                If Trim(MACT2) = vbNullString Then
                    MACT2 = "''"
                Else
                    MACT2 = "'" & MACT2 & "'"
                End If

                MACT3 = "'2'"
                STT = 1
                STT2 = i
                STTIN = "'" & CStr(i) & "'"

                .GetText .ColLetterToNumber("F"), .Row, SOKK
                If Trim(SOKK) = vbNullString Then
                    SOKK = "0"
                Else
                    SOKK = SOKK
                End If

                .GetText .ColLetterToNumber("G"), .Row, SODC
                If Trim(SODC) = vbNullString Then
                    SODC = "0"
                Else
                    SODC = SODC
                End If

                .GetText .ColLetterToNumber("H"), .Row, SOCL
                If Trim(SOCL) = vbNullString Then
                    SOCL = "0"
                Else
                    SOCL = SOCL
                End If

                sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
                          NGNOP & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
                          SOKK & "," & SODC & "," & SOCL & "," & SONGAYNC & "," & SOTIENNC & "," & LAN_QUET & "," & DANHAN

                sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                If MACT2 <> "''" Then
                    bln = clsDAO.ExecuteDLL(sSQL)
                End If

            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
        Loop Until .Text = "aa"
        ROWNUM = .Row
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

            sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
                      NGNOP & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
                      SOKK & "," & SODC & "," & SOCL & "," & SONGAYNC & "," & SOTIENNC & "," & LAN_QUET & "," & DANHAN


            sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
               
        .Row = ROWNUM + 3
         i = 1
         Do

                mact = "''"

                .GetText .ColLetterToNumber("E"), .Row, MACT2
                If Trim(MACT2) = vbNullString Then
                    MACT2 = "''"
                Else
                    MACT2 = "'" & MACT2 & "'"
                End If

                MACT3 = "'2'"
                STT = 1
                STT2 = i
                STTIN = "'" & CStr(i) & "'"

                .GetText .ColLetterToNumber("F"), .Row, SOKK
                If Trim(SOKK) = vbNullString Then
                    SOKK = "0"
                Else
                    SOKK = SOKK
                End If

                .GetText .ColLetterToNumber("G"), .Row, SODC
                If Trim(SODC) = vbNullString Then
                    SODC = "0"
                Else
                    SODC = SODC
                End If

                .GetText .ColLetterToNumber("H"), .Row, SOCL
                If Trim(SOCL) = vbNullString Then
                    SOCL = "0"
                Else
                    SOCL = SOCL
                End If

                sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
                          NGNOP & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
                          SOKK & "," & SODC & "," & SOCL & "," & SONGAYNC & "," & SOTIENNC & "," & LAN_QUET & "," & DANHAN

                sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

                If MACT2 <> "''" Then
                    bln = clsDAO.ExecuteDLL(sSQL)
                End If

            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
          
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
        .GetText .ColLetterToNumber("H"), 30, SOCL
        If Trim(SOCL) = vbNullString Then
            SOCL = "0"
        Else
            SOCL = SOCL
        End If
        
            sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
                      NGNOP & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
                      SOKK & "," & SODC & "," & SOCL & "," & SONGAYNC & "," & SOTIENNC & "," & LAN_QUET & "," & DANHAN

            sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
            
   End With

'   clsDAO.Disconnect
   InsertDTL_KHBS = vbNullString
   
End Function

Public Function InsertHDR() As String
    Dim TENNV      As String
    Dim CHUNGCHI   As String
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim MADTNT As Variant
    Dim MADLT As Variant
    Dim MATHUE As Variant
    Dim MATM As Variant
    Dim MAPP As Variant
    Dim KYLBO As Variant
    Dim NGNOP As Variant
    Dim LOAITIEN As Variant
    Dim MABPQL As Variant
    Dim SHTEP As Variant
    Dim KYKKHAI As Variant
    Dim NGNHAP As Variant
    Dim THUETKY As Variant
    Dim GTTINHTHUE As Variant
    Dim DAGHI As Variant
    Dim KTTRUOC As Variant
    Dim DoanhSo As Variant
    Dim MAPP2 As Variant
    Dim MAMUC As Variant
    Dim MAVACH As Variant
    Dim MATKHAI As Variant
    Dim THUETKY2 As Variant
    Dim HANNOP2 As Variant
    Dim TENDA As Variant
    Dim DAGHI1 As Variant
    Dim DAGHI2 As Variant
    Dim AUTO_CAL As Variant
    Dim NGANH As Variant
    Dim LANBS As Variant
    Dim THUECL As Variant
    Dim DSoCL As Variant
    Dim GTTTCL As Variant
    Dim NGAYPS As Variant
    Dim MANNKD As Variant
    Dim DANHAN As Variant
    Dim CHKGIAHAN As Variant
    Dim bln  As Boolean
    
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
    
    
    sSQLCol = "MADTNT, MADLT, MATHUE, MATM, MAPP, KYLBO, NGNOP, LOAITIEN, MABPQL, SHTEP, TTHTK,"
    sSQLCol = sSQLCol + "KYKKHAI, NGNHAP, THUETKY, GTTINHTHUE, DAGHI, KTTRUOC, DOANHSO, MAPP2, MAMUC, MAVACH, MATKHAI, HANNOP, THUETKY2, HANNOP2,"
    sSQLCol = sSQLCol + "TENDA, DAGHI1, DAGHI2, AUTO_CAL, NGANH, LANBS, THUECL, DSOCL, GTTTCL, LAN_QUET, NGAYPS, MANNKD, DANHAN,TENNV,CHUNGCHI"
    
    
       
    With fps
        .Sheet = 1
                'lay ten nhan vien dai ly thue, chung chi hanh nghe
        .GetText .ColLetterToNumber("H"), 57, TENNV
        .GetText .ColLetterToNumber("H"), 59, CHUNGCHI

        'ma doi tuong nop thue
        .GetText .ColLetterToNumber("G"), 8, MADTNT
        If Trim(MADTNT) = vbNullString Then
            MADTNT = "''"
        Else
            MADTNT = "'" & MADTNT & "'"
        End If
        'ma dai ly thue
        .GetText .ColLetterToNumber("G"), 15, MADLT
        If Trim(MADLT) = vbNullString Then
            MADLT = "''"
        Else
            MADLT = "'" & MADLT & "'"
        End If
        'ma thue
        MATHUE = "'01'"
        MAPP = "'6'"
        ' ky lap bo
        .GetText .ColLetterToNumber("E"), 22, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
        'ma tieu muc
        If Mid(Trim(KYLBO), 5, 4) < 2009 Then
            MATM = "'01'"
        Else
            MATM = "'1701'"
        End If
        'ngay nop
        .GetText .ColLetterToNumber("E"), 24, NGNOP
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
        End If
        ' loai tien
        LOAITIEN = "''"
        'ma phong quan ly
        MABPQL = strMaBPQL
        If Trim(MABPQL) = vbNullString Then
            MABPQL = "''"
        Else
            MABPQL = "'" & MABPQL & "'"
        End If
        'shtep
'        .GetText .ColLetterToNumber("M"), 22, SHTEP
'        If Trim(SHTEP) = vbNullString Then
'            SHTEP = "''"
'        Else
'            SHTEP = "'" & SHTEP & "'"
'        End If
        SHTEP = "''"
        'ky ke khai
        KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        'NGNHAP = Date
        .GetText .ColLetterToNumber("M"), 24, NGNHAP
        If Trim(NGNHAP) = vbNullString Then
            NGNHAP = "CTOD('')"
        Else
            NGNHAP = ToDate(Trim(NGNHAP), DDMMYYYY)
            NGNHAP = "CTOD('" & Format(NGNHAP, "mm/dd/yyyy") & "')"
        End If
        ' thue ky

        .GetText .ColLetterToNumber("J"), 48, THUETKY
        
        If Trim(THUETKY) = vbNullString Then
               THUETKY = "0"
        End If
        ' gia tri tinh thue
        'get gia tri tinh thue = CT 21 + CT 22
        .GetText .ColLetterToNumber("S"), 39, GTTINHTHUE
        If GTTINHTHUE = vbNullString Then
            GTTINHTHUE = "0"
        End If
        'da ghi
        DAGHI = ".F."
        'ky truoc
        KTTRUOC = "0"
        'doanh so
        .GetText .ColLetterToNumber("AC"), 43, DoanhSo
        If Trim(DoanhSo) = vbNullString Then
            DoanhSo = "0"
        End If
        'ma phuong phap
        MAPP2 = "'1'"
        'ma muc
        If Mid(Trim(KYLBO), 5, 4) < 2009 Then
            MAMUC = "'014'"
        Else
            MAMUC = "'1700'"
        End If
        'ma vach
        MAVACH = ".T."
        'ma to khai
        MATKHAI = "'05/GTGT'"
        'han nop
        If Trim(HANNOP) = vbNullString Then
            HANNOP = "CTOD('')"
        Else
            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
        End If
        'thue trong ky
        THUETKY2 = "0"
        HANNOP2 = HANNOP
        TENDA = "''"
        DAGHI1 = ".F."
        .GetText .ColLetterToNumber("D"), 28, CHKGIAHAN
        If Trim(CHKGIAHAN) = "1" Then
            DAGHI2 = ".T."
        Else
            DAGHI2 = ".F."
        End If
        
        AUTO_CAL = ".F."
        
        'tinh ma nganh
        NGANH = TinhMaNganh
        If Trim(CStr(NGANH)) = vbNullString Then
            NGANH = "''"
        Else
            NGANH = "'" & NGANH & "'"
        End If
        
        Dim solanBS As Variant
        ' so lan bo sung
         If TTHTK = "'" & 2 & "'" Then
         .GetText .ColLetterToNumber("T"), 5, solanBS
            LANBS = solanBS
         Else
            LANBS = "0"
         End If
        
        THUECL = "0"
        DSoCL = "0"
        GTTTCL = "0"
        NGAYPS = "CTOD('')"
        MANNKD = "''"
        DANHAN = "''"
                  
        sSQLVal = MADTNT & "," & MADLT & "," & MATHUE & "," & MATM & "," & MAPP & "," & KYLBO & "," & _
            NGNOP & ", " & LOAITIEN & "," & MABPQL & "," & SHTEP & "," & TTHTK & "," & _
            KYKKHAI & "," & NGNHAP & "," & THUETKY & "," & GTTINHTHUE & "," & DAGHI & "," & KTTRUOC & "," & _
            DoanhSo & "," & MAPP2 & "," & MAMUC & "," & MAVACH & ", " & MATKHAI & "," & HANNOP & "," & THUETKY2 & "," & _
            HANNOP2 & "," & TENDA & "," & DAGHI1 & "," & DAGHI2 & "," & AUTO_CAL & "," & NGANH & "," & LANBS & "," & THUECL & "," & _
            DSoCL & "," & GTTTCL & "," & LAN_QUET & "," & NGAYPS & "," & MANNKD & "," & DANHAN & "," & TENNV & "," & CHUNGCHI
       
        sSQL = "INSERT INTO TMP_TK" & sKyKeKhai & _
                "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
        bln = clsDAO.ExecuteDLL(sSQL)
        
    End With
     
    clsDAO.Disconnect
    InsertHDR = vbNullString
End Function

Public Function TKTT() As Boolean
    Dim sSQL            As String
    Dim rs              As ADODB.Recordset
    Dim fso             As New FileSystemObject
    Dim strFileNameHDR  As String
    Dim strFileNameDTL  As String
    Dim strFileNameBS   As String
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_DC" & sKyKeKhai & ".DBF"
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
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_DC.DBF"
        fso.CopyFile strTepmauDTL, strFileNameDTL, False
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
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyKeKhai & _
        " WHERE MATHUE = '01' AND  MAPP = '6' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='05/GTGT'"
    
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
   
   sSQL = "UPDATE TMP_TK" & sKyKeKhai & _
        " SET THUETKY2 = 0 " & _
        " WHERE MATHUE = '01' and MAPP = '6' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='05/GTGT'"

   clsDAO.ExecuteDLL sSQL

   UpdateTHUETKY2 = True

   clsDAO.Disconnect
End Function

Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL            As String
    Dim rs              As ADODB.Recordset
    Dim vKeKhai         As Variant
    Dim vKeKhai_temp    As Variant
    Dim vKyLB           As Variant
    Dim strTestFileName As String
    Dim fso             As New FileSystemObject
    Dim MATKHAI         As Variant
     
    Dim flag            As Boolean
    'kiem tra da co to khai chinh thuc chua
    flag = True
    vKyLB = DateSerial(Int(Year(Now())), Int(Month(Now())), 1)
    vKeKhai = DateSerial(Int(Val(TAX_Utilities_Svr_New.Year)), Int(Val(TAX_Utilities_Svr_New.Month)), 1)
    vKeKhai_temp = sKyKeKhai

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
        ' Truong hop khong phai PIT
        Do While flag
   
            strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & vKeKhai_temp & ".DBF"

            'kiem tra neu co db thi moi thuc hien truy van
            If fso.FileExists(strTestFileName) Then
                sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & vKeKhai_temp & " WHERE MATHUE = '01' AND MADTNT = " & "'" & strMST & "'" & " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'" & " AND (TTHTK = '1' or TTHTK = '3' or TTHTK = '4') AND MATKHAI='05/GTGT'"
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
   
   sSQL = "DELETE FROM TMP_TK" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='05/GTGT'"
         
   clsDAO.ExecuteDLL sSQL
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 22, KYLBO
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
   sSQL = "DELETE FROM TMP_DC" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "' AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' " & _
        " AND MATKHAI='05/GTGT'"
    
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
   Dim NGNOP As Variant
   Dim NGNOPDB As Variant
   Dim sSQL As String
   Dim THUEGTGT_KT As Double
   Dim rs As ADODB.Recordset
   Dim i As Integer
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 24, NGNOP
        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = ""
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
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
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TZ" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI <> '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
    
    
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
                If Trim(NGNOPDB) <> vbNullString And Trim(NGNOPDB) = Trim(NGNOP) Then
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
            .GetText Col, Row, varTemp
            If varTemp <> "" And varTemp <> "../../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    
                                            'check ngay nop khong lon hon ngay hien tai
'                    If DateDiff("D", Format(varTemp, "mm/dd/yyyy"), Format(Date, "mm/dd/yyyy")) < 0 Then
'                        DisplayMessage "0159", msOKOnly, miCriticalError
'                        .SetActiveCell Col, Row
'                    End If
                Else
                    .SetActiveCell Col, Row
                End If
            Else
             .SetText Col, Row, ""
            End If
           UpdateCell fps, .Col, .Row, .Text
        End If
        If Col = .ColLetterToNumber("M") And Row = 24 Then
            .GetText Col, Row, varTemp
            If varTemp <> "" And varTemp <> "../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
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
             sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_TK" & sKyKeKhai & _
                    " WHERE MADTNT = " & "'" & strMST & "'" & _
                    " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI = '05/GTGT'" & " AND (TTHTK = '1' or TTHTK = '2' or TTHTK = '3' or TTHTK = '4')"
    
                Set rs = clsDAO.Execute(sSQL)
                If Not rs Is Nothing Then
                     TinhSoLanQuet = rs.Fields("LAN_QUET")
                   
                      
                Else
                     TinhSoLanQuet = 0
                End If
                clsDAO.Disconnect
        Else
            TinhSoLanQuet = 0
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
'
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
'   sSQLCol = "INSERT INTO tmp_tncn_hdr (id,tin, ten_dtnt, dia_chi, loai_tkhai, ngay_nop, kylb_tu_ng, kylb_den_n, kykk_tu_ng, kykk_den_n, ngay_cap_n,"
'    sSQLCol = sSQLCol + " nguoi_cn, co_loi_dda, so_hieu_te, so_tt_tk, da_nhan, ghi_chu_lo, khoa_so, phong_xly, kkbs, tthtk, kylbo, kykkhai, ma_cqt) "
'    sSQLCol = sSQLCol + " values ("
'
'    sSQLVal = hdr.Id & "," & hdr.tin & "," & hdr.ten_dtnt & "," & hdr.DIA_CHI & "," & hdr.loai_tkhai & "," & hdr.ngay_nop & "," & kyLB_tu_ngay & "," & _
'    kyLB_den_ngay & "," & kykk_tu_ngay & "," & kykk_den_ngay & "," & hdr.ngay_cap_nhat & "," & hdr.nguoi_cn & "," & hdr.co_loi_dda & "," & _
'    hdr.so_hieu_tep & "," & hdr.so_tt_tk & "," & hdr.DA_NHAN & "," & hdr.ghi_chu_loi & "," & hdr.khoa_so & "," & hdr.phong_xly & "," & hdr.kkbs & "," & hdr.TTHTK & "," & hdr.KYLBO & "," & hdr.KYKKHAI & "," & hdr.ma_cqt
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
'    sSQLCol = "INSERT INTO tmp_tncn_dtl (id, hdr_id,matkhai, madtnt, kylbo, kykkhai, tthtk, ngnop, cttn, giatri, danhan, lan_quet, ky_hieu, ma_cqt) "
'    sSQLCol = sSQLCol + " values ("
'
'    sSQLVal = dtl.Id & "," & dtl.hdr_id & "," & dtl.MATKHAI & "," & dtl.MADTNT & "," & dtl.KYLBO & "," & dtl.KYKKHAI & "," & dtl.TTHTK & "," & dtl.NGNOP & "," & _
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
