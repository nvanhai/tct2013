VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_06TNCN10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const HEADER_SO_TT_TRONG_TEP_X = "V"
Const HEADER_SO_TT_TRONG_TEP_Y = 17 'Dong 16 da dc Hidden
Private clsConn As New TAX_Utilities_Svr_New.clsADO

Const KY_LAP_BO_Y = 30
Const KY_LAP_BO_X = "E"
Const NGAY_NHAN_TO_KHAI_Y = 32
Const NGAY_NHAN_TO_KHAI_X = "E"
Const NGUOI_SU_DUNG_Y = 32
Const NGUOI_SU_DUNG_X = "V"
Const PHONG_XU_LY_Y = 30
Const PHONG_XU_LY_X = "V"
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
    With fps
     .Sheet = 1
        .Col = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X)
        .Row = HEADER_SO_TT_TRONG_TEP_Y
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetNumeric
        .TypeMaxEditLen = 10

        'Ghi chu
        .Col = .ColLetterToNumber("M")
        .Row = 34
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100
        
        'Phong xu ly
        .Col = .ColLetterToNumber("V")
        .Row = 30
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
        
        ' Ngay que't
        SetDateFormat fps, 1, 32, .ColLetterToNumber("M"), DDMMYYYY
        .Row = 32
        .Col = .ColLetterToNumber("M")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
        'set ngay hop dong dai ly
        SetDateFormat fps, 1, 28, .ColLetterToNumber("N"), DDMMYYYY
        
        
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
        If rsPXL.Fields.Count > 0 Then
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
        
        'so thu tu to khai
        .Col = .ColLetterToNumber("M")
        .Row = 30
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        
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
        
        ' Dat lai check E_36 Duoc gian thue theo TT27/TT-BTC
        ' Thoi gian gian thue tu thang 1 den thang 6 nam 2009 cua to khai thang
        If Val(TAX_Utilities_Svr_New.Year) = 2009 And Val(TAX_Utilities_Svr_New.Month) <= 6 Then
            ' Neu trong hieu luc cua gian thue thi default la checked
            .Row = 36
            .Col = .ColLetterToNumber("E")
            .Text = "1"
            UpdateCell fps, .Col, .Row, "x"
        Else
            ' Neu ngoai hieu luc cua gian thue thi visible cac dieu kien check nay di
            .Row = 36
            .Col = .ColLetterToNumber("B")
            .CellType = CellTypeEdit
            .Lock = True
            .Text = ""
            
            .Row = 36
            .Col = .ColLetterToNumber("E")
            .CellType = CellTypeEdit
            .Lock = True
            .Text = ""
            UpdateCell fps, .Col, .Row, ""
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
    Dim ngayNop As Variant, ngayHienTai As Variant
    CheckValidData = True
    With fps
        .Sheet = 1
        
        ' check ngay nop khong dc lon hon ngay hien tai
        .GetText .ColLetterToNumber("E"), 32, ngayNop
        ngayHienTai = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        ngayHienTai = DateSerial(Int(Mid$(ngayHienTai, 7, 4)), Int(Mid$(ngayHienTai, 4, 2)), Int(Mid$(ngayHienTai, 1, 2)))
        ngayNop = DateSerial(Int(Mid$(ngayNop, 7, 4)), Int(Mid$(ngayNop, 4, 2)), Int(Mid$(ngayNop, 1, 2)))
        If ngayNop > ngayHienTai Then
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
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim MATKHAI As Variant
   Dim MADTNT As Variant
   Dim KYLBO As Variant
   Dim KYKKHAI As Variant
   Dim NAMQT As Variant
   Dim NGNOP As Variant
   Dim CTTN As Variant
   Dim GIATRI As Variant
   Dim GIATRI2 As Variant
   Dim MAMUC As Variant
   Dim MATM As Variant
   Dim TT As Variant
   Dim DANHAN As Variant
   Dim PNNQT As Variant
   Dim PNKK As Variant
   Dim CLECH As Variant
   Dim LANBS As Variant
   
   Dim bln  As Boolean
   
 
   If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
    
   sSQLCol = "NAMQT, MADTNT, MATKHAI, MAMUC, MATM, TTHTK, NGNOP, PNNQT, PNKK, CLECH, LAN_QUET, DANHAN"
  
   '-------------
   With fps
   
   .Sheet = 1
        NAMQT = TAX_Utilities_Svr_New.Year
        MATKHAI = "'06/KK-TNCN'"
        DANHAN = "''"
        .GetText .ColLetterToNumber("F"), 10, MADTNT
        If Trim(MADTNT) = vbNullString Then
            MADTNT = "''"
        Else
            MADTNT = "'" & MADTNT & "'"
        End If
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
      
        .GetText .ColLetterToNumber("E"), 32, NGNOP
        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
        End If
        KYKKHAI = "'" & TAX_Utilities_Svr_New.Year & "'"
        
        'get lanbs
        .GetText .ColLetterToNumber("S"), 6, LANBS
         If Trim(LANBS) = vbNullString Then
            LANBS = "0"
         End If
         If Val(CStr(LANBS)) > 0 Then
            TTHTK = "'" & 2 & "'"
         End If
        
        'MA MUC, TIEU MUC 1004
        MAMUC = "'1000'"
        MATM = "'1004'"
        Dim temp1 As Variant
        Dim TEMP2 As Variant
        .GetText .ColLetterToNumber("U"), 42, temp1
        .GetText .ColLetterToNumber("U"), 45, TEMP2
        PNNQT = temp1 + TEMP2
        PNKK = 0
        CLECH = PNNQT

        sSQLVal = NAMQT & "," & MADTNT & "," & MATKHAI & "," & MAMUC & "," & MATM & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & PNNQT & "," & PNKK & "," & CLECH & "," & LAN_QUET & "," & DANHAN
    
        sSQL = "INSERT INTO TMP_CNCT" & Right$(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        '---------------------
        'MA MUC, TIEU MUC 1008
        MAMUC = "'1000'"
        MATM = "'1008'"
        .GetText .ColLetterToNumber("U"), 48, temp1
        PNNQT = temp1
        PNKK = 0
        CLECH = PNNQT

        sSQLVal = NAMQT & "," & MADTNT & "," & MATKHAI & "," & MAMUC & "," & MATM & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & PNNQT & "," & PNKK & "," & CLECH & "," & LAN_QUET & "," & DANHAN
    
        sSQL = "INSERT INTO TMP_CNCT" & Right$(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
         '---------------------
        'MA MUC, TIEU MUC 1012
        MAMUC = "'1000'"
        MATM = "'1007'"
        .GetText .ColLetterToNumber("U"), 51, temp1
        PNNQT = temp1
        PNKK = 0
        CLECH = PNNQT

        sSQLVal = NAMQT & "," & MADTNT & "," & MATKHAI & "," & MAMUC & "," & MATM & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & PNNQT & "," & PNKK & "," & CLECH & "," & LAN_QUET & "," & DANHAN
    
        sSQL = "INSERT INTO TMP_CNCT" & Right$(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
         '---------------------
        'MA MUC, TIEU MUC 1003
        MAMUC = "'1000'"
        MATM = "'1003'"
        .GetText .ColLetterToNumber("U"), 55, temp1
        PNNQT = temp1
        PNKK = 0
        CLECH = PNNQT

        sSQLVal = NAMQT & "," & MADTNT & "," & MATKHAI & "," & MAMUC & "," & MATM & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & PNNQT & "," & PNKK & "," & CLECH & "," & LAN_QUET & "," & DANHAN
    
        sSQL = "INSERT INTO TMP_CNCT" & Right$(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
     End With
   
    clsDAO.Disconnect
   
   InsertDTL = vbNullString
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
    
    Dim MATKHAI    As Variant
    Dim KYLBO      As Variant
    Dim NAMQT      As Variant
    Dim MADTNT     As Variant
    Dim NGNOP      As Variant
    Dim TUKY       As Variant
    Dim DENKY      As Variant
    Dim NGNHAP     As Variant
    Dim MAMUC      As Variant
    Dim MATM       As Variant
    Dim PNNQT      As Variant
    Dim PNKK       As Variant
    Dim CLECH      As Variant
    Dim THUETKY2   As Variant
    Dim MAVACH     As Variant
    Dim AUTO_CAL   As Variant
    Dim DOANHSO    As Variant
    Dim GTTINHTHUE As Variant
    Dim THUETKY    As Variant
    Dim DSOCL      As Variant
    Dim GTTTCL     As Variant
    Dim THUECL     As Variant
    Dim MADLT      As Variant
    Dim LANBS      As Variant
    Dim DANHAN     As Variant
    Dim bln        As Boolean
    Dim varTemp    As Variant
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
    sSQLCol = " MATKHAI,KYLBO,NAMQT,MADTNT,TTHTK,NGNOP,TUKY,DENKY,NGNHAP,HANNOP,MAMUC,MATM,PNNQT,PNKK,CLECH,THUETKY2,MAVACH"
    sSQLCol = sSQLCol & ",AUTO_CAL,DOANHSO,GTTINHTHUE,THUETKY,DSOCL,GTTTCL,THUECL,DAGHI2,MADLT,LANBS,DANHAN,LAN_QUET"
       
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, MADTNT

        If Trim(MADTNT) = vbNullString Then
            MADTNT = "''"
        Else
            MADTNT = "'" & MADTNT & "'"
        End If
       
        NAMQT = TAX_Utilities_Svr_New.Year
        TUKY = "'01/" & TAX_Utilities_Svr_New.Year & "'"
        DENKY = "'" & Left$(sKyKeKhai, 2) & "/" & Right$(sKyKeKhai, 4) & "'"
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
        
        .GetText .ColLetterToNumber("E"), 32, NGNOP

        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
        End If
        
        .GetText .ColLetterToNumber("M"), 32, NGNHAP

        'NGNHAP = Date
        If Trim(NGNHAP) = vbNullString Then
            NGNHAP = "CTOD('')"
        Else
            NGNHAP = ToDate(Trim(NGNHAP), DDMMYYYY)
            NGNHAP = "CTOD('" & Format(NGNHAP, "mm/dd/yyyy") & "')"
        End If
        
        'set HANNOP
        If HANNOP = vbNullString Then
            HANNOP = "CTOD('')"
        Else
            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
        End If
        
        'get check loi dinh danh
        .GetText .ColLetterToNumber("E"), 34, varTemp
        If UCase(Trim(varTemp)) = "X" Or Trim(varTemp) = 1 Then
            DAGHI2 = ".T."
        Else
            DAGHI2 = ".F."
        End If
        
        MATKHAI = "'06/KK-TNCN'"
        DANHAN = "''"
        MAVACH = ".T."
        THUETKY2 = "'0'"
        
        'cac gia tri tam set
        AUTO_CAL = ".F."
        DOANHSO = 0
        GTTINHTHUE = 0
        THUETKY = 0
        DSOCL = 0
        GTTTCL = 0
        THUECL = 0
        
        'get ma dai ly thue
        .GetText .ColLetterToNumber("F"), 20, MADLT
        If MADLT = vbNullString Then
            MADLT = "''"
        Else
            MADLT = "'" & MADLT & "'"
        End If
        
        'get lanbs
        .GetText .ColLetterToNumber("S"), 6, LANBS
         If Trim(LANBS) = vbNullString Then
            LANBS = "0"
         End If
         If Val(CStr(LANBS)) > 0 Then
            TTHTK = "'" & 2 & "'"
         End If
         
         
        'Mã TM 1004
        MAMUC = "'1000'"
        MATM = "'1004'"
        Dim temp1 As Variant
        Dim TEMP2 As Variant
        .GetText .ColLetterToNumber("U"), 42, temp1
        .GetText .ColLetterToNumber("U"), 45, TEMP2
        PNNQT = temp1 + TEMP2
        THUETKY = PNNQT
        PNKK = 0
        CLECH = PNNQT
        'get CT[21] va CT[23]
        .GetText .ColLetterToNumber("U"), 41, temp1
        .GetText .ColLetterToNumber("U"), 44, TEMP2
        DOANHSO = temp1 + TEMP2
        GTTINHTHUE = DOANHSO
    
        sSQLVal = MATKHAI & "," & KYLBO & "," & NAMQT & "," & MADTNT & "," & TTHTK & "," & NGNOP & "," & TUKY & "," & DENKY
        sSQLVal = sSQLVal & "," & NGNHAP & "," & HANNOP & "," & MAMUC & "," & MATM & "," & PNNQT & "," & PNKK & "," & CLECH
        sSQLVal = sSQLVal & "," & THUETKY2 & "," & MAVACH & "," & AUTO_CAL & "," & DOANHSO & "," & GTTINHTHUE & "," & THUETKY
        sSQLVal = sSQLVal & "," & DSOCL & "," & GTTTCL & "," & THUECL & "," & DAGHI2 & "," & MADLT & "," & LANBS & "," & DANHAN & "," & LAN_QUET
    
        sSQL = "INSERT INTO TMP_QTCN" & Right$(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        'InsertHDR = sSQL
        '-----------------------------
        'Mã TM 1008
        MAMUC = "'1000'"
        MATM = "'1008'"
        .GetText .ColLetterToNumber("U"), 48, temp1
        PNNQT = temp1
        THUETKY = PNNQT
        PNKK = 0
        CLECH = PNNQT
        'get CT [25]
        .GetText .ColLetterToNumber("U"), 47, temp1
        DOANHSO = temp1
        GTTINHTHUE = DOANHSO
    
        sSQLVal = MATKHAI & "," & KYLBO & "," & NAMQT & "," & MADTNT & "," & TTHTK & "," & NGNOP & "," & TUKY & "," & DENKY
        sSQLVal = sSQLVal & "," & NGNHAP & "," & HANNOP & "," & MAMUC & "," & MATM & "," & PNNQT & "," & PNKK & "," & CLECH
        sSQLVal = sSQLVal & "," & THUETKY2 & "," & MAVACH & "," & AUTO_CAL & "," & DOANHSO & "," & GTTINHTHUE & "," & THUETKY
        sSQLVal = sSQLVal & "," & DSOCL & "," & GTTTCL & "," & THUECL & "," & DAGHI2 & "," & MADLT & "," & LANBS & "," & DANHAN & "," & LAN_QUET

        sSQL = "INSERT INTO TMP_QTCN" & Right$(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        'InsertHDR = sSQL
        bln = clsDAO.ExecuteDLL(sSQL)
        '-----------------------------
        'Mã TM 1012
        MAMUC = "'1000'"
        MATM = "'1007'"
        .GetText .ColLetterToNumber("U"), 51, temp1
        PNNQT = temp1
        THUETKY = PNNQT
        PNKK = 0
        CLECH = PNNQT
        'get CT[27]
        .GetText .ColLetterToNumber("U"), 50, temp1
        DOANHSO = temp1
        GTTINHTHUE = DOANHSO
    
        sSQLVal = MATKHAI & "," & KYLBO & "," & NAMQT & "," & MADTNT & "," & TTHTK & "," & NGNOP & "," & TUKY & "," & DENKY
        sSQLVal = sSQLVal & "," & NGNHAP & "," & HANNOP & "," & MAMUC & "," & MATM & "," & PNNQT & "," & PNKK & "," & CLECH
        sSQLVal = sSQLVal & "," & THUETKY2 & "," & MAVACH & "," & AUTO_CAL & "," & DOANHSO & "," & GTTINHTHUE & "," & THUETKY
        sSQLVal = sSQLVal & "," & DSOCL & "," & GTTTCL & "," & THUECL & "," & DAGHI2 & "," & MADLT & "," & LANBS & "," & DANHAN & "," & LAN_QUET

        sSQL = "INSERT INTO TMP_QTCN" & Right$(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        'InsertHDR = sSQL
        bln = clsDAO.ExecuteDLL(sSQL)
        '-----------------------------
        'Mã TM 1003
        MAMUC = "'1000'"
        MATM = "'1003'"
        .GetText .ColLetterToNumber("U"), 55, temp1
        PNNQT = temp1
        THUETKY = PNNQT
        PNKK = 0
        CLECH = PNNQT
        'get CT 30
        .GetText .ColLetterToNumber("U"), 54, temp1
        DOANHSO = temp1
        GTTINHTHUE = DOANHSO
    
        sSQLVal = MATKHAI & "," & KYLBO & "," & NAMQT & "," & MADTNT & "," & TTHTK & "," & NGNOP & "," & TUKY & "," & DENKY
        sSQLVal = sSQLVal & "," & NGNHAP & "," & HANNOP & "," & MAMUC & "," & MATM & "," & PNNQT & "," & PNKK & "," & CLECH
        sSQLVal = sSQLVal & "," & THUETKY2 & "," & MAVACH & "," & AUTO_CAL & "," & DOANHSO & "," & GTTINHTHUE & "," & THUETKY
        sSQLVal = sSQLVal & "," & DSOCL & "," & GTTTCL & "," & THUECL & "," & DAGHI2 & "," & MADLT & "," & LANBS & "," & DANHAN & "," & LAN_QUET

        sSQL = "INSERT INTO TMP_QTCN" & Right$(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        'InsertHDR = sSQL
        bln = clsDAO.ExecuteDLL(sSQL)
    End With

    clsDAO.Disconnect
End Function

Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QTCN" & Right$(sKyKeKhai, 4) & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_CNCT" & Right$(sKyKeKhai, 4) & ".DBF"
'    Dim strFileNameBS As String
'    strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"
    'file HDR
    If fso.FileExists(strFileNameHDR) = False Then
        Dim strTepmauHDR As String
        strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_QTCN.DBF"
        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    End If
    'file dtl
    If fso.FileExists(strFileNameDTL) = False Then
        Dim strTepmauDTL As String
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_CNCT.DBF"
        fso.CopyFile strTepmauDTL, strFileNameDTL, False
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

sSQL = "SELECT MADTNT, NGNOP FROM TMP_QTCN" & Right$(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND MATKHAI='06/KK-TNCN' "
    
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

'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'   End If
'
'   sSQL = "UPDATE TMP_TZ" & sKyKeKhai & _
'        " SET THUETKY2 = 1 " & _
'        " WHERE MATHUE = '09' and MAPP = '02' AND MADTNT = " & "'" & strMST & "'" & _
'        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='06/KK-TNCN'"
'
'   clsDAO.ExecuteDLL sSQL
'
'   UpdateTHUETKY2 = True
'
'   clsDAO.Disconnect
End Function

Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL As String
    

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   sSQL = "SELECT TMP_QTCN" & Right$(sKyKeKhai, 4) & _
        " WHERE TTHTK = '1' AND MADTNT = " & "'" & strMST & "'" & _
        " AND MATKHAI='06/KK-TNCN'"
   
   Set rs = clsDAO.Execute(sSQL)
   If rs Is Nothing Then
        isToKhaiChinhThuc = False
   Else
        isToKhaiChinhThuc = True
   End If
   
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
   
   sSQL = "DELETE FROM TMP_QTCN" & Right$(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND MATKHAI='06/KK-TNCN'"
     
    
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
   sSQL = "DELETE FROM TMP_QTCN" & Right$(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "' " & _
        " AND MATKHAI='06/KK-TNCN'"
    
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
        .GetText .ColLetterToNumber("E"), 12, NGNOP
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
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_QTCN" & Right$(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "'"
    
    
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
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QTCN" & Right$(sKyKeKhai, 4) & ".DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        End If
         sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_QTCN" & Right$(sKyKeKhai, 4) & _
                " WHERE MADTNT = " & "'" & strMST & "'" & _
                " AND MATKHAI='06/KK-TNCN'" & " AND(TTHTK = '1' OR TTHTK = '3' OR TTHTK = '4')"
        
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

