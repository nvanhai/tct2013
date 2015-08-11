VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_04TNDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_Y = 30
Const KY_LAP_BO_X = "E"
Const NGAY_NHAN_TO_KHAI_Y = 30
Const NGAY_NHAN_TO_KHAI_X = "M"
Const NGUOI_SU_DUNG_Y = 32
Const NGUOI_SU_DUNG_X = "V"
Const PHONG_XU_LY_Y = 30
Const PHONG_XU_LY_X = "V"
Const NGAY_NOP_Y = 32
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 30
Const NGAY_QUET_X = "M"
'Const TEN_BPQL_X = "V"
'Const TEN_BPQL_Y = 34
Const MA_BPQL_X = "V"
Const MA_BPQL_Y = 34
Private rs As ADODB.Recordset

' thong tin din vi
Const MA_SO_THUE_X = "F"
Const MA_SO_THUE_Y = 10
Const TEN_GOI_X = "H"
Const TEN_GOI_Y = 8
Const DIA_CHI_X = "F"
Const DIA_CHI_Y = 12
'Const QUAN_HUYEN_X = "F"
'Const QUAN_HUYEN_Y = 11
'Const TINH_TPHO_X = "T"
'Const TINH_TPHO_Y = 11
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
'end

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
Public isExistFile As Boolean
Public Lan_quet As Variant
Const STT_X = "M"
Const STT_Y = 30
'Public strMaSoTep As String
Public strNgayNhanToKhai As String
Public strNguoiSuDung As String
Public strPhongXuLy As String
Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long
Private LOAI_TO_KHAI As String

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
Private NGAYPS As Variant
Private MaMucKHBS As String

Private isTkQuy As Boolean


Public Function UpdateTHUETKY2() As Boolean
'    Dim sSQL As String
'
'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'   End If
'
'   sSQL = "UPDATE TMP_QT4" & Right(sKyKeKhai, 4) & _
'        " SET THUETKY2=1" & _
'        " WHERE AND MADTNT = " & "'" & strMST & "'" & _
'        " AND LOAITK='04/TNDN'"
'
'   clsDAO.ExecuteDLL sSQL

   UpdateTHUETKY2 = True

   'clsDAO.Disconnect

End Function

Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
        'Ghi chu
        .Col = .ColLetterToNumber("M")
        .Row = 36
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100

        ' Ky lap bo
        SetDateFormat fps, 1, 30, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft

        'Phong xu ly
        .Col = .ColLetterToNumber("V")
        .Row = 30
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60

        ' Ngay nop
        SetDateFormat fps, 1, 32, .ColLetterToNumber("E"), DDMMYYYY
        .Sheet = 1
        .Row = 32
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft

        ' Ngay quet
        SetDateFormat fps, 1, 32, .ColLetterToNumber("M"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 32
        .Col = .ColLetterToNumber("M")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft

   End With
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
    Dim i       As Integer
    Dim varMaDM As Variant
     
    i = 0
    ReDim Preserve larrPhongXuLy(0)
    larrPhongXuLy(0) = "00"
    ReDim Preserve larrid(0)
    larrid(0) = "00"
    Dim iCol As Long, iRow As Long

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
            
            'get ten chi tieu trong danh muc NNKD
            .Sheet = 1
            .Col = .ColLetterToNumber("B")
            .Row = 44

'            Do
'                .GetText .ColLetterToNumber("AV"), .Row, varMaDM
'
'                If Trim(varMaDM) <> vbNullString Then
'                    .Col = .ColLetterToNumber("C")
'
'                    .CellType = CellTypeStaticText
'                    .TypeTextWordWrap = True
'                    .RowHeight(.Row) = 20
'                    .ColWidth(.ColLetterToNumber("C")) = 9
'
'                    .Text = getTenDMNNKD(varMaDM)
'                End If
'
'                .Col = .ColLetterToNumber("B")
'                .Row = .Row + 1
'            Loop Until .Text = "aa"

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

    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1

        'set check loi dinh danh
        If checkDLTTonTai = False Then
            .Col = .ColLetterToNumber("E")
            .Row = 36
            .Value = 1
        End If

        '        'Set NgayNhanToKhai
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

        '        .Col = .ColLetterToNumber("V")
        '        .Row = 30
        '        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        '
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

        .Col = .ColLetterToNumber("E")
        .Row = 30

        If Len(Month(Now())) = 1 Then
            .Text = "0" & Month(Now()) & "/" & Year(Now())
            sKyKeKhai = "0" & Month(Now()) & Year(Now())
        Else
            .Text = Month(Now()) & "/" & Year(Now())
            sKyKeKhai = Month(Now()) & Year(Now())
        End If

        UpdateCell fps, .Col, .Row, .Text

        .Col = .ColLetterToNumber("E")
        .Row = 32
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        UpdateCell fps, .Col, .Row, .Text
        .Lock = False

        .Col = .ColLetterToNumber("M")
        .Row = 32
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        UpdateCell fps, .Col, .Row, .Text

        ' ten nguoi nop the
        .Col = .ColLetterToNumber(TEN_GOI_X)
        .Row = TEN_GOI_Y

        If strTenGoi <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenGoi, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If

        ' dia chi
        .Col = .ColLetterToNumber(DIA_CHI_X)
        .Row = DIA_CHI_Y

        If strDchi <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strDchi, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If

        ' dien thoai
        .Col = .ColLetterToNumber(DIEN_THOAI_X)
        .Row = DIEN_THOAI_Y

        If strDThoai <> "" Then
            .Text = strDThoai
            UpdateCell fps, .Col, .Row, .Text
        End If

        ' Fax
        .Col = .ColLetterToNumber(FAX_X)
        .Row = FAX_Y

        If strFax <> "" Then
            .Text = strFax
            UpdateCell fps, .Col, .Row, .Text
        End If

        'Mail
        .Col = .ColLetterToNumber(MAIL_X)
        .Row = MAIL_Y

        If strFax <> "" Then
            .Text = strFax
            UpdateCell fps, .Col, .Row, .Text
        End If

        '------------- them thong tin dai ly thue ------------------------------
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

        '        'dia chi
        .Col = .ColLetterToNumber(DIA_CHI_DL_X)
        .Row = DIA_CHI_DL_Y

        If strDchi_DLT <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strDchi_DLT, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        ''        ' quan huyen
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

        '        'dien thoai
        .Col = .ColLetterToNumber(DIEN_THOAI_DL_X)
        .Row = DIEN_THOAI_DL_Y

        If strDthoai_DLT <> vbNullString Then
            .Text = strDthoai_DLT
        End If

        '        'fax
        .Col = .ColLetterToNumber(FAX_DL_X)
        .Row = FAX_DL_Y

        If strFax_DLT <> "" Then
            .Text = strFax_DLT
        End If

        '        ' mail
        .Col = .ColLetterToNumber(MAIL_DL_X)
        .Row = MAIL_DL_Y

        If strMail_DLT <> "" Then
            .Text = strMail_DLT
        End If

        '        ' so hop dong
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

        ' ten bo phan quan ly
        '        .Col = .ColLetterToNumber(TEN_BPQL_X)
        '        .Row = TEN_BPQL_Y
        '        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        
        'so thu tu to khai
        .Col = .ColLetterToNumber(STT_X)
        .Row = STT_Y
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))

        ' lan quet
        .Col = .ColLetterToNumber("M")
        .Row = 30
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
'        .EventEnabled(EventAllEvents) = True
        
        ' ngay phat sinh
        .GetText .ColLetterToNumber("S"), 37, NGAYPS

        Dim tokhai As Variant
        .GetText .ColLetterToNumber("Q"), 37, tokhai

        If tokhai = "1" Then
            LOAI_TO_KHAI = "04/TNDN"
            TKTHANG = False
            TAX_Utilities_Svr_New.Month = Mid$(NGAYPS, 4, 2)
            TAX_Utilities_Svr_New.Year = Right$(NGAYPS, 4)
        Else
            LOAI_TO_KHAI = "04/TNDN"
            TKTHANG = True
        End If
        
        If Trim(NGAYPS) = vbNullString Then
            NGAYPS = "CTOD('')"
        Else
            NGAYPS = ToDate(Trim(NGAYPS), DDMMYYYY)
            NGAYPS = "CTOD('" & Format(NGAYPS, "mm/dd/yyyy") & "')"
        End If

    End With
    
    Prepared3 = True
End Function

Public Sub Prepared4(ByVal dNgayDauKy As Date)
Dim strKyLapBo As String
Dim blnValid As Boolean

    blnValid = True
    
'    With fps
'        .EventEnabled(EventAllEvents) = False
'        .Sheet = 1
'        .Col = .ColLetterToNumber(KY_LAP_BO_X)
'        .Row = KY_LAP_BO_Y
'        strKyLapBo = .Text
'
'        ' Get Phong xu ly
'        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
'        .Row = PHONG_XU_LY_Y
'        strPhongXuLy = .Text
'
'        ' Get NgayNhanToKhai
'        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
'        .Row = NGAY_NHAN_TO_KHAI_Y
'        strNgayNhanToKhai = .Text
'
'        'Go to last sheet (header sheet)
'        .Sheet = .SheetCount
'
'        If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
'            blnValid = False
'        End If
'
'        .EventEnabled(EventAllEvents) = True
 '   End With
End Sub

Public Function InsertDTL() As String
    Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
'   Dim MST As Variant
'   Dim KYKK As Variant
'   Dim IDHDR As Variant
'   Dim LOAITK As Variant
'   Dim mact As Variant
'   Dim giatri As Variant
   
   Dim madtnt   As Variant
   Dim ngnop    As Variant
   Dim NAMQT    As Variant
   Dim TTHTK    As Variant
   Dim MaMuc    As Variant
   Dim maTM     As Variant
   Dim matkhai  As Variant
   Dim kykkhai  As Variant
   Dim STT      As Variant
   Dim dthu1    As Variant
   Dim tyle1    As Variant
   Dim thue1    As Variant
   Dim dthu2    As Variant
   Dim tyle2    As Variant
   Dim thue2    As Variant
   Dim dthu3    As Variant
   Dim tyle3    As Variant
   Dim thue3    As Variant
   Dim tong     As Variant
'   Dim Lan_quet As Variant
   Dim danhan   As Variant
   
   Dim bln  As Boolean
   
   Dim rowLast   As Integer
   
   sSQLCol = "madtnt, ngnop,namqt, tthtk, mamuc, matm, matkhai,kykkhai, STT, dthu1, tyle1, thue1, dthu2, tyle2, thue2, dthu3, tyle3, thue3, tong, Lan_quet, danhan"
    
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
        
        'TTHTK = "''"
         'trang thai to khai
        Dim TTHTK1 As Variant

        .GetText .ColLetterToNumber("M"), 37, TTHTK1

        If Trim(TTHTK1) = vbNullString Then
            TTHTK = "'" & 2 & "'"
        Else
            TTHTK = "'" & 1 & "'"
        End If
        
        MaMuc = "'1050'"
        maTM = "'1052'"
        NAMQT = "2014"
        danhan = "''"
        kykkhai = "'" & TAX_Utilities_Svr_New.Year & "'"

        matkhai = "'" & LOAI_TO_KHAI & "'"
        
        .Col = .ColLetterToNumber("B")
        .Row = 45
        
        Do
            'Noi dung
            
            STT = STT + 1
            
            .GetText .ColLetterToNumber("C"), .Row, dthu1
            
            If Trim(dthu1) = vbNullString Then
                dthu1 = "0"
            Else
                dthu1 = dthu1
            End If
            
            .GetText .ColLetterToNumber("P"), .Row, dthu2
            
            If Trim(dthu2) = vbNullString Then
                dthu2 = "0"
            Else
                dthu2 = dthu2
            End If
            
            .GetText .ColLetterToNumber("AB"), .Row, dthu3
            
            If Trim(dthu3) = vbNullString Then
                dthu3 = "0"
            Else
                dthu3 = dthu3
            End If
            
            .GetText .ColLetterToNumber("H"), .Row, tyle1
            
            If Trim(tyle1) = vbNullString Then
                tyle1 = "0"
            Else
                tyle1 = tyle1
            End If
            
            .GetText .ColLetterToNumber("U"), .Row, tyle2
            
            If Trim(tyle2) = vbNullString Then
                tyle2 = "0"
            Else
                tyle2 = tyle2
            End If
            
             .GetText .ColLetterToNumber("AG"), .Row, tyle3
            
            If Trim(tyle3) = vbNullString Then
                tyle3 = "0"
            Else
                tyle3 = tyle3
            End If
            
            .GetText .ColLetterToNumber("K"), .Row, thue1
            
            If Trim(thue1) = vbNullString Then
                thue1 = "0"
            Else
                thue1 = thue1
            End If
            
            .GetText .ColLetterToNumber("W"), .Row, thue2
            
            If Trim(thue2) = vbNullString Then
                thue2 = "0"
            Else
                thue2 = thue2
            End If
            
            .GetText .ColLetterToNumber("AI"), .Row, thue3
            
            If Trim(thue3) = vbNullString Then
                thue3 = "0"
            Else
                thue3 = thue3
            End If
            
            tong = thue1 + thue2 + thue3
            
             'Lan quet
            If Trim(Lan_quet) = vbNullString Then
                Lan_quet = "1"
            End If
            
                sSQLVal = madtnt & "," & ngnop & "," & NAMQT & "," & TTHTK & "," & MaMuc & "," & maTM & ","
                sSQLVal = sSQLVal & matkhai & "," & kykkhai & "," & STT & "," & dthu1 & "," & tyle1 & "," & thue1 & ","
                sSQLVal = sSQLVal & dthu2 & "," & tyle2 & "," & thue2 & "," & dthu3 & "," & tyle3 & "," & thue3 & "," & tong & "," & Lan_quet & "," & danhan
   
                sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                bln = clsDAO.ExecuteDLL(sSQL)
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
            rowLast = .Row
        Loop Until .Text = "aa"

    End With

    clsDAO.Disconnect
   
End Function

Public Function InsertDTL_KHBS() As String
Dim madtnt As Variant
    Dim ngnop  As Variant
    Dim KYLBO  As Variant

    With fps
        .Sheet = 1
        
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 30, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If

    End With

    InsertDTL_KHBS = InsertDTL_KHBS156(fps, madtnt, "04/TNDN", ngnop, KYLBO, spathVat, TTHTK, Lan_quet, False, Right(sKyKeKhai, 4), MaMucKHBS)
    
    InsertHDR_KHBS
    
    
    InsertDTL_KHBS = vbNullString
   
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
        maTM = "'1052'"
        matkhai = "'04/TNDN'"
        

        
        MATHUE = "'23'"
        MAPP = "'2'"
       

        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 30, KYLBO

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
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 17, ngayhoan

        If Trim(ngayhoan) = vbNullString Then
            ngayhoan = "CTOD('')"
        Else
            ngayhoan = ToDate(Trim(ngayhoan), DDMMYYYY)
            ngayhoan = "CTOD('" & Format(ngayhoan, "mm/dd/yyyy") & "')"
        End If
        
        HANNOP = "CTOD('')"
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 13, cqt
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
        
        sSQL = "INSERT INTO TMP_ND" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        bln = clsDAO.ExecuteDLL(sSQL)
        
    End With

    InsertHDR_KHBS = vbNullString
    
    clsDAO.Disconnect
    
End Function

Public Function InsertHDR() As String
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    
    
    Dim madtnt  As Variant
    Dim madlt   As Variant
    Dim ngnop   As Variant
    Dim NAMQT   As Variant
    Dim kykkhai  As Variant
    Dim kytttu  As Variant
    Dim kyttden As Variant
    Dim tien    As Variant
    Dim MaVach  As Variant
    Dim TTHTK   As Variant
    Dim NGNHAP  As Variant
    Dim HANNOP  As Variant
    Dim dathay  As Variant
    Dim KYLBO   As Variant
    Dim MaMuc   As Variant
    Dim maTM    As Variant
    Dim SHTEP   As Variant
    Dim DaGhi   As Variant
    Dim DAGHI1  As Variant
    Dim DAGHI2  As Variant
    Dim AUTO_CAL    As Variant
    Dim lanbs   As Variant
    Dim tiencl   As Variant
    Dim TENNV   As Variant
    Dim CHUNGCHI    As Variant
    Dim NGANH   As Variant
    Dim HTHUC_NOP   As Variant
'    Dim Lan_quet As Variant
    Dim ITKHAI_ID As Variant
    
    Dim bln  As Boolean
    
    sSQLCol = "madtnt,madlt,ngnop,namqt,kykkhai,kytttu,kyttden,tien,mavach,tthtk,ngnhap,hannop,dathay,kylbo,mamuc,matm,shtep,daghi,daghi1,daghi2,auto_cal,lanbs,tiencl,tennv,chungchi,nganh,hthuc_nop,lan_quet,itkhai_id"
           
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
        
        .GetText .ColLetterToNumber("F"), 20, madlt
        If Trim(madlt) = vbNullString Then
            madlt = "''"
        Else
            madlt = "'" & madlt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
        'NAMQT = "2014"
        If NGAYPS <> vbNullString And NGAYPS <> "CTOD('')" Then
            NAMQT = TAX_Utilities_Svr_New.Year
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
            kytttu = kykkhai
            kyttden = kykkhai
        Else
            NAMQT = TAX_Utilities_Svr_New.Year
            kykkhai = "''"
             kytttu = "'01/" & TAX_Utilities_Svr_New.Year & "'"
            kyttden = "'12/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        
        
        'kykkhai = "'" & TAX_Utilities_Svr_New.Year & "'"
                
'        .GetText .ColLetterToNumber("E"), 10, kykkhai
'        If Trim(KYKK) = vbNullString Then
'            kykkhai = "''"
'        Else
'            kykkhai = "'" & kykkhai & "'"
'        End If
        
'        .GetText .ColLetterToNumber("M"), 12, NGAYQUET
'        If Trim(NGAYQUET) = vbNullString Then
'            NGAYQUET = "CTOD('')"
'        Else
'            NGAYQUET = "CTOD('" & Format(NGAYQUET, "mm/dd/yyyy") & "')"
'        End If
'
'        .GetText .ColLetterToNumber("M"), 14, GHICHU
'        If Trim(GHICHU) = vbNullString Then
'            GHICHU = "''"
'        Else
'            GHICHU = "'& GHICHU &'"
'        End If

       
        tien = "0"
        
        
        'TTHTK = "''"
        'trang thai to khai
        Dim TTHTK1 As Variant

        .GetText .ColLetterToNumber("M"), 37, TTHTK1

        If Trim(TTHTK1) = vbNullString Then
            TTHTK = "'" & 2 & "'"
        Else
            TTHTK = "'" & 1 & "'"
        End If
        
        
        
        
        MaVach = ".T."
        
        .GetText .ColLetterToNumber("M"), 32, NGNHAP
        'NGNOP = Date
        If Trim(NGNHAP) = vbNullString Then
            NGNHAP = "CTOD('')"
        Else
            NGNHAP = ToDate(Trim(NGNHAP), DDMMYYYY)
            NGNHAP = "CTOD('" & Format(NGNHAP, "mm/dd/yyyy") & "')"
        End If
        
        'HANNOP = "CTOD('')"
        
        ' ---------- han nop -----------------------
        HANNOP = "CTOD('')"
'        If NGAYPS <> "CTOD('')" Then
'            HANNOP = DateAdd("d", 10, Ng)
'            HANNOP = ToString(HANNOP, "DD/MM/YYYY")
'        Else
'           If Trim(HANNOP) = vbNullString Then
'                HANNOP = "CTOD('')"
'            Else
'                HANNOP = DateSerial(Int(Mid$(HANNOP, 7, 4)), Int(Mid$(HANNOP, 4, 2)), Int(Mid$(HANNOP, 1, 2)))
'                HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
'            End If
'        End If
        
        
        dathay = ".T."
        
        .GetText .ColLetterToNumber("E"), 30, KYLBO
        
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
        
        MaMuc = "'1050'"
        maTM = "'1052'"
        
        SHTEP = "''"
        DaGhi = ".T."
        DAGHI1 = ".F."
        .GetText .ColLetterToNumber("E"), 36, DAGHI2
        If DAGHI2 = vbNullString Or DAGHI2 = "" Or DAGHI2 = "0" Then
            DAGHI2 = ".F."
        Else
            DAGHI2 = ".T."
        End If
        
        AUTO_CAL = ".F."
             
        '  LanBS
        .GetText .ColLetterToNumber("S"), 6, lanbs
        If lanbs = vbNullString Or lanbs = "0" Then
            lanbs = "0"
        End If
        
        tiencl = "0"
        'TENNV
        .GetText .ColLetterToNumber("F"), .MaxRows - 10, TENNV
        If Trim(TENNV) = vbNullString Then
            TENNV = "''"
        Else
            TENNV = "'" & TENNV & "'"
        End If
        'CHUNGCHI
        .GetText .ColLetterToNumber("F"), .MaxRows - 8, CHUNGCHI
        If Trim(CHUNGCHI) = vbNullString Then
            CHUNGCHI = "''"
        Else
            CHUNGCHI = "'" & CHUNGCHI & "'"
        End If
        
        'Lay truong Nganh trong DTNT2 thuoc folder DTNT
        NGANH = GetMaNganhTheoDTNT(madtnt)
    
        If Trim(NGANH) = vbNullString Then
            NGANH = "''"
        Else
            NGANH = Split(NGANH, "--")(0)
            NGANH = "'" & NGANH & "'"
        End If
        
        HTHUC_NOP = "''"
        'Lan quet
        If Trim(Lan_quet) = vbNullString Then
            Lan_quet = "1"
        End If
        ITKHAI_ID = "0"
        
        
'        .GetText .ColLetterToNumber("E"), 14, LOIDD
'        If Trim(LOIDD) = 0 Then
'            LOIDD = ".F."
'        ElseIf Trim(LOIDD) = 1 Then
'            LOIDD = ".T."
'        End If
'
'        Id = "'" & Mid(MST, 2, Len(MST) - 2) & Mid(KYKK, 2, 2) & Mid(KYKK, 5, 4) & "'"
'        LOAITK = "'04TNDN'"
        
    End With
    
    sSQLVal = madtnt & "," & madlt & "," & ngnop & "," & NAMQT & "," & kykkhai & "," & kytttu & ","
    sSQLVal = sSQLVal & kyttden & "," & tien & "," & MaVach & "," & TTHTK & "," & NGNHAP & "," & HANNOP & ","
    sSQLVal = sSQLVal & dathay & "," & KYLBO & "," & MaMuc & "," & maTM & "," & SHTEP & "," & DaGhi & "," & DAGHI1 & "," & DAGHI2 & ","
    sSQLVal = sSQLVal & AUTO_CAL & "," & lanbs & "," & tiencl & "," & TENNV & "," & CHUNGCHI & ","
    sSQLVal = sSQLVal & NGANH & "," & HTHUC_NOP & "," & Lan_quet & "," & ITKHAI_ID
    
    sSQL = "INSERT INTO TMP_QT4" & Right(sKyKeKhai, 4) & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    clsDAO.Disconnect
     
    InsertHDR = vbNullString
    
End Function
Private Function GetMaNganhTheoDTNT(ByVal MaDTNopThue As Variant) As String
    Dim strDataFileName As String
    Dim pSQL            As String
    Dim strName         As String
    
    Dim rcord           As ADODB.Recordset
    Dim fso             As New FileSystemObject
    Dim clsDAOs As New TAX_Utilities_Svr_New.clsADO
    
    strDataFileName = spathVat & "\DTNT\DTNT2.DBF"

    If fso.FileExists(strDataFileName) Then
        If clsDAOs.Connected = False Then
            ' Lay mamuc theo matm
            clsDAOs.CreateConnectionString spathVat & "\DTNT\"
            clsDAOs.Connect
        End If
    
        pSQL = "SELECT Nganh,Macap,Machuong FROM DTNT2 WHERE MADTNT = " & MaDTNopThue
       
        Set rcord = clsDAOs.Execute(pSQL)

        If Not rcord Is Nothing Then
            strName = rcord.Fields("Nganh") & "--" & rcord.Fields("Macap") & "--" & rcord.Fields("Machuong")
        End If
    End If
    
    GetMaNganhTheoDTNT = strName
    
    clsDAOs.Disconnect
End Function

Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso            As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String

    
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QT4" & Right(sKyKeKhai, 4) & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QT4D" & Right(sKyKeKhai, 4) & ".DBF"
    Dim strFileNameBS As String
    strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & Right(sKyKeKhai, 4) & ".DBF"
    Dim strFileNameBSHDR As String
    strFileNameBSHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_ND" & Right(sKyKeKhai, 4) & ".DBF"
    'file HDR
    If fso.FileExists(strFileNameHDR) = False Then
        Dim strTepmauHDR As String
        strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_QT4.DBF"
        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    End If

    'file dtl
    If fso.FileExists(strFileNameDTL) = False Then
        Dim strTepmauDTL As String
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_QT4D.DBF"
        fso.CopyFile strTepmauDTL, strFileNameDTL, False
    End If
    'file bo sung
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
    
    isExistFile = True
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    If TKTHANG = True Then
        sSQL = "SELECT MADTNT, NGNOP FROM TMP_QT4" & Right(sKyKeKhai, 4) & " WHERE MADTNT = " & "'" & strMST & "'" & " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Year & "' "
    
    Else
        sSQL = "SELECT MADTNT, NGNOP FROM TMP_QT4" & Right(sKyKeKhai, 4) & " WHERE MADTNT = " & "'" & strMST & "' "
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

Public Function XoaTKTT() As Boolean
    Dim sSQL As String
    Dim KYLBO As Variant
    Dim rs As ADODB.Recordset
    Dim idTKHDR As String 'Tam thoi ID cua to khai nay la MST & KYKK
    If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
    End If
   idTKHDR = strMST & sKyKeKhai
   sSQL = "DELETE FROM TMP_QT4" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='04/TNDN'"
    
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
   sSQL = "DELETE FROM TMP_QT4" & sKyKeKhai & _
        " WHERE AND MADTNT = " & "'" & strMST & "' AND KYLBO = " & KYLBO & _
        " AND MATKHAI='04/TNDN' AND KYKKHAI ='" & TAX_Utilities_Svr_New.Year & "'"
    
   clsDAO.ExecuteDLL sSQL
   
   XoaTKTT = True
   
   clsDAO.Disconnect

End Function


Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String) 'ByRef strPMaSoTep, ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
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
    Dim varCheckValue As Variant, varNoteValue As Variant
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
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
   sSQL = "SELECT MST, NGAYNOP FROM TMP_QT4 WHERE LOAITK = '04TNDN' AND ID = '" & strMST & sKyKeKhai & "'"
    
   Set rs = clsDAO.Execute(sSQL)
   If Not rs Is Nothing Then
        For i = 1 To rs.RecordCount
            If Not rs.EOF Then
                NGNOPDB = rs.Fields("NGAYNOP")
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

Public Function TKRB() As Boolean

   TKRB = True

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

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
 Dim varTemp As Variant
 Dim vNgayQyet As Variant
 With fps
    If .ActiveSheet = 1 Then
        If Col = .ColLetterToNumber("E") And Row = 25 Then
            .GetText Col, Row, varTemp
            If varTemp <> "" And varTemp <> "../../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                  .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                    checkNgayNop varTemp, vNgayQyet
                Else
                    .SetActiveCell Col, Row
                End If
            Else
             .SetText Col, Row, ""
            End If
           UpdateCell fps, .Col, .Row, .Text
        End If
        If Col = .ColLetterToNumber("R") And Row = 25 Then
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
        If Col = .ColLetterToNumber("R") And Row = 52 Then
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
   
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QT4" & Right(sKyKeKhai, 4) & ".DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        End If
            
'            sSQL = "SELECT max(LAN_QUET) as LAN_QUET  FROM TMP_QT4" & Right(sKyKeKhai, 4) & _
'                " WHERE MADTNT = " & "'" & strMST & "'" & _
'                " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Year & "'"
            
            If TKTHANG = True Then
                sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_QT4" & Right(sKyKeKhai, 4) & " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT = '" & TAX_Utilities_Svr_New.Year & "'"
            Else
                ' ngay phat sinh
                Dim NGAY_PS As Variant
                fps.Sheet = 1
                fps.GetText fps.ColLetterToNumber("S"), 37, NGAY_PS
'
'                If Trim(NGAY_PS) = vbNullString Then
'                    NGAY_PS = "CTOD('')"
'                Else
'                    NGAY_PS = ToDate(Trim(NGAY_PS), DDMMYYYY)
'                    NGAY_PS = "CTOD('" & Format(NGAY_PS, "mm/dd/yyyy") & "')"
'                End If
    
                'sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_QT4" & Right(sKyKeKhai, 4) & " WHERE MADTNT = " & "'" & strMST & "' AND NGAY_PS = " & NGAY_PS
                sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_QT4" & Right(sKyKeKhai, 4) & " WHERE MADTNT = " & "'" & strMST & "' AND KYKKHAI = '" & Mid$(NGAY_PS, 4) & "'"
    
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

'Private Function getTenDMNNKD(maDM As Variant) As String
'    Dim maDM1           As String
'    Dim strDuongDanFile As String
'    Dim xmlDOMdata      As New MSXML.DOMDocument
'    Dim xmlNodelist     As MSXML.IXMLDOMNodeList
'    Dim xmlNode         As MSXML.IXMLDOMNode
'    Dim arrStrName()    As String
'
'    maDM1 = CStr(Trim(maDM))
'    strDuongDanFile = GetCatalogueFileName
'    'lay file chua tenDM
'    strDuongDanFile = "..\InterfaceTemplates\Catalogue_TyLeThue_TNDN.xml"
'
'    strDuongDanFile = GetAbsolutePath(strDuongDanFile)
'    xmlDOMdata.Load (strDuongDanFile)
'    getTenDMNNKD = ""
'    Set xmlNodelist = xmlDOMdata.getElementsByTagName("Item")
'
'    For Each xmlNode In xmlNodelist
'
'        If Left(GetAttribute(xmlNode, "Value"), 2) = maDM1 Then
'            arrStrName = Split(GetAttribute(xmlNode, "Value"), "###")
'            getTenDMNNKD = arrStrName(1)
'        End If
'
'    Next
'
'End Function


Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim vKeKhai As Variant
    Dim vKeKhai_temp As Variant
    Dim vKyLB As Variant
    'them bien de kiem tra su ton tai cua bang dl
    Dim strTestFileName As String
    Dim fso As New FileSystemObject
    
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
        strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QT4" & Right(vKeKhai_temp, 4) & ".DBF"
        'kiem tra neu ton tai bang thi moi thuc hien truy van
        If fso.FileExists(strTestFileName) Then
            If TKTHANG = True Then
                sSQL = "SELECT * FROM TMP_QT4" & Right(vKeKhai_temp, 4) & _
                " WHERE  MADTNT = " & "'" & strMST & "'" & _
                " AND NAMQT = " & TAX_Utilities_Svr_New.Year
            Else
                ' ngay phat sinh
                Dim NGAY_PS As Variant
                fps.Sheet = 1
                fps.GetText fps.ColLetterToNumber("S"), 37, NGAY_PS
    
'                If Trim(NGAY_PS) = vbNullString Then
'                    NGAY_PS = "CTOD('')"
'                Else
'                    NGAY_PS = ToDate(Trim(NGAY_PS), DDMMYYYY)
'                    NGAY_PS = "CTOD('" & Format(NGAY_PS, "mm/dd/yyyy") & "')"
'                End If
                sSQL = "SELECT * FROM TMP_QT4" & Right(vKeKhai_temp, 4) & _
                " WHERE  MADTNT = " & "'" & strMST & "'" & _
                " AND KYKKHAI ='" & Mid$(NGAY_PS, 4) & "'"
            End If
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


