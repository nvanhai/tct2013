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
Private clsConn As New TAX_Utilities_Svr_New.clsADO

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

'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
Public isExistFile As Boolean
Public lan_quet As Variant
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

Const HEADER_SO_TT_TRONG_TEP_X = "V"
Const HEADER_SO_TT_TRONG_TEP_Y = 10
'Const PHONG_XU_LY_Y = 16
'Const PHONG_XU_LY_X = "E"

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'Public strMaSoTep As String
'Public strNgayNhanToKhai As String
'Public strNguoiSuDung As String
Public strMST As String
'Public strPhongXuLy As String
Public strTenGoi As String
Public strDchi As String
Public strNganh As String
Public strMaBPQL As String
Public strDThoai As String
Public strFax As String
Public spathVat As String

'Private larrid() As String
'Private larrPhongXuLy() As String
'Private lSoPhongXL As Long
Public sSaiCT11 As String
Public sKyKeKhai As String
Public dNgayNopDB As Variant
Public hannop As String
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
Private TKTHANG As Variant
Private NgayPs As Variant



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
        If rsPXL.Fields.count > 0 Then
            Do While Not rsPXL.EOF
                .TypeComboBoxIndex = -1
                .TypeComboBoxString = TAX_Utilities_Svr_New.Convert(rsPXL.Fields(1).Value, TCVN, UNICODE)
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
    'strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

'Public Function Prepared3() As Boolean
'    With fps
'        '.EventEnabled(EventAllEvents) = False
'        'Set MaSoTep
'        .Sheet = 1
''        .Col = .ColLetterToNumber(MA_SO_TEP_X)
''        .Row = MA_SO_TEP_Y
''        .Text = strMaSoTep
''        UpdateCell fps, .Col, .Row, .Text
'
'        'Set NgayNhanToKhai
'        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
'        .Row = NGAY_NHAN_TO_KHAI_Y
'        If strNgayNhanToKhai <> "" Then
'            .Text = strNgayNhanToKhai
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'
''**************************************************************
''       ThanhDX remove
''       Date: 14/09/2006
''       Reason: Doi Nguoi su dung -> Phong quan ly.
'
''        'Set NguoiSuDung
''        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_X)
''        .Row = NGUOI_SU_DUNG_Y
''        If strNguoiSuDung <> "" Then
''            .Text = strNguoiSuDung
''            UpdateCell fps, .Col, .Row, .Text
''        End If
''**************************************************************
'
'         'set Phong xu ly
'        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
'        .Row = PHONG_XU_LY_Y
'        If strPhongXuLy <> "" Then
'            .Text = strPhongXuLy
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'
'        .Col = .ColLetterToNumber(MA_SO_THUE_X)
'        .Row = MA_SO_THUE_Y
'        If strMST <> "" Then
'            .Text = strMST
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'
'        .Col = .ColLetterToNumber(KY_LAP_BO_X)
'        .Row = KY_LAP_BO_Y
'        If Len(Month(Now())) = 1 Then
'            .Text = "0" & Month(Now()) & "/" & Year(Now())
'            sKyKeKhai = "0" & Month(Now()) & Year(Now())
'        Else
'            .Text = Month(Now()) & "/" & Year(Now())
'            sKyKeKhai = Month(Now()) & Year(Now())
'        End If
'        UpdateCell fps, .Col, .Row, .Text
'
'        .Col = .ColLetterToNumber(NGAY_NOP_X)
'        .Row = NGAY_NOP_Y
'        .Value = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = .ColLetterToNumber(NGAY_QUET_X)
'        .Row = NGAY_QUET_Y
'        .Value = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = .ColLetterToNumber(TEN_GOI_X)
'        .Row = TEN_GOI_Y
'        If strTenGoi <> "" Then
'            .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenGoi, TCVN, UNICODE))
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'
'        .Col = .ColLetterToNumber(DIA_CHI_X)
'        .Row = DIA_CHI_Y
'        If strDchi <> "" Then
'            .Text = Trim(TAX_Utilities_Svr_New.Convert(strDchi, TCVN, UNICODE))
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'
'        .Col = .ColLetterToNumber(DIEN_THOAI_X)
'        .Row = DIEN_THOAI_Y
'        If strDThoai <> "" Then
'            .Text = strDThoai
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'
'        .Col = .ColLetterToNumber(FAX_X)
'        .Row = FAX_Y
'        If strFax <> "" Then
'            .Text = strFax
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'
'        .Col = .ColLetterToNumber(MA_BPQL_X)
'        .Row = MA_BPQL_Y
'        If strMaBPQL <> "" Then
'            .Text = strMaBPQL
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'
'        '.EventEnabled(EventAllEvents) = True
'    End With
'
'    Prepared3 = True
'End Function
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

'        If Len(Month(Now())) = 1 Then
'            .Text = Year(Now())
'            sKyKeKhai = Year(Now())
'        Else
            .Text = Year(Now())
            sKyKeKhai = Year(Now())
        'End If

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
        .Col = .ColLetterToNumber("R")
        .Row = 37
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        .EventEnabled(EventAllEvents) = True
        
        ' ngay phat sinh
        .GetText .ColLetterToNumber("S"), 37, NgayPs

        Dim tokhai As Variant
        .GetText .ColLetterToNumber("Q"), 37, tokhai

        If tokhai = "1" Then
            LOAI_TO_KHAI = "03/PSNTNN"
            TKTHANG = False
            TAX_Utilities_Svr_New.Month = Mid$(NgayPs, 4, 2)
            TAX_Utilities_Svr_New.Year = Right$(NgayPs, 4)
        Else
            LOAI_TO_KHAI = "03/NTNN"
            TKTHANG = True
        End If
        
        If Trim(NgayPs) = vbNullString Then
            NgayPs = "CTOD('')"
        Else
            NgayPs = ToDate(Trim(NgayPs), DDMMYYYY)
            NgayPs = "CTOD('" & Format(NgayPs, "mm/dd/yyyy") & "')"
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

Public Function TinhSoLanQuet() As Integer
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim LAN_QUET1 As Variant
   
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        End If
            sSQL = "SELECT max(LAN_QUET) as LAN_QUET  FROM TMP_TK" & sKyKeKhai & _
                " WHERE MADTNT = " & "'" & strMST & "'" & _
                " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='" & LOAI_TO_KHAI & "'"
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
    With fps
        If .ActiveSheet = 1 And .ActiveCol = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) _
           And .ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then
            If KeyAscii = 46 Or KeyAscii = 45 Then ' KeyAscii of '.' character
                KeyAscii = 0
            End If
        End If
    End With
End Sub

Public Function CheckValidData() As Boolean
    Dim varCheckValue As Variant, varNoteValue As Variant
    
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 14, varCheckValue
        .GetText .ColLetterToNumber("M"), 14, varNoteValue
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

Public Function GenerateSQL_Header(xmlDOMdata As MSXML.DOMDocument, strSQL_HDR As String, vHdrID As Variant, ByVal dNgayDauKy As Date) As String
    Dim xmlList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlAttribute As MSXML.IXMLDOMAttribute
    Dim iRowID As Long, strSQL As String, strTempSQL As String
    Dim dDate As Date, strDate() As String
    Dim vTIN, vTEN_DTNT, vDIA_CHI, vLOAI_TKHAI, vNGAY_NOP, vKyLB
    Dim vKYKK, vNGAY_CAP_NHAT, vNGUOI_CAP_NHAT, vCO_LOI_DDANH
    Dim vSO_HIEU_TEP, vSO_TT_TK, vDA_NHAN, vGHI_CHU_LOI, vCO_GTRINH_02A
    Dim vCO_GTRINH_02B, vCO_GTRINH_02C
    Dim vPHONG_XU_LY
    Dim i As Long, j As Long
    Dim strMaPhongXuLy As String
    
On Error GoTo ErrHandle
    strSQL = strSQL_HDR
    Set xmlList = xmlDOMdata.getElementsByTagName("Cell")
    For Each xmlNode In xmlList
        With xmlNode.Attributes
        
        If Trim(GetAttribute(xmlNode, "MCT")) = vbNullString Then
            Select Case Trim(GetAttribute(xmlNode, "CellID"))
                Case "G_4"
                    vTIN = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "G_5"
                    vTEN_DTNT = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "G_6"
                    vDIA_CHI = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_12"
                    vNGAY_NOP = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_10"
                    vKyLB = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "M_12"
                    vNGAY_CAP_NHAT = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
'                Case "V_12"
'                    vNGUOI_CAP_NHAT = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_14"
                    vCO_LOI_DDANH = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "M_10"
                    vSO_HIEU_TEP = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case HEADER_SO_TT_TRONG_TEP_X & "_" & HEADER_SO_TT_TRONG_TEP_Y
                    vSO_TT_TK = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "M_14"
                    vGHI_CHU_LOI = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case PHONG_XU_LY_X & "_" & PHONG_XU_LY_Y
                    vPHONG_XU_LY = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
            End Select
        End If
        End With
    Next
    
'**************************************************************
'       ThanhDX added
'       Date: 14/09/2006
'       Reason: Lay nguoi cap nhat tu bien Nguoi su dung
     
    vNGUOI_CAP_NHAT = Trim(TAX_Utilities_Svr_New.Convert(strNguoiSuDung, UNICODE, TCVN))
'**************************************************************

    strSQL = strSQL & "'" & vHdrID & "',"
    strSQL = strSQL & "'" & vTIN & "',"
    strSQL = strSQL & "'" & vTEN_DTNT & "',"
    strSQL = strSQL & "'" & vDIA_CHI & "',"
    vLOAI_TKHAI = TAX_Utilities_Svr_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue
    strSQL = strSQL & "'" & vLOAI_TKHAI & "',"
    strSQL = strSQL & "To_date('" & vNGAY_NOP & "','dd/mm/yyyy'),"
    
    'Ky/Quy LB
    If Trim(TAX_Utilities_Svr_New.Month) <> "" Then
        'Ngay dau ky lap bo va ngay cuoi ky lap bo
        strDate = Split(vKyLB, "/")
        dDate = DateSerial(Val(strDate(1)), Val(strDate(0)), 1)
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
        dDate = DateAdd("m", 1, dDate)
        dDate = DateAdd("d", -1, dDate)
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    End If
    
    'Ky/ Quy KK
    If Trim(TAX_Utilities_Svr_New.Month) <> "" Then
        'Ngay dau ky ke khai va ngay cuoi ky ke khai
        'strDate = Split(TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "/")
        'dDate = DateSerial(Val(strDate(1)), Val(strDate(0)), 1)
        dDate = dNgayDauKy
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
        dDate = DateAdd("m", 1, dDate)
        dDate = DateAdd("d", -1, dDate)
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    End If
    
    strSQL = strSQL & "To_date('" & vNGAY_CAP_NHAT & "','dd/mm/yyyy'),"
    strSQL = strSQL & "'" & vNGUOI_CAP_NHAT & "',"
    strSQL = strSQL & "'" & vCO_LOI_DDANH & "',"
    strSQL = strSQL & "'" & vSO_HIEU_TEP & "',"
    strSQL = strSQL & "'" & vSO_TT_TK & "',"
    
    strSQL = strSQL & "'" & vDA_NHAN & "',"
    strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'','',"
    
    With fps
        For i = 1 To lSoPhongXL
            If vPHONG_XU_LY = TAX_Utilities_Svr_New.Convert(larrPhongXuLy(i), UNICODE, TCVN) Then
                strMaPhongXuLy = larrid(i)
                Exit For
            End If
        Next
    End With
    strSQL = strSQL & "'" & strMaPhongXuLy & "')"
    GenerateSQL_Header = strSQL
    
    Exit Function
ErrHandle:
    SaveErrorLog "cls001", "GenerateSQL_Header", Err.Number, Err.Description
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
   Dim MST As Variant
   Dim KYKK As Variant
   Dim IDHDR As Variant
   Dim LOAITK As Variant
   Dim MACT As Variant
   Dim GIATRI As Variant
   
   Dim madtnt   As Variant
   Dim ngnop    As Variant
   Dim namqt    As Variant
   Dim tthtk    As Variant
   Dim mamuc    As Variant
   Dim matm     As Variant
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
   
   Dim bln  As Boolean
   
   Dim rowLast   As Integer
   
   sSQLCol = "madtnt, ngnop,namqt, tthtk, mamuc, matm, matkhai,kykkhai, STT, dthu1, tyle1, thue1, dthu2, tyle2, thue2, dthu3, tyle3, thue3, tong"
   
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
        
        tthtk = "''"
        mamuc = "''"
        matm = "''"
        namqt = "2014"
        
        kykkhai = "'" & TAX_Utilities_Svr_New.Year & "'"

        matkhai = "'04TNDN'"
    
     Do
            'Noi dung
            STT = "1"
            
            .GetText .ColLetterToNumber("C"), 45, dthu1
            
            If Trim(dthu1) = vbNullString Then
                dthu1 = "0"
            Else
                dthu1 = dthu1
            End If
            
            .GetText .ColLetterToNumber("P"), 45, dthu2
            
            If Trim(dthu2) = vbNullString Then
                dthu2 = "0"
            Else
                dthu2 = dthu2
            End If
            
            .GetText .ColLetterToNumber("AB"), 45, dthu3
            
            If Trim(dthu3) = vbNullString Then
                dthu3 = "0"
            Else
                dthu3 = dthu3
            End If
            
            .GetText .ColLetterToNumber("H"), 45, tyle1
            
            If Trim(tyle1) = vbNullString Then
                tyle1 = "0"
            Else
                tyle1 = tyle1
            End If
            
            .GetText .ColLetterToNumber("U"), 45, tyle2
            
            If Trim(tyle2) = vbNullString Then
                tyle2 = "0"
            Else
                tyle2 = tyle2
            End If
            
             .GetText .ColLetterToNumber("AG"), 45, tyle3
            
            If Trim(tyle3) = vbNullString Then
                tyle3 = "0"
            Else
                tyle3 = tyle3
            End If
            
            .GetText .ColLetterToNumber("AB"), 45, thue1
            
            If Trim(thue1) = vbNullString Then
                thue1 = "0"
            Else
                thue1 = thue1
            End If
            
            .GetText .ColLetterToNumber("AB"), 45, thue2
            
            If Trim(thue2) = vbNullString Then
                thue2 = "0"
            Else
                thue2 = thue2
            End If
            
            .GetText .ColLetterToNumber("AB"), 45, thue3
            
            If Trim(thue3) = vbNullString Then
                thue3 = "0"
            Else
                thue3 = thue3
            End If
            
            tong = thue1 + thue2 + thue3
            
            
                sSQLVal = madtnt & "," & ngnop & "," & namqt & "," & tthtk & "," & mamuc & "," & matm & ","
                sSQLVal = sSQLVal & matkhai & "," & kykkhai & "," & STT & "," & dthu1 & "," & tyle1 & "," & thue1 & ","
                sSQLVal = dthu2 & "," & tyle2 & "," & thue2 & "," & dthu3 & "," & tyle3 & "," & thue3 & "," & tong
   
                sSQL = "INSERT INTO TKDTL( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                bln = clsConn.ExecuteDLL(sSQL)
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
            rowLast = .Row
        Loop Until .Text = "aa"

    End With
   clsDAO.Disconnect

End Function

Public Function InsertHDR() As String
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    
    
    Dim madtnt  As Variant
    Dim madlt   As Variant
    Dim ngnop   As Variant
    Dim namqt   As Variant
    Dim kykkhai  As Variant
    Dim kytttu  As Variant
    Dim kyttden As Variant
    Dim tien    As Variant
    Dim mavach  As Variant
    Dim tthtk   As Variant
    Dim ngnhap  As Variant
    Dim hannop  As Variant
    Dim dathay  As Variant
    Dim kylbo   As Variant
    Dim mamuc   As Variant
    Dim matm    As Variant
    Dim shtep   As Variant
    Dim daghi   As Variant
    Dim daghi1  As Variant
    Dim daghi2  As Variant
    Dim auto_cal    As Variant
    Dim lanbs   As Variant
    Dim tiencl   As Variant
    Dim tennv   As Variant
    Dim chungchi    As Variant
    Dim nganh   As Variant
    Dim hthuc_nop   As Variant
    Dim lan_quet As Variant
    Dim ITKHAI_ID As Variant
    
    Dim bln  As Boolean
    
    sSQLCol = "madtnt,madlt,ngnop,namqt,kykhai,kytttu,kyttden,tien,mavach,tthtk,ngnhap,hannop,dathay,kylbo,mamuc,matm,shtep,daghi,daghi1,daghi2,auto_cal,lanbs,tiencl,tennv,chungchi,nganh,hthuc_nop,lan_quet,itkhai_id"
           
    If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & "\TK_NEW\" 'GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
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
        
        namqt = "2014"
        
        kykkhai = "'" & TAX_Utilities_Svr_New.Year & "'"
                
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

        kytttu = "''"
        kyttden = "''"
        tien = "0"
        tthtk = "''"
        mavach = ".T."
        
        .GetText .ColLetterToNumber("M"), 32, ngnhap
        'NGNOP = Date
        If Trim(ngnhap) = vbNullString Then
            ngnhap = "CTOD('')"
        Else
            ngnhap = ToDate(Trim(ngnhap), DDMMYYYY)
            ngnhap = "CTOD('" & Format(ngnhap, "mm/dd/yyyy") & "')"
        End If
        
        hannop = "CTOD('')"
        
        dathay = ".T."
        
        .GetText .ColLetterToNumber("E"), 30, kylbo
        
        If Trim(kylbo) = vbNullString Then
            kylbo = "''"
        Else
            kylbo = "'" & kylbo & "'"
        End If
        
        mamuc = "'1050'"
        matm = "'1052'"
        
        shtep = "''"
        daghi = ".T."
        daghi1 = ".F."
        .GetText .ColLetterToNumber("E"), 36, daghi2
        If daghi2 = vbNullString Or daghi2 = "" Or daghi2 = "0" Then
            daghi2 = ".F."
        Else
            daghi2 = ".T."
        End If
        
        auto_cal = ".F."
             
        '  LanBS
        .GetText .ColLetterToNumber("S"), 6, lanbs
        If lanbs = vbNullString Or lanbs = "0" Then
            lanbs = "0"
        End If
        
        tiencl = "0"
        'TENNV
        .GetText .ColLetterToNumber("F"), .MaxRows - 10, tennv
        If Trim(tennv) = vbNullString Then
            tennv = "''"
        Else
            tennv = "'" & tennv & "'"
        End If
        'CHUNGCHI
        .GetText .ColLetterToNumber("F"), .MaxRows - 8, chungchi
        If Trim(chungchi) = vbNullString Then
            chungchi = "''"
        Else
            chungchi = "'" & chungchi & "'"
        End If
        
        'Lay truong Nganh trong DTNT2 thuoc folder DTNT
        nganh = GetMaNganhTheoDTNT(madtnt)
    
        If Trim(nganh) = vbNullString Then
            nganh = "''"
        Else
            nganh = Split(nganh, "--")(0)
            nganh = "'" & nganh & "'"
        End If
        
        hthuc_nop = "''"
        If Trim(lan_quet) = vbNullString Then
            lan_quet = "1"
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
    
    sSQLVal = madtnt & "," & madlt & "," & ngnop & "," & namqt & "," & kykkhai & "," & kytttu & ","
    sSQLVal = sSQLVal & kyttden & "," & tien & "," & mavach & "," & tthtk & "," & ngnhap & "," & hannop & ","
    sSQLVal = sSQLVal & dathay & "," & kylbo & "," & mamuc & "," & matm & "," & shtep & "," & daghi & "," & daghi1 & "," & daghi2 & ","
    sSQLVal = sSQLVal & auto_cal & "," & lanbs & "," & tiencl & "," & tennv & "," & chungchi & ","
    sSQLVal = sSQLVal & nganh & "," & hthuc_nop & "," & lan_quet & "," & ITKHAI_ID
    
    sSQL = "INSERT INTO TMP_QT4 " & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsConn.ExecuteDLL(sSQL)
    
    clsConn.Disconnect
    InsertHDR = vbNullString
End Function
' Tra ve gia tri "Nganh--macap--machuong"
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

    
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QT4" & sKyKeKhai & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QT4D" & sKyKeKhai & ".DBF"
    
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
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
'    If clsConn.Connected = False Then
''        Me.MousePointer = vbHourglass
''        frmSystem.MousePointer = vbHourglass
'        clsConn.CreateConnectionString spathVat & "\TK_NEW\" 'GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsConn.Connect
''        frmSystem.MousePointer = vbDefault
''        Me.MousePointer = vbDefault
'   End If
        
'   sSQL = "SELECT MST, NGAYNOP FROM TKHDR WHERE LOAITK = '04TNDN' AND ID = '" & strMST & sKyKeKhai & "'"
   Set rs = clsDAO.Execute(sSQL)
   
   If Not rs Is Nothing Then
        dNgayNopDB = rs.Fields("NGAYNOP")
        If Trim(dNgayNopDB) = vbNullString Then
            dNgayNopDB = "CTOD('')"
        Else
            dNgayNopDB = "CTOD('" & Format(dNgayNopDB, "mm/dd/yyyy") & "')"
        End If
        TKTT = True
   Else
        TKTT = False
   End If
   clsConn.Disconnect

End Function

Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim kylbo As Variant
    Dim rs As ADODB.Recordset
    Dim idTKHDR As String 'Tam thoi ID cua to khai nay la MST & KYKK
    If clsConn.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsConn.CreateConnectionString spathVat & "\TK_NEW\" 'GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   idTKHDR = strMST & sKyKeKhai
   sSQL = "DELETE FROM TKHDR WHERE LOAITK = '04TNDN' AND ID = '" & idTKHDR & "'"
    
   clsConn.ExecuteDLL sSQL
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 10, kylbo
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
   sSQL = vbNullString
   sSQL = "DELETE FROM TKDTL WHERE LOAITK = '04TNDN' IDHDR = '" & idTKHDR & "'"
   
   clsConn.ExecuteDLL sSQL
   
   sSQL = vbNullString
   sSQL = "DELETE FROM PLDTL WHERE LOAITK = '04TNDN' IDHDR = '" & idTKHDR & "'"
   
   clsConn.ExecuteDLL sSQL
   
   XoaTKTT = True
   
   clsConn.Disconnect

End Function

Public Function TKRB() As Boolean

   TKRB = True

End Function

Public Function TKTTNGNOP() As Boolean
   Dim ngnop As Variant
   Dim NGNOPDB As Variant
   Dim sSQL As String
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
   
'   If clsConn.Connected = False Then
''        Me.MousePointer = vbHourglass
''        frmSystem.MousePointer = vbHourglass
'        clsConn.CreateConnectionString spathVat & "\TK_NEW\" 'GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsConn.Connect
''        frmSystem.MousePointer = vbDefault
''        Me.MousePointer = vbDefault
'   End If
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
   sSQL = "SELECT MST, NGAYNOP FROM TKHDR WHERE LOAITK = '04TNDN' AND ID = '" & strMST & sKyKeKhai & "'"
    
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
                    clsConn.Disconnect
                    TKTTNGNOP = False
                    Exit Function
                Else
                    TKTTNGNOP = True
                End If
            End If
            rs.MoveNext
        Next
   Else
        clsConn.Disconnect
        TKTTNGNOP = True
   End If

End Function



