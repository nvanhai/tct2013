VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_04GTGT"
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
Private LOAI_TO_KHAI As String

Private isTkQuy As Boolean

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
        
    End With
    
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
    Dim varMaDM  As Variant
    Dim i        As Integer
    Dim LoaiKyKK As Variant
    Dim NgayPs As Variant
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
        .GetText .ColLetterToNumber("K"), 39, LoaiKyKK
        If LoaiKyKK = "0" Then
            LOAI_TO_KHAI = "04/GTGT"
            isTkQuy = False
        ElseIf LoaiKyKK = "1" Then
            LOAI_TO_KHAI = "04Q/GTGT"
            isTkQuy = True
        Else
            LOAI_TO_KHAI = "04/PSGTGT"
            isTkQuy = False
            .GetText .ColLetterToNumber("L"), 39, NgayPs
            TAX_Utilities_Svr_New.Month = Mid$(NgayPs, 4, 2)
            TAX_Utilities_Svr_New.Year = Right$(NgayPs, 4)
        End If
        
'        'get ten chi tieu trong danh muc NNKD
'        .Sheet = 1
'        .Col = .ColLetterToNumber("B")
'        .Row = 45
'        i = 0
'
'        Do
'            .GetText .ColLetterToNumber("AE"), .Row, varMaDM
'
'            If Trim(varMaDM) <> vbNullString Then
'                '                    .SetText .ColLetterToNumber("C"), .Row, getTenDMNNKD(varMaDM)
'                .Col = .ColLetterToNumber("C")
'                .CellType = CellTypeStaticText
'                .TypeTextWordWrap = True
'                .RowHeight(.Row) = 20
'                .Text = getTenDMNNKD(varMaDM)
'            End If
'
'            i = i + 1
'            .Col = .ColLetterToNumber("B")
'            .Row = 45 + i
'        Loop Until UCase(.Text) = "DD"
            
    End With
    
    Prepared2 = True
End Function
Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String) 'ByRef strPMaSoTep, ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

Public Function Prepared3() As Boolean
    Dim sTT1, STT2 As Integer
    
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
        
        'set STT to khai
        'so thu tu to khai
        .Col = .ColLetterToNumber("M")
        .Row = 30
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        
        'st phong quan ly
        .Col = .ColLetterToNumber(PHONQ_QUAN_LY_Y)
        .Row = PHONG_QUAN_LY_X
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        UpdateCell fps, .Col, .Row, .Text
        
        'set lai STT cac dong
        .Sheet = 1
        sTT1 = 1
        STT2 = 1
        .Col = .ColLetterToNumber("B")
        .Row = 45

        Do
            .Text = sTT1 & "." & STT2
            .Row = .Row + 1
            STT2 = STT2 + 1

            If .Text = "aa" Then
                sTT1 = 2
                STT2 = 1
                .Row = .Row + 3
            ElseIf .Text = "bb" Then
                sTT1 = 3
                STT2 = 1
                .Row = .Row + 3
            ElseIf .Text = "cc" Then
                sTT1 = 4
                STT2 = 1
                .Row = .Row + 3
            End If

            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "dd"
        
        'end set STT
        
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
    With fps
        If .ActiveSheet = 1 And .ActiveCol = .ColLetterToNumber("M") _
           And .ActiveRow = 14 Then
            If KeyAscii = 13 Then
                KeyAscii = 0
            End If
        End If
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
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("E_10"), "Value"), "dd/mm/yyyy")
    
    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function

Public Function InsertDTL() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
   
    Dim madtnt     As Variant
    Dim MATHUE     As Variant
    Dim maTM       As Variant
    Dim MAPP       As Variant
    Dim KYLBO      As Variant
    Dim NGANH      As Variant
    Dim ngnop      As Variant
    Dim LoaiTien   As Variant
    Dim TKHODKY    As Variant
    Dim GTRIMV     As Variant
    Dim GTRIMVGTGT As Variant
    Dim GTRIBR     As Variant
    Dim GTTT       As Variant
    Dim THUE       As Variant
    Dim THUESUAT   As Variant
    Dim TYLE       As Variant
    Dim MaMuc      As Variant
    Dim DSO As Variant
    Dim CTIEUTNDN  As Variant
    Dim SOTIEN     As Variant
    Dim SOTIEN2    As Variant
    Dim tt         As Variant
    Dim THUEPS As Variant
    Dim ThuePN As Variant
    Dim matkhai    As Variant
    Dim kykkhai    As Variant
    Dim DANHAN     As Variant
    Dim THUECL     As Variant
   Dim DSoCL As Variant
   Dim THUEPSCL As Variant
   Dim thuepncl As Variant
    Dim bln        As Boolean
    Dim countTemp  As Integer
    'bien ktra la dong chi tieu tong
    Dim isRowSum   As Boolean
    'tinh so lan quet
    'vttoan
    'ngay 28/09/10
    '------------------------------------------------

    ' set value
    isRowSum = True

    If clsDAO.Connected = False Then
        '        Me.MousePointer = vbHourglass
        '        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        '        frmSystem.MousePointer = vbDefault
        '        Me.MousePointer = vbDefault
    End If
    
    sSQLCol = "MATKHAI,MADTNT,KYLBO,KYKKHAI,NGNOP,TTHTK,MAMUC,MATM,NGANH,DSO,TYLE,THUEPS,THUEPN,DSOCL,THUEPSCL,THUEPNCL,MACT,STT,STT2,STTIN,DANHAN,LAN_QUET"
    '--------------------------------
       With fps
        .Sheet = 1
        
        DANHAN = "''"
        'set THUECL chung cho cac chi tieu con lai
        THUECL = "0"
        
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        MATHUE = "'01'"
        MAPP = "'3'"
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
      
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        LoaiTien = "''"
        DANHAN = "''"

        tt = "''"
        ' sua MAMUC & MATM tu nam 2009
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        matkhai = "'" & LOAI_TO_KHAI & "'"
        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        NGANH = "''"
        
        CTIEUTNDN = "'1000'"
        .GetText .ColLetterToNumber("K"), 43, DSO
        
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        
        THUEPS = "0"
        ThuePN = "0"
        THUEPSCL = "0"
        thuepncl = "0"
        DSoCL = DSO
        TYLE = "0"
        sSQLVal = matkhai & "," & madtnt & "," & KYLBO & "," & kykkhai & "," & ngnop & "," & TTHTK & "," & MaMuc & "," & maTM & "," & NGANH & "," & DSO & "," & TYLE & "," & THUEPS & "," & ThuePN & "," & DSoCL & "," & THUEPSCL & "," & thuepncl & "," & CTIEUTNDN & ",1,0,'I'," & DANHAN & "," & LAN_QUET
        sSQL = "INSERT INTO TMP_GT04" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        bln = clsDAO.ExecuteDLL(sSQL)
        
        CTIEUTNDN = "'2000'"
        .GetText .ColLetterToNumber("Q"), 43, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If
        
        .GetText .ColLetterToNumber("Z"), 43, ThuePN

        If Trim(ThuePN) = vbNullString Then
            ThuePN = "0"
        End If

        NGANH = "'0100000'"
        thuepncl = ThuePN
        DSoCL = DSO
        TYLE = "1"

        sSQLVal = matkhai & "," & madtnt & "," & KYLBO & "," & kykkhai & "," & ngnop & "," & TTHTK & "," & MaMuc & "," & maTM & "," & NGANH & "," & DSO & "," & TYLE & "," & THUEPS & "," & ThuePN & "," & DSoCL & "," & THUEPSCL & "," & thuepncl & "," & CTIEUTNDN & ",2,0,'II'," & DANHAN & "," & LAN_QUET
        sSQL = "INSERT INTO TMP_GT04" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        bln = clsDAO.ExecuteDLL(sSQL)
        
        CTIEUTNDN = "'3000'"
        .GetText .ColLetterToNumber("Q"), 44, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If
        
        .GetText .ColLetterToNumber("Z"), 44, ThuePN

        If Trim(ThuePN) = vbNullString Then
            ThuePN = "0"
        End If

        NGANH = "'0200000'"
        thuepncl = ThuePN
        DSoCL = DSO
        TYLE = "5"

        sSQLVal = matkhai & "," & madtnt & "," & KYLBO & "," & kykkhai & "," & ngnop & "," & TTHTK & "," & MaMuc & "," & maTM & "," & NGANH & "," & DSO & "," & TYLE & "," & THUEPS & "," & ThuePN & "," & DSoCL & "," & THUEPSCL & "," & thuepncl & "," & CTIEUTNDN & ",3,0,'III'," & DANHAN & "," & LAN_QUET
        sSQL = "INSERT INTO TMP_GT04" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        bln = clsDAO.ExecuteDLL(sSQL)
        
        CTIEUTNDN = "'4000'"
        .GetText .ColLetterToNumber("Q"), 45, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If
        
        .GetText .ColLetterToNumber("Z"), 45, ThuePN

        If Trim(ThuePN) = vbNullString Then
            ThuePN = "0"
        End If

        NGANH = "'0300000'"
        thuepncl = ThuePN
        DSoCL = DSO
        TYLE = "3"

        sSQLVal = matkhai & "," & madtnt & "," & KYLBO & "," & kykkhai & "," & ngnop & "," & TTHTK & "," & MaMuc & "," & maTM & "," & NGANH & "," & DSO & "," & TYLE & "," & THUEPS & "," & ThuePN & "," & DSoCL & "," & THUEPSCL & "," & thuepncl & "," & CTIEUTNDN & ",4,0,'IV'," & DANHAN & "," & LAN_QUET
        sSQL = "INSERT INTO TMP_GT04" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        bln = clsDAO.ExecuteDLL(sSQL)
        
        CTIEUTNDN = "'5000'"
        .GetText .ColLetterToNumber("Q"), 46, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If
        
        .GetText .ColLetterToNumber("Z"), 46, ThuePN

        If Trim(ThuePN) = vbNullString Then
            ThuePN = "0"
        End If

        NGANH = "'0400000'"
        thuepncl = ThuePN
        DSoCL = DSO
        TYLE = "2"

        sSQLVal = matkhai & "," & madtnt & "," & KYLBO & "," & kykkhai & "," & ngnop & "," & TTHTK & "," & MaMuc & "," & maTM & "," & NGANH & "," & DSO & "," & TYLE & "," & THUEPS & "," & ThuePN & "," & DSoCL & "," & THUEPSCL & "," & thuepncl & "," & CTIEUTNDN & ",5,0,'V'," & DANHAN & "," & LAN_QUET
        sSQL = "INSERT INTO TMP_GT04" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        bln = clsDAO.ExecuteDLL(sSQL)
   End With
   
   '--------------------------------
   
  
    clsDAO.Disconnect
    
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL1
    End If
    InsertDTL = vbNullString
    'InsertDTL = sSQL
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
    Dim i           As Integer, iCol As Long, iRow As Long
    Dim xmlNode     As MSXML.IXMLDOMNode
    
    Dim strFileName As String
    Dim fso         As New FileSystemObject
    Dim DANHAN      As Variant
    'them bien So Tien nop cham va so ngay nop cham trong TT38
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
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        MATHUE = "'01'"
        DANHAN = "''"
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
        
        ' sua MAMUC & MATM tu nam 2009
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If

        MAPP = "'5'"
        
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
            ElseIf MACT2 = "30" Then
                MACT2 = "'010'"
            ElseIf MACT2 = "31" Then
                MACT2 = "'011'"
            ElseIf MACT2 = "32" Then
                MACT2 = "'012'"
            ElseIf MACT2 = "34" Then
                MACT2 = "'014'"
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
            ElseIf MACT2 = "30" Then
                MACT2 = "'010'"
            ElseIf MACT2 = "31" Then
                MACT2 = "'011'"
            ElseIf MACT2 = "32" Then
                MACT2 = "'012'"
            ElseIf MACT2 = "34" Then
                MACT2 = "'014'"
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
        
        i = 1
        .Row = .Row + 3
       
        mact = "''"
        .GetText .ColLetterToNumber("BE"), .Row, MACT2

        If Trim(MACT2) = vbNullString Then
            MACT2 = "''"
        ElseIf MACT2 = "30" Then
            MACT2 = "'010'"
        ElseIf MACT2 = "31" Then
            MACT2 = "'011'"
        ElseIf MACT2 = "32" Then
            MACT2 = "'012'"
        ElseIf MACT2 = "34" Then
            MACT2 = "'014'"
        Else
            MACT2 = "'000'"
        End If
                
        MACT3 = "'2'"
        STT = 3
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

    End With
   
    clsDAO.Disconnect
    InsertDTL_KHBS = vbNullString
   
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
   
    sSQLCol = "MATKHAI, KYLBO, KYKKHAI, MADTNT, TTHTK, NGNOP, LOAIHD, KYHIEU, SOHDON, NGAYPH, HANG, NGMUA," & "MST, DSOCHUA, THUESUAT, THUEGTGT, MACT, STT, STT2, STTIN, GHICHU, MAPL,LAN_QUET, DANHAN,MAUHD "
   
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
        mact = "'1000'"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        GhiChu = "''"
        MAPL = "'PL04GT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & ",''"
                       
        sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = 8

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
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL04GT'"
            mact = "''"
            STT2 = i
            STTIN = "'" & i & "'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                 
            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"
             
        'insert phan phoi cung cap hang hoa thue suat 1%
        LOAIHD = "''"
        SOHDON = "''"
        KYHIEU = "''"
        NGAYPH = "CTOD('')"
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2000'"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        GhiChu = "''"
        MAPL = "'PL04GT'"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0161" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & ",''"
                       
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

            THUESUAT = 1
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL04GT'"
            mact = "''"
            STT2 = i
            STTIN = "'" & i & "'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                           
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                
            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "bb"
             
        'insert thue 5%
        LOAIHD = "''"
        SOHDON = "''"
        KYHIEU = "''"
        NGAYPH = "CTOD('')"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0162" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'3000'"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        GhiChu = "''"
        MAPL = "'PL04GT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & ",''"
                       
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

            THUESUAT = 5
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL04GT'"
            mact = "''"
            STT2 = i
            STTIN = "'" & i & "'"
            
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                           
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "cc"
             
        'insert thue 3%
        LOAIHD = "''"
        SOHDON = "''"
        KYHIEU = "''"
        NGAYPH = "CTOD('')"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0163" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'4000'"
        STT = 4
        STT2 = 0
        STTIN = "'IV'"
        GhiChu = "''"
        MAPL = "'PL04GT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & ",''"
                       
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

            THUESUAT = 3
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL04GT'"
            mact = "''"
            STT2 = i
            STTIN = "'" & i & "'"
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                           
            sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "dd"
            
        'insert thue 2%
        LOAIHD = "''"
        SOHDON = "''"
        KYHIEU = "''"
        NGAYPH = "CTOD('')"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0164" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'5000'"
        STT = 5
        STT2 = 0
        STTIN = "'V'"
        GhiChu = "''"
        MAPL = "'PL04GT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & ",''"
                       
        sSQL = "INSERT INTO TMP_P1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = .Row + 4

        Do
            LOAIHD = "''"
            THUESUAT = 2
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

            THUESUAT = 2
                
            .GetText .ColLetterToNumber("N"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL04GT'"
            mact = "''"
            STT2 = i
            STTIN = "'" & i & "'"
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

Public Function InsertHDR() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    
    Dim LoaiKyKK   As Variant
    
    Dim madtnt     As Variant
    Dim MATHUE     As Variant
    Dim maTM       As Variant
    Dim MAPP       As Variant
    Dim KYLBO      As Variant
    Dim ngnop      As Variant
    Dim LoaiTien   As Variant
    Dim MABPQL     As Variant
    Dim SHTEP      As Variant
    Dim kykkhai    As Variant
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
    Dim matkhai    As Variant
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
    Dim TENNV      As Variant
    Dim CHUNGCHI   As Variant
    Dim NGAYNOP As Variant
    
    sSQLCol = "MADTNT, MATHUE, MATM, MAPP, KYLBO, NGNOP, "
    sSQLCol = sSQLCol & "LOAITIEN, MABPQL, SHTEP, TTHTK, KYKKHAI, "
    sSQLCol = sSQLCol & "NGNHAP, THUETKY, GTTINHTHUE, DAGHI, "
    sSQLCol = sSQLCol & "KTTRUOC, DOANHSO, MAPP2, MAMUC, MAVACH,MATKHAI,HANNOP, THUETKY2,HANNOP2,TENDA,NGANH,DAGHI1,DAGHI2,AUTO_CAL,LAN_QUET, DANHAN,MADLT ,LANBS,MANNKD,THUECL,DSOCL,GTTTCL,NGAYPS,TENNV,CHUNGCHI"
    
    With fps
        .Sheet = 1
        'lay ten nhan vien dai ly thue, chung chi hanh nghe
        .GetText .ColLetterToNumber("H"), 63, TENNV
        'nshung- bo sung
        If (Trim(TENNV) <> vbNullString) Then
            TENNV = TAX_Utilities_Svr_New.Convert(Trim(TENNV), UNICODE, TCVN)
        End If
        
        .GetText .ColLetterToNumber("H"), 65, CHUNGCHI
        'nshung- bo sung
        If (Trim(CHUNGCHI) <> vbNullString) Then
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
        MAPP = "'3'"
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
            NGAYNOP = ngnop
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
            SHTEP = "'" & SHTEP & "'"
        End If

        '.GetText .ColLetterToNumber("M"), 12, NGNHAP
        NgNhap = Date

        If Trim(NgNhap) = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If
        
        ThueTKy = "0"
        .GetText .ColLetterToNumber("O"), 52, ThueTKy

        If Trim(ThueTKy) = "0" Then
            ThueTKy = "0"
        End If
        
        ' Lay loai ky ke khai(thang/quy)
        .GetText .ColLetterToNumber("K"), 39, LoaiKyKK

        If LoaiKyKK = "0" Then
            LoaiKyKK = ".T."
            isTkQuy = False
            
        ElseIf LoaiKyKK = "1" Then
            LoaiKyKK = ".F."
            isTkQuy = True
            
        Else
            LoaiKyKK = ".T."
            .GetText .ColLetterToNumber("L"), 39, NgayPs
            HANNOP = DateAdd("d", 10, NGAYNOP)
            HANNOP = ToString(HANNOP, "DD/MM/YYYY")
            isTkQuy = False
        End If
        
        If Trim(NgayPs) = vbNullString Then
            NgayPs = "CTOD('')"
        Else
            NgayPs = ToDate(Trim(NgayPs), DDMMYYYY)
            NgayPs = "CTOD('" & Format(NgayPs, "mm/dd/yyyy") & "')"
        End If
        
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        If Trim(MANNKD) = vbNullString Then
            MANNKD = "''"
        Else
            MANNKD = "'" & MANNKD & "'"
        End If
        
        .GetText .ColLetterToNumber("S"), 50, CT27

        If Trim(CT27) = vbNullString Then
            CT27 = 0
        End If
        
        GTTinhThue = CT27

        If GTTinhThue = 0 Then
            GTTinhThue = "0"
        Else
            GTTinhThue = Trim(GTTinhThue)
        End If

        DaGhi = ".T."
        KTTRUOC = "0"

        'DOANHSO = "0"
        DoanhSo = CT27

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
    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & ngnop & ", "
    sSQLVal = sSQLVal & LoaiTien & "," & MABPQL & "," & SHTEP & "," & TTHTK & "," & kykkhai & ", "
    sSQLVal = sSQLVal & NgNhap & "," & ThueTKy & "," & GTTinhThue & "," & DaGhi & ", "
    sSQLVal = sSQLVal & KTTRUOC & "," & DoanhSo & "," & MAPP2 & "," & MaMuc & "," & MaVach & ", "
    sSQLVal = sSQLVal & matkhai & "," & HANNOP & "," & ThueTKy2 & "," & HANNOP2 & "," & TENDA & ",'" & NGANH & "'," & DAGHI1 & "," & DAGHI2 & "," & AUTO_CAL & "," & LAN_QUET & "," & DANHAN & "," & madlt & "," & LanBS & "," & MANNKD & "," & THUECL & "," & DSoCL & "," & GTTTCL & "," & NgayPs & ",'" & TENNV & "','" & CHUNGCHI & "'"
       
    sSQL = "INSERT INTO TMP_TK" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    InsertHDR = sSQL
End Function

Public Function TKTT() As Boolean
    Dim sSQL           As String
    Dim rs             As ADODB.Recordset
    Dim fso            As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim strFileNameIDX As String
    Dim strFileNameBS  As String
    Dim strFileNamePL As String
    
    '    'kiem tra ton tai tep *.dbf chua
    '    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "QTGT" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileNameHDR) = False Then
    ''    Dim strTepmauHDR As String
    ''        strTepmauHDR = spathVat & "\tepmau\QTGT.DBF"
    ''        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    '        isExistFile = False
    '        DisplayMessage "0109", msOKOnly, miCriticalError
    '        Exit Function
    '    Else
    '        isExistFile = True
    '    End If
    ''
    '    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "QT2A" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileNameDTL) = False Then
    '    Dim strTepmauDTL As String
    '        strTepmauDTL = spathVat & "\tepmau\QT2A.DBF"
    '        fso.CopyFile strTepmauDTL, strFileNameDTL, False
    '    End If
    'Ghi vao DB trung gian kiem tra file .dbf
    'vttoan  sua
    'ngay 28/09/2010
    '-----------------------------
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_GT04" & sKyKeKhai & ".DBF"
    strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"
    strFileNamePL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_P1" & sKyKeKhai & ".DBF"
    'file HDR
    If fso.FileExists(strFileNameHDR) = False Then
        Dim strTepmauHDR As String
        strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_TK.DBF"
        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    End If

    'file dtl
    If fso.FileExists(strFileNameDTL) = False Then
        Dim strTepmauDTL As String
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_GT04.DBF"
        fso.CopyFile strTepmauDTL, strFileNameDTL, False
    End If
        'file phu luc
    If fso.FileExists(strFileNamePL) = False Then
        Dim strTepmauPL As String
        strTepmauPL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_P1.DBF"
        fso.CopyFile strTepmauPL, strFileNamePL, False
    End If
    'file bo xung
    If fso.FileExists(strFileNameBS) = False Then
        Dim strTepmauBS As String
        strTepmauBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_BS.DBF"
        fso.CopyFile strTepmauBS, strFileNameBS, False
    End If
        
    isExistFile = True

    '    'Xoa file *.IDX
    '    strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "QTGT" & sKyKeKhai & ".IDX"
    '    If fso.FileExists(strFileNameIDX) = True Then
    '        fso.DeleteFile strFileNameIDX, True
    '    End If
    '
    '        strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "QT2A" & sKyKeKhai & ".IDX"
    '    If fso.FileExists(strFileNameIDX) = True Then
    '        fso.DeleteFile strFileNameIDX, True
    '    End If
    '
    '        strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "BS" & sKyKeKhai & ".IDX"
    '    If fso.FileExists(strFileNameIDX) = True Then
    '        fso.DeleteFile strFileNameIDX, True
    '    End If
    
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

    sSQL = "SELECT MADTNT, NGNOP FROM  TMP_TK" & sKyKeKhai & " WHERE MATHUE = '01' AND MADTNT = " & "'" & strMST & "'" & " AND KYKKHAI = '" & vKyKK & "'" & " AND MATKHAI = '" & LOAI_TO_KHAI & "'"
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
   
    sSQL = "UPDATE TMP_QTGT" & sKyKeKhai & " SET DATHAY=.T." & " WHERE LOAITK = '" & LOAI_TO_KHAI & "' AND MADTNT = " & "'" & strMST & "' AND NAMQT = " & TAX_Utilities_Svr_New.Year

    clsDAO.ExecuteDLL sSQL
   
    UpdateTHUETKY2 = True
   
    clsDAO.Disconnect

End Function

Public Function XoaTKTT() As Boolean

    Dim sSQL  As String
    Dim KYLBO As Variant
    
    Dim rs    As ADODB.Recordset
    
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
    
    If clsDAO.Connected = False Then
        '        Me.MousePointer = vbHourglass
        '        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        '        frmSystem.MousePointer = vbDefault
        '        Me.MousePointer = vbDefault
    End If
   
   sSQL = "DELETE FROM TMP_QTGT" & sKyKeKhai & _
         " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT = " & TAX_Utilities_Svr_New.Year & " AND KYLBO=" & KYLBO & _
         " AND LOAITK='" & LOAI_TO_KHAI & "'"
    
    clsDAO.ExecuteDLL sSQL
   
   sSQL = "DELETE FROM TMP_QT2A" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT = " & TAX_Utilities_Svr_New.Year
    
    clsDAO.ExecuteDLL sSQL
   
    XoaTKTT = True
   
    clsDAO.Disconnect

End Function

Public Function TKRB() As Boolean

    TKRB = True

End Function

Public Function TKTTNGNOP() As Boolean
    Dim ngnop       As Variant
    Dim NGNOPDB     As Variant
    Dim sSQL        As String
    Dim THUEGTGT_KT As Double
    Dim rs          As ADODB.Recordset
    Dim i           As Integer
   
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
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_QTGT" & sKyKeKhai & _
        " WHERE LOAITK = '" & LOAI_TO_KHAI & "' AND MADTNT = " & "'" & strMST & "' AND NAMQT <> " & TAX_Utilities_Svr_New.Year
    
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
        If Col = .ColLetterToNumber("E") And Row = 32 Then
            .GetText Col, Row, varTemp
            If varTemp <> "" And varTemp <> "../../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                        
                        'check ngay nop khong lon hon ngay hien tai
                        '                    varTemp = ToDate(CStr(varTemp), DDMMYYYY)
                        '                    If DateDiff("D", Format(varTemp, "mm/dd/yyyy"), Format(Date, "mm/dd/yyyy")) < 0 Then
                        '                       DisplayMessage "0159", msOKOnly, miCriticalError
                        '                       .SetActiveCell Col, Row
                        '                    End If
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

Public Function TinhSoLanQuet() As Integer
    Dim sSQL           As String
    Dim rs             As ADODB.Recordset
    Dim fso            As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim LAN_QUET1      As Variant
    
    '-----------------------------
    'kiem tra ton tai tep *.dbf chua
    'Tinh so lan quyet de ghi vao DB trung gian
    'vttoan sua
    'ngay 28/09/2010
    '-----------------------------
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"

    If fso.FileExists(strFileNameHDR) = True Then
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

    '------------------------------------------
End Function

Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL            As String
    Dim rs              As ADODB.Recordset
    Dim vKeKhai         As Variant
    Dim vKeKhai_temp    As Variant
    Dim vKyLB           As Variant
    'them bien de kiem tra su ton tai cua bang dl
    Dim strTestFileName As String
    Dim fso             As New FileSystemObject
    
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

    Do While flag
        strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & vKeKhai_temp & ".DBF"

        'kiem tra neu ton tai bang thi moi thuc hien truy van
        
        Dim vKyKK As String
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

Public Function TinhMaNganh() As String
    Dim sSQL           As String
    Dim rs             As ADODB.Recordset
    Dim fso            As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim NGANH          As Variant
    Dim TEMP           As New clsADO

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

'ham lay ten DM NNKD
Private Function getTenDMNNKD(maDM As Variant) As String
    Dim maDM1           As String
    Dim strDuongDanFile As String
    Dim xmlDOMdata      As New MSXML.DOMDocument
    Dim xmlNodelist     As MSXML.IXMLDOMNodeList
    Dim xmlNode         As MSXML.IXMLDOMNode
    Dim arrStrName()    As String
    
    maDM1 = CStr(Trim(maDM))
    strDuongDanFile = GetCatalogueFileName
    'lay file chua tenDM
    strDuongDanFile = "..\InterfaceTemplates\Catalogue_DM_nhom_NNKD.xml"
    
    strDuongDanFile = GetAbsolutePath(strDuongDanFile)
    xmlDOMdata.Load (strDuongDanFile)
    getTenDMNNKD = ""
    Set xmlNodelist = xmlDOMdata.getElementsByTagName("Item")

    For Each xmlNode In xmlNodelist

        If Left(GetAttribute(xmlNode, "Value"), 2) = maDM1 Then
            arrStrName = Split(GetAttribute(xmlNode, "Value"), "###")
            getTenDMNNKD = arrStrName(1)
        End If

    Next

End Function
