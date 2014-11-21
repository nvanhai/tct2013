VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_02NTNN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_Y = 30
Const KY_LAP_BO_X = "E"
Const NGAY_NHAN_TO_KHAI_Y = 32
Const NGAY_NHAN_TO_KHAI_X = "M"
Const NGUOI_SU_DUNG_Y = 25
Const NGUOI_SU_DUNG_X = "AF"
Const PHONG_XU_LY_Y = 30
Const PHONG_XU_LY_X = "V"
Const NGAY_NOP_Y = 32
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 32
Const NGAY_QUET_X = "M"
'Const TEN_BPQL_X = "V"
'Const TEN_BPQL_Y = 34
Const MA_BPQL_X = "V"
Const MA_BPQL_Y = 34
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
Const MA_SO_THUE_DL_Y = 20
Const TEN_GOI_DL_X = "H"
Const TEN_GOI_DL_Y = 118
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
Public sKyKeKhai As String
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
Const STT_Y = 30

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
'Public namKyLapBo As String

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
'nkhoan
Public KyTTTu As Variant
Public KyTTDen As Variant
Public namQT As Variant
Private MaMucKHBS As String


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
        .Sheet = 1
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
        
        'set ky tinh thue
        .Row = 37
        .Col = .ColLetterToNumber("Y")
        .Text = TAX_Utilities_Svr_New.FirstDay
        UpdateCell fps, .Col, .Row, .Text
        
        .Col = .ColLetterToNumber("Z")
        .Text = TAX_Utilities_Svr_New.LastDay
        UpdateCell fps, .Col, .Row, .Text
'
'        'ky tinh thue tu ngay
'        .GetText .ColLetterToNumber("Y"), 37, KyTTTu
'        namQT = Val(Right$(KyTTTu, 4))
'        If Trim(KyTTTu) = vbNullString Then
'            KyTTTu = "CTOD('')"
'        Else
'            KyTTTu = "CTOD('" & Format(KyTTTu, "mm/dd/yyyy") & "')"
'        End If
'
'        'ky tinh thue den ngay
'        .GetText .ColLetterToNumber("Z"), 37, KyTTDen
'
'        If Trim(KyTTDen) = vbNullString Then
'            KyTTDen = "CTOD('')"
'        Else
'            KyTTDen = "CTOD('" & Format(KyTTDen, "mm/dd/yyyy") & "')"
'        End If
    End With

    Prepared2 = True
End Function

Public Function Prepared3() As Boolean

    With fps

        .Sheet = 1

        'set check loi dinh danh
        If checkDLTTonTai = False Then
            .Col = .ColLetterToNumber("E")
            .Row = 36
            .Value = 1
        End If
        
        'Set NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
        .Row = NGAY_NHAN_TO_KHAI_Y

        If strNgayNhanToKhai <> "" Then
            .Text = strNgayNhanToKhai
            UpdateCell fps, .Col, .Row, .Text
        End If
        
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

        ' lan quet
 
        'so thu tu to khai
        .Col = .ColLetterToNumber(STT_X)
        .Row = STT_Y
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        
        'ky tinh thue tu ngay
        .GetText .ColLetterToNumber("Y"), 37, KyTTTu
        namQT = Val(Right$(KyTTTu, 4))
        If Trim(KyTTTu) = vbNullString Then
            KyTTTu = "CTOD('')"
        Else
            KyTTTu = "CTOD('" & Format(KyTTTu, "mm/dd/yyyy") & "')"
        End If
        
        'ky tinh thue den ngay
        .GetText .ColLetterToNumber("Z"), 37, KyTTDen
        
        If Trim(KyTTDen) = vbNullString Then
            KyTTDen = "CTOD('')"
        Else
            KyTTDen = "CTOD('" & Format(KyTTDen, "mm/dd/yyyy") & "')"
        End If
        
'        .GetText .ColLetterToNumber("E"), 30, namKyLapBo
'        namKyLapBo = Right$(namKyLapBo, 4)
        
        'hannop
        .Col = .ColLetterToNumber("AD")
        .Row = 37
        .Text = HANNOP
        UpdateCell fps, .Col, .Row, .Text
        
        ' lan quet
        .Col = .ColLetterToNumber("M")
        .Row = 30
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        
    End With

    '
    Prepared3 = True
End Function

Public Sub Prepared4(ByVal dNgayDauKy As Date)

End Sub

Public Function InsertDTL() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    Dim bln        As Boolean
   
    Dim LoaiTK     As Variant
    Dim madtnt     As Variant

    Dim ngnop      As Variant
    Dim CTTN       As Variant
    Dim KeKhai     As Variant
    Dim QuyetToan  As Variant
    Dim KeKhai2    As Variant
    Dim QuyetToan2 As Variant
    Dim KKHAICL    As Variant
    Dim QTOANCL    As Variant
    Dim GhiChu     As Variant
    Dim tt         As Variant
    Dim DANHAN     As Variant
   
    sSQLCol = "LoaiTK, MaDTNT, NamQT, TTHTK, NgNop, CTTN, KeKhai, QuyetToan, KeKhai2, QuyetToan2, TT "
    sSQLCol = sSQLCol + ",KKHAICL,QTOANCL,GHICHU,DANHAN, LAN_QUET"

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If

    With fps
        .Sheet = 1
        'ma to khai
        LoaiTK = "'02/NTNN'"
        ' ma doi tuong nop thue
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        ' ngay nop
        .GetText .ColLetterToNumber("E"), 32, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

'        'trang thai to khai
'        Dim TTHTK1 As Variant
'        .GetText .ColLetterToNumber("M"), 6, TTHTK1
'
'        If Trim(TTHTK1) = vbNullString Then
'            TTHTK = "'" & 3 & "'"
'        Else
'            TTHTK = "'" & 1 & "'"
'        End If
        
        tt = "' '"
        DANHAN = "' '"
        
        .Sheet = 1
        .Col = .ColLetterToNumber("B")
        .Row = 43
        Dim i As Integer
        i = 0
        Dim arrMaHieu() As Variant
                
        arrMaHieu() = Array("001", "002", "003", "004", "005", "006", "007", "008", "009", "010", "011", "012", "013")
      
        Do
            .GetText .ColLetterToNumber("O"), .Row, KeKhai

            If Trim(KeKhai) = vbNullString Then
                KeKhai = "0"
            End If

            .GetText .ColLetterToNumber("S"), .Row, QuyetToan

            If Trim(QuyetToan) = vbNullString Then
                QuyetToan = "0"
            End If
            
            .GetText .ColLetterToNumber("AA"), .Row, GhiChu
    
            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
            
            KeKhai2 = KeKhai
            QuyetToan2 = QuyetToan
            KKHAICL = KeKhai
            QTOANCL = QuyetToan
            CTTN = "'" & arrMaHieu(i) & "'"
            sSQLVal = LoaiTK & "," & madtnt & "," & namQT & "," & TTHTK & "," & ngnop & "," & CTTN & "," & KeKhai & "," & QuyetToan & "," & KeKhai2 & "," & QuyetToan2 & "," & tt & "," & KKHAICL & "," & QTOANCL & "," & GhiChu & "," & DANHAN & "," & LAN_QUET
            sSQL = "INSERT INTO TMP_QTNTD" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
            i = i + 1
        Loop Until .Row = 56

    End With
    
    clsDAO.Disconnect
    
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL1
    End If
    
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2
    End If

    InsertDTL = vbNullString
   
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
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If

    End With

    InsertDTL_KHBS = InsertDTL_KHBS156(fps, madtnt, "02/NTNN", ngnop, KYLBO, spathVat, TTHTK, LAN_QUET, False, sKyKeKhai, MaMucKHBS)
    
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
    Dim DANHAN As Variant
    
    With fps
        .Sheet = 1
        
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        kykkhai = "'" & TAX_Utilities_Svr_New.Year & "'"
        MaMuc = "'1550'"
        MaMucKHBS = MaMuc
        maTM = "''"
        matkhai = "'02/NTNN'"
        
        MATHUE = "'23'"
        MAPP = "'1'"
 
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
    
        DANHAN = "''"
        
        If clsDAO.Connected = False Then
            clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
            clsDAO.Connect
        End If
    
        sSQLCol = "madtnt,kykkhai,matkhai,MATHUE,MaMuc,maTM,MAPP,ngnop,TTHTK,KYLBO,songaycn,sotiencn,tienhoan," & _
            "lenhhoan,ngayhoan,HANNOP,cqt,songay,SOTIEN,lydokhac,DANHAN,LAN_QUET"
            
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & songaycn & "," & sotiencn & "," & tienhoan & "," & _
            lenhhoan & "," & ngayhoan & "," & HANNOP & "," & cqt & "," & songay & "," & SOTIEN & "," & lydokhac & "," & DANHAN & "," & LAN_QUET
        
        sSQL = "INSERT INTO TMP_ND" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        bln = clsDAO.ExecuteDLL(sSQL)
        
    End With

    InsertHDR_KHBS = vbNullString
    
    clsDAO.Disconnect
    
End Function
Public Function InsertHDR() As String
    Dim sSQL    As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim bln     As Boolean
    
    Dim LoaiTK  As Variant
    Dim madtnt  As Variant
    Dim madlt   As Variant

    Dim ngnop   As Variant
    Dim NgNhap  As Variant
    Dim HANNOP  As Variant
    Dim MaVach  As Variant
    Dim MaMuc   As Variant
    Dim maTM    As Variant
    Dim DATHAY  As Variant
    Dim KYLBO   As Variant
    Dim ThueTKy As Variant
    Dim THUECL  As Variant
    Dim LanBS   As Variant
    Dim DANHAN  As Variant
    Dim TENNV      As Variant
    Dim CHUNGCHI   As Variant
    Dim SOHD As Variant
    Dim NGAYHD As Variant
    
    sSQLCol = "LOAITK,MADTNT,MaDLT,NamQT,TTHTK,NGNOP,KyTTTu,KyTTDen,NgNhap,HanNop,"
    sSQLCol = sSQLCol & "MaVach,MaMuc,MaTM,Dathay,KyLBo,ThueTKy,ThueCL,LanBS,DANHAN,LAN_QUET, TENNV, CHUNGCHI,sohd,ngayhd"

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
       
    With fps
        .Sheet = 1
        'loai to khai
        LoaiTK = "'02/NTNN'"
        
        .GetText .ColLetterToNumber("H"), 59, TENNV
         If Trim(TENNV) <> vbNullString Then
            TENNV = TAX_Utilities_Svr_New.Convert(CStr(TENNV), UNICODE, TCVN)
        End If
        .GetText .ColLetterToNumber("H"), 61, CHUNGCHI
        If Trim(CHUNGCHI) <> vbNullString Then
            CHUNGCHI = TAX_Utilities_Svr_New.Convert(CStr(CHUNGCHI), UNICODE, TCVN)
        End If
        ' ma doi tuong not thue
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        ' ma dai ly thue
        .GetText .ColLetterToNumber("F"), 20, madlt

        If Trim(madlt) = vbNullString Then
            madlt = "''"
        Else
            madlt = "'" & madlt & "'"
        End If
        
'        'trang thai to khai
'        Dim TTHTK1 As Variant
'        .GetText .ColLetterToNumber("M"), 6, TTHTK1
'
'        If Trim(TTHTK1) = vbNullString Then
'            TTHTK = "'" & 2 & "'"
'        Else
'            TTHTK = "'" & 1 & "'"
'        End If

        ' han nop
        .GetText .ColLetterToNumber("AD"), 37, HANNOP
        If Trim(HANNOP) = vbNullString Then
            HANNOP = "CTOD('')"
        Else
            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
        End If
        
        ' ngay nop
        .GetText .ColLetterToNumber("E"), 32, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
        'ngay nhap
        .GetText .ColLetterToNumber("M"), 32, NgNhap

        If Trim(NgNhap) = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = ToDate(Trim(NgNhap), DDMMYYYY)
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If
        
        'sohd
        .GetText .ColLetterToNumber("H"), 29, SOHD
        'ngayhd
        .GetText .ColLetterToNumber("N"), 29, NGAYHD
        If Trim(NGAYHD) = vbNullString Then
            NGAYHD = "CTOD('')"
        Else
            NGAYHD = ToDate(Trim(NGAYHD), DDMMYYYY)
            NGAYHD = "CTOD('" & Format(NGAYHD, "mm/dd/yyyy") & "')"
        End If
        
        MaVach = ".T."
       
        DATHAY = ".F."
        
        ' ky lap bo
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
        
        ' thue chenh lech
        THUECL = "0"

        ' so lan bo sung
        If TTHTK = "'" & 2 & "'" Then
            .GetText .ColLetterToNumber("O"), 37, LanBS
                     
        Else
            LanBS = "0"
        End If

        'to khai da chuyen sang VAT
        DANHAN = "''"
            
        ' -------------- ban ghi thu nhat ----------------------
        MaMuc = "'1700'"
        maTM = "'1701'"
        
        'ThueTK
        .Sheet = 1
        .GetText .ColLetterToNumber("S"), 48, ThueTKy
        
        sSQLVal = LoaiTK & "," & madtnt & "," & madlt & "," & namQT & "," & TTHTK & "," & ngnop & "," & KyTTTu & ", "
        sSQLVal = sSQLVal & KyTTDen & "," & NgNhap & "," & HANNOP & "," & MaVach & "," & MaMuc & ", "
        sSQLVal = sSQLVal & maTM & "," & DATHAY & "," & KYLBO & "," & ThueTKy & ", "
        sSQLVal = sSQLVal & THUECL & "," & LanBS & "," & DANHAN & "," & LAN_QUET & ",'" & TENNV & "','" & CHUNGCHI & "','" & SOHD & "'," & NGAYHD
           
        sSQL = "INSERT INTO TMP_QTNT" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        '------------- ban ghi thu hai ------------------------
        MaMuc = "'1050'"
        maTM = "'1052'"
       
       'ThueTK
        .Sheet = 1
        .GetText .ColLetterToNumber("S"), 49, ThueTKy
                
        sSQLVal = LoaiTK & "," & madtnt & "," & madlt & "," & namQT & "," & TTHTK & "," & ngnop & "," & KyTTTu & ", "
        sSQLVal = sSQLVal & KyTTDen & "," & NgNhap & "," & HANNOP & "," & MaVach & "," & MaMuc & ", "
        sSQLVal = sSQLVal & maTM & "," & DATHAY & "," & KYLBO & "," & ThueTKy & ", "
        sSQLVal = sSQLVal & THUECL & "," & LanBS & "," & DANHAN & "," & LAN_QUET & ",'" & TENNV & "','" & CHUNGCHI & "','" & SOHD & "'," & NGAYHD
           
        sSQL = "INSERT INTO TMP_QTNT" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
       
    End With

    clsDAO.Disconnect
    InsertHDR = vbNullString
    
End Function

Public Function TKTT() As Boolean
    Dim sSQL           As String
    Dim rs             As ADODB.Recordset
    Dim fso            As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim strFileNameBS As String
    Dim strFileNameBSHDR As String
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QTNT" & Right(sKyKeKhai, 4) & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QTNTD" & Right(sKyKeKhai, 4) & ".DBF"
    strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"
    strFileNameBSHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_ND" & sKyKeKhai & ".DBF"
    
    
    'file HDR
    If fso.FileExists(strFileNameHDR) = False Then
        Dim strTepmauHDR As String
        strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tepmau\TMP_QTNT.DBF"
        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    End If

    'file dtl
    If fso.FileExists(strFileNameDTL) = False Then
        Dim strTepmauDTL As String
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tepmau\TMP_QTNTD.DBF"
        fso.CopyFile strTepmauDTL, strFileNameDTL, False
    End If
    
    'file phu luc 01
    Dim strFileNamePL As String
    Dim strTepmauPL As String
    strFileNamePL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_NT02PL1" & Right(sKyKeKhai, 4) & ".DBF"

    If fso.FileExists(strFileNamePL) = False Then
        
        strTepmauPL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_NT02PL1.DBF"
        fso.CopyFile strTepmauPL, strFileNamePL, False
    End If
    
    'file phu luc 02
    strFileNamePL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_NT02PL2" & Right(sKyKeKhai, 4) & ".DBF"
    If fso.FileExists(strFileNamePL) = False Then
        strTepmauPL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_NT02PL2.DBF"
        fso.CopyFile strTepmauPL, strFileNamePL, False
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
    isExistFile = True
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
     
    sSQL = "SELECT MADTNT, NGNOP FROM TMP_QTNT" & Right(sKyKeKhai, 4) & " WHERE MADTNT = '" & strMST & "' AND KyTTTu = CTOD('" & Format(TAX_Utilities_Svr_New.FirstDay, "mm/dd/yyyy") & "') AND KyTTDen = CTOD('" & Format(TAX_Utilities_Svr_New.LastDay, "mm/dd/yyyy") & "') AND LOAITK = '02/NTNN'"
    
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

Public Function InsertDTL_PL1() As String
    Dim sSQL      As String
    Dim sSQLCol   As String
    Dim sSQLVal   As String
    Dim bln       As Boolean
   
    Dim madtnt    As Variant

    'Dim TTHTK     As Variant
    Dim ngnop     As Variant
    Dim LanBS     As Variant
    Dim STT       As Variant
    Dim TENNTNN   As Variant
    Dim NUOC      As Variant
    Dim MADTNTVN  As Variant
    Dim MADTNTNN  As Variant
    Dim SOHD      As Variant
    Dim NOIDUNGHD As Variant
    Dim DIADIEM   As Variant
    Dim THOIHAN   As Variant
    Dim GTHDNT    As Variant
    Dim GTHDVN    As Variant
    Dim GTQTHDNT  As Variant
    Dim GTQTHDVN  As Variant
    Dim SOLD      As Variant
    Dim DANHAN    As Variant
    
    sSQLCol = "MADTNT, NAMQT, TTHTK, NGNOP, LANBS, STT, TENNTNN, NUOC, MADTNTVN, MADTNTNN, SOHD "
    sSQLCol = sSQLCol + ", NOIDUNGHD, DIADIEM, THOIHAN, GTHDNT, GTHDVN, GTQTHDNT, GTQTHDVN, SOLD, DANHAN, LAN_QUET"
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
    With fps
        
        .Sheet = 1
        
        ' ma doi tuong nop thue
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
'        'trang thai to khai
'        Dim TTHTK1 As Variant
'        .GetText .ColLetterToNumber("M"), 6, TTHTK1
'
'        If Trim(TTHTK1) = vbNullString Then
'            TTHTK = "'" & 3 & "'"
'        Else
'            TTHTK = "'" & 1 & "'"
'        End If

        ' ngay nop
        .GetText .ColLetterToNumber("E"), 32, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
        ' so lan bo sung
        If TTHTK = "'" & 2 & "'" Then
            .GetText .ColLetterToNumber("O"), 37, LanBS
                     
        Else
            LanBS = "0"
        End If
        
        DANHAN = "''"

        .Sheet = 2
        .Col = .ColLetterToNumber("B")
        .Row = 23

        Do
        
            'so thu tu
            .GetText .ColLetterToNumber("B"), .Row, STT

            If Trim(STT) = vbNullString Then
                STT = "0"
            End If

            'ten nha thau nuoc ngoai
            .GetText .ColLetterToNumber("E"), .Row, TENNTNN

            If Trim(TENNTNN) = vbNullString Then
                TENNTNN = "''"
            Else
                
                TENNTNN = "'" & TAX_Utilities_Svr_New.Convert(Trim(TENNTNN), UNICODE, TCVN) & "'"
            End If

            ' Nuoc
           
            .GetText .ColLetterToNumber("J"), .Row, NUOC

            If Trim(NUOC) = vbNullString Then
                NUOC = "''"
            Else
                
                 NUOC = "'" & TAX_Utilities_Svr_New.Convert(Trim(NUOC), UNICODE, TCVN) & "'"
                 
            End If

            ' ma doi tuong nop thue VN
          
            .GetText .ColLetterToNumber("O"), .Row, MADTNTVN

            If Trim(MADTNTVN) = vbNullString Then
                MADTNTVN = "''"
            Else
                MADTNTVN = "'" & MADTNTVN & "'"
            End If

            ' ma doi tuong nop thue nuoc ngoai
            .GetText .ColLetterToNumber("W"), .Row, MADTNTNN

            If Trim(MADTNTNN) = vbNullString Then
                MADTNTNN = "''"
            Else
                MADTNTNN = "'" & MADTNTNN & "'"
            End If

            ' so hop dong
            .GetText .ColLetterToNumber("AE"), .Row, SOHD

            If Trim(SOHD) = vbNullString Then
                SOHD = "''"
            Else
                SOHD = "'" & SOHD & "'"
            End If

            'noi dung
            .GetText .ColLetterToNumber("AJ"), .Row, NOIDUNGHD

            If Trim(NOIDUNGHD) = vbNullString Then
                NOIDUNGHD = "''"
            Else
                NOIDUNGHD = "'" & TAX_Utilities_Svr_New.Convert(Trim(NOIDUNGHD), UNICODE, TCVN) & "'"
            End If

            ' dia diem
            .GetText .ColLetterToNumber("AO"), .Row, DIADIEM

            If Trim(DIADIEM) = vbNullString Then
                DIADIEM = "''"
            Else
                DIADIEM = "'" & TAX_Utilities_Svr_New.Convert(Trim(DIADIEM), UNICODE, TCVN) & "'"
            End If

            'thoi gian
            .GetText .ColLetterToNumber("AS"), .Row, THOIHAN

            If Trim(THOIHAN) = vbNullString Then
                THOIHAN = "''"
            Else
                
                THOIHAN = "'" & THOIHAN & "'"
            End If
            
            ' gia tri hop dong nuoc ngoai
            .GetText .ColLetterToNumber("AY"), .Row, GTHDNT

            If Trim(GTHDNT) = vbNullString Then
                GTHDNT = "'0'"
            Else
                GTHDNT = "'" & GTHDNT & "'"
            End If
            'gia tri hop dong viet nam

            .GetText .ColLetterToNumber("BG"), .Row, GTHDVN

            If Trim(GTHDVN) = vbNullString Then
                GTHDVN = "0"
            End If

            ' gia tri quyet toan hop dong nuoc ngoai
             .GetText .ColLetterToNumber("BQ"), .Row, GTQTHDNT

            If Trim(GTQTHDNT) = vbNullString Then
                GTQTHDNT = "'0'"
                
                Else
                
                    GTQTHDNT = "'" & GTQTHDNT & "'"
                
                End If
           
            
            'gia tri quyet toan hop dong viet nam
            
            .GetText .ColLetterToNumber("AY"), .Row, GTQTHDVN

            If Trim(GTQTHDVN) = vbNullString Then
                GTQTHDVN = "0"
            End If

            'so lao dong
            .GetText .ColLetterToNumber("CJ"), .Row, SOLD

            If Trim(SOLD) = vbNullString Then
                SOLD = "0"
            End If
        
            sSQLVal = madtnt & "," & namQT & "," & TTHTK & "," & ngnop & "," & LanBS & "," & STT & "," & TENNTNN & "," & NUOC & "," & MADTNTVN & "," & MADTNTNN & "," & SOHD & "," & NOIDUNGHD & "," & DIADIEM & "," & THOIHAN & "," & GTHDNT & "," & GTHDVN & "," & GTQTHDNT & "," & GTQTHDVN & "," & SOLD & "," & DANHAN & "," & LAN_QUET
            sSQL = "INSERT INTO TMP_NT02PL1" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            bln = clsDAO.ExecuteDLL(sSQL)
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
            
        Loop Until .Text = "aa"
        
    End With
    
    InsertDTL_PL1 = vbNullString
End Function

Public Function InsertDTL_PL2() As String
    Dim sSQL      As String
    Dim sSQLCol   As String
    Dim sSQLVal   As String
    Dim bln       As Boolean
   
    Dim madtnt    As Variant

    Dim TTHTK     As Variant
    Dim ngnop     As Variant
    Dim LanBS     As Variant
    Dim STT       As Variant
    Dim TENNTVN   As Variant
    Dim MADTNTVN  As Variant
    Dim MADTNTNN  As Variant
    Dim SOHD      As Variant
    Dim NOIDUNGHD As Variant
    Dim DIADIEM   As Variant
    Dim THOIHAN   As Variant
    Dim GTHDNT    As Variant
    Dim GTHDVN    As Variant
    Dim GTQTHDNT  As Variant
    Dim GTQTHDVN  As Variant
    Dim DANHAN    As Variant
    
    sSQLCol = "MADTNT, NAMQT, TTHTK, NGNOP, LANBS, STT, TENNTVN, MADTNTVN, MADTNTNN, SOHD "
    sSQLCol = sSQLCol + ", NOIDUNGHD, DIADIEM, THOIHAN, GTHDNT, GTHDVN, GTQTHDNT, GTQTHDVN, DANHAN, LAN_QUET"
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
    With fps
        .Sheet = 1
        ' ma doi tuong nop thue
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        'trang thai to khai
        Dim TTHTK1 As Variant
        .GetText .ColLetterToNumber("M"), 6, TTHTK1

        If Trim(TTHTK1) = vbNullString Then
            TTHTK = "'" & 3 & "'"
        Else
            TTHTK = "'" & 1 & "'"
        End If

        ' ngay nop
        .GetText .ColLetterToNumber("E"), 32, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
        ' so lan bo sung
        If TTHTK = "'" & 3 & "'" Then
            .GetText .ColLetterToNumber("O"), 37, LanBS
                     
        Else
            LanBS = "0"
        End If
        
        DANHAN = "''"

        .Sheet = 3
        .Col = .ColLetterToNumber("B")
        .Row = 24

        Do
        
            'so thu tu
            .GetText .ColLetterToNumber("B"), .Row, STT

            If Trim(STT) = vbNullString Then
                STT = "0"
            End If

            'ten nha thau nuoc ngoai
            .GetText .ColLetterToNumber("E"), .Row, TENNTVN

            If Trim(TENNTVN) = vbNullString Then
                TENNTVN = "''"
            Else
                TENNTVN = "'" & TAX_Utilities_Svr_New.Convert(Trim(TENNTVN), UNICODE, TCVN) & "'"
            End If

            ' ma doi tuong nop thue VN
          
            .GetText .ColLetterToNumber("L"), .Row, MADTNTVN

            If Trim(MADTNTVN) = vbNullString Then
                MADTNTVN = "''"
            Else
                MADTNTVN = "'" & MADTNTVN & "'"
            End If

            ' ma doi tuong nop thue nuoc ngoai
            .GetText .ColLetterToNumber("R"), .Row, MADTNTNN

            If Trim(MADTNTNN) = vbNullString Then
                MADTNTNN = "''"
                
            Else
                MADTNTNN = "'" & MADTNTNN & "'"
            End If

            ' so hop dong
            .GetText .ColLetterToNumber("Y"), .Row, SOHD

            If Trim(SOHD) = vbNullString Then
                SOHD = "''"
                Else
                    
               SOHD = "'" & SOHD & "'"
            End If

            'noi dung
            .GetText .ColLetterToNumber("AG"), .Row, NOIDUNGHD

            If Trim(NOIDUNGHD) = vbNullString Then
                NOIDUNGHD = "''"
            Else
                NOIDUNGHD = "'" & TAX_Utilities_Svr_New.Convert(Trim(NOIDUNGHD), UNICODE, TCVN) & "'"
            End If

            ' dia diem
            .GetText .ColLetterToNumber("AM"), .Row, DIADIEM

            If Trim(DIADIEM) = vbNullString Then
                DIADIEM = "''"
            Else
                DIADIEM = "'" & TAX_Utilities_Svr_New.Convert(Trim(DIADIEM), UNICODE, TCVN) & "'"
            End If

            'thoi gian
            .GetText .ColLetterToNumber("AR"), .Row, THOIHAN

            If Trim(THOIHAN) = vbNullString Then
                THOIHAN = "''"
            Else
                
                THOIHAN = "'" & THOIHAN & "'"
            End If
            
            ' gia tri hop dong nuoc ngoai
           
            .GetText .ColLetterToNumber("AW"), .Row, GTHDNT

            If Trim(GTHDNT) = vbNullString Then
                GTHDNT = "'0'"
             Else
                GTHDNT = "'" & GTHDNT & "'"
            End If
            
             

            'gia tri hop dong viet nam

            .GetText .ColLetterToNumber("BD"), .Row, GTHDVN

            If Trim(GTHDVN) = vbNullString Then
                GTHDVN = "0"
            End If

            ' gia tri quyet toan hop dong nuoc ngoai
            .GetText .ColLetterToNumber("BL"), .Row, GTQTHDNT

            If Trim(GTQTHDNT) = vbNullString Then
                GTQTHDNT = "'0'"
            Else
                GTQTHDNT = "'" & GTQTHDNT & "'"
            End If
            
            
            'gia tri quyet toan hop dong viet nam
            
            .GetText .ColLetterToNumber("BS"), .Row, GTQTHDVN

            If Trim(GTQTHDVN) = vbNullString Then
                GTQTHDVN = "0"
            End If
        
            sSQLVal = madtnt & "," & namQT & "," & TTHTK & "," & ngnop & "," & LanBS & "," & STT & "," & TENNTVN & "," & MADTNTVN & "," & MADTNTNN & "," & SOHD & "," & NOIDUNGHD & "," & DIADIEM & "," & THOIHAN & "," & GTHDNT & "," & GTHDVN & "," & GTQTHDNT & "," & GTQTHDVN & "," & DANHAN & "," & LAN_QUET
            sSQL = "INSERT INTO TMP_NT02PL2" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            bln = clsDAO.ExecuteDLL(sSQL)
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
            
        Loop Until .Text = "aa"
        
    End With

    InsertDTL_PL2 = vbNullString
End Function

Public Function UpdateTHUETKY2() As Boolean

   UpdateTHUETKY2 = True
   
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
    
    Dim flag As Boolean
    'kiem tra da co to khai chinh thuc chua
    flag = True
    vKyLB = DateSerial(Int(Year(Now())), Int(Month(Now())), 1)
    vKeKhai = DateSerial(Int(Val(TAX_Utilities_Svr_New.Year)), Int(Val(TAX_Utilities_Svr_New.Month)), 1)
    vKeKhai_temp = Right(sKyKeKhai, 4)
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   Do While flag
        strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QTNT" & vKeKhai_temp & ".DBF"
        'kiem tra neu ton tai bang thi moi thuc hien truy van
        If fso.FileExists(strTestFileName) Then
             sSQL = "SELECT MADTNT, NGNOP FROM TMP_QTNT" & vKeKhai_temp & " WHERE MADTNT = " & "'" & strMST & "'" & " AND KyTTTu = " & KyTTTu & " AND KyTTDen = " & KyTTDen & "AND LOAITK = '02/NTNN'"
        
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
'            If Len(Month(vKyLB)) = 1 Then
'                vKeKhai_temp = "0" & Month(vKyLB) & Year(vKyLB)
'            Else
'                vKeKhai_temp = Month(vKyLB) & Year(vKyLB)
'            End If

            vKeKhai_temp = Year(vKyLB)
            flag = True
        Else
            flag = False
        End If
   Loop
   clsDAO.Disconnect
End Function

Public Function XoaTKTT() As Boolean
   XoaTKTT = True
End Function


Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String) 'ByRef strPMaSoTep, ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
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
Public Function TKTTNGNOP() As Boolean
   
        TKTTNGNOP = True

End Function

Public Function TKRB() As Boolean

   TKRB = True

End Function

Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date

    KiemTraKhoaSo = True
    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("E_32"), "Value"), "dd/mm/yyyy")

    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
 Dim varTemp As Variant
 Dim vNgayQyet As Variant
 With fps
    If .ActiveSheet = 1 Then
          If (Col = .ColLetterToNumber("E") Or Col = .ColLetterToNumber("M")) And Row = 32 Then
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
   
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QTNT" & Right(sKyKeKhai, 4) & ".DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        End If
            sSQL = "SELECT max(LAN_QUET) as LAN_QUET  FROM TMP_QTNT" & Right(sKyKeKhai, 4) & " WHERE MADTNT = " & "'" & strMST & "'" & " AND KyTTTu = CTOD('" & Format(TAX_Utilities_Svr_New.FirstDay, "mm/dd/yyyy") & "') AND KyTTDen = CTOD('" & Format(TAX_Utilities_Svr_New.LastDay, "mm/dd/yyyy") & "') AND LoaiTK = '02/NTNN'"
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
