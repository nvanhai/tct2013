VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_BC21_AC_BLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const NGAY_NHAN_TO_KHAI_Y = 13
Const NGAY_NHAN_TO_KHAI_X = "E"
Const NGUOI_SU_DUNG_Y = 15
Const NGUOI_SU_DUNG_X = "AC"
Const PHONG_QL_X = "AC"
Const PHONG_QL_Y = 17
Const STT_X = "S"
Const STT_Y = 13
Const PHONG_XU_LY_Y = 13
Const PHONG_XU_LY_X = "AC"
Private rs As ADODB.Recordset
Const MA_SO_TEP_Y = 13
Const MA_SO_TEP_X = "S"
'Longvh
Const MA_SO_THUE_X = "F"
Const MA_SO_THUE_Y = 7
Const NGAY_NOP_Y = 13
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 15
Const NGAY_QUET_X = "S"

Const TEN_GOI_X = "O"
Const TEN_GOI_Y = 5
Const DIA_CHI_X = "F"
Const DIA_CHI_Y = 9
'Const TEN_BPQL_X = "W"
'Const TEN_BPQL_Y = 16
Const DIEN_THOAI_X = "E"
Const DIEN_THOAI_Y = 11
Const FAX_X = "S"
Const FAX_Y = 11
Const MAIL_X = "AC"
Const MAIL_Y = 11

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
Public isExistFile As Boolean
Public LAN_QUET As Variant

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


Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
         'Ma so tep
        .Col = .ColLetterToNumber("S")
        .Row = 13
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetAlphanumeric
        .TypeMaxEditLen = 20
        'Ghi chu
        .Col = .ColLetterToNumber("S")
        .Row = 19
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100
        
        'Phong xu ly
        .Col = .ColLetterToNumber("AC")
        .Row = 13
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
        ' ngay may huy
        SetDateFormat fps, 1, 15, .ColLetterToNumber("N"), DDMMYYYY
        .TypeHAlign = TypeHAlignLeft
        ' Ngay nop
        SetDateFormat fps, 1, 13, .ColLetterToNumber("E"), DDMMYYYY
        .Sheet = 1
        .Row = 13
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        ' Ngay quet
        SetDateFormat fps, 1, 15, .ColLetterToNumber("S"), DDMMYYYY
        .Sheet = 1
        .Row = 15
        .Col = .ColLetterToNumber("S")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
    End With
    hdrId = GetHdrIdAC(spathVat)
    
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
    Dim tempValue As Variant
    With fps
        .Sheet = 1
        'set phong quan ly
        .Col = .ColLetterToNumber(PHONG_QL_X)
        .Row = PHONG_QL_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        
        'set nguoi su dung
        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_X)
        .Row = NGUOI_SU_DUNG_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strNguoiSuDung, TCVN, UNICODE))
       
        ' set STT to khai
        .Col = .ColLetterToNumber(STT_X)
        .Row = STT_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        
        'Set NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
        .Row = NGAY_NHAN_TO_KHAI_Y
        If strNgayNhanToKhai <> "" Then
            .Text = strNgayNhanToKhai
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' thong tin header
        .Col = .ColLetterToNumber(MA_SO_THUE_X)
        .Row = MA_SO_THUE_Y
        If strMST <> "" Then
            .Text = strMST
            UpdateCell fps, .Col, .Row, .Text
        End If
        
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
        
        .Col = .ColLetterToNumber(MAIL_X)
        .Row = MAIL_Y
        If strMail <> "" Then
            strMail = .Text
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' set lai cot ghi chu
        .Row = 41
        .Col = .ColLetterToNumber("B")
        Do
            .GetText .ColLetterToNumber("AI"), .Row, tempValue
            If CStr(tempValue) = "02" Then
                .SetText .ColLetterToNumber("AF"), .Row, GetAttribute(GetMessageCellById("0171"), "Msg")
            ElseIf CStr(tempValue) = "04" Then
                .SetText .ColLetterToNumber("AF"), .Row, GetAttribute(GetMessageCellById("0172"), "Msg")
            End If
            .Row = .Row + 1
        Loop Until .Text = "aa"
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
    Dim varCheckValue As Variant, varNoteValue As Variant
    Dim vNgayNop, vNgayQuet As Variant
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 14, varCheckValue
        .GetText .ColLetterToNumber("M"), 14, varNoteValue
        .GetText .ColLetterToNumber("E"), 13, vNgayNop
        .GetText .ColLetterToNumber("S"), 15, vNgayQuet
        ' Kiem tra ngay nop khong duoc lon hon ngay quet
        If ToDate(CStr(vNgayNop), "dd/mm/yyyy") > ToDate(CStr(vNgayQuet), "dd/mm/yyyy") Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 12, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 12, "1"
        End If
        
'        If CStr(varCheckValue) = "1" And Trim(CStr(varNoteValue)) = "" Then
'            .Sheet = .SheetCount
'            .SetText 2, 14, "0"
'            CheckValidData = False
'        Else
'            .Sheet = .SheetCount
'            .SetText 2, 14, "1"
'        End If
        
    End With
End Function


Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
    KiemTraKhoaSo = False
'    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
'    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("E_10"), "Value"), "dd/mm/yyyy")
'
'    If dNgayKhoaSo < dKyLapBo Then
'        KiemTraKhoaSo = False
'    End If
End Function

Public Function InsertDTL() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   
   Dim Id As Variant
   Dim Hdr_id As Variant
   Dim So_tt As Variant
   Dim Ten_bl As Variant
   Dim Mau_so As Variant
   Dim Ky_hieu_bl As Variant
   Dim Tu_so As Variant
   Dim Den_so As Variant
   Dim So_luong As Variant
   Dim vLienHD, vGhiChu As Variant
   Dim Loai_bl As Variant
   
   Dim bln  As Boolean
  
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
    
   sSQLCol = "Id, Hdr_id, So_tt, Ten_bl, Mau_so, Ky_hieu_bl,  Tu_so, Den_so, So_luong, lien_bl, ghi_chu, Loai_bl "
  
   '-------------
   With fps
   .Sheet = 1
        Hdr_id = hdrId
        .Col = .ColLetterToNumber("B")
        .Row = 42
            Do
                Id = GetDtlIdAC(spathVat)
                'so thu tu
                .GetText .ColLetterToNumber("B"), .Row, So_tt
                If Trim(So_tt) = vbNullString Then
                    So_tt = "0"
                Else
                    So_tt = So_tt
                End If
                'ten bien lai
                .GetText .ColLetterToNumber("C"), .Row, Ten_bl
                If Trim(Ten_bl) = vbNullString Then
                    Ten_bl = "''"
                Else
                    Ten_bl = TAX_Utilities_Svr_New.Convert(CStr(Ten_bl), UNICODE, TCVN)
                    Ten_bl = "'" & Ten_bl & "'"
                End If
                ' mau so bien lai
                .GetText .ColLetterToNumber("O"), .Row, Mau_so
                If Trim(Mau_so) = vbNullString Then
                    Mau_so = "''"
                Else
                    Mau_so = TAX_Utilities_Svr_New.Convert(CStr(Mau_so), UNICODE, TCVN)
                    Mau_so = "'" & Mau_so & "'"
                End If
                'ky hieu bien lai
                .GetText .ColLetterToNumber("S"), .Row, Ky_hieu_bl
                If Trim(Ky_hieu_bl) = vbNullString Then
                    Ky_hieu_bl = "''"
                Else
                    Ky_hieu_bl = "'" & Ky_hieu_bl & "'"
                End If
                'tu so
                .GetText .ColLetterToNumber("V"), .Row, Tu_so
                If Trim(Tu_so) = vbNullString Then
                    Tu_so = "''"
                Else
                    Tu_so = "'" & Tu_so & "'"
                End If
                ' den so
                .GetText .ColLetterToNumber("Y"), .Row, Den_so
                If Trim(Den_so) = vbNullString Then
                    Den_so = "''"
                Else
                    Den_so = "'" & Den_so & "'"
                End If
                'so luong
                .GetText .ColLetterToNumber("AA"), .Row, So_luong
                If Trim(So_luong) = vbNullString Then
                    So_luong = "0"
                End If
                ' Lien bien lai
                .GetText .ColLetterToNumber("AE"), .Row, vLienHD
                If Trim(vLienHD) = vbNullString Then
                    vLienHD = "''"
                Else
                    vLienHD = TAX_Utilities_Svr_New.Convert(CStr(vLienHD), UNICODE, TCVN)
                    If Len(vLienHD) > 30 Then
                        vLienHD = Left(CStr(vLienHD), 30)
                    End If
                    vLienHD = "'" & vLienHD & "'"
                End If
                ' GHi chu
                .GetText .ColLetterToNumber("AI"), .Row, vGhiChu
                If Trim(vGhiChu) = vbNullString Then
                    vGhiChu = "''"
                Else
                    vGhiChu = TAX_Utilities_Svr_New.Convert(CStr(vGhiChu), UNICODE, TCVN)
                    vGhiChu = "'" & vGhiChu & "'"
                End If
                
                ' loai bl
                .GetText .ColLetterToNumber("AH"), .Row, Loai_bl
                If Trim(Loai_bl) = vbNullString Then
                    Loai_bl = "''"
                Else
                    Loai_bl = "'" & Loai_bl & "'"
                End If
                '------------------
                
                sSQLVal = Id & "," & Hdr_id & "," & So_tt & "," & Ten_bl & "," & Mau_so & ","
                sSQLVal = sSQLVal & Ky_hieu_bl & "," & Tu_so & "," & Den_so & "," & So_luong & "," & vLienHD & "," & vGhiChu & "," & Loai_bl
                
            
                sSQL = "INSERT INTO TMP_DTL_BC21_AC_BLP( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                bln = clsDAO.ExecuteDLL(sSQL)

                .Col = .ColLetterToNumber("B")
                .Row = .Row + 1
            Loop Until .Text = "aa"

     End With
   
    clsDAO.Disconnect
    InsertDTL = vbNullString
End Function


Public Function InsertDTL_KHBS() As String
   InsertDTL_KHBS = vbNullString
   
End Function



Public Function InsertHDR() As String
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    
    Dim Id As Variant
    Dim Tin As Variant
    Dim Loai_bc As Variant
    Dim Ngay_nop As Variant
    Dim Ngay_cn As Variant
    Dim nguoi_cn As Variant
    Dim So_tt_tk As Variant
    Dim Phong_xly As Variant
    Dim Phong_qly As Variant
    Dim Nguoi_dd As Variant
    Dim Ngay_bc As Variant
    Dim Cq_nhan As Variant
    Dim ngay_Thang_Nam_Mat As Variant
    Dim DA_NHAN As Variant
    Dim ly_DO_MAT_CHAY_HONG As Variant
    Dim Ghi_chu As Variant

    'bien xu ly combobox
    Dim strComboBox As Variant
    Dim combovalue() As String
    
    Dim maCQT As Variant
    Dim bln  As Boolean
    
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
    
    sSQLCol = "Id, Tin, Loai_bc,Ngay_nop, Ngay_cn,Nguoi_cn,So_tt_tk,Phong_xly,Phong_qly,"
    sSQLCol = sSQLCol & "Ngay_bc, Nguoi_dd, ngay_mat_h,ma_cqt,ghi_chu,da_nhan, ly_do_mat "
     
    'set maCQT
    maCQT = GetMaCqt
    With fps
        .Sheet = 1
        
        Id = hdrId
        'lay stt tren to khai
        .Row = 42
        'mst
        .GetText .ColLetterToNumber(MA_SO_THUE_X), MA_SO_THUE_Y, Tin
        If Trim(Tin) = vbNullString Then
            Tin = "''"
        Else
            Tin = Trim(CStr(Tin))
            If Len(Tin) = 13 Then
                Tin = Left(CStr(Tin), 10) & "-" & Right(CStr(Tin), 3)
            End If
            Tin = "'" & Tin & "'"
        End If
        Loai_bc = "'BC21_AC_BLP'"
        
        .GetText .ColLetterToNumber("E"), 13, Ngay_nop
        'NGNOP = Date
        If Trim(Ngay_nop) = vbNullString Then
            Ngay_nop = "CTOD('')"
        Else
            Ngay_nop = ToDate(Trim(Ngay_nop), DDMMYYYY)
            Ngay_nop = "CTOD('" & Format(Ngay_nop, "mm/dd/yyyy") & "')"
        End If
        'ngay cap nhat
        .GetText .ColLetterToNumber("S"), 15, Ngay_cn
        If Trim(Ngay_cn) = vbNullString Then
            Ngay_cn = "CTOD('')"
        Else
            Ngay_cn = ToDate(Trim(Ngay_cn), DDMMYYYY)
            Ngay_cn = "CTOD('" & Format(Ngay_cn, "mm/dd/yyyy") & "')"
        End If
        ' nguoi cap nhat
        .GetText .ColLetterToNumber("AC"), 15, nguoi_cn
        If Trim(nguoi_cn) = vbNullString Then
            nguoi_cn = "''"
        Else
            nguoi_cn = TAX_Utilities_Svr_New.Convert(CStr(nguoi_cn), UNICODE, TCVN)
            nguoi_cn = "'" & nguoi_cn & "'"
        End If
        'set ly do mat chay hong
        .GetText .ColLetterToNumber("C"), 31, ly_DO_MAT_CHAY_HONG
        If ly_DO_MAT_CHAY_HONG = vbNullString Then
            ly_DO_MAT_CHAY_HONG = "''"
        Else
            ly_DO_MAT_CHAY_HONG = TAX_Utilities_Svr_New.Convert(CStr(ly_DO_MAT_CHAY_HONG), UNICODE, TCVN)
            ly_DO_MAT_CHAY_HONG = "'" & ly_DO_MAT_CHAY_HONG & "'"
        End If
        
        ' so thu tu to khai
        .GetText .ColLetterToNumber("S"), 13, So_tt_tk
        If Trim(So_tt_tk) = vbNullString Then
            So_tt_tk = 0
        End If
        
        'set PhongXuLy
        .Col = .ColLetterToNumber("AC")
        .Row = 13
        .GetText .ColLetterToNumber("AC"), 13, strComboBox
        If strComboBox = vbNullString Then
            Phong_xly = "''"
        Else
            combovalue = Split(strComboBox, "{")
            Phong_xly = Mid(combovalue(1), 1, Len(combovalue(1)) - 1)
            Phong_xly = "'" & Phong_xly & "'"
        End If
        'phong quan ly
         If Trim(strMaBPQL) = vbNullString Then
            Phong_qly = "''"
        Else
            Phong_qly = "'" & Trim(strMaBPQL) & "'"
        End If
    
        ' ngay bao cao
        .GetText .ColLetterToNumber("P"), 31, Ngay_bc
        If Trim(Ngay_bc) = vbNullString Then
            Ngay_bc = "CTOD('')"
        Else
            Ngay_bc = ToDate(Trim(Ngay_bc), DDMMYYYY)
            Ngay_bc = "CTOD('" & Format(Ngay_bc, "mm/dd/yyyy") & "')"
        End If
        'nguoi dai dien
        .GetText .ColLetterToNumber("L"), 31, Nguoi_dd
        If Trim(Nguoi_dd) = vbNullString Then
            Nguoi_dd = "''"
        Else
            Nguoi_dd = TAX_Utilities_Svr_New.Convert(CStr(Nguoi_dd), UNICODE, TCVN)
            Nguoi_dd = "'" & Nguoi_dd & "'"
        End If
        'set da nhan
        DA_NHAN = "''"
        ' ngay mat chay hong
        .GetText .ColLetterToNumber("N"), 15, ngay_Thang_Nam_Mat
        If Trim(ngay_Thang_Nam_Mat) = vbNullString Then
            ngay_Thang_Nam_Mat = "CTOD('')"
        Else
            ngay_Thang_Nam_Mat = ToDate(Trim(ngay_Thang_Nam_Mat), DDMMYYYY)
            ngay_Thang_Nam_Mat = "CTOD('" & Format(ngay_Thang_Nam_Mat, "mm/dd/yyyy") & "')"
        End If
        
        'set ghi chu
        .GetText .ColLetterToNumber("S"), 19, Ghi_chu
        
        If Ghi_chu = vbNullString Then
                Ghi_chu = "''"
            Else
                Ghi_chu = "'" & TAX_Utilities_Svr_New.Convert(CStr(Ghi_chu), UNICODE, TCVN) & "'"
        End If
     sSQLVal = Id & "," & Tin & "," & Loai_bc & "," & Ngay_nop & "," & Ngay_cn & "," & nguoi_cn & "," & So_tt_tk & "," & _
                Phong_xly & "," & Phong_qly & "," & Ngay_bc & "," & Nguoi_dd & "," & ngay_Thang_Nam_Mat & "," & maCQT & "," & Ghi_chu & "," & DA_NHAN & "," & ly_DO_MAT_CHAY_HONG
       
        sSQL = "INSERT INTO TMP_BCAO_HDR_AC( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
    End With
     
    InsertHDR = vbNullString
    clsDAO.Disconnect
End Function

Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
'     strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BCAO_HDR_AC" & ".DBF"
'     strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_DTL_BC21_AC_BLP" & ".DBF"
'    'file HDR
'    If fso.FileExists(strFileNameHDR) = False Then
'        Dim strTepmauHDR As String
'        strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_BCAO_HDR_AC.DBF"
'        fso.CopyFile strTepmauHDR, strFileNameHDR, False
'    End If
'    'file dtl
'    If fso.FileExists(strFileNameDTL) = False Then
'        Dim strTepmauDTL As String
'        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_DTL_BC21_AC_BLP.DBF"
'        fso.CopyFile strTepmauDTL, strFileNameDTL, False
'    End If
'    isExistFile = True

'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'   End If
'   sSQL = "SELECT TIN, NGAY_NOP FROM TMP_BCAO_HDR_AC" & Right(sKyKeKhai, 4) & _
'        " WHERE TIN = " & "'" & strMST & "'" & _
'        " AND LOAI_BC='01_TBAC'"
'
'   Set rs = clsDAO.Execute(sSQL)
'   If Not rs Is Nothing Then
'        dNgayNopDB = rs.Fields("NGAY_NOP")
'        If Trim(dNgayNopDB) = vbNullString Then
'            dNgayNopDB = "CTOD('')"
'        Else
'            dNgayNopDB = "CTOD('" & Format(dNgayNopDB, "mm/dd/yyyy") & "')"
'        End If
'        TKTT = True
'   Else
'        TKTT = False
'   End If
'   LAN_QUET = TinhSoLanQuet + 1
'   clsDAO.Disconnect
    isExistFile = True
    TKTT = isTKTonTai
End Function

Public Function UpdateTHUETKY2() As Boolean
    Dim sSQL As String

'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'    End If
'
'    sSQL = "UPDATE TMP_CNTK" & Right(sKyKeKhai, 4) & _
'        " SET THUETKY2 = 1 " & _
'        " WHERE MATHUE = '09' and MAPP = '03' AND MADTNT = " & "'" & strMST & "'" & _
'        " AND KYKKHAI = '" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='03Q/KK-TNCN'"
'
'    clsDAO.ExecuteDLL sSQL

    UpdateTHUETKY2 = True

 '   clsDAO.Disconnect
End Function


Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL As String
    
'
'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'   End If
'
'   sSQL = "SELECT * FROM TMP_BCAO_HDR_AC WHERE  TIN = " & "'" & strMST & "'" & _
'        " AND LOAI_BC='BC21_AC_BLP'"
'
'   Set rs = clsDAO.Execute(sSQL)
'   If rs Is Nothing Then
'        isToKhaiChinhThuc = False
'   Else
        isToKhaiChinhThuc = True
'   End If
   
   clsDAO.Disconnect
End Function


Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim KYLBO As Variant
    Dim rs As ADODB.Recordset
    
'    If clsDAO.Connected = False Then
''        Me.MousePointer = vbHourglass
''        frmSystem.MousePointer = vbHourglass
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
''        frmSystem.MousePointer = vbDefault
''        Me.MousePointer = vbDefault
'   End If
'
'   sSQL = "DELETE FROM TMP_BCAO_HDR_AC WHERE TIN = " & "'" & strMST & "'" & _
'        " AND LOAI_BC='01_TBAC'"
'
'
'   clsDAO.ExecuteDLL sSQL
'
'
'   sSQL = "DELETE FROM TMP_DTL_01_TBAC" & Right(sKyKeKhai, 4) & _
'        " WHERE TIN = " & "'" & strMST & "'" & _
'        " AND LOAI_BC='01_TBAC'"
'
'   clsDAO.ExecuteDLL sSQL
   
   XoaTKTT = True
   
 '  clsDAO.Disconnect

End Function

Public Function TKRB() As Boolean
   
   TKRB = True
'   clsDAO.Disconnect

End Function

Public Function TKTTNGNOP() As Boolean
   Dim Ngay_nop As Variant
   Dim NGNOPDB As Variant
   Dim sSQL As String
   Dim THUEGTGT_KT As Double
   Dim rs As ADODB.Recordset
   Dim i As Integer
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 13, Ngay_nop
        'NGNOP = Date
        If Trim(Ngay_nop) = vbNullString Then
            Ngay_nop = ""
        Else
            Ngay_nop = ToDate(Trim(Ngay_nop), DDMMYYYY)
            Ngay_nop = "CTOD('" & Format(Ngay_nop, "mm/dd/yyyy") & "')"
        End If
   End With
   
'   If clsDAO.Connected = False Then
''        Me.MousePointer = vbHourglass
''        frmSystem.MousePointer = vbHourglass
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
''        frmSystem.MousePointer = vbDefault
''        Me.MousePointer = vbDefault
'   End If
'
'   sSQL = "SELECT MADTNT, NGNOP FROM TMP_BCAO_HDR_AC WHERE TIN = " & "'" & strMST & "'"
'
'   Set rs = clsDAO.Execute(sSQL)
'   If Not rs Is Nothing Then
'        For i = 1 To rs.RecordCount
'            If Not rs.EOF Then
'                NGNOPDB = rs.Fields("NGAY_NOP")
'                If Trim(NGNOPDB) = vbNullString Then
'                    NGNOPDB = ""
'                Else
'                    NGNOPDB = "CTOD('" & Format(NGNOPDB, "mm/dd/yyyy") & "')"
'                End If
'                If Trim(NGNOPDB) <> vbNullString And Trim(Ngay_nop) = Trim(Ngay_nop) Then
'                    clsDAO.Disconnect
'                    TKTTNGNOP = False
'                    Exit Function
'                Else
'                    TKTTNGNOP = True
'                End If
'            End If
'            rs.MoveNext
'        Next
'   Else
'        clsDAO.Disconnect
'        TKTTNGNOP = True
'   End If
    TKTTNGNOP = False
    
End Function
Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
 Dim varTemp As Variant
 With fps
    If .ActiveSheet = 1 Then
        If Col = .ColLetterToNumber("E") And Row = 13 Then
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
        Else
            TEMP = "''"
        End If
        clsConn.Disconnect
        GetMaCqt = TEMP
End Function

