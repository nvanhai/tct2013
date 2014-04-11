VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01_TBAC_BLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const NGAY_NHAN_TO_KHAI_Y = 10
Const NGAY_NHAN_TO_KHAI_X = "E"
Const NGUOI_SU_DUNG_Y = 12
Const NGUOI_SU_DUNG_X = "Q"
Const PHONG_XU_LY_Y = 10
Const PHONG_XU_LY_X = "Q"
Private rs As ADODB.Recordset

'Longvh
Const MA_SO_THUE_X = "E"
Const MA_SO_THUE_Y = 5
Const NGAY_NOP_Y = 10
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 12
Const NGAY_QUET_X = "K"

Const TEN_GOI_X = "J"
Const TEN_GOI_Y = 4
Const DIA_CHI_X = "E"
Const DIA_CHI_Y = 6
Const TEN_BPQL_X = "Q"
Const TEN_BPQL_Y = 14
Const DIEN_THOAI_X = "E"
Const DIEN_THOAI_Y = 8
Const FAX_X = "K"
Const FAX_Y = 8
Const MAIL_X = "Q"
Const MAIL_Y = 8
Const STT_X = "K"
Const STT_Y = 10

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
                
        'Ghi chu
        .Col = .ColLetterToNumber("E")
        .Row = 14
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100
        
        'Phong xu ly
        .Col = .ColLetterToNumber("Q")
        .Row = 10
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
           
        ' Ngay nop
        SetDateFormat fps, 1, 10, .ColLetterToNumber("E"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 10
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 12, .ColLetterToNumber("K"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 12
        .Col = .ColLetterToNumber("K")
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
        .TypeComboBoxCurSel = 1
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
        
        'set phong su ly
        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y
        If strPhongXuLy <> "" Then
            .Text = strPhongXuLy
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(TEN_BPQL_X)
        .Row = TEN_BPQL_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
            
        'set nguoi su dung
        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_X)
        .Row = NGUOI_SU_DUNG_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strNguoiSuDung, TCVN, UNICODE))
        
        'set STT
        
        
        .Col = .ColLetterToNumber(STT_X)
        .Row = STT_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        
     
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
        
        .GetText .ColLetterToNumber("E"), 10, vNgayNop
        .GetText .ColLetterToNumber("K"), 12, vNgayQuet



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
   Dim Ngay_bd_sd As Variant
   Dim Ten_datin As Variant
   Dim Ngay_bl_in As Variant
   Dim So_bl_in As Variant
   Dim Mst_datin As Variant
   Dim Loai_bl As Variant
   
   Dim bln  As Boolean
  
    If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
    
   sSQLCol = "Id, Hdr_id, So_tt, Ten_bl, Mau_so, Ky_hieu_bl, So_luong, Tu_so, Den_so, Ngay_bd_sd, Ten_datin, Ngay_bl_in, So_bl_in, Mst_datin, Loai_bl "
  
   '-------------
   With fps
   .Sheet = 1
        .Col = .ColLetterToNumber("B")
        .Row = 29
            Do
                Id = GetDtlIdAC(spathVat)
                Hdr_id = hdrId
                'so thu tu
                .GetText .ColLetterToNumber("B"), .Row, So_tt
                If Trim(So_tt) = vbNullString Then
                    So_tt = "0"
                End If
                'ten bien lai
                .GetText .ColLetterToNumber("D"), .Row, Ten_bl
                If Trim(Ten_bl) = vbNullString Then
                    Ten_bl = "''"
                Else
                    Ten_bl = TAX_Utilities_Svr_New.Convert(CStr(Ten_bl), UNICODE, TCVN)
                    If Len(Ten_bl) > 200 Then
                        Ten_bl = Left(CStr(Ten_bl), 200)
                    End If
                    Ten_bl = "'" & Ten_bl & "'"
                End If
                ' mau so bien lai
                .GetText .ColLetterToNumber("I"), .Row, Mau_so
                If Trim(Mau_so) = vbNullString Then
                    Mau_so = "''"
                Else
                    Mau_so = TAX_Utilities_Svr_New.Convert(CStr(Mau_so), UNICODE, TCVN)
                    Mau_so = "'" & Mau_so & "'"
                End If
                'ky hieu bien lai
                .Col = .ColLetterToNumber("J")
                Ky_hieu_bl = .Text
                If Trim(Ky_hieu_bl) = vbNullString Then
                    Ky_hieu_bl = "''"
                Else
                    Ky_hieu_bl = "'" & Ky_hieu_bl & "'"
                End If
                'tu so
                .Col = .ColLetterToNumber("M")
                Tu_so = .Value
                If Trim(Tu_so) = vbNullString Then
                    Tu_so = "''"
                Else
                    Tu_so = "'" & Tu_so & "'"
                End If
                ' den so
                .Col = .ColLetterToNumber("O")
                Den_so = .Value
                If Trim(Den_so) = vbNullString Then
                    Den_so = "''"
                Else
                    Den_so = "'" & Den_so & "'"
                End If
                'so luong
                .Col = .ColLetterToNumber("K")
                So_luong = .Value
                If Trim(Tu_so) <> "" And Trim(Den_so) <> "" Then
                    So_luong = So_luong
                Else
                    So_luong = "0"
                End If
                ' ngay bat dau su dung
                .GetText .ColLetterToNumber("Q"), .Row, Ngay_bd_sd
                If Trim(Ngay_bd_sd) = vbNullString Then
                     Ngay_bd_sd = "CTOD('')"
                Else
                     Ngay_bd_sd = ToDate(Trim(Ngay_bd_sd), DDMMYYYY)
                     Ngay_bd_sd = "CTOD('" & Format(Ngay_bd_sd, "mm/dd/yyyy") & "')"
                End If
                ' ten doanh nghiep in
                .Col = .ColLetterToNumber("Z")
                Ten_datin = .Text
                If Trim(Ten_datin) = vbNullString Then
                    Ten_datin = "''"
                Else
                    Ten_datin = TAX_Utilities_Svr_New.Convert(CStr(Ten_datin), UNICODE, TCVN)
                    If Len(Ten_datin) > 250 Then
                        Ten_datin = Left(CStr(Ten_datin), 250)
                    End If
                    Ten_datin = "'" & Ten_datin & "'"
                End If
                'ma so thue doanh nghiep in
                .Col = .ColLetterToNumber("AC")
                Mst_datin = .Text
                If Trim(Mst_datin) = vbNullString Then
                    Mst_datin = "''"
                Else
                    Mst_datin = Trim(CStr(Mst_datin))
                    If Len(Mst_datin) = 13 Then
                        Mst_datin = Left(CStr(Mst_datin), 10) & "-" & Right(CStr(Mst_datin), 3)
                    End If
                    Mst_datin = "'" & Mst_datin & "'"
                End If
                ' so hop dong in
                .Col = .ColLetterToNumber("S")
                So_bl_in = .Text
                If Trim(So_bl_in) = vbNullString Then
                    So_bl_in = "''"
                Else
                    So_bl_in = "'" & So_bl_in & "'"
                End If
                ' ngay hop dong dat in
                .GetText .ColLetterToNumber("V"), .Row, Ngay_bl_in
                If Trim(Ngay_bl_in) = vbNullString Then
                   Ngay_bl_in = "CTOD('')"
                Else
                    Ngay_bl_in = ToDate(Trim(Ngay_bl_in), DDMMYYYY)
                    Ngay_bl_in = "CTOD('" & Format(Ngay_bl_in, "mm/dd/yyyy") & "')"
                End If
                ' loai bl
                ' truong an
                .Col = .ColLetterToNumber("AF")
                Loai_bl = .Text
                If Trim(Loai_bl) = vbNullString Then
                    Loai_bl = "''"
                Else
                    Loai_bl = "'" & Loai_bl & "'"
                End If
                '------------------
                
                sSQLVal = Id & "," & Hdr_id & "," & So_tt & "," & Ten_bl & "," & Mau_so & ","
                sSQLVal = sSQLVal & Ky_hieu_bl & "," & So_luong & "," & Tu_so & "," & Den_so & "," & Ngay_bd_sd & "," & Ten_datin & "," & Ngay_bl_in & "," & So_bl_in & "," & Mst_datin & "," & Loai_bl
            
                sSQL = "INSERT INTO TMP_DTL_01_TBAC_BLP( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                sSQL = Replace(sSQL, Chr(10), "")
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
'    Dim Da_nhan As Variant
    Dim Phong_xly As Variant
    Dim Phong_qly As Variant
    Dim Ten_dv_cq As Variant
    Dim Tin_dv_cq As Variant
    Dim Nguoi_dd As Variant
    Dim Ngay_bc As Variant
    Dim Cq_nhan As Variant
    Dim DUNG_DN_CQ As Variant
    Dim Ghi_chu As Variant
'    Dim Ma_cqt As Variant
    Dim maCQT As Variant
    Dim bln  As Boolean
    Dim combovalue() As String
    
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
    
    
    sSQLCol = "Id, Tin, Loai_bc,Ngay_nop, Ngay_cn,Nguoi_cn,So_tt_tk,da_nhan, Phong_xly,Phong_qly,"
    sSQLCol = sSQLCol & "Ten_dv_cq, Tin_dv_cq,Ngay_bc, Nguoi_dd,Cq_nhan, Ghi_chu, MA_CQT,dung_dn_cq"
    
    
    'set maCQT
    maCQT = GetMaCqt
       
    With fps
    
        .Sheet = 1
        Id = hdrId
        
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
        Loai_bc = "'01_TBAC_BLP'"
        
        .GetText .ColLetterToNumber("E"), 10, Ngay_nop
        If Trim(Ngay_nop) = vbNullString Then
            Ngay_nop = "CTOD('')"
        Else
            Ngay_nop = ToDate(Trim(Ngay_nop), DDMMYYYY)
            Ngay_nop = "CTOD('" & Format(Ngay_nop, "mm/dd/yyyy") & "')"
        End If
        'ngay cap nhat
        .GetText .ColLetterToNumber("K"), 12, Ngay_cn
        If Trim(Ngay_cn) = vbNullString Then
            Ngay_cn = "CTOD('')"
        Else
            Ngay_cn = ToDate(Trim(Ngay_cn), DDMMYYYY)
            Ngay_cn = "CTOD('" & Format(Ngay_cn, "mm/dd/yyyy") & "')"
        End If
        ' nguoi cap nhat
        .GetText .ColLetterToNumber("Q"), 12, nguoi_cn
        If Trim(nguoi_cn) = vbNullString Then
            nguoi_cn = "''"
        Else
            nguoi_cn = TAX_Utilities_Svr_New.Convert(CStr(nguoi_cn), UNICODE, TCVN)
            nguoi_cn = "'" & nguoi_cn & "'"
        End If
        ' so thu tu to khai
        .GetText .ColLetterToNumber("K"), 10, So_tt_tk
        If Trim(So_tt_tk) = vbNullString Then
            So_tt_tk = "0"
        End If
        ' phong xu ly
        .GetText .ColLetterToNumber("Q"), 10, Phong_xly
        If Trim(Phong_xly) = vbNullString Then
            Phong_xly = "''"
        Else
            combovalue = Split(Phong_xly, "{")
            Phong_xly = Mid(combovalue(1), 1, Len(combovalue(1)) - 1)
            Phong_xly = "'" & Phong_xly & "'"
        End If
        'phong quan ly
        If Trim(strMaBPQL) = vbNullString Then
            Phong_qly = "''"
        Else
            Phong_qly = "'" & Trim(strMaBPQL) & "'"
        End If
        'ten don vi chu quan
        .GetText .ColLetterToNumber("E"), 22, Ten_dv_cq
        If Trim(Ten_dv_cq) = vbNullString Then
            Ten_dv_cq = "''"
        Else
            Ten_dv_cq = TAX_Utilities_Svr_New.Convert(CStr(Ten_dv_cq), UNICODE, TCVN)
            Ten_dv_cq = "'" & Ten_dv_cq & "'"
        End If
        'mst don vi chu quan
        .GetText .ColLetterToNumber("I"), 22, Tin_dv_cq
        If Trim(Tin_dv_cq) = vbNullString Then
            Tin_dv_cq = "''"
        Else
            Tin_dv_cq = Trim(CStr(Tin_dv_cq))
            If Len(Tin_dv_cq) = 13 Then
                Tin_dv_cq = Left(CStr(Tin_dv_cq), 10) & "-" & Right(CStr(Tin_dv_cq), 3)
            End If
            Tin_dv_cq = "'" & Tin_dv_cq & "'"
        End If
        
        
           ' Lay thong tin don vi chu quan dat in, tu in
        If Trim(Tin) = Trim(Tin_dv_cq) Then
            DUNG_DN_CQ = "'N'"
        Else
            DUNG_DN_CQ = "'Y'"
        End If
        
        
        ' ngay bao cao
        .GetText .ColLetterToNumber("O"), 22, Ngay_bc
        If Trim(Ngay_bc) = vbNullString Then
            Ngay_bc = "CTOD('')"
        Else
            Ngay_bc = ToDate(Trim(Ngay_bc), DDMMYYYY)
            Ngay_bc = "CTOD('" & Format(Ngay_bc, "mm/dd/yyyy") & "')"
        End If
        'nguoi dai dien
        .GetText .ColLetterToNumber("R"), 22, Nguoi_dd
        If Trim(Nguoi_dd) = vbNullString Then
            Nguoi_dd = "''"
        Else
            Nguoi_dd = TAX_Utilities_Svr_New.Convert(CStr(Nguoi_dd), UNICODE, TCVN)
            Nguoi_dd = "'" & Nguoi_dd & "'"
        End If
        ' ten co quan tiep nhan
        .GetText .ColLetterToNumber("L"), 22, Cq_nhan
        If Trim(Cq_nhan) = vbNullString Then
            Cq_nhan = "''"
        Else
            Cq_nhan = TAX_Utilities_Svr_New.Convert(CStr(Cq_nhan), UNICODE, TCVN)
            Cq_nhan = "'" & Cq_nhan & "'"
        End If
       
        'ghi chu
        .GetText .ColLetterToNumber("E"), 14, Ghi_chu
        If Trim(Ghi_chu) = vbNullString Then
            Ghi_chu = "''"
        Else
            Ghi_chu = TAX_Utilities_Svr_New.Convert(CStr(Ghi_chu), UNICODE, TCVN)
            Ghi_chu = "'" & Ghi_chu & "'"
        End If
        
        sSQLVal = Id & "," & Tin & "," & Loai_bc & "," & Ngay_nop & "," & Ngay_cn & "," & nguoi_cn & "," & So_tt_tk & ",''," & _
                Phong_xly & "," & Phong_qly & "," & Ten_dv_cq & "," & Tin_dv_cq & "," & Ngay_bc & "," & Nguoi_dd & "," & Cq_nhan & "," & _
               Ghi_chu & "," & maCQT & "," & DUNG_DN_CQ
       
        sSQL = "INSERT INTO TMP_BCAO_HDR_AC( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
        bln = clsDAO.ExecuteDLL(sSQL)
        
    End With
     
    InsertHDR = vbNullString
    clsDAO.Disconnect
End Function

Public Function TKTT() As Boolean
 
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
'
'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'   End If
'
'   sSQL = "SELECT * FROM TMP_BCAO_HDR_AC WHERE  TIN = " & "'" & strMST & "'" & _
'        " AND LOAI_BC='01_TBAC_BLP'"
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
'        " AND LOAI_BC='01_TBAC_BLP'"
'
'
'   clsDAO.ExecuteDLL sSQL
'
'
'   sSQL = "DELETE FROM TMP_DTL_01_TBAC_BLP" & Right(sKyKeKhai, 4) & _
'        " WHERE TIN = " & "'" & strMST & "'" & _
'        " AND LOAI_BC='01_TBAC_BLP'"
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
        .GetText .ColLetterToNumber("E"), 10, Ngay_nop
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
        If Col = .ColLetterToNumber("E") And Row = 10 Then
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
    
    'kiem tra ton tai tep *.dbf chua
    'Tinh so lan quyet de ghi vao DB trung gian
    'dhdang sua
    'ngay 24/09/2010
'-----------------------------
'    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_CNTK" & Right(sKyKeKhai, 4) & ".DBF"
'    If fso.FileExists(strFileNameHDR) = True Then
'        If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'        End If
'
'             sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_BCAO_HDR_AC" & Right(sKyKeKhai, 4) & _
'                  " WHERE MADTNT = " & "'" & strMST & "'" & _
'                  " AND KYKKHAI = '" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='03Q/KK-TNCN'" & " AND (TTHTK = '1' OR TTHTK = '3' OR TTHTK = '4')"
'
'
'            Set rs = clsDAO.Execute(sSQL)
'            If Not rs Is Nothing Then
'                 LAN_QUET1 = rs.Fields("LAN_QUET")
'                 TinhSoLanQuet = LAN_QUET1
'            Else
'                 TinhSoLanQuet = 0
'            End If
'            clsDAO.Disconnect
'    Else
'    TinhSoLanQuet = 0
'    End If
End Function



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


