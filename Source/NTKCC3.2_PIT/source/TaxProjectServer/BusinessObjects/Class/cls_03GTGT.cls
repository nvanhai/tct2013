VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_03GTGT"
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

Private isTkQuy As String

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
      
'        'Ma so tep
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
        
        ' Ky lap bo
        SetDateFormat fps, 1, 30, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 32, .ColLetterToNumber("E"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
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
        
        ' Ngay hop dong
        SetDateFormat fps, 1, 28, .ColLetterToNumber("N"), DDMMYYYY
        
    End With
    
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
    Dim i As Integer
    Dim LoaiKyKK As Variant

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
        
        'set to khai thang/quy
        .GetText .ColLetterToNumber("L"), 61, LoaiKyKK

        If LoaiKyKK = "0" Then
            LOAI_TO_KHAI = "03/GTGT"
            isTkQuy = False
        Else
            LOAI_TO_KHAI = "03Q/GTGT"
            isTkQuy = True
        End If

    End With
    
    Prepared2 = True
End Function

Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String) 'ByRef strPMaSoTep, ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
    'strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

Public Function Prepared3() As Boolean
    With fps
        '.EventEnabled(EventAllEvents) = False
        'Set MaSoTep
        .Sheet = 1
'        .Col = .ColLetterToNumber(MA_SO_TEP_X)
'        .Row = MA_SO_TEP_Y
'        .Text = strMaSoTep
'        UpdateCell fps, .Col, .Row, .Text
        
        'Set NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
        .Row = NGAY_NHAN_TO_KHAI_Y
        If strNgayNhanToKhai <> "" Then
            .Text = strNgayNhanToKhai
            UpdateCell fps, .Col, .Row, .Text
        End If
        
'**************************************************************
'       ThanhDX remove
'       Date: 14/09/2006
'       Reason: Doi Nguoi su dung -> Phong quan ly.
        
'        'Set NguoiSuDung
'        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_X)
'        .Row = NGUOI_SU_DUNG_Y
'        If strNguoiSuDung <> "" Then
'            .Text = strNguoiSuDung
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'**************************************************************

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
        
'        .Col = .ColLetterToNumber(MA_BPQL_X)
'        .Row = MA_BPQL_Y
'        If strMaBPQL <> "" Then
'            .Text = strMaBPQL
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'
'        .Col = .ColLetterToNumber(TEN_BPQL_X)
'        .Row = TEN_BPQL_Y
'        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        '.EventEnabled(EventAllEvents) = True
        
                'so thu tu to khai
        .Col = .ColLetterToNumber("M")
        .Row = 30
        strSoTTTKhai = Trim(TinhSoLanQuet)
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
        .Col = .ColLetterToNumber(KY_LAP_BO_X)
        .Row = KY_LAP_BO_Y
        strKyLapBo = .Text
        
        ' Get MaSoTep
'        .Col = .ColLetterToNumber(MA_SO_TEP_X)
'        .Row = MA_SO_TEP_Y
'        strMaSoTep = .Text
        
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
        
'        If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
'            blnValid = False
'        End If
'
'        .Col = .ColLetterToNumber(HEADER_KY_LAP_BO_X)
'        .Row = HEADER_KY_LAP_BO_Y
'
'        If Not blnValid Then
'            .Formula = "0"
'        Else
'            .Formula = "1"
'        End If
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
'    If fps.ActiveCol = fps.ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) And fps.ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then _
'        CellEditFormatNumber fps, fps.ActiveSheet, fps.ActiveCol, fps.ActiveRow, KeyAscii
'    With fps
'        If .ActiveSheet = 1 And .ActiveCol = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) _
'           And .ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then
'            If KeyAscii = 46 Or KeyAscii = 45 Then ' KeyAscii of '.' character
'                KeyAscii = 0
'            End If
'        End If
'    End With
End Sub

Public Function CheckValidData() As Boolean
    Dim varCheckValue As Variant, varNoteValue As Variant
    
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 36, varCheckValue
        .GetText .ColLetterToNumber("M"), 36, varNoteValue
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



Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
    KiemTraKhoaSo = True
    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("E_30"), "Value"), "dd/mm/yyyy")
    
    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function

Public Function InsertDTL() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   
   Dim madtnt As Variant
   Dim MATHUE As Variant
   Dim maTM As Variant
   Dim MAPP As Variant
   Dim KYLBO As Variant
   Dim NGANH As Variant
   Dim ngnop As Variant
   Dim LoaiTien As Variant
   Dim TKHODKY As Variant
   Dim GTRIMV As Variant
   Dim GTRIMVGTGT As Variant
   Dim GTRIBR As Variant
   Dim GTTT As Variant
   Dim THUE As Variant
   Dim THUESUAT As Variant
   Dim TYLE As Variant
   Dim MaMuc As Variant
   
   Dim CTIEUTNDN As Variant
   Dim SOTIEN As Variant
   Dim SOTIEN2 As Variant
   Dim tt As Variant

   Dim matkhai As Variant
   Dim kykkhai As Variant
   Dim DANHAN As Variant
   Dim THUECL As Variant
   
   Dim bln  As Boolean
   'tinh so lan quet
    'vttoan
    'ngay 28/09/10
    '------------------------------------------------
     
  
    
     

   If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
    
   sSQLCol = "MADTNT , MATHUE, MATM, MAPP, KYLBO, NGNOP, LOAITIEN, "
   sSQLCol = sSQLCol + "CTIEUTNDN, SOTIEN, TT, SOTIEN2, MAMUC ,TTHTK, MATKHAI, KYKKHAI, NGANH, LAN_QUET, DANHAN,THUECL"
     
   THUECL = "0"
   DANHAN = "''"
   NGANH = "''"
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        MATHUE = "'01'"
        MAPP = "'2'"
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
        
      '  If Trim(strNganh) = vbNullString Then
      '      NGANH = "''"
      '  Else
      '      NGANH = "'" & strNganh & "'"
      '  End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        LoaiTien = "''"
        
        'Chi tieu
        '
        CTIEUTNDN = "'001'"
        .GetText .ColLetterToNumber("Q"), 42, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
         End If
        
        SOTIEN2 = SOTIEN
        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
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
       
        
        
   End With
   sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LoaiTien & ","
   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   '--------------------------------
   
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        MATHUE = "'01'"
        MAPP = "'2'"
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
        
      '  If Trim(strNganh) = vbNullString Then
      '      NGANH = "''"
      '  Else
      '      NGANH = "'" & strNganh & "'"
      '  End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        LoaiTien = "''"
        
        'Chi tieu
        '
        CTIEUTNDN = "'002'"
        .GetText .ColLetterToNumber("Q"), 43, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
         End If
        
        SOTIEN2 = SOTIEN
        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
        
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
   End With
    
   sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LoaiTien & ","
   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   '--------------------------------
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        MATHUE = "'01'"
        MAPP = "'2'"
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
        
      '  If Trim(strNganh) = vbNullString Then
      '      NGANH = "''"
      '  Else
      '      NGANH = "'" & strNganh & "'"
      '  End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        LoaiTien = "''"
        
        'Chi tieu
        '
        CTIEUTNDN = "'003'"
        .GetText .ColLetterToNumber("Q"), 44, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
         End If
        
        SOTIEN2 = SOTIEN
        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
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
   End With
    
   sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LoaiTien & ","
   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   '--------------------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        MATHUE = "'01'"
        MAPP = "'2'"
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
        
      '  If Trim(strNganh) = vbNullString Then
      '      NGANH = "''"
      '  Else
      '      NGANH = "'" & strNganh & "'"
      '  End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        LoaiTien = "''"
        
        'Chi tieu
        '
        CTIEUTNDN = "'004'"
        .GetText .ColLetterToNumber("Q"), 45, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
         End If
        
        SOTIEN2 = SOTIEN
        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
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
   End With
    
   sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LoaiTien & ","
   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   '--------------------------------

   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        MATHUE = "'01'"
        MAPP = "'2'"
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
        'Chi tieu
        CTIEUTNDN = "'005'"
        .GetText .ColLetterToNumber("Q"), 46, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
         End If
        SOTIEN2 = SOTIEN
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
   End With
   sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LoaiTien & ","
   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
    '-------------------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        MATHUE = "'01'"
        MAPP = "'2'"
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
        
      '  If Trim(strNganh) = vbNullString Then
      '      NGANH = "''"
      '  Else
      '      NGANH = "'" & strNganh & "'"
      '  End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        LoaiTien = "''"
        
        'Chi tieu
        '
        CTIEUTNDN = "'006'"
        .GetText .ColLetterToNumber("Q"), 47, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
         End If
        
        SOTIEN2 = SOTIEN
        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
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
   End With
    
   sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LoaiTien & ","
   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   '--------------------------------
   
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        MATHUE = "'01'"
        MAPP = "'2'"
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
        
      '  If Trim(strNganh) = vbNullString Then
      '      NGANH = "''"
      '  Else
      '      NGANH = "'" & strNganh & "'"
      '  End If
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        LoaiTien = "''"
        
        'Chi tieu
        '
        CTIEUTNDN = "'007'"
        .GetText .ColLetterToNumber("Q"), 48, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
         End If
        
        SOTIEN2 = SOTIEN
        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
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
   End With
    
   sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LoaiTien & ","
   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
'   With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("F"), 10, MADTNT
'        If Trim(MADTNT) = vbNullString Then
'            MADTNT = "''"
'        Else
'            MADTNT = "'" & MADTNT & "'"
'        End If
'        MATHUE = "'01'"
'        MAPP = "'2'"
'        .GetText .ColLetterToNumber("E"), 30, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "''"
'        Else
'            KYLBO = "'" & KYLBO & "'"
'        End If
'
'        .GetText .ColLetterToNumber("E"), 32, NGNOP
'
'        If Trim(NGNOP) = vbNullString Then
'            NGNOP = "CTOD('')"
'        Else
'            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
'            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
'        End If
'        LOAITIEN = "''"
'
'        'Chi tieu
'        CTIEUTNDN = "'028'"
'        SOTIEN = "0"
'        SOTIEN2 = SOTIEN
'        TT = "''"
'        ' sua MAMUC & MATM tu nam 2009
'        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
'            MATM = "'01'"
'            MAMUC = "'014'"
'        Else
'            MAMUC = "'1700'"
'            MATM = "'1701'"
'        End If
'
'
'   End With
'   sSQLVal = MADTNT & "," & MATHUE & "," & MATM & "," & MAPP & ","
'   sSQLVal = sSQLVal & KYLBO & "," & NGNOP & "," & LOAITIEN & ","
'   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & TT & "," & SOTIEN2 & ","
'   sSQLVal = sSQLVal & MAMUC & "," & TTHTK & "," & MATKHAI & "," & KYKKHAI & "," & NGANH & "," & LAN_QUET & "," & DANHAN
'   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
'     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'   bln = clsDAO.ExecuteDLL(sSQL)
   
   '--------------------------------
   
   '--------------------------------
'    With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("F"), 10, madtnt
'        If Trim(madtnt) = vbNullString Then
'            madtnt = "''"
'        Else
'            madtnt = "'" & madtnt & "'"
'        End If
'        MATHUE = "'01'"
'        MAPP = "'2'"
'        .GetText .ColLetterToNumber("E"), 30, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "''"
'        Else
'            If Len(Trim(KYLBO)) = 6 Then
'                KYLBO = "'0" & KYLBO & "'"
'            Else
'                KYLBO = "'" & KYLBO & "'"
'            End If
'
'        End If
'
'      '  If Trim(strNganh) = vbNullString Then
'      '      NGANH = "''"
'      '  Else
'      '      NGANH = "'" & strNganh & "'"
'      '  End If
'
'        .GetText .ColLetterToNumber("E"), 32, ngnop
'        'NGNOP = Date
'        If Trim(ngnop) = vbNullString Then
'            ngnop = "CTOD('')"
'        Else
'            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
'            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
'        End If
'        LOAITIEN = "''"
'
'        'Chi tieu
'        '
'        CTIEUTNDN = "'066'"
'        .GetText .ColLetterToNumber("S"), 42, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'         End If
'
'        SOTIEN2 = SOTIEN
'        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
'        ' sua MAMUC & MATM tu nam 2009
'        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
'            MATM = "'01'"
'            MaMuc = "'014'"
'        Else
'            MaMuc = "'1700'"
'            MATM = "'1701'"
'        End If
'
'        MATKHAI = "'" & LOAI_TO_KHAI & "'"
'        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'         If isTkQuy = True Then
'            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
'        Else
'            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'        End If
'   End With
'
'   sSQLVal = madtnt & "," & MATHUE & "," & MATM & "," & MAPP & ","
'   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LOAITIEN & ","
'   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
'   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & MATKHAI & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
'   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
'     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'   bln = clsDAO.ExecuteDLL(sSQL)
'
'   '--------------------------------
'   With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("F"), 10, madtnt
'        If Trim(madtnt) = vbNullString Then
'            madtnt = "''"
'        Else
'            madtnt = "'" & madtnt & "'"
'        End If
'        MATHUE = "'01'"
'        MAPP = "'2'"
'        .GetText .ColLetterToNumber("E"), 30, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "''"
'        Else
'            If Len(Trim(KYLBO)) = 6 Then
'                KYLBO = "'0" & KYLBO & "'"
'            Else
'                KYLBO = "'" & KYLBO & "'"
'            End If
'
'        End If
'
'      '  If Trim(strNganh) = vbNullString Then
'      '      NGANH = "''"
'      '  Else
'      '      NGANH = "'" & strNganh & "'"
'      '  End If
'
'        .GetText .ColLetterToNumber("E"), 32, ngnop
'        'NGNOP = Date
'        If Trim(ngnop) = vbNullString Then
'            ngnop = "CTOD('')"
'        Else
'            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
'            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
'        End If
'        LOAITIEN = "''"
'
'        'Chi tieu
'        '
'        CTIEUTNDN = "'068'"
'        .GetText .ColLetterToNumber("S"), 43, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'         End If
'
'        SOTIEN2 = SOTIEN
'        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
'        ' sua MAMUC & MATM tu nam 2009
'        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
'            MATM = "'01'"
'            MaMuc = "'014'"
'        Else
'            MaMuc = "'1700'"
'            MATM = "'1701'"
'        End If
'        MATKHAI = "'" & LOAI_TO_KHAI & "'"
'        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'         If isTkQuy = True Then
'            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
'        Else
'            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'        End If
'   End With
'
'   sSQLVal = madtnt & "," & MATHUE & "," & MATM & "," & MAPP & ","
'   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LOAITIEN & ","
'   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
'   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & MATKHAI & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
'   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
'     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'   bln = clsDAO.ExecuteDLL(sSQL)
'
'   '--------------------------------
'
'
'   With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("F"), 10, madtnt
'        If Trim(madtnt) = vbNullString Then
'            madtnt = "''"
'        Else
'            madtnt = "'" & madtnt & "'"
'        End If
'        MATHUE = "'01'"
'        MAPP = "'2'"
'        .GetText .ColLetterToNumber("E"), 30, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "''"
'        Else
'            If Len(Trim(KYLBO)) = 6 Then
'                KYLBO = "'0" & KYLBO & "'"
'            Else
'                KYLBO = "'" & KYLBO & "'"
'            End If
'
'        End If
'
'        .GetText .ColLetterToNumber("E"), 32, ngnop
'
'        If Trim(ngnop) = vbNullString Then
'            ngnop = "CTOD('')"
'        Else
'            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
'            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
'        End If
'        LOAITIEN = "''"
'        DANHAN = "''"
'        'Chi tieu
'        CTIEUTNDN = "'070'"
'        SOTIEN = "0"
'        SOTIEN2 = SOTIEN
'        tt = "''"
'        ' sua MAMUC & MATM tu nam 2009
'        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
'            MATM = "'01'"
'            MaMuc = "'014'"
'        Else
'            MaMuc = "'1700'"
'            MATM = "'1701'"
'        End If
'        MATKHAI = "'" & LOAI_TO_KHAI & "'"
'        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'         If isTkQuy = True Then
'            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
'        Else
'            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'        End If
'        NGANH = "''"
'   End With
'   sSQLVal = madtnt & "," & MATHUE & "," & MATM & "," & MAPP & ","
'   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LOAITIEN & ","
'   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
'   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & MATKHAI & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
'   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
'     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'   bln = clsDAO.ExecuteDLL(sSQL)
'
'    '-------------------------------
'
'
'
'      With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("F"), 10, madtnt
'        If Trim(madtnt) = vbNullString Then
'            madtnt = "''"
'        Else
'            madtnt = "'" & madtnt & "'"
'        End If
'        MATHUE = "'01'"
'        MAPP = "'2'"
'        .GetText .ColLetterToNumber("E"), 30, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "''"
'        Else
'            If Len(Trim(KYLBO)) = 6 Then
'                KYLBO = "'0" & KYLBO & "'"
'            Else
'                KYLBO = "'" & KYLBO & "'"
'            End If
'
'        End If
'
'      '  If Trim(strNganh) = vbNullString Then
'      '      NGANH = "''"
'      '  Else
'      '      NGANH = "'" & strNganh & "'"
'      '  End If
'
'        .GetText .ColLetterToNumber("E"), 32, ngnop
'        'NGNOP = Date
'        If Trim(ngnop) = vbNullString Then
'            ngnop = "CTOD('')"
'        Else
'            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
'            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
'        End If
'        LOAITIEN = "''"
'
'        'Chi tieu
'        '
'        CTIEUTNDN = "'072'"
'        .GetText .ColLetterToNumber("X"), 40, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'         End If
'
'        SOTIEN2 = SOTIEN
'        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
'        ' sua MAMUC & MATM tu nam 2009
'        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
'            MATM = "'01'"
'            MaMuc = "'014'"
'        Else
'            MaMuc = "'1700'"
'            MATM = "'1701'"
'        End If
'        MATKHAI = "'" & LOAI_TO_KHAI & "'"
'        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'         If isTkQuy = True Then
'            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
'        Else
'            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'        End If
'   End With
'
'   sSQLVal = madtnt & "," & MATHUE & "," & MATM & "," & MAPP & ","
'   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LOAITIEN & ","
'   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
'   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & MATKHAI & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
'   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
'     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'   bln = clsDAO.ExecuteDLL(sSQL)
'
'   '--------------------------------
'      With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("F"), 10, madtnt
'        If Trim(madtnt) = vbNullString Then
'            madtnt = "''"
'        Else
'            madtnt = "'" & madtnt & "'"
'        End If
'        MATHUE = "'01'"
'        MAPP = "'2'"
'        .GetText .ColLetterToNumber("E"), 30, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "''"
'        Else
'            If Len(Trim(KYLBO)) = 6 Then
'                KYLBO = "'0" & KYLBO & "'"
'            Else
'                KYLBO = "'" & KYLBO & "'"
'            End If
'
'        End If
'
'      '  If Trim(strNganh) = vbNullString Then
'      '      NGANH = "''"
'      '  Else
'      '      NGANH = "'" & strNganh & "'"
'      '  End If
'
'        .GetText .ColLetterToNumber("E"), 32, ngnop
'        'NGNOP = Date
'        If Trim(ngnop) = vbNullString Then
'            ngnop = "CTOD('')"
'        Else
'            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
'            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
'        End If
'        LOAITIEN = "''"
'
'        'Chi tieu
'        '
'        CTIEUTNDN = "'074'"
'        .GetText .ColLetterToNumber("X"), 41, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'         End If
'
'        SOTIEN2 = SOTIEN
'        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
'        ' sua MAMUC & MATM tu nam 2009
'        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
'            MATM = "'01'"
'            MaMuc = "'014'"
'        Else
'            MaMuc = "'1700'"
'            MATM = "'1701'"
'        End If
'        MATKHAI = "'" & LOAI_TO_KHAI & "'"
'        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'         If isTkQuy = True Then
'            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
'        Else
'            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'        End If
'   End With
'
'   sSQLVal = madtnt & "," & MATHUE & "," & MATM & "," & MAPP & ","
'   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LOAITIEN & ","
'   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
'   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & MATKHAI & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
'   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
'     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'   bln = clsDAO.ExecuteDLL(sSQL)
'
'   '--------------------------------
'      With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("F"), 10, madtnt
'        If Trim(madtnt) = vbNullString Then
'            madtnt = "''"
'        Else
'            madtnt = "'" & madtnt & "'"
'        End If
'        MATHUE = "'01'"
'        MAPP = "'2'"
'        .GetText .ColLetterToNumber("E"), 30, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "''"
'        Else
'            If Len(Trim(KYLBO)) = 6 Then
'                KYLBO = "'0" & KYLBO & "'"
'            Else
'                KYLBO = "'" & KYLBO & "'"
'            End If
'
'        End If
'
'      '  If Trim(strNganh) = vbNullString Then
'      '      NGANH = "''"
'      '  Else
'      '      NGANH = "'" & strNganh & "'"
'      '  End If
'
'        .GetText .ColLetterToNumber("E"), 32, ngnop
'        'NGNOP = Date
'        If Trim(ngnop) = vbNullString Then
'            ngnop = "CTOD('')"
'        Else
'            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
'            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
'        End If
'        LOAITIEN = "''"
'
'        'Chi tieu
'        '
'        CTIEUTNDN = "'076'"
'        .GetText .ColLetterToNumber("X"), 42, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'         End If
'
'        SOTIEN2 = SOTIEN
'        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
'        ' sua MAMUC & MATM tu nam 2009
'        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
'            MATM = "'01'"
'            MaMuc = "'014'"
'        Else
'            MaMuc = "'1700'"
'            MATM = "'1701'"
'        End If
'        MATKHAI = "'" & LOAI_TO_KHAI & "'"
'        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'         If isTkQuy = True Then
'            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
'        Else
'            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'        End If
'   End With
'
'   sSQLVal = madtnt & "," & MATHUE & "," & MATM & "," & MAPP & ","
'   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LOAITIEN & ","
'   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
'   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & MATKHAI & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
'   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
'     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'   bln = clsDAO.ExecuteDLL(sSQL)
'
'   '--------------------------------
'      With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("F"), 10, madtnt
'        If Trim(madtnt) = vbNullString Then
'            madtnt = "''"
'        Else
'            madtnt = "'" & madtnt & "'"
'        End If
'        MATHUE = "'01'"
'        MAPP = "'2'"
'        .GetText .ColLetterToNumber("E"), 30, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "''"
'        Else
'            If Len(Trim(KYLBO)) = 6 Then
'                KYLBO = "'0" & KYLBO & "'"
'            Else
'                KYLBO = "'" & KYLBO & "'"
'            End If
'
'        End If
'
'      '  If Trim(strNganh) = vbNullString Then
'      '      NGANH = "''"
'      '  Else
'      '      NGANH = "'" & strNganh & "'"
'      '  End If
'
'        .GetText .ColLetterToNumber("E"), 32, ngnop
'        'NGNOP = Date
'        If Trim(ngnop) = vbNullString Then
'            ngnop = "CTOD('')"
'        Else
'            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
'            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
'        End If
'        LOAITIEN = "''"
'
'        'Chi tieu
'        '
'        CTIEUTNDN = "'078'"
'        .GetText .ColLetterToNumber("X"), 43, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'         End If
'
'        SOTIEN2 = SOTIEN
'        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
'        ' sua MAMUC & MATM tu nam 2009
'        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
'            MATM = "'01'"
'            MaMuc = "'014'"
'        Else
'            MaMuc = "'1700'"
'            MATM = "'1701'"
'        End If
'        MATKHAI = "'" & LOAI_TO_KHAI & "'"
'        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'         If isTkQuy = True Then
'            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
'        Else
'            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'        End If
'   End With
'
'   sSQLVal = madtnt & "," & MATHUE & "," & MATM & "," & MAPP & ","
'   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LOAITIEN & ","
'   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
'   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & MATKHAI & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
'   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
'     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'   bln = clsDAO.ExecuteDLL(sSQL)
'
'   '--------------------------------
'      With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("F"), 10, madtnt
'        If Trim(madtnt) = vbNullString Then
'            madtnt = "''"
'        Else
'            madtnt = "'" & madtnt & "'"
'        End If
'        MATHUE = "'01'"
'        MAPP = "'2'"
'        .GetText .ColLetterToNumber("E"), 30, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "''"
'        Else
'            If Len(Trim(KYLBO)) = 6 Then
'                KYLBO = "'0" & KYLBO & "'"
'            Else
'                KYLBO = "'" & KYLBO & "'"
'            End If
'
'        End If
'
'      '  If Trim(strNganh) = vbNullString Then
'      '      NGANH = "''"
'      '  Else
'      '      NGANH = "'" & strNganh & "'"
'      '  End If
'
'        .GetText .ColLetterToNumber("E"), 32, ngnop
'        'NGNOP = Date
'        If Trim(ngnop) = vbNullString Then
'            ngnop = "CTOD('')"
'        Else
'            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
'            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
'        End If
'        LOAITIEN = "''"
'
'        'Chi tieu
'        '
'        CTIEUTNDN = "'080'"
'        .GetText .ColLetterToNumber("R"), 45, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'         End If
'
'        SOTIEN2 = SOTIEN
'        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
'        ' sua MAMUC & MATM tu nam 2009
'        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
'            MATM = "'01'"
'            MaMuc = "'014'"
'        Else
'            MaMuc = "'1700'"
'            MATM = "'1701'"
'        End If
'        MATKHAI = "'" & LOAI_TO_KHAI & "'"
'        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'         If isTkQuy = True Then
'            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
'        Else
'            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'        End If
'   End With
'
'   sSQLVal = madtnt & "," & MATHUE & "," & MATM & "," & MAPP & ","
'   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LOAITIEN & ","
'   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
'   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & MATKHAI & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
'   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
'     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'   bln = clsDAO.ExecuteDLL(sSQL)
   
   '--------------------------------

'      With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("F"), 10, madtnt
'        If Trim(madtnt) = vbNullString Then
'            madtnt = "''"
'        Else
'            madtnt = "'" & madtnt & "'"
'        End If
'        MATHUE = "'01'"
'        MAPP = "'2'"
'        .GetText .ColLetterToNumber("E"), 30, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "''"
'        Else
'            If Len(Trim(KYLBO)) = 6 Then
'                KYLBO = "'0" & KYLBO & "'"
'            Else
'                KYLBO = "'" & KYLBO & "'"
'            End If
'
'        End If
'
'      '  If Trim(strNganh) = vbNullString Then
'      '      NGANH = "''"
'      '  Else
'      '      NGANH = "'" & strNganh & "'"
'      '  End If
'
'        .GetText .ColLetterToNumber("E"), 32, ngnop
'        'NGNOP = Date
'        If Trim(ngnop) = vbNullString Then
'            ngnop = "CTOD('')"
'        Else
'            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
'            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
'        End If
'        LOAITIEN = "''"
'
'        'Chi tieu
'        '
'        CTIEUTNDN = "'082'"
'        .GetText .ColLetterToNumber("R"), 47, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'         End If
'
'        SOTIEN2 = SOTIEN
'        tt = IIf(Trim(sSaiCT11) = vbNullString, "''", sSaiCT11)
'        ' sua MAMUC & MATM tu nam 2009
'        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
'            MATM = "'01'"
'            MaMuc = "'014'"
'        Else
'            MaMuc = "'1700'"
'            MATM = "'1701'"
'        End If
'        MATKHAI = "'" & LOAI_TO_KHAI & "'"
'        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'         If isTkQuy = True Then
'            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
'        Else
'            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'        End If
'   End With
'
'   sSQLVal = madtnt & "," & MATHUE & "," & MATM & "," & MAPP & ","
'   sSQLVal = sSQLVal & KYLBO & "," & ngnop & "," & LOAITIEN & ","
'   sSQLVal = sSQLVal & CTIEUTNDN & "," & SOTIEN & "," & tt & "," & SOTIEN2 & ","
'   sSQLVal = sSQLVal & MaMuc & "," & TTHTK & "," & MATKHAI & "," & kykkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & "," & THUECL
'   sSQL = "INSERT INTO TMP_DG" & sKyKeKhai & _
'     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'   bln = clsDAO.ExecuteDLL(sSQL)
   
   '--------------------------------

   clsDAO.Disconnect
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
    'SONGAYNC,SOTIENNC
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
        SOTIENNC = "0"
        SONGAYNC = "0"
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
            ElseIf MACT2 = "14" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "16" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "23" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "24" Then
                MACT2 = "'050'"
            ElseIf MACT2 = "21" Then
                MACT2 = "'039'"
            ElseIf MACT2 = "31" Then
                MACT2 = "'031'"
            ElseIf MACT2 = "32" Then
                MACT2 = "'032'"
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
            ElseIf MACT2 = "14" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "16" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "23" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "24" Then
                MACT2 = "'050'"
            ElseIf MACT2 = "21" Then
                MACT2 = "'039'"
            ElseIf MACT2 = "31" Then
                MACT2 = "'031'"
            ElseIf MACT2 = "32" Then
                MACT2 = "'032'"
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
        ElseIf MACT2 = "14" Then
            MACT2 = "'041'"
        ElseIf MACT2 = "16" Then
            MACT2 = "'041'"
        ElseIf MACT2 = "23" Then
            MACT2 = "'041'"
        ElseIf MACT2 = "24" Then
            MACT2 = "'050'"
        ElseIf MACT2 = "21" Then
            MACT2 = "'039'"
        ElseIf MACT2 = "31" Then
            MACT2 = "'031'"
        ElseIf MACT2 = "32" Then
            MACT2 = "'032'"
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

Public Function InsertHDR() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    
    Dim LoaiKyKK   As Variant
    Dim TENNV      As Variant
    Dim CHUNGCHI   As Variant
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
    Dim MuaVao     As Variant
    'bien them trong tt28
    Dim vTT28      As Variant
    'end
    Dim BanRa      As Variant
    Dim MaVach     As Variant
    Dim matkhai    As Variant
    Dim ThueTKy2   As Variant
    Dim HANNOP2    As Variant
    Dim TENDA      As Variant
    Dim NGANH      As Variant
    Dim DAGHI1     As Variant
    Dim AUTO_CAL   As Variant
    Dim DANHAN     As Variant
    
    'MADLT
    Dim madlt      As Variant
    Dim LanBS      As Variant
    Dim THUECL     As Variant
    Dim DSoCL      As Variant
    Dim GTTTCL     As Variant
    Dim NgayPs     As Variant
    Dim MANNKD     As Variant

    sSQLCol = "MADTNT, MATHUE, MATM, MAPP, KYLBO, NGNOP, "
    sSQLCol = sSQLCol & "LOAITIEN, MABPQL, SHTEP, TTHTK, KYKKHAI, "
    sSQLCol = sSQLCol & "NGNHAP, THUETKY, GTTINHTHUE, DAGHI, "
    sSQLCol = sSQLCol & "KTTRUOC, DOANHSO, MAPP2, MAMUC, MAVACH, MATKHAI, HANNOP, THUETKY2, HANNOP2,TENDA, NGANH,DAGHI1, DAGHI2,AUTO_CAL, LAN_QUET, DANHAN,MADLT ,LANBS,THUECL,DSOCL,GTTTCL,NGAYPS,MANNKD,TENNV,CHUNGCHI"
    NGANH = TinhMaNganh
    
    With fps
        .Sheet = 1
        'lay ten nhan vien dai ly thue, chung chi hanh nghe
        .GetText .ColLetterToNumber("H"), 56, TENNV
        'nshung- bo sung
        If (Trim(TENNV) <> vbNullString) Then
            TENNV = TAX_Utilities_Svr_New.Convert(Trim(TENNV), UNICODE, TCVN)
        End If
        
        .GetText .ColLetterToNumber("H"), 58, CHUNGCHI
        'nshung- bo sung
        If (Trim(CHUNGCHI) <> vbNullString) Then
            CHUNGCHI = TAX_Utilities_Svr_New.Convert(Trim(CHUNGCHI), UNICODE, TCVN)
        End If
        
        'set DAGHI2 theo chk loi dinh danh
        .Col = .ColLetterToNumber("E")
        .Row = 36

        If .Value = 1 Then
            DAGHI2 = ".T."
        Else
            DAGHI2 = ".F."
        End If
        
        'set cac chi tieu
        DSoCL = "0"
        GTTTCL = "0"
        NgayPs = "CTOD('')"
        MANNKD = "''"
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'2'"
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
        
        ' Lay loai ky ke khai(thang/quy)
        .GetText .ColLetterToNumber("L"), 61, LoaiKyKK

        If LoaiKyKK = "0" Then
            LoaiKyKK = ".T."
            isTkQuy = False
        Else
            LoaiKyKK = ".F."
            isTkQuy = True
        End If

        '        .GetText .ColLetterToNumber("M"), 10, SHTEP
        If Trim(SHTEP) = vbNullString Then
            SHTEP = "''"
        Else
            SHTEP = "'" & TAX_Utilities_Svr_New.Convert(Trim(SHTEP), UNICODE, TCVN) & "'"
        End If

        'kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

        NgNhap = Date

        If Trim(NgNhap) = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If
        
        .GetText .ColLetterToNumber("Q"), 48, ThueTKy

        If Trim(ThueTKy) = vbNullString Then
            ThueTKy = "0"
        End If

        .GetText .ColLetterToNumber("Q"), 47, GTTinhThue

        If Trim(GTTinhThue) = vbNullString Then
            GTTinhThue = "0"
        End If
        
        DaGhi = ".F."
        DAGHI1 = ".F."
        AUTO_CAL = ".F."
        KTTRUOC = 0
        DANHAN = "''"
        
        .GetText .ColLetterToNumber("Q"), 43, DoanhSo

        If Trim(DoanhSo) = vbNullString Then
            DoanhSo = "0"
        End If
        MAPP2 = "'2'"

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
        
        '       lay MADLT
        .GetText .ColLetterToNumber("F"), 20, madlt

        If Trim(madlt) = vbNullString Then
            madlt = "''"
        Else
            madlt = "'" & madlt & "'"
        End If

        '       lay so lan bs
        .GetText .ColLetterToNumber("S"), 6, LanBS

        If Trim(LanBS) = vbNullString Then
            LanBS = "0"
        End If

        '       .GetText .ColLetterToNumber("M"), 30, LAN_QUET
        If Trim(LAN_QUET) = vbNullString Then
            LAN_QUET = "1"
        End If

        DANHAN = "''"
        THUECL = "0"
    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & ngnop & ", "
    sSQLVal = sSQLVal & LoaiTien & "," & MABPQL & "," & SHTEP & "," & TTHTK & "," & kykkhai & ", "
    sSQLVal = sSQLVal & NgNhap & "," & ThueTKy & "," & GTTinhThue & "," & DaGhi & ", "
    sSQLVal = sSQLVal & KTTRUOC & "," & DoanhSo & "," & MAPP2 & "," & MaMuc & "," & MaVach & ","
    sSQLVal = sSQLVal & matkhai & "," & HANNOP & "," & ThueTKy2 & "," & HANNOP2 & "," & TENDA & ",'" & NGANH & "'," & DAGHI1 & "," & DAGHI2 & "," & AUTO_CAL & "," & LAN_QUET & "," & DANHAN & "," & madlt & "," & LanBS & "," & THUECL & "," & DSoCL & "," & GTTTCL & "," & NgayPs & "," & MANNKD & ",'" & TENNV & "','" & CHUNGCHI & "'"
    
    sSQL = "INSERT INTO TMP_TK" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    InsertHDR = sSQL
End Function


Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim strFileNameIDX As String
    
'    'kiem tra ton tai tep *.dbf chua
'    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TK" & sKyKeKhai & ".DBF"
'    If fso.FileExists(strFileNameHDR) = False Then
''    Dim strTepmauHDR As String
''        strTepmauHDR = spathVat & "\tepmau\TK.DBF"
''        fso.CopyFile strTepmauHDR, strFileNameHDR, False
'        isExistFile = False
'        DisplayMessage "0109", msOKOnly, miCriticalError
'        Exit Function
'    Else
'        isExistFile = True
'    End If

    'Ghi vao DB trung gian kiem tra file .dbf
    'vttoan  sua
    'ngay 28/09/2010
    '-----------------------------
        strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
        strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_DG" & sKyKeKhai & ".DBF"
        Dim strFileNameBS As String
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
            strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_DG.DBF"
            fso.CopyFile strTepmauDTL, strFileNameDTL, False
        End If
        'file bo xung
        If fso.FileExists(strFileNameBS) = False Then
            Dim strTepmauBS As String
            strTepmauBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_BS.DBF"
            fso.CopyFile strTepmauBS, strFileNameBS, False
        End If

        isExistFile = True

'
'    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "DG" & sKyKeKhai & ".DBF"
'    If fso.FileExists(strFileNameDTL) = False Then
'    Dim strTepmauDTL As String
'        strTepmauDTL = spathVat & "\tepmau\DG.DBF"
'        fso.CopyFile strTepmauDTL, strFileNameDTL, False
'    End If
'
''    'Xoa file *.IDX
'    strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TK" & sKyKeKhai & ".IDX"
'    If fso.FileExists(strFileNameIDX) = True Then
'        fso.DeleteFile strFileNameIDX, True
'    End If
'
'        strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "DG" & sKyKeKhai & ".IDX"
'    If fso.FileExists(strFileNameIDX) = True Then
'        fso.DeleteFile strFileNameIDX, True
'    End If
'
'        strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "BS" & sKyKeKhai & ".IDX"
'    If fso.FileExists(strFileNameIDX) = True Then
'        fso.DeleteFile strFileNameIDX, True
'    End If
'
'
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
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyKeKhai & _
        " WHERE MAPP = '2' AND MATHUE = '01' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & vKyKK & "'" & _
        " AND TTHTK = '1'" & _
        " AND MATKHAI = '" & LOAI_TO_KHAI & "'"
        
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
   
   Dim vKyKK As String
    If isTkQuy = True Then
        vKyKK = Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year
    Else
        vKyKK = TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
    End If
   
   sSQL = "UPDATE TMP_TK" & sKyKeKhai & _
        " SET THUETKY2=1" & _
        " WHERE MAPP = '2' AND  MATHUE = '01' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & vKyKK & "'" & _
        " AND MATKHAI = '" & LOAI_TO_KHAI & "'"
   clsDAO.ExecuteDLL sSQL
   
   UpdateTHUETKY2 = True
   
   clsDAO.Disconnect

End Function

Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim KYLBO As Variant
    Dim rs As ADODB.Recordset
    
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 10, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                If Len(Trim(KYLBO)) = 6 Then
                    KYLBO = "'0" & KYLBO & "'"
                Else
                    KYLBO = "'" & KYLBO & "'"
                End If

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
   
   Dim vKyKK As String
    If isTkQuy = True Then
        vKyKK = Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year
    Else
        vKyKK = TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
    End If
   sSQL = "DELETE FROM TMP_TK" & sKyKeKhai & _
         " WHERE MAPP = '2' AND  MATHUE = '01' AND MADTNT = " & "'" & strMST & "'" & _
         " AND KYKKHAI = '" & vKyKK & "'" & _
         " AND MATKHAI = '" & LOAI_TO_KHAI & "' AND KYLBO = " & KYLBO
    
   clsDAO.ExecuteDLL sSQL
   
   
   sSQL = "DELETE FROM TMP_DG" & sKyKeKhai & _
        " WHERE MAPP = '2' AND  MATHUE = '01' AND MADTNT = " & "'" & strMST & "' AND KYLBO = " & KYLBO & _
        " AND MATKHAI='" & LOAI_TO_KHAI & "' AND KYKKHAI= '" & vKyKK & "'"
    
   clsDAO.ExecuteDLL sSQL
   
   XoaTKTT = True
   
   clsDAO.Disconnect

End Function

Public Function TKRB() As Boolean

   TKRB = True

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
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
   Dim vKyKK As String
    If isTkQuy = True Then
        vKyKK = Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year
    Else
        vKyKK = TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
    End If
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyKeKhai & _
        " WHERE MAPP = '2' AND  MATHUE = '01' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI <> '" & vKyKK & "'" & _
        " AND MATKHAI='" & LOAI_TO_KHAI & "'"
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
 Dim vNgayNop As Variant
 With fps
    If .ActiveSheet = 1 Then
        If Col = .ColLetterToNumber("E") And Row = 32 Then
            .GetText Col, Row, varTemp
            .GetText .ColLetterToNumber("M"), Row, vNgayNop
            If varTemp <> "" And varTemp <> "../../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                    checkNgayNop vNgayNop, varTemp
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
        
        If .CellType = CellTypeComboBox Then
            .Sheet = 1
            .Col = Col
            .Row = Row
            Dim TEMP As Variant
            TEMP = .TypeComboBoxIndex
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
            
            sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_TK" & sKyKeKhai & _
                  " WHERE MAPP = '2' AND MATHUE = '01' AND MADTNT = " & "'" & strMST & "'" & _
                  " AND KYKKHAI = '" & vKyKK & "'" & _
                  " AND (TTHTK = '1' OR TTHTK = '3' OR TTHTK = '4')"
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
    Dim sSQL            As String
    Dim rs              As ADODB.Recordset
    Dim vKeKhai         As Variant
    Dim vKeKhai_temp    As Variant
    Dim vKyLB           As Variant
    'them bien de kiem tra su ton tai cua bang dl
    Dim strTestFileName As String
    Dim fso             As New FileSystemObject
    Dim kykkhai         As String
    Dim flag            As Boolean
    
    If isTkQuy = True Then
        kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
    Else
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
    End If

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

        'ktr neu ton tai db moi thuc hien truy van
        If fso.FileExists(strTestFileName) Then
            sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & vKeKhai_temp & " WHERE MAPP = '2' AND MATHUE = '01' AND MADTNT = " & "'" & strMST & "'" & " AND KYKKHAI = " & kykkhai & "" & " AND (TTHTK = '1' or TTHTK = '3' or TTHTK = '4')"
            
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
