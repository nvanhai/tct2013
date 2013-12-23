VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_03TNDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_Y = 30
Const KY_LAP_BO_X = "E"
'Const MA_SO_TEP_Y = 10
'Const MA_SO_TEP_X = "M"
Const NGAY_NHAN_TO_KHAI_Y = 32
Const NGAY_NHAN_TO_KHAI_X = "E"
Const NGUOI_SU_DUNG_Y = 32
Const NGUOI_SU_DUNG_X = "V"
Const HEADER_KY_LAP_BO_Y = 13
Const HEADER_KY_LAP_BO_X = "B"
'Const MA_BPQL_X = "V"
'Const MA_BPQL_Y = 34 'Dong 16 da dc Hidden
Const PHONG_XU_LY_Y = 30
Const PHONG_XU_LY_X = "V"


Const MATM_Y = 28
Const MATM_X = "V"

Const MA_SO_THUE_X = "F"
Const MA_SO_THUE_Y = 10
Const NGAY_NOP_Y = 32
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 32
Const NGAY_QUET_X = "M"

Const TEN_GOI_X = "H"
Const TEN_GOI_Y = 8
Const DIA_CHI_X = "F"
Const DIA_CHI_Y = 12
Const TEN_BPQL_X = "V"
Const TEN_BPQL_Y = 34
Const DIEN_THOAI_X = "F"
Const DIEN_THOAI_Y = 16
Const FAX_X = "N"
Const FAX_Y = 16
Private rs As ADODB.Recordset
Private LAN_QUET As Variant

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String

Public MATM_Y1 As String
Public MATM_X1 As String

Public isExistFile As Boolean
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

'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
'Public strMaSoTep As String
Public strNgayNhanToKhai As String
Public strNguoiSuDung As String
Public strPhongXuLy As String
Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long
Private DAGHI2 As Variant
Private kyLapBo As Variant

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
Public strMaTM As String
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


Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
        
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
    Dim varMaDM As Variant
    Dim i, j As Integer
    i = 0
    j = 0
    Dim rsMATM As ADODB.Recordset
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
        '-------------------------------------------
        'lay maTM
        'dhdang sua ngay 21/03/2011
        Dim rsReturn1 As New ADODB.Recordset
        Dim strSQL1 As String
        Dim clsDAO1 As New TAX_Utilities_Svr_New.clsADO
        'dhdang sua lay matm ngay 21/03/11
        'connect to database
        If clsDAO1.Connected = False Then
            clsDAO1.CreateConnectionString spathVat & "\DB_HT\"
            clsDAO1.Connect
        End If
       
        strSQL1 = "SELECT matm, tengoi FROM dmmt2 WHERE (mamuc ='1050')and (mamuc<>matm) order by matm"
        Set rsReturn1 = clsDAO1.Execute(strSQL1)
        Set rsMATM = rsReturn1
        Set rsReturn1 = Nothing
        
        .Col = .ColLetterToNumber(MATM_X)
        .Row = MATM_Y
        MATM_X1 = MATM_X
        MATM_Y1 = MATM_Y
        If rsMATM.Fields.count > 0 Then
            Do While Not rsMATM.EOF
                'dhdang sua combox ngay xu ly
                'ngay 05-08
                .TypeComboBoxIndex = -1
                .TypeComboBoxString = TAX_Utilities_Svr_New.Convert(Trim(rsMATM.Fields(0).Value), TCVN, UNICODE) + "-" + TAX_Utilities_Svr_New.Convert(rsMATM.Fields(1).Value, TCVN, UNICODE)
                .TextTip = TextTipOff
                j = j + 1
                ReDim Preserve larrPhongXuLy(j)
                larrPhongXuLy(j) = TAX_Utilities_Svr_New.Convert(rsMATM.Fields(1).Value, TCVN, UNICODE)
                ReDim Preserve larrid(j)
                larrid(j) = rsMATM.Fields(0).Value
                rsMATM.MoveNext
            Loop
        Else
            'DisplayMessage "0077", msOKOnly
            Exit Function
        End If
        .TypeComboBoxCurSel = 1
        UpdateCell fps, .Col, .Row, .Text
        'lSoPhongXL = i
        'clsDAO1.Disconnect
        
        'lay ten DM tren PL 07-7
        If TAX_Utilities_Svr_New.NodeValidity.childNodes(13).Attributes.getNamedItem("Active").nodeValue = 1 Then
            .Sheet = 14
            .Col = .ColLetterToNumber("B")
            .Row = 37
            i = 0
            Do
                .GetText .ColLetterToNumber("AH"), .Row, varMaDM
                If Trim(varMaDM) <> vbNullString Then
                    .SetText .ColLetterToNumber("Q"), .Row, getTenDMPPDG(varMaDM)
                End If
                .GetText .ColLetterToNumber("AI"), .Row, varMaDM
                If Trim(varMaDM) <> vbNullString Then
                    .SetText .ColLetterToNumber("AB"), .Row, getTenDMPPDG(varMaDM)
                End If
                i = i + 1
                .Col = .ColLetterToNumber("B")
                .Row = 37 + i
            Loop Until UCase(.Text) = "II"
        End If
    End With
    
    Prepared2 = True
End Function
Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String) 'ByRef strPMaSoTep, ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
' If fps.ActiveCol = fps.ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) And fps.ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then _
'        CellEditFormatNumber fps, fps.ActiveSheet, fps.ActiveCol, fps.ActiveRow, KeyAscii
'    With fps
'        If .ActiveSheet = 1 And .ActiveCol = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) _
'           And .ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then
'            If KeyAscii = 46 Or KeyAscii = 45 Then ' KeyAscii of '.' character
'                KeyAscii = 0
'            End If
'        End If
'    End With
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

Public Function Prepared3() As Boolean
    Dim STT As Integer
    
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
'
'        'Set NguoiSuDung
'        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_X)
'        .Row = NGUOI_SU_DUNG_Y
'        If strNguoiSuDung <> "" Then
'            .Text = strNguoiSuDung
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'**************************************************************
        .Sheet = 1
         'set Phong xu ly
        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y
        If strPhongXuLy <> "" Then
            .Text = strPhongXuLy
            UpdateCell fps, .Col, .Row, .Text
        End If
        'set maTM
        .Col = .ColLetterToNumber(MATM_X)
        .Row = MATM_Y
        If strMaTM <> "" Then
            .Text = strMaTM
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Sheet = 1
         .Col = .ColLetterToNumber(MA_SO_THUE_X)
        .Row = MA_SO_THUE_Y
        If strMST <> "" Then
            .Text = strMST
            UpdateCell fps, .Col, .Row, .Text
        End If
        .Sheet = 1
        .Col = .ColLetterToNumber(KY_LAP_BO_X)
        .Row = KY_LAP_BO_Y
        If Len(Month(Now())) = 1 Then
            .Text = "0" & Month(Now()) & "/" & Year(Now())
            sKyKeKhai = Year(Now())
        Else
            .Text = Month(Now()) & "/" & Year(Now())
            sKyKeKhai = Year(Now())
        End If
        UpdateCell fps, .Col, .Row, .Text
        .Sheet = 1
        .Col = .ColLetterToNumber(NGAY_NOP_X)
        .Row = NGAY_NOP_Y
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
         UpdateCell fps, .Col, .Row, .Text
        .Lock = False
        .Col = .ColLetterToNumber(NGAY_QUET_X)
        .Row = NGAY_QUET_Y
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        UpdateCell fps, .Col, .Row, .Text
       .Sheet = 1
        .Col = .ColLetterToNumber(TEN_GOI_X)
        .Row = TEN_GOI_Y
        If strTenGoi <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenGoi, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        .Sheet = 1
        .Col = .ColLetterToNumber(DIA_CHI_X)
        .Row = DIA_CHI_Y
        If strDchi <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strDchi, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        .Sheet = 1
        .Col = .ColLetterToNumber(DIEN_THOAI_X)
        .Row = DIEN_THOAI_Y
        If strDThoai <> "" Then
            .Text = strDThoai
            UpdateCell fps, .Col, .Row, .Text
        End If
        .Sheet = 1
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
        
'        .Sheet = 1
'        .Col = .ColLetterToNumber(MA_BPQL_X)
'        .Row = MA_BPQL_Y
'        If strMaBPQL <> "" Then
'            .Text = strMaBPQL
'            UpdateCell fps, .Col, .Row, .Text
'        End If
        .Col = .ColLetterToNumber(TEN_BPQL_X)
        .Row = TEN_BPQL_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))

        
        ' Dat lai check E_15 Duoc gian thue theo NQ30/NQ-CP
        ' Thoi gian gian thue trong ca nam 2009 cua to khai TNDN
        If Trim(TAX_Utilities_Svr_New.Year) = "2009" Then
            .Row = 37
            .Col = .ColLetterToNumber("E")
            UpdateCell fps, .Col, .Row, .Text
            .Row = 37
            .Col = .ColLetterToNumber("F")
            .Text = "GIA HAN NOP THEO NQ30/NQ-CP"
        ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" Then
            .Row = 37
            .Col = .ColLetterToNumber("E")
            UpdateCell fps, .Col, .Row, .Text
            .Row = 37
            .Col = .ColLetterToNumber("F")
            .Text = "GIA HAN THEO QD 21/2011/QD-TTG"
        ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2011" Then
            .Row = 37
            .Col = .ColLetterToNumber("E")
            UpdateCell fps, .Col, .Row, .Text
            .Row = 37
            .Col = .ColLetterToNumber("F")
            .Text = "GIA HAN THEO QD 21/2011/QD-TTG; QD 54/2011/QD-TTg"
        Else
            ' Neu ngoai hieu luc cua gian thue thi visible cac dieu kien check nay di
            .Row = 37
            .Col = .ColLetterToNumber("B")
            .CellType = CellTypeEdit
            .Lock = True
            .Text = ""
            
            .Row = 37
            .Col = .ColLetterToNumber("E")
            .CellType = CellTypeEdit
            .Lock = True
            .Text = ""
            UpdateCell fps, .Col, .Row, ""
            
            .Row = 37
            .Col = .ColLetterToNumber("F")
            .CellType = CellTypeEdit
            .Lock = True
            .Text = ""
        End If
        '.EventEnabled(EventAllEvents) = True
        
        'set STT tren PL_05
        
            .Sheet = 12
            If .SheetVisible = True Then
                STT = 0
                .Row = 39
                       Do
                            .Col = .ColLetterToNumber("C")
                            If (.Row = 39 + STT * 3) Then
                               .Text = str(STT + 3)
                               STT = STT + 1
                            End If
                            .Col = .ColLetterToNumber("C")
                            .Row = .Row + 1
                       Loop Until .Text = "aa"
            End If
            .Sheet = 1
            
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
        
        ' Get Phong xu ly
        .Col = .ColLetterToNumber(MATM_X)
        .Row = MATM_Y
        strMaTM = .Text
        
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
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   
   Dim sSQLCol_CTD As String
   Dim sSQLVal_CTD As String
   
   Dim madtnt As Variant
   Dim ngnop As Variant
   Dim CTTN As Variant
   Dim SOTIEN As Variant
   Dim SOTIEN2 As Variant
   Dim CK As Variant
   Dim bln  As Boolean
   Dim namQT As Variant
   Dim TT As Variant
   Dim DANHAN As Variant
   'NTKCC 3.1
   Dim SOTIENCL As Variant
   
   Dim MaMuc As Variant
   Dim maTM As Variant


   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
    
   sSQLCol = "MADTNT, NGNOP, CTTN, SOTIEN, SOTIEN2, NAMQT, TTHTK, TT, MAMUC, MATM, LAN_QUET, DANHAN,SOTIENCL"
   '----------------------
   
   namQT = TAX_Utilities_Svr_New.Year
   DANHAN = "''"
   MaMuc = "'1050'"
   'MATM = "'1052'"
   With fps
        .Sheet = 1
        'lay MATM theo MANNKD
        .GetText .ColLetterToNumber("T"), 5, maTM
        If Trim(maTM) = vbNullString Then
            maTM = "''"
        Else
            maTM = "'" & maTM & "'"
        End If
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
        'Chi tieu
        CTTN = "'001'"
        .GetText .ColLetterToNumber("U"), 41, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
        SOTIENCL = "0"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'002'"
        .GetText .ColLetterToNumber("U"), 42, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'003'"
        .GetText .ColLetterToNumber("U"), 43, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'004'"
        .GetText .ColLetterToNumber("U"), 44, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'005'"
        .GetText .ColLetterToNumber("U"), 45, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
'--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'006'"
        .GetText .ColLetterToNumber("U"), 46, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
 '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'007'"
        .GetText .ColLetterToNumber("U"), 47, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)

   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'008'"
        .GetText .ColLetterToNumber("U"), 48, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'009'"
        .GetText .ColLetterToNumber("U"), 49, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'010'"
        .GetText .ColLetterToNumber("U"), 50, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'011'"
        .GetText .ColLetterToNumber("U"), 51, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'012'"
        .GetText .ColLetterToNumber("U"), 52, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'013'"
        .GetText .ColLetterToNumber("U"), 53, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'014'"
        .GetText .ColLetterToNumber("U"), 54, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'015'"
        .GetText .ColLetterToNumber("U"), 55, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'016'"
        .GetText .ColLetterToNumber("U"), 56, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'017'"
        .GetText .ColLetterToNumber("U"), 57, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'018'"
        .GetText .ColLetterToNumber("U"), 58, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'019'"
        .GetText .ColLetterToNumber("U"), 59, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'020'"
        .GetText .ColLetterToNumber("U"), 60, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'021'"
        .GetText .ColLetterToNumber("U"), 61, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'022'"
        .GetText .ColLetterToNumber("U"), 62, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'023'"
        .GetText .ColLetterToNumber("U"), 63, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'024'"
        .GetText .ColLetterToNumber("U"), 64, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'025'"
        .GetText .ColLetterToNumber("U"), 65, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'026'"
        .GetText .ColLetterToNumber("U"), 66, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'027'"
        .GetText .ColLetterToNumber("U"), 67, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'028'"
        .GetText .ColLetterToNumber("U"), 68, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'029'"
        .GetText .ColLetterToNumber("U"), 69, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'030'"
        .GetText .ColLetterToNumber("U"), 70, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'031'"
        .GetText .ColLetterToNumber("U"), 71, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'032'"
        .GetText .ColLetterToNumber("U"), 72, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 1
        'Chi tieu
        CTTN = "'033'"
        .GetText .ColLetterToNumber("U"), 73, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        TT = "''"
   End With
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & TT & "," & MaMuc & "," & maTM & "," & LAN_QUET & "," & DANHAN & "," & SOTIENCL
   sSQL = "INSERT INTO TMP_QT6D" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)

'---------------------- end ------------
'insert tai lieu dinh kem
   insert_TLDK

   clsDAO.Disconnect
   
   If TAX_Utilities_Svr_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL1
   End If
   If TAX_Utilities_Svr_New.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL1B
   End If
   If TAX_Utilities_Svr_New.NodeValidity.childNodes(3).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL1C
   End If
   If TAX_Utilities_Svr_New.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2
   End If
   If TAX_Utilities_Svr_New.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2B
   End If
   If TAX_Utilities_Svr_New.NodeValidity.childNodes(6).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL3_3A
   End If
   If TAX_Utilities_Svr_New.NodeValidity.childNodes(7).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL3_3B
   End If
   If TAX_Utilities_Svr_New.NodeValidity.childNodes(8).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL3_3C
   End If
   If TAX_Utilities_Svr_New.NodeValidity.childNodes(9).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL03_4
   End If
   If TAX_Utilities_Svr_New.NodeValidity.childNodes(10).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL03_5
   End If

  
   InsertDTL = vbNullString
End Function

Public Function InsertDTL_KHBS() As String
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim bln  As Boolean
    Dim madtnt As Variant
    Dim ngnop As Variant
    Dim kylbo As Variant
    Dim kykkhai As Variant
    Dim matkhai As Variant
    Dim MATHUE As Variant
    Dim MaMuc As Variant
    Dim maTM As Variant
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
'        'Lay ds ma tieu muc
'       ' layt ma tieu muc theo yeu cau chi Em
'        'ngay 14/09/2010
'        .Col = .ColLetterToNumber("B")
'        .Row = 50
'        Do
'            ' Lay ma tieu muc tren cel
'            .Col = .ColLetterToNumber("AO")
'            vTemp = vTemp & .Text & "~"
'
'            .Col = .ColLetterToNumber("B")
'            .Row = .Row + 1
'         Loop Until .Text = "aa"
'         arrMTM = Split(vTemp, "~")
'         vTemp = ""
'         ReDim arrMTMSave(1)
'         arrMTMSave(0) = arrMTM(0)
'         For idx = 1 To UBound(arrMTM) - 1
'            countMTM = 0
'            vTemp = arrMTM(idx)
'            For i = 0 To UBound(arrMTMSave)
'                If vTemp = arrMTMSave(i) Then
'                    countMTM = countMTM + 1
'                End If
'            Next
'            If countMTM = 0 Then
'                ReDim Preserve arrMTMSave(UBound(arrMTMSave) + 1)
'                arrMTMSave(UBound(arrMTMSave) - 1) = vTemp
'            End If
'         Next
         'the end
         
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        matkhai = "'03/TNDN'"
        MATHUE = "'01'"
        .GetText .ColLetterToNumber("E"), 30, kylbo
        If Trim(kylbo) = vbNullString Then
            kylbo = "''"
        Else
             If Len(Trim(kylbo)) = 6 Then
                kylbo = "'0" & kylbo & "'"
            Else
                kylbo = "'" & kylbo & "'"
            End If
        End If
        If CLng(Mid(Right(kylbo, 5), 1, 4)) < 2009 Then
            maTM = "'03'"
            MaMuc = "'012'"
        Else
            MaMuc = "'1550'"
        '    matm = "''"
        End If
'        vTemp = arrMTMSave(i)
        maTM = "''"
        MAPP = "'2'"
        
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
        SONGAYNC = "0"
        SOTIENNC = "0"
        DANHAN = "''"
            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & _
                      ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
                      SOKK & "," & SODC & "," & SOCL & "," & SONGAYNC & "," & SOTIENNC & "," & LAN_QUET & "," & DANHAN


            sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)

        .Sheet = .SheetCount - 1
        ' so ngay nop cham
        .GetText .ColLetterToNumber("E"), 5, SONGAYNC
        If Trim(SONGAYNC) = vbNullString Then
            SONGAYNC = "0"
        Else
            SONGAYNC = SONGAYNC
        End If
        ' so tien nop cham
        .GetText .ColLetterToNumber("G"), 5, SOTIENNC
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

                sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & _
                          ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
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

            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & _
                      ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
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

                sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & _
                          ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
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
        
            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & _
                      ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
                      SOKK & "," & SODC & "," & SOCL & "," & SONGAYNC & "," & SOTIENNC & "," & LAN_QUET & "," & DANHAN

            sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
            
   End With

'   clsDAO.Disconnect
   InsertDTL_KHBS = vbNullString
   
End Function



Public Function InsertHDR() As String
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    
    Dim madtnt As Variant
    Dim ngnop As Variant
    Dim NGNOP_D As Variant
    Dim namQT As Variant
    Dim KyTTTu As Variant
    Dim KyTTDen As Variant
    Dim kykkhai As Variant
    Dim SOKKQT As Variant
    Dim SOTHUECT18 As Variant
    Dim SOTNCN As Variant
    Dim MaVach As Variant
    Dim kylbo As Variant
    Dim MaMuc As Variant
    Dim maTM As Variant
    Dim NgNhap As Variant
    Dim DATHAY As Variant
    Dim CHKGIAHAN As Variant
    Dim DaGhi As Variant
    Dim DAGHI1 As Variant
    Dim DAGHI2 As Variant
    Dim AUTO_CAL As Variant
    Dim SHTEP As String
    Dim DANHAN As Variant
    
       Dim TENNV As Variant
    Dim CHUNGCHI As Variant
 
    '08/11/2011
    Dim MADLT As Variant
    Dim LanBS As Variant
    Dim SOKKQTCL As Variant
    Dim THUECT18CL As Variant
    Dim SOTNCNCL As Variant
    Dim MANNKD  As Variant
    Dim bln     As Boolean
    
    Dim vCT_E As Variant
    Dim vCT_C11, vCT_D As Variant
    
    sSQLCol = "MADTNT , NGNOP, NAMQT, KYTTTU, KYTTDEN, "
    sSQLCol = sSQLCol & "KYKKHAI , SOKKQT, SOTHUECT18, SOTNCN, MAVACH, TTHTK, HANNOP, KYLBO, MAMUC, MATM,SHTEP, NGNHAP, DATHAY ,DAGHI,DAGHI1,DAGHI2,AUTO_CAL,LAN_QUET, DANHAN,"
    sSQLCol = sSQLCol & "MADLT,LANBS,SOKKQTCL,THUECT18CL,SOTNCNCL,MANNKD,TENNV,CHUNGCHI"

   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If

    With fps
        .Sheet = 1
                'lay ten nhan vien dai ly thue, chung chi hanh nghe
        .GetText .ColLetterToNumber("G"), 83, TENNV
        .GetText .ColLetterToNumber("G"), 85, CHUNGCHI

        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        '.GetText .ColLetterToNumber("E"), 12, NGNOP
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            NGNOP_D = ngnop
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
        .GetText .ColLetterToNumber("E"), 30, kylbo
        If Trim(kylbo) = vbNullString Then
            kylbo = "''"
        Else
            kylbo = "'" & kylbo & "'"
        End If
        
        'set Ngay nhap
        .GetText .ColLetterToNumber("M"), 32, NgNhap
        If NgNhap = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = ToDate(Trim(NgNhap), DDMMYYYY)
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If
        namQT = TAX_Utilities_Svr_New.Year
        
        KyTTTu = "'01/" & TAX_Utilities_Svr_New.Year & "'"
        KyTTDen = "'12/" & TAX_Utilities_Svr_New.Year & "'"
       ' KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        kykkhai = kylbo
        SOKKQT = "0"
        SOTHUECT18 = "0"
        SOTNCN = "0"
        MaVach = ".T."
        DANHAN = "''"
        
    '   get MANNKD
        .GetText .ColLetterToNumber("T"), 5, MANNKD
        If Trim(MANNKD) = vbNullString Then
            MANNKD = "''"
        Else
            MANNKD = "'" & MANNKD & "'"
        End If
 ' sua MAMUC & MATM tu nam 2009
        MaMuc = "'1050'"
        maTM = "'1053'"
'        MATM = "'" & Left$(strMaTM, 4) & "'"
        'MAVACH = ".F."
        DATHAY = ".F."
        SHTEP = "''"
        DAGHI1 = ".F."
        DAGHI2 = ".F."
        AUTO_CAL = ".F."
        'Check gia han nop
        .GetText .ColLetterToNumber("E"), 37, CHKGIAHAN
        If Trim(CHKGIAHAN) = "1" Or Trim(CHKGIAHAN) = "x" Then
            DaGhi = ".T."
            If Trim(TAX_Utilities_Svr_New.Year) = "2009" Then
                HANNOP = "02/11/2010"
            ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" Then
                HANNOP = "02/04/2012"
            ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2011" Then
                HANNOP = "01/04/2013"
            End If
            If Trim(HANNOP) = vbNullString Then
                HANNOP = "CTOD('')"
            Else
                HANNOP = DateSerial(Int(Mid$(HANNOP, 7, 4)), Int(Mid$(HANNOP, 4, 2)), Int(Mid$(HANNOP, 1, 2)))
                If tthtk = "'2'" And DateDiff("d", HANNOP, NGNOP_D) > 0 And (Trim(TAX_Utilities_Svr_New.Year) = "2011" Or Trim(TAX_Utilities_Svr_New.Year) = "2010") Then
                HANNOP = NGNOP_D
                End If
                HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
            End If
        Else
            DaGhi = ".F."
            If Trim(HANNOP) = vbNullString Then
                HANNOP = "CTOD('')"
            Else
                HANNOP = DateSerial(Int(Mid$(HANNOP, 7, 4)), Int(Mid$(HANNOP, 4, 2)), Int(Mid$(HANNOP, 1, 2)))
                If tthtk = "'2'" And DateDiff("d", HANNOP, NGNOP_D) > 0 And (Trim(TAX_Utilities_Svr_New.Year) = "2011" Or Trim(TAX_Utilities_Svr_New.Year) = "2010") Then
                HANNOP = NGNOP_D
                End If
                HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
            End If
        End If
        'set daghi
        DAGHI1 = DaGhi
        'get MADLT
        .GetText .ColLetterToNumber("F"), 20, MADLT
        If MADLT = vbNullString Then
            MADLT = "''"
        Else
            MADLT = "'" & MADLT & "'"
        End If
        ' get LANBS
        .GetText .ColLetterToNumber("U"), 6, LanBS
        If LanBS = vbNullString Then
            LanBS = "0"
        End If
        
        'set cac Chi tieu CL
        SOKKQTCL = "0"
        THUECT18CL = "0"
        SOTNCNCL = "0"
        
        'get chi tieu D
'        .GetText .ColLetterToNumber("U"), 70, vCT_D
'        If Trim(vCT_D) = vbNullString Then
'            vCT_D = "0"
'        End If

        'get chi tieu B14
        .GetText .ColLetterToNumber("U"), 57, vCT_C11
        If Trim(vCT_C11) = vbNullString Then
            vCT_C11 = "0"
        End If
        'get chi tieu E2
        .GetText .ColLetterToNumber("U"), 73, vCT_E
        If Trim(vCT_E) = vbNullString Then
            vCT_E = "0"
        End If
        SOTNCN = vCT_E
        'set SOKKQT
        SOKKQT = Val(vCT_C11)
'        'set sothuect18
'        SOTHUECT18 = vCT_E
         
     End With

    sSQLVal = madtnt & "," & ngnop & "," & namQT & "," & KyTTTu & "," & KyTTDen & ", "
    sSQLVal = sSQLVal & kykkhai & "," & SOKKQT & "," & SOTHUECT18 & "," & SOTNCN & "," & MaVach & "," & tthtk & "," & HANNOP & "," & kylbo & "," & MaMuc & "," & maTM & "," & SHTEP & "," & NgNhap & "," & DATHAY & "," & DaGhi & "," & DAGHI1 & "," & DAGHI2 & "," & AUTO_CAL & "," & LAN_QUET & "," & DANHAN
    sSQLVal = sSQLVal & "," & MADLT & "," & LanBS & "," & SOKKQTCL & "," & THUECT18CL & "," & SOTNCNCL & "," & MANNKD & ",'" & TENNV & "','" & CHUNGCHI & "'"
    
    
    sSQL = "INSERT INTO TMP_QT6" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
     
    'insert theo matm lay theo mannkd
    With fps

        maTM = MANNKD
    
        'get chi tieu B13
        .GetText .ColLetterToNumber("U"), 56, vCT_C11
        If Trim(vCT_C11) = vbNullString Then
            vCT_C11 = "0"
        End If
        'get chi tieu E1
        .GetText .ColLetterToNumber("U"), 72, vCT_E
        If Trim(vCT_E) = vbNullString Then
            vCT_E = "0"
        End If
        SOTNCN = vCT_E
        'set SOKKQT
        SOKKQT = Val(vCT_C11)
'        'set sothuect18
'        SOTHUECT18 = vCT_E
        
    End With

    sSQLVal = madtnt & "," & ngnop & "," & namQT & "," & KyTTTu & "," & KyTTDen & ", "
    sSQLVal = sSQLVal & kykkhai & "," & SOKKQT & "," & SOTHUECT18 & "," & SOTNCN & "," & MaVach & "," & tthtk & "," & HANNOP & "," & kylbo & "," & MaMuc & "," & maTM & "," & SHTEP & "," & NgNhap & "," & DATHAY & "," & DaGhi & "," & DAGHI1 & "," & DAGHI2 & "," & AUTO_CAL & "," & LAN_QUET & "," & DANHAN
    sSQLVal = sSQLVal & "," & MADLT & "," & LanBS & "," & SOKKQTCL & "," & THUECT18CL & "," & SOTNCNCL & "," & MANNKD
    
    
    sSQL = "INSERT INTO TMP_QT6" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
     
    InsertHDR = vbNullString
    
    clsDAO.Disconnect
    
End Function

Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim strFileDinhKem As String
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QT6" & sKyKeKhai & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QT6D" & sKyKeKhai & ".DBF"
    strFileDinhKem = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_Q3D" & sKyKeKhai & ".DBF"
    Dim strFileNameBS As String
    strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"
    'file HDR
    If fso.FileExists(strFileNameHDR) = False Then
        Dim strTepmauHDR As String
        strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_QT6.DBF"
        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    End If
    'file dtl
    If fso.FileExists(strFileNameDTL) = False Then
        Dim strTepmauDTL As String
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_QT6D.DBF"
        fso.CopyFile strTepmauDTL, strFileNameDTL, False
    End If
    If fso.FileExists(strFileDinhKem) = False Then
        Dim strTepMauDinhKem As String
        strTepMauDinhKem = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_Q3D.DBF"
        fso.CopyFile strTepMauDinhKem, strFileDinhKem, False
    End If
    'filedinh kiem
    'Cac file fu luc
    Dim strFileNamePL1 As String
    Dim strTepmauPL1 As String
    'Them bang phu luc PL03-1A/TNDN
    strFileNamePL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_Q31A" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileNamePL1) = False Then
            strTepmauPL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_Q31A.DBF"
            fso.CopyFile strTepmauPL1, strFileNamePL1, False
    End If
    'Them bang phu luc PL03-1B/TNDN
    'Them bang phu luc PL03-2/TNDN
    strFileNamePL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_Q322" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileNamePL1) = False Then
            strTepmauPL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_Q322.DBF"
            fso.CopyFile strTepmauPL1, strFileNamePL1, False
    End If
    'Them bang phu luc PL03-4A/TNDN
    strFileNamePL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_Q3ABC" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileNamePL1) = False Then
            strTepmauPL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_Q3ABC.DBF"
            fso.CopyFile strTepmauPL1, strFileNamePL1, False
    End If
    'Them bang phu luc PL03-5/TNDN
    strFileNamePL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_Q35" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileNamePL1) = False Then
            strTepmauPL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_Q35.DBF"
            fso.CopyFile strTepmauPL1, strFileNamePL1, False
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
   
'   sSQL = "SELECT MADTNT, NGNOP FROM QT6" & sKyKeKhai & _
'        " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT = " & TAX_Utilities_Svr_New.Year & " AND TTHTK = '1'" & " AND KYKKHAI = '" & Month(Now) & "/" & Year(Now) & "'"
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_QT6" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT = " & TAX_Utilities_Svr_New.Year
    
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
   
   sSQL = "UPDATE TMP_QT6" & sKyKeKhai & _
        " SET DATHAY=.T." & _
        " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT = " & TAX_Utilities_Svr_New.Year
    
   clsDAO.ExecuteDLL sSQL
   
   UpdateTHUETKY2 = True
   
   clsDAO.Disconnect

End Function


Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim ngnop As Variant
    Dim rs As ADODB.Recordset
    
    If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
     
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 12, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
   End With
   sSQL = "DELETE FROM TMP_QT6D" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
    
   clsDAO.ExecuteDLL sSQL
   
   ' xoa hrd
   sSQL = "DELETE FROM TMP_QT6" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT=" & TAX_Utilities_Svr_New.Year
    
   clsDAO.ExecuteDLL sSQL
   
'   With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("E"), 12, NGNOP
'        'NGNOP = Date
'        If Trim(NGNOP) = vbNullString Then
'            NGNOP = "CTOD('')"
'        Else
'            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
'            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
'        End If
'   End With
'   sSQL = "DELETE FROM Q6D" & sKyKeKhai & _
'        " WHERE MADTNT = " & "'" & strMST & "' AND NGNOP = " & dNgayNopDB
'
'   clsDAO.ExecuteDLL sSQL
   '------------
   'Xoa phu luc
   'Phu luc PL1A
   sSQL = "DELETE FROM TMP_Q31A" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   
   'Phu luc PL1B
   sSQL = "DELETE FROM TMP_Q31B" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   
   ' phuc luc PL1C
   sSQL = "DELETE FROM TMP_Q31C" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' xoa phu luc PL3
   sSQL = "DELETE FROM TMP_Q33" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   
   ' xoa phuc luc PL4A
    sSQL = "DELETE FROM TMP_Q34A" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' xoa phu luc PL4B
   sSQL = "DELETE FROM TMP_Q34B" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   'xoa phu luc PL4C
   sSQL = "DELETE FROM TMP_Q34C" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' xoa phu luc PL4D
   sSQL = "DELETE FROM TMP_Q34D" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' xoa phu luc PL4DD
   sSQL = "DELETE FROM TMP_Q34Z" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' Xoa phu luc PL4E
   sSQL = "DELETE FROM TMP_Q34E" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' Xoa phu luc PL4G
   sSQL = "DELETE FROM TMP_Q34G" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' Xoa phu luc PL4H
   sSQL = "DELETE FROM TMP_Q34H" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   ' Xoa phu luc PL5
   sSQL = "DELETE FROM TMP_Q35" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   'Phu luc PL2A
   sSQL = "DELETE FROM TMP_Q321" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL
   'Xoa du lieu bang Q322
   sSQL = "DELETE FROM TMP_Q322" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'"
   clsDAO.ExecuteDLL sSQL

   XoaTKTT = True
   clsDAO.Disconnect
End Function

Public Function TKRB() As Boolean

   TKRB = True

End Function

Public Function InsertDTL_PL1C() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim namQT As Variant
   Dim madtnt As Variant
   Dim ngnop As Variant
   Dim CTTN As Variant
   Dim SOTIEN As Variant
   Dim SOTIEN2 As Variant
   Dim DANHAN As Variant
   Dim bln  As Boolean
   
   Dim MAPL As Variant
   Dim TT   As Variant
   
   
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
    
   sSQLCol = "MADTNT, NGNOP, CTTN, SOTIEN, SOTIEN2,NAMQT,TTHTK,LAN_QUET,DANHAN,MAPL,TT"
     
    '---------------
    MAPL = "'PL03-1C/TNDN'"
    TT = "''"
    namQT = TAX_Utilities_Svr_New.Year
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
        .Sheet = 4
        ' Chi tieu
        CTTN = "'001'"
        .GetText .ColLetterToNumber("E"), 23, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   DANHAN = "''"
     
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'002'"
        .GetText .ColLetterToNumber("E"), 24, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
     
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'003'"
        .GetText .ColLetterToNumber("E"), 25, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'004'"
        .GetText .ColLetterToNumber("E"), 26, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'005'"
        .GetText .ColLetterToNumber("E"), 27, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'006'"
        .GetText .ColLetterToNumber("E"), 28, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'007'"
        .GetText .ColLetterToNumber("E"), 29, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'008'"
        .GetText .ColLetterToNumber("E"), 30, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'009'"
        .GetText .ColLetterToNumber("E"), 31, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'010'"
        .GetText .ColLetterToNumber("E"), 32, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'011'"
        .GetText .ColLetterToNumber("E"), 33, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'012'"
        .GetText .ColLetterToNumber("E"), 34, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'013'"
        .GetText .ColLetterToNumber("E"), 35, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'014'"
        .GetText .ColLetterToNumber("E"), 36, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'015'"
        .GetText .ColLetterToNumber("E"), 37, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'016'"
        .GetText .ColLetterToNumber("E"), 38, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
     
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'017'"
        .GetText .ColLetterToNumber("E"), 39, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'018'"
        .GetText .ColLetterToNumber("E"), 40, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'019'"
        .GetText .ColLetterToNumber("E"), 41, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'020'"
        .GetText .ColLetterToNumber("E"), 42, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
   
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'021'"
        .GetText .ColLetterToNumber("E"), 43, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'022'"
        .GetText .ColLetterToNumber("E"), 44, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'023'"
        .GetText .ColLetterToNumber("E"), 45, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'024'"
        .GetText .ColLetterToNumber("E"), 46, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'025'"
        .GetText .ColLetterToNumber("E"), 47, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'026'"
        .GetText .ColLetterToNumber("E"), 48, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   '-----------------------
    With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'027'"
        .GetText .ColLetterToNumber("E"), 49, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '-----------------------
   
   
       With fps
        .Sheet = 4
        ' Chi tieu
        CTTN = "'028'"
        .GetText .ColLetterToNumber("E"), 50, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   
   
   
   
   '-----------------------
   
   
   clsDAO.Disconnect
   InsertDTL_PL1C = vbNullString
   End Function

Public Function InsertDTL_PL1B() As String
    Dim sSQL    As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim namQT   As Variant
    Dim madtnt  As Variant
    Dim ngnop   As Variant
    Dim CTTN    As Variant
    Dim SOTIEN  As Variant
    Dim SOTIEN2 As Variant
    Dim CK      As Variant
    Dim DANHAN  As Variant
    Dim bln     As Boolean
   
    Dim MAPL    As Variant
    Dim TT      As Variant
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
    sSQLCol = "MADTNT, NGNOP, CTTN, SOTIEN, SOTIEN2,LAN_QUET,DANHAN,MAPL,NAMQT,TTHTK,TT"
     
    '---------------
    With fps
        .Sheet = 1
        MAPL = "'PL03-1B/TNDN'"
        namQT = TAX_Utilities_Svr_New.Year
        TT = "''"
        
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

        .Sheet = 3
        ' Chi tieu
        CTTN = "'001'"
        .GetText .ColLetterToNumber("E"), 23, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With

    DANHAN = "''"
     
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'002'"
        .GetText .ColLetterToNumber("E"), 24, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'003'"
        .GetText .ColLetterToNumber("E"), 25, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
       
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'004'"
        .GetText .ColLetterToNumber("E"), 26, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'005'"
        .GetText .ColLetterToNumber("E"), 27, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'006'"
        .GetText .ColLetterToNumber("E"), 28, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'007'"
        .GetText .ColLetterToNumber("E"), 29, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'008'"
        .GetText .ColLetterToNumber("E"), 30, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'009'"
        .GetText .ColLetterToNumber("E"), 31, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'010'"
        .GetText .ColLetterToNumber("E"), 32, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'011'"
        .GetText .ColLetterToNumber("E"), 33, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With

    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'012'"
        .GetText .ColLetterToNumber("E"), 34, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'013'"
        .GetText .ColLetterToNumber("E"), 35, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'014'"
        .GetText .ColLetterToNumber("E"), 36, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'015'"
        .GetText .ColLetterToNumber("E"), 37, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
    With fps
        .Sheet = 3
        ' Chi tieu
        CTTN = "'016'"
        .GetText .ColLetterToNumber("E"), 38, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & namQT & "," & tthtk & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
   
    '--------------------
   
    clsDAO.Disconnect
    InsertDTL_PL1B = vbNullString
   
End Function
   



Public Function InsertDTL_PL1() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   
   Dim sSQLCol_CTD As String
   Dim sSQLVal_CTD As String
   Dim namQT As Variant
   Dim madtnt As Variant
   Dim ngnop As Variant
   Dim CTTN As Variant
   Dim SOTIEN As Variant
   Dim SOTIEN2 As Variant
   Dim CK As Variant
   Dim DANHAN As Variant
   Dim bln  As Boolean
   
   Dim MAPL As Variant
   Dim TT   As Variant
   
   
   If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
    
   sSQLCol = "MADTNT, NGNOP, CTTN, SOTIEN, SOTIEN2, NAMQT , TTHTK, LAN_QUET, DANHAN,MAPL,TT"
  ' sSQLCol_CTD = "MADTNT, NGNOP, CTTN, CK "
     
   '--------------
   Dim sSQLCol_PL1I As String
   Dim sSQLVal_PL1I As String
   
   'Dim MADTNT As Variant
   'Dim NGNOP As Variant
   Dim NAMPS As Variant
   Dim NAMCL As Variant
   'Dim SOTIEN As Variant
   
   'sSQLCol_PL1I = "MADTNT, NGNOP, NamPS, NamCL, SOTIEN "
   '---------------
   
   namQT = TAX_Utilities_Svr_New.Year
   DANHAN = "''"
   MAPL = "'PL03-1A/TNDN'"
   TT = "''"
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
        .Sheet = 2
        .GetText .ColLetterToNumber("E"), 23, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        ' Chi tieu
        CTTN = "'001'"
'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
     
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'002'"
        .GetText .ColLetterToNumber("E"), 24, SOTIEN
         If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
        
'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'003'"
        .GetText .ColLetterToNumber("E"), 25, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'004'"
        .GetText .ColLetterToNumber("E"), 26, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'005'"
        .GetText .ColLetterToNumber("E"), 27, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'006'"
        .GetText .ColLetterToNumber("E"), 28, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'007'"
        .GetText .ColLetterToNumber("E"), 29, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'008'"
        .GetText .ColLetterToNumber("E"), 30, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'009'"
        .GetText .ColLetterToNumber("E"), 31, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'010'"
        .GetText .ColLetterToNumber("E"), 32, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'011'"
        .GetText .ColLetterToNumber("E"), 33, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'012'"
        .GetText .ColLetterToNumber("E"), 34, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
     
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'013'"
        .GetText .ColLetterToNumber("E"), 35, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        ' Chi tieu
        CTTN = "'014'"
        .GetText .ColLetterToNumber("E"), 36, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'015'"
        .GetText .ColLetterToNumber("E"), 37, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
     
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'016'"
        .GetText .ColLetterToNumber("E"), 38, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'017'"
        .GetText .ColLetterToNumber("E"), 39, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"

'        .GetText .ColLetterToNumber("F"), 4, NamCL
'        If Trim(NamCL) = vbNullString Then
'            NamCL = "0"
'        End If
'        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
        
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'018'"
        .GetText .ColLetterToNumber("E"), 40, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
      
      With fps
        .Sheet = 2
        ' Chi tieu
        CTTN = "'019'"
        .GetText .ColLetterToNumber("E"), 41, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        SOTIEN2 = "0"
   End With
   
   sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
   sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------
   clsDAO.Disconnect
   
   InsertDTL_PL1 = vbNullString
End Function

Public Function InsertDTL_PL2() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   
   Dim sSQLCol_CTD As String
   Dim sSQLVal_CTD As String
   
   Dim madtnt As Variant
   Dim ngnop As Variant
   Dim NAMPS As Variant
   Dim NAMCL As Variant
   Dim SOTIEN As Variant
   Dim NAMCL3 As Variant
   Dim NAMCL4 As Variant
   Dim NAMCL5 As Variant
   Dim NAMCL6 As Variant
   Dim NAMCL7 As Variant
   Dim NAMCL8 As Variant
   
   Dim NAMPS1 As Variant
   Dim NAMPS2 As Variant
   Dim NAMPS3 As Variant
   Dim NAMPS4 As Variant
   Dim NAMPS5 As Variant
   Dim NAMPS6 As Variant
   
   
   Dim SOTIEN1 As Variant
   Dim SOTIEN2 As Variant
   Dim SOTIEN3 As Variant
   Dim SOTIEN4 As Variant
   Dim DANHAN As Variant
   Dim namQT  As Variant
   Dim MAPL   As Variant

   Dim bln  As Boolean
   
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   'Ghi du lieu chuye trong ky tinh thue
   sSQLCol = "MADTNT, NGNOP, NAMPS, SOTIEN1, SOTIEN2, SOTIEN3, SOTIEN4,LAN_QUET, DANHAN,NAMQT,TTHTK,MAPL"
    '-----------------
      With fps
        .Sheet = 1
        
        MAPL = "'PL03-2A/TNDN'"
        
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
        
        'set nam quyet toan
        namQT = TAX_Utilities_Svr_New.Year
        DANHAN = "''"
       End With
   'Nam 1
   With fps
        .Sheet = 5
        ' Chi tieu
       ' get namps
       .GetText .ColLetterToNumber("B"), 26, NAMPS
       If NAMPS = vbNullString Then
            NAMPS = "0"
       End If
       .GetText .ColLetterToNumber("C"), 26, SOTIEN1
       If Trim(SOTIEN1) = vbNullString Then
            SOTIEN1 = "0"
       End If
       
       .GetText .ColLetterToNumber("H"), 26, SOTIEN2
       If Trim(SOTIEN2) = vbNullString Then
            SOTIEN2 = "0"
       End If
       
        .GetText .ColLetterToNumber("M"), 26, SOTIEN3
        If Trim(SOTIEN3) = vbNullString Then
                SOTIEN3 = "0"
        End If
        
       .GetText .ColLetterToNumber("R"), 26, SOTIEN4
        If Trim(SOTIEN4) = vbNullString Then
                SOTIEN4 = "0"
        End If

   End With
   sSQLVal = madtnt & "," & ngnop & "," & NAMPS & "," & SOTIEN1 & "," & SOTIEN2 & "," & _
   SOTIEN3 & "," & SOTIEN4 & "," & LAN_QUET & "," & DANHAN & "," & namQT & "," & tthtk & "," & MAPL
   sSQL = "INSERT INTO TMP_Q322" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------------
   
   'Nam 2
   With fps
        .Sheet = 5
        ' Chi tieu
       ' get namps
       .GetText .ColLetterToNumber("B"), 27, NAMPS
       If NAMPS = vbNullString Then
            NAMPS = "0"
       End If
       .GetText .ColLetterToNumber("C"), 27, SOTIEN1
       If Trim(SOTIEN1) = vbNullString Then
            SOTIEN1 = "0"
       End If
       
       .GetText .ColLetterToNumber("H"), 27, SOTIEN2
       If Trim(SOTIEN2) = vbNullString Then
            SOTIEN2 = "0"
       End If
       
        .GetText .ColLetterToNumber("M"), 27, SOTIEN3
        If Trim(SOTIEN3) = vbNullString Then
                SOTIEN3 = "0"
        End If
        
       .GetText .ColLetterToNumber("R"), 27, SOTIEN4
        If Trim(SOTIEN4) = vbNullString Then
                SOTIEN4 = "0"
        End If

   End With
   sSQLVal = madtnt & "," & ngnop & "," & NAMPS & "," & SOTIEN1 & "," & SOTIEN2 & "," & _
   SOTIEN3 & "," & SOTIEN4 & "," & LAN_QUET & "," & DANHAN & "," & namQT & "," & tthtk & "," & MAPL
   sSQL = "INSERT INTO TMP_Q322" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------------
   'Nam 3
   With fps
        .Sheet = 5
        ' Chi tieu
       ' get namps
       .GetText .ColLetterToNumber("B"), 28, NAMPS
       If NAMPS = vbNullString Then
            NAMPS = "0"
       End If
       .GetText .ColLetterToNumber("C"), 28, SOTIEN1
       If Trim(SOTIEN1) = vbNullString Then
            SOTIEN1 = "0"
       End If
       
       .GetText .ColLetterToNumber("H"), 28, SOTIEN2
       If Trim(SOTIEN2) = vbNullString Then
            SOTIEN2 = "0"
       End If
       
        .GetText .ColLetterToNumber("M"), 28, SOTIEN3
        If Trim(SOTIEN3) = vbNullString Then
                SOTIEN3 = "0"
        End If
        
       .GetText .ColLetterToNumber("R"), 28, SOTIEN4
        If Trim(SOTIEN4) = vbNullString Then
                SOTIEN4 = "0"
        End If

   End With
   sSQLVal = madtnt & "," & ngnop & "," & NAMPS & "," & SOTIEN1 & "," & SOTIEN2 & "," & _
   SOTIEN3 & "," & SOTIEN4 & "," & LAN_QUET & "," & DANHAN & "," & namQT & "," & tthtk & "," & MAPL
   sSQL = "INSERT INTO TMP_Q322" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------------
   
   'Nam 4
   With fps
        .Sheet = 5
        ' Chi tieu
       ' get namps
       .GetText .ColLetterToNumber("B"), 29, NAMPS
       If NAMPS = vbNullString Then
            NAMPS = "0"
       End If
       .GetText .ColLetterToNumber("C"), 29, SOTIEN1
       If Trim(SOTIEN1) = vbNullString Then
            SOTIEN1 = "0"
       End If
       
       .GetText .ColLetterToNumber("H"), 29, SOTIEN2
       If Trim(SOTIEN2) = vbNullString Then
            SOTIEN2 = "0"
       End If
       
        .GetText .ColLetterToNumber("M"), 29, SOTIEN3
        If Trim(SOTIEN3) = vbNullString Then
                SOTIEN3 = "0"
        End If
        
       .GetText .ColLetterToNumber("R"), 29, SOTIEN4
        If Trim(SOTIEN4) = vbNullString Then
                SOTIEN4 = "0"
        End If

   End With
   sSQLVal = madtnt & "," & ngnop & "," & NAMPS & "," & SOTIEN1 & "," & SOTIEN2 & "," & _
   SOTIEN3 & "," & SOTIEN4 & "," & LAN_QUET & "," & DANHAN & "," & namQT & "," & tthtk & "," & MAPL
   sSQL = "INSERT INTO TMP_Q322" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------------
   'Nam 5
   With fps
        .Sheet = 5
        ' Chi tieu
       ' get namps
       .GetText .ColLetterToNumber("B"), 30, NAMPS
       If NAMPS = vbNullString Then
            NAMPS = "0"
       End If
       .GetText .ColLetterToNumber("C"), 30, SOTIEN1
       If Trim(SOTIEN1) = vbNullString Then
            SOTIEN1 = "0"
       End If
       
       .GetText .ColLetterToNumber("H"), 30, SOTIEN2
       If Trim(SOTIEN2) = vbNullString Then
            SOTIEN2 = "0"
       End If
       
        .GetText .ColLetterToNumber("M"), 30, SOTIEN3
        If Trim(SOTIEN3) = vbNullString Then
                SOTIEN3 = "0"
        End If
        
       .GetText .ColLetterToNumber("R"), 30, SOTIEN4
        If Trim(SOTIEN4) = vbNullString Then
                SOTIEN4 = "0"
        End If

   End With
   sSQLVal = madtnt & "," & ngnop & "," & NAMPS & "," & SOTIEN1 & "," & SOTIEN2 & "," & _
   SOTIEN3 & "," & SOTIEN4 & "," & LAN_QUET & "," & DANHAN & "," & namQT & "," & tthtk & "," & MAPL
   sSQL = "INSERT INTO TMP_Q322" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------------
   
   clsDAO.Disconnect
   
   
   InsertDTL_PL2 = vbNullString
End Function

Public Function InsertDTL_PL2B() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   
   Dim sSQLCol_CTD As String
   Dim sSQLVal_CTD As String
   
   Dim madtnt As Variant
   Dim ngnop As Variant
   Dim NAMPS As Variant
   Dim NAMCL As Variant
   Dim SOTIEN As Variant
   Dim NAMCL3 As Variant
   Dim NAMCL4 As Variant
   Dim NAMCL5 As Variant
   Dim NAMCL6 As Variant
   Dim NAMCL7 As Variant
   Dim NAMCL8 As Variant
   
   Dim NAMPS1 As Variant
   Dim NAMPS2 As Variant
   Dim NAMPS3 As Variant
   Dim NAMPS4 As Variant
   Dim NAMPS5 As Variant
   Dim NAMPS6 As Variant
   
   
   Dim SOTIEN1 As Variant
   Dim SOTIEN2 As Variant
   Dim SOTIEN3 As Variant
   Dim SOTIEN4 As Variant
   Dim DANHAN As Variant
   Dim namQT  As Variant
   Dim MAPL    As Variant

   Dim bln  As Boolean
   
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   'Ghi du lieu chuye trong ky tinh thue
   sSQLCol = "MADTNT, NGNOP, NAMPS, SOTIEN1, SOTIEN2, SOTIEN3, SOTIEN4,LAN_QUET, DANHAN,NAMQT,TTHTK,MAPL"
    '-----------------
      With fps
        .Sheet = 1
        
        
        MAPL = "'PL03-2B/TNDN'"
        
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
        
        'set nam quyet toan
        namQT = TAX_Utilities_Svr_New.Year
        DANHAN = "''"
       End With
   'Nam 1
   With fps
        .Sheet = 6
        ' Chi tieu
       ' get namps
       .GetText .ColLetterToNumber("B"), 26, NAMPS
       If NAMPS = vbNullString Then
            NAMPS = "0"
       End If
       .GetText .ColLetterToNumber("C"), 26, SOTIEN1
       If Trim(SOTIEN1) = vbNullString Then
            SOTIEN1 = "0"
       End If
       
       .GetText .ColLetterToNumber("H"), 26, SOTIEN2
       If Trim(SOTIEN2) = vbNullString Then
            SOTIEN2 = "0"
       End If
       
        .GetText .ColLetterToNumber("M"), 26, SOTIEN3
        If Trim(SOTIEN3) = vbNullString Then
                SOTIEN3 = "0"
        End If
        
       .GetText .ColLetterToNumber("R"), 26, SOTIEN4
        If Trim(SOTIEN4) = vbNullString Then
                SOTIEN4 = "0"
        End If

   End With
   sSQLVal = madtnt & "," & ngnop & "," & NAMPS & "," & SOTIEN1 & "," & SOTIEN2 & "," & _
   SOTIEN3 & "," & SOTIEN4 & "," & LAN_QUET & "," & DANHAN & "," & namQT & "," & tthtk & "," & MAPL
   sSQL = "INSERT INTO TMP_Q322" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------------
   
   'Nam 2
   With fps
        .Sheet = 6
        ' Chi tieu
       ' get namps
       .GetText .ColLetterToNumber("B"), 27, NAMPS
       If NAMPS = vbNullString Then
            NAMPS = "0"
       End If
       .GetText .ColLetterToNumber("C"), 27, SOTIEN1
       If Trim(SOTIEN1) = vbNullString Then
            SOTIEN1 = "0"
       End If
       
       .GetText .ColLetterToNumber("H"), 27, SOTIEN2
       If Trim(SOTIEN2) = vbNullString Then
            SOTIEN2 = "0"
       End If
       
        .GetText .ColLetterToNumber("M"), 27, SOTIEN3
        If Trim(SOTIEN3) = vbNullString Then
                SOTIEN3 = "0"
        End If
        
       .GetText .ColLetterToNumber("R"), 27, SOTIEN4
        If Trim(SOTIEN4) = vbNullString Then
                SOTIEN4 = "0"
        End If

   End With
   sSQLVal = madtnt & "," & ngnop & "," & NAMPS & "," & SOTIEN1 & "," & SOTIEN2 & "," & _
   SOTIEN3 & "," & SOTIEN4 & "," & LAN_QUET & "," & DANHAN & "," & namQT & "," & tthtk & "," & MAPL
   sSQL = "INSERT INTO TMP_Q322" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------------
   'Nam 3
   With fps
        .Sheet = 6
        ' Chi tieu
       ' get namps
       .GetText .ColLetterToNumber("B"), 28, NAMPS
       If NAMPS = vbNullString Then
            NAMPS = "0"
       End If
       .GetText .ColLetterToNumber("C"), 28, SOTIEN1
       If Trim(SOTIEN1) = vbNullString Then
            SOTIEN1 = "0"
       End If
       
       .GetText .ColLetterToNumber("H"), 28, SOTIEN2
       If Trim(SOTIEN2) = vbNullString Then
            SOTIEN2 = "0"
       End If
       
        .GetText .ColLetterToNumber("M"), 28, SOTIEN3
        If Trim(SOTIEN3) = vbNullString Then
                SOTIEN3 = "0"
        End If
        
       .GetText .ColLetterToNumber("R"), 28, SOTIEN4
        If Trim(SOTIEN4) = vbNullString Then
                SOTIEN4 = "0"
        End If

   End With
   sSQLVal = madtnt & "," & ngnop & "," & NAMPS & "," & SOTIEN1 & "," & SOTIEN2 & "," & _
   SOTIEN3 & "," & SOTIEN4 & "," & LAN_QUET & "," & DANHAN & "," & namQT & "," & tthtk & "," & MAPL
   sSQL = "INSERT INTO TMP_Q322" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------------
   
   'Nam 4
   With fps
        .Sheet = 6
        ' Chi tieu
       ' get namps
       .GetText .ColLetterToNumber("B"), 29, NAMPS
       If NAMPS = vbNullString Then
            NAMPS = "0"
       End If
       .GetText .ColLetterToNumber("C"), 29, SOTIEN1
       If Trim(SOTIEN1) = vbNullString Then
            SOTIEN1 = "0"
       End If
       
       .GetText .ColLetterToNumber("H"), 29, SOTIEN2
       If Trim(SOTIEN2) = vbNullString Then
            SOTIEN2 = "0"
       End If
       
        .GetText .ColLetterToNumber("M"), 29, SOTIEN3
        If Trim(SOTIEN3) = vbNullString Then
                SOTIEN3 = "0"
        End If
        
       .GetText .ColLetterToNumber("R"), 29, SOTIEN4
        If Trim(SOTIEN4) = vbNullString Then
                SOTIEN4 = "0"
        End If

   End With
   sSQLVal = madtnt & "," & ngnop & "," & NAMPS & "," & SOTIEN1 & "," & SOTIEN2 & "," & _
   SOTIEN3 & "," & SOTIEN4 & "," & LAN_QUET & "," & DANHAN & "," & namQT & "," & tthtk & "," & MAPL
   sSQL = "INSERT INTO TMP_Q322" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------------
   'Nam 5
   With fps
        .Sheet = 6
        ' Chi tieu
       ' get namps
       .GetText .ColLetterToNumber("B"), 30, NAMPS
       If NAMPS = vbNullString Then
            NAMPS = "0"
       End If
       .GetText .ColLetterToNumber("C"), 30, SOTIEN1
       If Trim(SOTIEN1) = vbNullString Then
            SOTIEN1 = "0"
       End If
       
       .GetText .ColLetterToNumber("H"), 30, SOTIEN2
       If Trim(SOTIEN2) = vbNullString Then
            SOTIEN2 = "0"
       End If
       
        .GetText .ColLetterToNumber("M"), 30, SOTIEN3
        If Trim(SOTIEN3) = vbNullString Then
                SOTIEN3 = "0"
        End If
        
       .GetText .ColLetterToNumber("R"), 30, SOTIEN4
        If Trim(SOTIEN4) = vbNullString Then
                SOTIEN4 = "0"
        End If

   End With
   sSQLVal = madtnt & "," & ngnop & "," & NAMPS & "," & SOTIEN1 & "," & SOTIEN2 & "," & _
   SOTIEN3 & "," & SOTIEN4 & "," & LAN_QUET & "," & DANHAN & "," & namQT & "," & tthtk & "," & MAPL
   sSQL = "INSERT INTO TMP_Q322" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    
   bln = clsDAO.ExecuteDLL(sSQL)
   '--------------------
   
   clsDAO.Disconnect
   
   
   InsertDTL_PL2B = vbNullString
End Function

Public Function InsertDTL_PL3_3A() As String
    Dim sSQL         As String
    Dim sSQLCol      As String
    Dim sSQLVal      As String

    Dim MAPL         As Variant
    Dim madtnt       As Variant
    Dim ngnop        As Variant
    Dim namQT        As Variant
    Dim CTTN         As Variant
    Dim CK           As Variant
    Dim SOTIEN       As Variant
    Dim TT           As Variant
    Dim DANHAN       As Variant
    Dim SOTIEN2      As Variant
    Dim giatri       As Variant
    
    Dim bln          As Boolean
   
    Dim xmlNodelist  As MSXML.IXMLDOMNode
    Dim xmlNodeChild As MSXML.IXMLDOMNode
    Dim xmlNodeList2 As MSXML.IXMLDOMNode
    Dim lRow         As Long, lCol As Long
   
    Dim count, count2, countRow As Integer
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    sSQLCol = "MAPL,MADTNT,NGNOP,NAMQT,TTHTK,CTTN,CK,SOTIEN,TT,DANHAN,LAN_QUET,SOTIEN2,GIATRI"
    
    namQT = TAX_Utilities_Svr_New.Year
    DANHAN = "''"
    giatri = "''"
    MAPL = "'PL03-3A/TNDN'"
    bln = False

    '----------------------
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
              
        namQT = TAX_Utilities_Svr_New.Year
        
        'demo
        count = 0
        count2 = 0
        'countRow dem dong dung de set GT cho CK va SOTIEN
        countRow = 0
        CTTN = "000"
        Set xmlNodelist = TAX_Utilities_Svr_New.Data(6).lastChild.lastChild

        For count = 0 To xmlNodelist.childNodes.length - 1
            Set xmlNodeList2 = xmlNodelist.childNodes(count)
            
            countRow = 1

            For count2 = 0 To xmlNodeList2.childNodes.length - 1
                Set xmlNodeChild = xmlNodeList2.childNodes(count2)
                ParserCellID fps, GetAttribute(xmlNodeChild, "CellID"), lCol, lRow
                .Sheet = 7
                
                If countRow < 12 Then

                    'ghi cac row title
                    If countRow = 1 Then
                        CTTN = "'000'"
                        TT = "'0'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                
                        CTTN = "'001'"
                        TT = "'0'"
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    End If
                    
                    TT = "'1'"
                    
                    .GetText lCol, lRow, CK

                    If Trim(CK) = "1" Then
                        CK = "'C'"
                    Else
                        CK = "'K'"
                    End If

                    SOTIEN = "0"
                Else

                    'ghi cac row title
                    If countRow = 12 Then
                        CTTN = "'013'"
                        TT = "'0'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    ElseIf countRow = 13 Then
                        CTTN = "'015'"
                        TT = "'0'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    ElseIf countRow = 15 Then
                        CTTN = "'018'"
                        TT = "'0'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                        'ghi row title
                    ElseIf countRow = 17 Then
                        CTTN = "'021'"
                        TT = "'0'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    ElseIf countRow = 19 Then
                        CTTN = "'024'"
                        TT = "'0'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                        
                        CTTN = "'025'"
                        TT = "'0'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                        
                    ElseIf countRow = 23 Then

                        CTTN = "'030'"
                        TT = "'0'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    End If
                    
                    'set lai CTTN

                    
                    TT = "'2'"
                    CK = "''"
                    .GetText lCol, lRow, SOTIEN
                    
                    If SOTIEN = vbNullString Then
                        SOTIEN = "0"
                    End If
                End If
                
'                CTTN = "'" & IIf(Val(CStr(CTTN)) < 10, "00", "0") & CTTN & "'"
                CTTN = CStr(Val(Replace(CTTN, "'", "")) + 1)
                CTTN = "'" & IIf(Val(CStr(CTTN)) < 10, "00", "0") & Val(CStr(CTTN)) & "'"
                
                SOTIEN2 = SOTIEN

                sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                bln = clsDAO.ExecuteDLL(sSQL)
                
                countRow = countRow + 1
            Next
        Next

    End With

    '---------------
    clsDAO.Disconnect
   
    InsertDTL_PL3_3A = vbNullString
End Function

Public Function InsertDTL_PL3_3B() As String
    Dim sSQL         As String
    Dim sSQLCol      As String
    Dim sSQLVal      As String

    Dim MAPL         As Variant
    Dim madtnt       As Variant
    Dim ngnop        As Variant
    Dim namQT        As Variant
    Dim CTTN         As Variant
    Dim CK           As Variant
    Dim SOTIEN       As Variant
    Dim TT           As Variant
    Dim DANHAN       As Variant
    Dim SOTIEN2      As Variant
    Dim giatri       As Variant
    
    Dim bln          As Boolean
   
    Dim xmlNodelist  As MSXML.IXMLDOMNode
    Dim xmlNodeChild As MSXML.IXMLDOMNode
    Dim xmlNodeList2 As MSXML.IXMLDOMNode
    Dim lRow         As Long, lCol As Long
   
    Dim count, count2, countRow As Integer
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    sSQLCol = "MAPL,MADTNT,NGNOP,NAMQT,TTHTK,CTTN,CK,SOTIEN,TT,DANHAN,LAN_QUET,SOTIEN2,GIATRI"
    
    namQT = TAX_Utilities_Svr_New.Year
    DANHAN = "''"
    giatri = "''"
    MAPL = "'PL03-3B/TNDN'"
    bln = False

    '----------------------
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
              
        namQT = TAX_Utilities_Svr_New.Year
        
        'demo
        count = 0
        count2 = 0
        'countRow dem dong dung de set GT cho CK va SOTIEN
        countRow = 0
        CTTN = "000"
        Set xmlNodelist = TAX_Utilities_Svr_New.Data(7).lastChild.lastChild

        For count = 0 To xmlNodelist.childNodes.length - 1
            Set xmlNodeList2 = xmlNodelist.childNodes(count)
            
            countRow = 1

            For count2 = 0 To xmlNodeList2.childNodes.length - 1
                Set xmlNodeChild = xmlNodeList2.childNodes(count2)
                ParserCellID fps, GetAttribute(xmlNodeChild, "CellID"), lCol, lRow
                .Sheet = 8
                
                If countRow < 8 Then

                    'ghi cac row title
                    If countRow = 1 Then
                        CTTN = "'000'"
                        TT = "'0'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                
                        CTTN = "'001'"
                        TT = "'0'"
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                        
                        CTTN = "'002'"
                        TT = "'0'"
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    ElseIf countRow = 6 Then
                        
                        CTTN = "'009'"
                        TT = "'0'"
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    
                    End If
                    
                    TT = "'1'"
                    
                    .GetText lCol, lRow, CK

                    If Trim(CK) = "1" Then
                        CK = "'C'"
                    Else
                        CK = "'K'"
                    End If

                    SOTIEN = "0"
                ElseIf countRow > 7 And countRow < 12 Then
                     If countRow = 8 Then
                        
                        CTTN = "'012'"
                        TT = "'0'"
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    End If
                    TT = "'4'"
                    CK = "''"
                    SOTIEN = "0"
                    .GetText lCol, lRow, giatri
                    If giatri = vbNullString Then
                        giatri = "''"
                    Else
                        If Len(CStr(giatri)) > 11 Then
                            giatri = Left(CStr(giatri), 11)
                        End If
                        giatri = "'" & giatri & "'"
                    End If
                Else
                    'set lai bien GIATRI
                    giatri = "''"
                    'ghi cac row title
                    If countRow = 12 Then
                        CTTN = "'017'"
                        TT = "'4'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    ElseIf countRow = 13 Then

                        CTTN = "'019'"
                        TT = "'4'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    ElseIf countRow = 15 Then

                        CTTN = "'022'"
                        TT = "'4'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    ElseIf countRow = 17 Then

                        CTTN = "'025'"
                        TT = "'4'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    ElseIf countRow = 19 Then
                        CTTN = "'028'"
                        TT = "'4'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                        
                        CTTN = "'029'"
                        TT = "'4'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                        
                        'ghi row title
                    ElseIf countRow = 23 Then

                        CTTN = "'034'"
                        TT = "'4'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    ElseIf countRow = 26 Then

                        CTTN = "'038'"
                        TT = "'4'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    End If
                    
                    TT = "'2'"
                    CK = "''"
                    .GetText lCol, lRow, SOTIEN

                    If SOTIEN = vbNullString Then
                        SOTIEN = "0"
                    End If
                End If

                CTTN = CStr(Val(Replace(CTTN, "'", "")) + 1)
                CTTN = "'" & IIf(Val(CStr(CTTN)) < 10, "00", "0") & Val(CStr(CTTN)) & "'"

                SOTIEN2 = SOTIEN
                sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                bln = clsDAO.ExecuteDLL(sSQL)
                
                countRow = countRow + 1
            Next
        Next

    End With
   
    '---------------
    clsDAO.Disconnect
   
    InsertDTL_PL3_3B = vbNullString
End Function

Public Function InsertDTL_PL3_3C() As String
    Dim sSQL         As String
    Dim sSQLCol      As String
    Dim sSQLVal      As String

    Dim MAPL         As Variant
    Dim madtnt       As Variant
    Dim ngnop        As Variant
    Dim namQT        As Variant
    Dim CTTN         As Variant
    Dim CK           As Variant
    Dim SOTIEN       As Variant
    Dim TT           As Variant
    Dim DANHAN       As Variant
    Dim SOTIEN2      As Variant
    Dim giatri       As Variant
    
    Dim bln          As Boolean
   
    Dim xmlNodelist  As MSXML.IXMLDOMNode
    Dim xmlNodeChild As MSXML.IXMLDOMNode
    Dim xmlNodeList2 As MSXML.IXMLDOMNode
    Dim lRow         As Long, lCol As Long
   
    Dim count, count2, countRow As Integer
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    sSQLCol = "MAPL,MADTNT,NGNOP,NAMQT,TTHTK,CTTN,CK,SOTIEN,TT,DANHAN,LAN_QUET,SOTIEN2,GIATRI"
    
    namQT = TAX_Utilities_Svr_New.Year
    DANHAN = "''"
    giatri = "''"
    MAPL = "'PL03-3C/TNDN'"
    bln = False

    '----------------------
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
              
        namQT = TAX_Utilities_Svr_New.Year
        
        'demo
        count = 0
        count2 = 0
        'countRow dem dong dung de set GT cho CK va SOTIEN
        countRow = 0
        CTTN = "000"
        Set xmlNodelist = TAX_Utilities_Svr_New.Data(8).lastChild.lastChild

        For count = 0 To xmlNodelist.childNodes.length - 1
            Set xmlNodeList2 = xmlNodelist.childNodes(count)
            
            countRow = 1

            For count2 = 0 To xmlNodeList2.childNodes.length - 1
                Set xmlNodeChild = xmlNodeList2.childNodes(count2)
                ParserCellID fps, GetAttribute(xmlNodeChild, "CellID"), lCol, lRow
                .Sheet = 9
                
                If countRow < 11 Then

                    'ghi cac row title
                    If countRow = 1 Then
                        CTTN = "'000'"
                        TT = "'0'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                
                        CTTN = "'001'"
                        TT = "'0'"
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    ElseIf countRow = 9 Then
                        CTTN = "'012'"
                        TT = "'0'"
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                    End If

                    If countRow = 1 Or countRow = 5 Or countRow = 9 Or countRow = 10 Then
                        TT = "'1'"
                        .GetText lCol, lRow, CK

                        If Trim(CK) = "1" Then
                            CK = "'C'"
                        Else
                            CK = "'K'"
                        End If

                        SOTIEN = "0"
                        giatri = "''"
                    ElseIf countRow = 2 Or countRow = 3 Or countRow = 6 Or countRow = 7 Then
                        TT = "'2'"
                        giatri = "''"
                        CK = "''"
                        .GetText lCol, lRow, SOTIEN
                        If SOTIEN = vbNullString Then
                            SOTIEN = "0"
                        End If
                    ElseIf countRow = 4 Or countRow = 8 Then
                        TT = "'4'"
                        CK = "''"
                        SOTIEN = "0"
                        .GetText lCol, lRow, giatri
                        If giatri = vbNullString Then
                            giatri = "''"
                        Else
                            If Len(giatri) > 11 Then
                                giatri = Left(CStr(giatri), 11)
                            End If
                            giatri = "'" & giatri & "'"
                        End If
                    End If

                Else

                    'ghi cac row title
                    If countRow = 11 Then
                        CTTN = "'015'"
                        TT = "'0'"
                        CK = "''"
                        giatri = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                        
                        CTTN = "'016'"
                        TT = "'0'"
                        CK = "''"
                        SOTIEN = "0"
                        SOTIEN2 = SOTIEN
                        sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                        sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                        sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                        bln = clsDAO.ExecuteDLL(sSQL)
                        
                    End If
                    
                    TT = "'2'"
                    CK = "''"
                    giatri = "''"
                    .GetText lCol, lRow, SOTIEN

                    If SOTIEN = vbNullString Then
                        SOTIEN = "0"
                    End If
                End If

                CTTN = CStr(Val(Replace(CTTN, "'", "")) + 1)
                CTTN = "'" & IIf(Val(CStr(CTTN)) < 10, "00", "0") & Val(CStr(CTTN)) & "'"

                SOTIEN2 = SOTIEN
                sSQLVal = MAPL & "," & madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & CTTN & "," & CK & "," & SOTIEN
                sSQLVal = sSQLVal & "," & TT & "," & DANHAN & "," & LAN_QUET & "," & SOTIEN & "," & giatri
                sSQL = "INSERT INTO TMP_Q3ABC" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                bln = clsDAO.ExecuteDLL(sSQL)
                
                countRow = countRow + 1
            Next
        Next

    End With

    '---------------
    InsertDTL_PL3_3C = vbNullString
    clsDAO.Disconnect

End Function
   

Public Function InsertDTL_PL03_4() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim madtnt As Variant
   Dim ngnop As Variant
  
   Dim CK As Variant
   Dim bln  As Boolean
   
   Dim STT As Variant
   Dim TENCSKD As Variant
   Dim SOTIEN1 As Variant
   Dim SOTIEN2 As Variant
   Dim SOTIEN3 As Variant
   Dim SOTIEN4 As Variant
   Dim SOTIEN5 As Variant
   Dim SOTIEN6 As Variant
   Dim SOTIEN7 As Variant
   Dim SOTIEN8 As Variant
   Dim SOTIEN9 As Variant
   Dim DANHAN As Variant
   Dim namQT  As Variant
   
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   sSQLCol = "MADTNT, NGNOP, STT, TENCSKD, SOTIEN1, SOTIEN2, SOTIEN3, SOTIEN4, SOTIEN5, SOTIEN6 "
   sSQLCol = sSQLCol + ", SOTIEN7, SOTIEN8, SOTIEN9,LAN_QUET,DANHAN,NAMQT,TTHTK"
   
   '--------------
   'set NAMQT
   namQT = TAX_Utilities_Svr_New.Year
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
        .Sheet = 10
            .Col = .ColLetterToNumber("C")
             .Row = 26
                 Do
                 
                      .Col = .ColLetterToNumber("C")
                      STT = .Text
                      If Trim(STT) = vbNullString Then
                            STT = "''"
                      Else
                            STT = "'" & STT & "'"
                      End If
                        
                     .Col = .ColLetterToNumber("E")
                     TENCSKD = .Text
                     If Trim(TENCSKD) = vbNullString Then
                        TENCSKD = "''"
                     Else
                        TENCSKD = "'" & TAX_Utilities_Svr_New.Convert(CStr(TENCSKD), UNICODE, TCVN) & "'"
                     End If
                     
                     .Col = .ColLetterToNumber("O")
                     SOTIEN1 = .Value
                     If Trim(SOTIEN1) = vbNullString Then
                            SOTIEN1 = "0"
                     End If
                    
                     .Col = .ColLetterToNumber("R")
                     SOTIEN2 = .Value
                     If Trim(SOTIEN2) = vbNullString Then
                            SOTIEN2 = "0"
                     End If
                     
                     .Col = .ColLetterToNumber("U")
                     SOTIEN3 = .Value
                     If Trim(SOTIEN3) = vbNullString Then
                            SOTIEN3 = "0"
                     End If
                     
                     .Col = .ColLetterToNumber("Z")
                     SOTIEN4 = .Value
                     If Trim(SOTIEN4) = vbNullString Then
                            SOTIEN4 = "0"
                     End If
                     
                     .Col = .ColLetterToNumber("AE")
                     SOTIEN5 = .Value
                     If Trim(SOTIEN5) = vbNullString Then
                            SOTIEN5 = "0"
                     End If
                     
                     .Col = .ColLetterToNumber("AH")
                     SOTIEN6 = .Value
                     If Trim(SOTIEN6) = vbNullString Then
                            SOTIEN6 = "0"
                     End If
                    
                    .Col = .ColLetterToNumber("AK")
                    SOTIEN7 = .Value
                    If Trim(SOTIEN7) = vbNullString Then
                            SOTIEN7 = "0"
                    End If
                    
                    .Col = .ColLetterToNumber("AN")
                    SOTIEN8 = .Value
                    If Trim(SOTIEN8) = vbNullString Then
                            SOTIEN8 = "0"
                    End If
                    
                    .Col = .ColLetterToNumber("AR")
                    SOTIEN9 = .Value
                    If Trim(SOTIEN9) = vbNullString Then
                            SOTIEN9 = "0"
                    End If
                    DANHAN = "''"
                     
                       sSQLVal = madtnt & "," & ngnop & "," & STT & "," & TENCSKD & "," & SOTIEN1 & "," & _
                       SOTIEN2 & "," & SOTIEN3 & "," & SOTIEN4 & "," & SOTIEN5 & "," & SOTIEN6 & "," & SOTIEN7 & _
                       "," & SOTIEN8 & "," & SOTIEN9 & "," & LAN_QUET & "," & DANHAN & "," & namQT & "," & tthtk
                       sSQL = "INSERT INTO TMP_Q35" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                       bln = clsDAO.ExecuteDLL(sSQL)

                     .Col = .ColLetterToNumber("C")
                     .Row = .Row + 1
                  Loop Until .Text = "aa"
    End With
   '---------------
   clsDAO.Disconnect
   
   InsertDTL_PL03_4 = vbNullString
End Function

Public Function InsertDTL_PL03_5() As String
    Dim sSQL        As String
    Dim sSQLCol     As String
    Dim sSQLVal     As String
   
    Dim sSQLCol_CTD As String
    Dim sSQLVal_CTD As String
    Dim namQT       As Variant
    Dim madtnt      As Variant
    Dim ngnop       As Variant
    Dim CTTN        As Variant
    Dim SOTIEN      As Variant
    Dim SOTIEN2     As Variant
    Dim CK          As Variant
    Dim DANHAN      As Variant
    Dim bln         As Boolean
   
    Dim MAPL        As Variant
    Dim TT          As Variant
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
    sSQLCol = "MADTNT, NGNOP, CTTN, SOTIEN, SOTIEN2, NAMQT , TTHTK, LAN_QUET, DANHAN,MAPL,TT"
     
    '--------------
    Dim sSQLCol_PL1I As String
    Dim sSQLVal_PL1I As String
   
    'Dim MADTNT As Variant
    'Dim NGNOP As Variant
    Dim NAMPS        As Variant
    Dim NAMCL        As Variant
    'Dim SOTIEN As Variant
   
    'sSQLCol_PL1I = "MADTNT, NGNOP, NamPS, NamCL, SOTIEN "
    '---------------
   
    namQT = TAX_Utilities_Svr_New.Year
    DANHAN = "''"
    MAPL = "'PL03-5/TNDN'"
    TT = "''"

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

        .Sheet = 11
        
        .GetText .ColLetterToNumber("E"), 21, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
        ' Chi tieu
        CTTN = "'001'"
    End With
     
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'002'"
        .GetText .ColLetterToNumber("E"), 22, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'003'"
        .GetText .ColLetterToNumber("E"), 23, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'004'"
        .GetText .ColLetterToNumber("E"), 24, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'005'"
        .GetText .ColLetterToNumber("E"), 25, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'006'"
        .GetText .ColLetterToNumber("E"), 26, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'007'"
        .GetText .ColLetterToNumber("E"), 27, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'008'"
        .GetText .ColLetterToNumber("E"), 28, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'009'"
        .GetText .ColLetterToNumber("E"), 29, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'010'"
        .GetText .ColLetterToNumber("E"), 30, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'011'"
        .GetText .ColLetterToNumber("E"), 31, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'012'"
        .GetText .ColLetterToNumber("E"), 32, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'013'"
        .GetText .ColLetterToNumber("E"), 33, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'014'"
        .GetText .ColLetterToNumber("E"), 34, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'015'"
        .GetText .ColLetterToNumber("E"), 35, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'016'"
        .GetText .ColLetterToNumber("E"), 36, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'017'"
        .GetText .ColLetterToNumber("E"), 37, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '--------------
   
    With fps
        .Sheet = 11
        ' Chi tieu
        CTTN = "'018'"
        .GetText .ColLetterToNumber("E"), 38, SOTIEN

        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If

        SOTIEN2 = "0"
    End With
   
    sSQLVal = madtnt & "," & ngnop & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & tthtk & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)


'------------------ end ---------------------


    clsDAO.Disconnect
   
    InsertDTL_PL03_5 = vbNullString
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
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_QT6" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT <> " & TAX_Utilities_Svr_New.Year
    
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
        If (Col = .ColLetterToNumber("E") Or Col = .ColLetterToNumber("F")) And Row = 32 Then
            .GetText .ColLetterToNumber("E"), Row, varTemp
            If varTemp <> "" And varTemp <> "../../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    
                    'check ngay nop khong lon hon ngay hien tai
'                      varTemp = ToDate(CStr(varTemp), DDMMYYYY)
'                      If DateDiff("D", Format(varTemp, "mm/dd/yyyy"), Format(Date, "mm/dd/yyyy")) < 0 Then
'                         .Sheet = 1
'                         DisplayMessage "0159", msOKOnly, miCriticalError
'                        .SetActiveCell Col, Row
'                      End If
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
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QT6" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        End If
        
            sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_QT6" & sKyKeKhai & _
            " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT = " & TAX_Utilities_Svr_New.Year & " AND (TTHTK = '1' OR TTHTK = '3' OR TTHTK = '4')"
            
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

Private Function GetMATM() As Object
    Dim rsReturn1 As New ADODB.Recordset
    Dim strSQL1 As String
    Dim clsDAO1 As New TAX_Utilities_Svr_New.clsADO
    'connect to database
    If clsDAO1.Connected = False Then
        clsDAO1.CreateConnectionString spathVat & "\DB_HT\"
        clsDAO1.Connect
    End If
   
    strSQL1 = "SELECT matm, tengoi FROM dmmt2 WHERE mamuc ='1050' order by matm"
    Set rsReturn1 = clsDAO1.Execute(strSQL1)

    Set GetMATM = rsReturn1
    
    Set rsReturn1 = Nothing
    
    'Connect DB success
    'blnSuccess = True
    clsDAO1.Disconnect
End Function
Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL            As String
    Dim rs              As ADODB.Recordset
    Dim vKeKhai         As Variant
    Dim vKeKhai_temp    As Variant
    Dim vKyLB           As Variant
    Dim strTestFileName As String
    Dim fso             As New FileSystemObject
    Dim matkhai         As Variant
     
    Dim flag            As Boolean
    'kiem tra da co to khai chinh thuc chua
    flag = True
    vKyLB = DateSerial(Int(Year(Now())), Int(Month(Now())), 1)
    vKeKhai = DateSerial(Int(Val(TAX_Utilities_Svr_New.Year)), Int(Val(TAX_Utilities_Svr_New.Month)), 1)
    vKeKhai_temp = kyLapBo

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If


        ' Truong hop khong phai PIT
     Do While flag

         strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QT6" & vKeKhai_temp & ".DBF"

         'kiem tra neu co db thi moi thuc hien truy van
         If fso.FileExists(strTestFileName) Then
             
            sSQL = "SELECT MADTNT, NGNOP FROM TMP_QT6" & sKyKeKhai & _
                   " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT = " & TAX_Utilities_Svr_New.Year & " AND (TTHTK = '1' or TTHTK = '4')"
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
             vKeKhai_temp = Year(vKyLB)

             flag = True
         Else
             flag = False
         End If

     Loop

    clsDAO.Disconnect
   
End Function
Private Sub insert_TLDK()
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   
   Dim madtnt As Variant
   Dim ngnop  As Variant
   Dim namQT  As Variant
   Dim SOTT   As Variant
   Dim TENTL  As Variant
   Dim DANHAN As Variant
   Dim bln    As Variant
   
   'khai bao bien de lay gia tri tren form nhap
   Dim SOTT2  As Variant
   Dim TenTL2  As Variant
   Dim i As Integer
   
   sSQLCol = "MADTNT,NGNOP,NAMQT,TTHTK,SOTT,TENTL,DANHAN,LAN_QUET"
   
   
   bln = False
   
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   With fps
        .Sheet = 1
        'kiem tra neu k co du lieu thi k ghi
        .GetText .ColLetterToNumber("C"), 79, TenTL2
        If TenTL2 = vbNullString Then
            Exit Sub
        End If
        'get MADTNT
        .GetText .ColLetterToNumber("F"), 10, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        'get ngay nop
        .GetText .ColLetterToNumber("E"), 32, ngnop
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        'get NAMQT
        namQT = TAX_Utilities_Svr_New.Year
        DANHAN = "''"
        
        .Row = 79
        .Col = .ColLetterToNumber("B")
        i = 0
        Do
            .Col = .ColLetterToNumber("B")
            SOTT2 = .Text
            If SOTT2 = vbNullString Then
                SOTT2 = "''"
            Else
                SOTT2 = "'" & SOTT2 & "'"
            End If
            SOTT = SOTT2
            
            .Col = .ColLetterToNumber("C")
            TenTL2 = .Text
            If TenTL2 = vbNullString Then
                TenTL2 = "''"
            Else
                TenTL2 = Trim(TAX_Utilities_Svr_New.Convert(CStr(TenTL2), UNICODE, TCVN))
                TenTL2 = "'" & TenTL2 & "'"
            End If
            TENTL = TenTL2
            
            sSQLVal = madtnt & "," & ngnop & "," & namQT & "," & tthtk & "," & SOTT & "," & TENTL & "," & DANHAN & "," & LAN_QUET
            
            sSQL = "INSERT INTO TMP_Q3D" & sKyKeKhai & _
                    "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                    
            bln = clsDAO.ExecuteDLL(sSQL)
            
            .Col = .ColLetterToNumber("B")
            i = i + 1
            .Row = i + 79
        Loop Until .Text = "aa"
   End With
   clsDAO.Disconnect
End Sub

''' ParserCellID description
''' Parser CellID string to column and row value
''' Parameter1 pGrid    : the current fpSpread grid
''' Parameter2 pCellID  : the CellID value of the xmlNode need parser
''' Parameter2 pCol     : Column value of cell (Output value)
''' Parameter2 pRow     : Row number of cell (Output value)
Private Sub ParserCellID(pGrid As fpSpread, pCellID As String, pCol As Long, pRow As Long)
    On Error GoTo ErrorHandle
    
    Dim lPos As Long
    
    lPos = InStr(1, pCellID, "_", vbTextCompare)
    
    If lPos > 0 Then
        pCol = pGrid.ColLetterToNumber(Left(pCellID, lPos - 1))
        pRow = Val(Right(pCellID, Len(pCellID) - lPos))
    'ThanhDX added
    Else
        pCol = 0
        pRow = 0
    End If
    
    Exit Sub
ErrorHandle:
    SaveErrorLog "Errors", "ParserCellID", Err.Number, Err.Description
End Sub


'ham lay ten DM cho PL 03-7
Private Function getTenDMPPDG(maDM As Variant) As String
    Dim maDM1 As String
    Dim strDuongDanFile As String
    Dim xmlDomData      As New MSXML.DOMDocument
    Dim xmlNodelist     As MSXML.IXMLDOMNodeList
    Dim xmlNode         As MSXML.IXMLDOMNode
    Dim arrStrName()    As String
    
    maDM1 = CStr(Trim(maDM))
    strDuongDanFile = GetCatalogueFileName
    'lay file chua tenDM
    strDuongDanFile = "..\InterfaceTemplates\Catalogue_PPDG_TNDN.xml"
    
    strDuongDanFile = GetAbsolutePath(strDuongDanFile)
    xmlDomData.Load (strDuongDanFile)
    getTenDMPPDG = ""
    Set xmlNodelist = xmlDomData.getElementsByTagName("Item")
        For Each xmlNode In xmlNodelist
            If Left(GetAttribute(xmlNode, "Value"), 3) = maDM1 Then
                arrStrName = Split(GetAttribute(xmlNode, "Value"), "###")
                getTenDMPPDG = arrStrName(1)
            End If
        Next
End Function
