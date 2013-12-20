VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01ATNDN"
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
        Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
        Dim strDataFileName As String
        Dim xmlNodeListItems As MSXML.IXMLDOMNodeList
        Dim strCboLyDo As String
        Dim strCboLyDoAn As String
        Dim xmlNode As MSXML.IXMLDOMNode
        Dim fldList() As String
    With fps
        .Sheet = 1
                
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
        ' ngay ky hop dong
'        SetDateFormat fps, 1, 19, .ColLetterToNumber("S"), DDMMYYYY
'        .TypeHAlign = TypeHAlignLeft
        ' Ngay nop
        SetDateFormat fps, 1, 32, .ColLetterToNumber("E"), DDMMYYYY
        .Row = 32
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        '.Text = Day(Date) & "/" & Month(Date) & "/" & Year(Date)
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 32, .ColLetterToNumber("M"), DDMMYYYY
        .Row = 32
        .Col = .ColLetterToNumber("M")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
        'ngay hop dong dai ly thue
        SetDateFormat fps, 1, 28, .ColLetterToNumber("N"), DDMMYYYY
        
        
        ' set gia tri gia han
        ' Ky dc gia han moi set gia tri ly do gia han
        strDataFileName = GetAbsolutePath("..\InterfaceTemplates\Catalogue_DM_Lydo_GH_TT16.xml")
        If xmlDomData.Load(strDataFileName) Then
             Set xmlNodeListItems = xmlDomData.getElementsByTagName("Item")
             For Each xmlNode In xmlNodeListItems
                 fldList = Split(GetAttribute(xmlNode, "Value"), "###")
                 strCboLyDo = strCboLyDo + fldList(1) + Chr$(9)
                 strCboLyDoAn = strCboLyDoAn + fldList(0) + Chr$(9)
             Next
         End If
          'set ly do gia han
        .Row = 64
        .Col = .ColLetterToNumber("M")
        .TypeComboBoxList = strCboLyDo
        ' set Ma ly do gia han
        .Row = 60
        .Col = .ColLetterToNumber("K")
        .TypeComboBoxList = strCboLyDoAn
        
    End With
    
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
    Dim i As Integer
    Dim varMaDM As Variant
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
        
        'get ten chi tieu trong danh muc NNKD
        .Sheet = 1
        .Col = .ColLetterToNumber("K")
        .Row = 60
         .GetText .ColLetterToNumber("K"), .Row, varMaDM
        If Trim(varMaDM) <> vbNullString Then
            .Col = .ColLetterToNumber("M")
            .Row = 64
            .Text = getTenLyDoGD(varMaDM)
        End If
        
        
        ' Neu khong co gia han se khong cho chon combobox
         ' Neu check gia han moi cho phep chon GT
        .Col = .ColLetterToNumber("K")
        .Row = 62
        If .Text = "1" Or UCase$(.Text) = "X" Then
            .Col = .ColLetterToNumber("M")
            .Row = 64
            .Lock = False
        Else
            .Col = .ColLetterToNumber("M")
            .Row = 64
            .Lock = True
        End If
        ' End set gia han
    End With
    
    Prepared2 = True
End Function
Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String) 'ByRef strPMaSoTep, ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
    'strPMaSoTep = strMaSoTep
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

        'Set NguoiSuDung
'        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_X)
'        .Row = NGUOI_SU_DUNG_Y
'        If strNguoiSuDung <> "" Then
'            .Text = strNguoiSuDung
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'**************************************************************

         'set Phong xu ly
        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_X)
        .Row = NGUOI_SU_DUNG_Y
        If strNguoiSuDung <> "" Then
            .Text = strNguoiSuDung
            UpdateCell fps, .Col, .Row, .Text
        End If
        
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
        .Text = Format(Date, "dd/mm/yyyy")
         UpdateCell fps, .Col, .Row, .Text
        .Lock = False
        
        .Col = .ColLetterToNumber(NGAY_QUET_X)
        .Row = NGAY_QUET_Y
        '.Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .Text = Format(Date, "dd/mm/yyyy")
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
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strMail_DLT, TCVN, UNICODE))
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
        
        'set text check gia han
        If Trim(TAX_Utilities_Svr_New.Year) = "2011" Then
            .Row = 52
            .Col = .ColLetterToNumber("C")
            .Text = "GIA HAN THEO QD 21/2011/QD-TTG; QD 54/2011/QD-TTg"
        ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" And Trim(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
            .Row = 52
            .Col = .ColLetterToNumber("C")
            .Text = "GIA HAN THEO QD 21/2011/QD-TTG"
        ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" And Trim(TAX_Utilities_Svr_New.ThreeMonths) <> 4 Then
            .Row = 52
            .Col = .ColLetterToNumber("C")
            .Text = "GIA HAN THEO QD 12/2010/QD-TTG"
        ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2009" Then
            .Row = 52
            .Col = .ColLetterToNumber("C")
            .Text = "GIA HAN THEO NQ 30/2008/NQ-CP"
        Else
            .Row = 52
            .Col = .ColLetterToNumber("C")
            .Text = ""
            .Lock = True
        End If
'set STT tren PL_05

    .Sheet = 2
    If .SheetVisible = True Then
        STT = 0
        .Row = 35
               Do
                    .Col = .ColLetterToNumber("C")
                    If (.Row = 35 + STT * 3) Then
                       .Text = str(STT + 3)
                       STT = STT + 1
                    End If
                    .Col = .ColLetterToNumber("C")
                    .Row = .Row + 1
               Loop Until .Text = "aa"
    End If

'        .Col = .ColLetterToNumber(MA_BPQL_X)
'        .Row = MA_BPQL_Y
'        If strMaBPQL <> "" Then
'            .Text = strMaBPQL
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'        .Col = .ColLetterToNumber(TEN_BPQL_X)
'        .Row = TEN_BPQL_Y
'        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        '.EventEnabled(EventAllEvents) = True
        
        ' Dat lai check E_17 Duoc gian thue theo NQ30/NQ-CP
        ' Thoi gian gian thue trong ca nam 2009 cua to khai TNDN
'''        If Trim(TAX_Utilities_Svr_New.Year) = "2009" Then
'''            .Row = 17
'''            .Col = .ColLetterToNumber("E")
'''            UpdateCell fps, .Col, .Row, .Text
'''            .Row = 17
'''            .Col = .ColLetterToNumber("F")
'''            .Text = "GIA HAN NOP THEO NQ30/NQ-CP"
'''        ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" And Trim(TAX_Utilities_Svr_New.ThreeMonths) <> "04" Then
'''            .Row = 17
'''            .Col = .ColLetterToNumber("E")
'''            UpdateCell fps, .Col, .Row, .Text
'''            .Row = 17
'''            .Col = .ColLetterToNumber("F")
'''            .Text = "GIA HAN THEO QD 12/2010/QD-TTG"
'''        ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2011" Or (Trim(TAX_Utilities_Svr_New.Year) = "2010" And Trim(TAX_Utilities_Svr_New.ThreeMonths) = "04") Then
'''            .Row = 17
'''            .Col = .ColLetterToNumber("E")
'''            UpdateCell fps, .Col, .Row, .Text
'''            .Row = 17
'''            .Col = .ColLetterToNumber("F")
'''            .Text = "GIA HAN THEO QD 21/2011/QD-TTG"
'''        Else
'''            ' Neu ngoai hieu luc cua gian thue thi visible cac dieu kien check nay di
'''            .Row = 17
'''            .Col = .ColLetterToNumber("B")
'''            .CellType = CellTypeEdit
'''            .Lock = True
'''            .Text = ""
'''
'''            .Row = 17
'''            .Col = .ColLetterToNumber("E")
'''            .CellType = CellTypeEdit
'''            .Lock = True
'''            .Text = ""
'''            UpdateCell fps, .Col, .Row, ""
'''
'''            .Row = 17
'''            .Col = .ColLetterToNumber("F")
'''            .CellType = CellTypeEdit
'''            .Lock = True
'''            .Text = ""
'''        End If
        
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

Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
' set gia lock cell ly do gia han
   fps.EventEnabled(EventAllEvents) = False
    If fps.ActiveSheet = 1 Then
            ' Neu check gia han moi cho phep chon GT
            If Col = fps.ColLetterToNumber("K") And Row = 62 Then
                fps.Col = fps.ColLetterToNumber("K")
                fps.Row = 62
                If fps.Text = "1" Or UCase$(fps.Text) = "X" Then
                    fps.Col = fps.ColLetterToNumber("M")
                    fps.Row = 64
                    fps.Lock = False
                Else
                    fps.Col = fps.ColLetterToNumber("M")
                    fps.Row = 64
                    fps.Lock = True
                    fps.SetText fps.ColLetterToNumber("M"), 64, ""
                    fps.SetText fps.ColLetterToNumber("K"), 60, "00"
                    UpdateCell fps, fps.ColLetterToNumber("K"), 62, "00"
                    UpdateCell fps, fps.ColLetterToNumber("M"), 64, ""
                End If
            End If
    End If
    fps.EventEnabled(EventAllEvents) = True
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


Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim intIndexCombo As Integer
    With fps
       If .ActiveSheet = 1 Then
            .Col = Col
            .Row = Row
             If Col = .ColLetterToNumber("M") And Row = 64 Then
                intIndexCombo = .TypeComboBoxCurSel
                .Col = .ColLetterToNumber("K")
                .Row = 60
                .TypeComboBoxCurSel = intIndexCombo
                'Update combo C
                UpdateCell fps, .Col, .Row, .Text
            End If
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
        .GetText .ColLetterToNumber("E"), 34, varCheckValue
        .GetText .ColLetterToNumber("M"), 34, varNoteValue
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

Public Function TKRB() As Boolean

   TKRB = True

End Function

Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim strFileNameIDX As String
    
    'kiem tra ton tai tep *.dbf chua
    '    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TKDN" & Right(sKyKeKhai, 4) & ".DBF"
    '    If fso.FileExists(strFileNameHDR) = False Then
    ''    Dim strTepmauHDR As String
    ''        strTepmauHDR = spathVat & "\tepmau\TKDN.DBF"
    ''        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    '        isExistFile = False
    '        DisplayMessage "0109", msOKOnly, miCriticalError
    '        Exit Function
    '    Else
    '        isExistFile = True
    '    End If
    '
    'Ghi vao DB trung gian kiem tra file .dbf
    'dhdang sua
    'ngay 24/09/2010
    '-----------------------------
        strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TKDN" & Right(sKyKeKhai, 4) & ".DBF"
        strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TNDN" & Right(sKyKeKhai, 4) & ".DBF"
        Dim strFileNamePL1 As String
        strFileNamePL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_PN1" & sKyKeKhai & ".DBF"
        Dim strFileNameBS As String
        strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"
        'file HDR
        If fso.FileExists(strFileNameHDR) = False Then
            Dim strTepmauHDR As String
            strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_TKDN.DBF"
            fso.CopyFile strTepmauHDR, strFileNameHDR, False
        End If
        'file dtl
        If fso.FileExists(strFileNameDTL) = False Then
            Dim strTepmauDTL As String
            strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_TNDN.DBF"
            fso.CopyFile strTepmauDTL, strFileNameDTL, False
        End If
        'file phu luc
        If fso.FileExists(strFileNamePL1) = False Then
            Dim strTepmauPL1 As String
            strTepmauPL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_PN1.DBF"
            fso.CopyFile strTepmauPL1, strFileNamePL1, False
        End If
        
        'file bo xung
        If fso.FileExists(strFileNameBS) = False Then
            Dim strTepmauBS As String
            strTepmauBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_BS.DBF"
            fso.CopyFile strTepmauBS, strFileNameBS, False
        End If


        isExistFile = True

'------------------------------------------
'
'    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TNDN" & Right(sKyKeKhai, 4) & ".DBF"
'    If fso.FileExists(strFileNameDTL) = False Then
'    Dim strTepmauDTL As String
'        strTepmauDTL = spathVat & "\tepmau\TNDN.DBF"
'        fso.CopyFile strTepmauDTL, strFileNameDTL, False
'    End If
    
'    'Xoa file *.IDX
'    strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TKDN" & Right(sKyKeKhai, 4) & ".IDX"
'    If fso.FileExists(strFileNameIDX) = True Then
'        fso.DeleteFile strFileNameIDX, True
'    End If
'
'    strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TNDN" & Right(sKyKeKhai, 4) & ".IDX"
'    If fso.FileExists(strFileNameIDX) = True Then
'        fso.DeleteFile strFileNameIDX, True
'    End If
'
'    strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "BS" & Right(sKyKeKhai, 4) & ".IDX"
'    If fso.FileExists(strFileNameIDX) = True Then
'        fso.DeleteFile strFileNameIDX, True
'    End If
    
    If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
   End If
        
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TKDN" & Right(sKyKeKhai, 4) & _
          " WHERE MATKHAI = '01A/TNDN' AND MAPP='1' AND MADTNT= " & "'" & strMST & "'" & _
          " AND MATKHAI='01A/TNDN' AND QUI = '" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'" & _
          " AND TTHTK = '1'"
   Set rs = clsConn.Execute(sSQL)
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
   clsConn.Disconnect
    
    
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
   
   If clsConn.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsConn.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
   
   
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TKDN" & Right(sKyKeKhai, 4) & " WHERE MATKHAI = '01A/TNDN' AND MADTNT = '" & strMST & "'"
   
    
    
    
    
   Set rs = clsConn.Execute(sSQL)
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

Public Function UpdateTHUETKY2() As Boolean
    Dim sSQL As String
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   sSQL = "UPDATE TMP_TKDN" & Right(sKyKeKhai, 4) & _
        " SET THUETKY2=1" & _
        " WHERE MATKHAI = '01A/TNDN' AND  MAPP='1' AND MADTNT= " & "'" & strMST & "'" & _
        " AND QUI = '" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
    
   clsDAO.ExecuteDLL sSQL
   
   UpdateTHUETKY2 = True
   
   clsDAO.Disconnect

End Function


Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim kylbo As Variant
    Dim rs As ADODB.Recordset
    Dim idTKHDR As String 'Tam thoi ID cua to khai nay la MST & KYKK
    If clsConn.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsConn.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
  ' idTKHDR = strMST & sKyKeKhai
   sSQL = "DELETE FROM TMP_TKDN" & Right(sKyKeKhai, 4) & _
   " WHERE MATKHAI = '01A/TNDN' AND MAPP='1' AND MADTNT= " & "'" & strMST & "'" & _
   " AND QUI='" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
    
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
   sSQL = "DELETE FROM TMP_TNDN" & Right(sKyKeKhai, 4) & " WHERE MATKHAI = '01A/TNDN' AND MADTNT = '" & strMST & "'" & _
   " AND QUI='" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
   
   
   clsConn.ExecuteDLL sSQL
   
  
   
   XoaTKTT = True
   
   clsConn.Disconnect

End Function

Public Function InsertDTL() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim sSQLColPL As String
   Dim sSQLValPL As String
   Dim bln  As Boolean
   Dim i As Integer
   
   Dim IDHDR As Variant
   Dim MST As Variant
   Dim KYKK As Variant
   Dim LoaiTK As Variant
   ' Cac truong thuoc to khai
   Dim mact As Variant
   Dim giatri As Variant
   ' Cac truong thuoc phu luc
   Dim SOHD As Variant
   Dim NGAYHD As Variant
   Dim NGUOINHAN As Variant
   Dim DCLODAT As Variant
   Dim DIENTICH As Variant
   Dim GIA1M As Variant
   Dim GTCQ As Variant
   Dim GIA1MUB As Variant
   Dim GTCQUB As Variant
   
   Dim QUI As Variant
   Dim madtnt As Variant
   Dim MATHUE As Variant
   Dim MaMuc As Variant
   Dim maTM As Variant
   Dim MAPP As Variant
   Dim ngnop As Variant
   Dim CTIEUTNDN As Variant
   Dim SOTIEN As Variant
   Dim SOTIEN2 As Variant
   Dim tt As Variant
   Dim LoaiTien As Variant
   Dim matkhai As Variant
   Dim DANHAN As Variant

   sSQLCol = "QUI, MADTNT, MATHUE, MAMUC, MATM, MAPP, NGNOP, CTIEUTNDN, SOTIEN, SOTIEN2, TT, LOAITIEN ,TTHTK, MATKHAI, LAN_QUET, DANHAN"
   
    'tinh so lan quet
    'dhdang
    'ngay 24/09/10
    '------------------------------------------------
     DANHAN = "''"
   
   
   If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
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
        
        
         .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        MATHUE = "'08'"
        
        ' sua MAMUC & MATM tu nam 2009
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 10
        CTIEUTNDN = "'001'"
        
        .GetText .ColLetterToNumber("U"), 40, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
        
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
'-----------------------

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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 11
        CTIEUTNDN = "'002'"
        
        .GetText .ColLetterToNumber("U"), 41, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)

'--------------------------

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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 12
        CTIEUTNDN = "'003'"
        
        .GetText .ColLetterToNumber("U"), 42, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        
        matkhai = "'01A/TNDN'"
   
   End With

   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
'---------------------

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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 13
        CTIEUTNDN = "'004'"
        
        .GetText .ColLetterToNumber("U"), 43, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
        
   End With
 

   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
'-------------------------

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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 14
        CTIEUTNDN = "'005'"
        
        .GetText .ColLetterToNumber("U"), 44, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
    
   End With
 

   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
'---------------------


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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 15
        CTIEUTNDN = "'006'"
        
        .GetText .ColLetterToNumber("U"), 45, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
    
   End With

   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 16
        CTIEUTNDN = "'007'"
        
        .GetText .ColLetterToNumber("U"), 46, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
        
   End With

   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)

'---------------

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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 17
        CTIEUTNDN = "'008'"
        
        .GetText .ColLetterToNumber("U"), 47, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
        
   End With

   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)

   '--------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 18
        CTIEUTNDN = "'009'"
        
        .GetText .ColLetterToNumber("U"), 48, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
   End With

   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
   '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'010'"
        
        .GetText .ColLetterToNumber("U"), 49, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)

   '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'011'"
        
        .GetText .ColLetterToNumber("U"), 50, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)

   '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'012'"
        
        .GetText .ColLetterToNumber("U"), 51, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
       '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'013'"
        
        .GetText .ColLetterToNumber("U"), 52, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
      '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'014'"
        
        .GetText .ColLetterToNumber("U"), 53, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
      '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'015'"
        
        .GetText .ColLetterToNumber("U"), 54, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
      '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'016'"
        
        .GetText .ColLetterToNumber("U"), 55, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
      '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'017'"
        
        .GetText .ColLetterToNumber("U"), 56, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
      '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'018'"
        
        .GetText .ColLetterToNumber("U"), 57, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
      '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'019'"
        
        .GetText .ColLetterToNumber("U"), 58, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
      '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'020'"
        
        .GetText .ColLetterToNumber("U"), 59, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
      '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'021'"
        
        .GetText .ColLetterToNumber("M"), 66, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
      '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'022'"
        
        .GetText .ColLetterToNumber("M"), 68, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
      '-------------------
   
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
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        tt = "''"
        LoaiTien = "''"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' Chi tieu 19
        CTIEUTNDN = "'023'"
        
        .GetText .ColLetterToNumber("M"), 70, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        Else
            SOTIEN = SOTIEN
        End If
        SOTIEN2 = SOTIEN
        matkhai = "'01A/TNDN'"
       
   End With
   sSQLVal = QUI & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ","
   sSQLVal = sSQLVal & ngnop & "," & CTIEUTNDN & "," & SOTIEN & "," & SOTIEN2 & "," & tt & "," & LoaiTien & ","
   sSQLVal = sSQLVal & tthtk & "," & matkhai & "," & LAN_QUET & "," & DANHAN
   
    sSQL = "INSERT INTO TMP_TNDN" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
'   If TAX_Utilities_Svr_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1 Then
'        InsertDTL = InsertDTL_PL1
'   End If
   clsConn.Disconnect
   
   InsertDTL = vbNullString
   
End Function

'Public Function InsertDTL_PL1() As String
'    Dim sSQL     As String
'    Dim sSQLCol  As String
'    Dim sSQLVal  As String
'
'    Dim MADTNT   As Variant
'    Dim NGNOP    As Variant
'    Dim MATKHAI  As Variant
'    Dim TTHTK    As Variant
'    Dim QUI      As Variant
'    Dim MATM     As Variant
'    Dim MAMUC    As Variant
'    Dim MATHUE   As Variant
'    Dim MAPP     As Variant
'    Dim SOHDON   As Variant
'    Dim NGHDON   As Variant
'    Dim NGUOINH  As Variant
'    Dim DCHILO   As Variant
'    Dim DIENTICH As Variant
'    Dim GIADAT1  As Variant
'    Dim GIATRI1  As Variant
'    Dim GIADAT2  As Variant
'    Dim GIATRI2  As Variant
'
'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'    End If
'
'    sSQLCol = "MADTNT,NGNOP,MATKHAI,TTHTK,QUI ,MATM ,MAMUC ,MATHUE ,MAPP ,SOHDON ,NGHDON ,NGUOINH ,DCHILO ,DIENTICH ,GIADAT1 ,GIATRI1 ,GIADAT2 ,GIATRI2"
'
'    DANHAN = "''"
'    MAPL = "'PL01-5A/TNDN'"
'    TT = "''"
'
'    With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("F"), 10, MADTNT
'
'        If Trim(MADTNT) = vbNullString Then
'            MADTNT = "''"
'        Else
'            MADTNT = "'" & MADTNT & "'"
'        End If
'
'        .GetText .ColLetterToNumber("E"), 32, NGNOP
'
'        'NGNOP = Date
'        If Trim(NGNOP) = vbNullString Then
'            NGNOP = "CTOD('')"
'        Else
'            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
'            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
'        End If
'
'        .Sheet = 2
'        .GetText .ColLetterToNumber("F"), 7, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'        CTTN = "'001'"
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'002'"
'        .GetText .ColLetterToNumber("E"), 24, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'003'"
'        .GetText .ColLetterToNumber("E"), 25, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'004'"
'        .GetText .ColLetterToNumber("E"), 26, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'005'"
'        .GetText .ColLetterToNumber("E"), 27, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'006'"
'        .GetText .ColLetterToNumber("E"), 28, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'007'"
'        .GetText .ColLetterToNumber("E"), 29, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'008'"
'        .GetText .ColLetterToNumber("E"), 30, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'009'"
'        .GetText .ColLetterToNumber("E"), 31, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'010'"
'        .GetText .ColLetterToNumber("E"), 32, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'011'"
'        .GetText .ColLetterToNumber("E"), 33, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'012'"
'        .GetText .ColLetterToNumber("E"), 34, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'013'"
'        .GetText .ColLetterToNumber("E"), 35, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        ' Chi tieu
'        CTTN = "'014'"
'        .GetText .ColLetterToNumber("E"), 36, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'015'"
'        .GetText .ColLetterToNumber("E"), 37, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'016'"
'        .GetText .ColLetterToNumber("E"), 38, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'
'    '--------------
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'017'"
'        .GetText .ColLetterToNumber("E"), 39, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'
'        '        .GetText .ColLetterToNumber("F"), 4, NamCL
'        '        If Trim(NamCL) = vbNullString Then
'        '            NamCL = "0"
'        '        End If
'        '        .GetText .ColLetterToNumber("F"), 6, SOTIEN
'        '        If Trim(SOTIEN) = vbNullString Then
'        '            SOTIEN = "0"
'        '        End If
'
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'018'"
'        .GetText .ColLetterToNumber("E"), 40, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'
'    With fps
'        .Sheet = 2
'        ' Chi tieu
'        CTTN = "'019'"
'        .GetText .ColLetterToNumber("E"), 41, SOTIEN
'
'        If Trim(SOTIEN) = vbNullString Then
'            SOTIEN = "0"
'        End If
'
'        SOTIEN2 = "0"
'    End With
'
'    sSQLVal = MADTNT & "," & NGNOP & "," & CTTN & "," & SOTIEN & "," & SOTIEN2 & "," & namQT & "," & TTHTK & "," & LAN_QUET & "," & DANHAN & "," & MAPL & "," & TT
'    sSQL = "INSERT INTO TMP_Q31A" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'    bln = clsDAO.ExecuteDLL(sSQL)
'    '--------------
'    clsDAO.Disconnect
'
'    InsertDTL_PL1 = vbNullString
'End Function


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
    Dim DANHAN As Variant
    Dim i As Integer, iCol As Long, iRow As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim strFileName As String
    Dim fso As New FileSystemObject
    
    'khai bao them 2 bien luu so ngay nop chap va so tien nop cham
    Dim SONGAYNC As Variant
    Dim SOTIENNC As Variant
    
   'kiem tra ton tai tep *.dbf chua
'    strFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "BS" & sKyKeKhai & ".DBF"
'    If fso.FileExists(strFileName) = False Then
'    Dim strTepmau As String
'        strTepmau = spathVat & "\tepmau\BS.DBF"
'        fso.CopyFile strTepmau, strFileName, False
'    End If
   
   
   sSQLCol = "MADTNT, KYKKHAI, MATKHAI, MATHUE, MAMUC, MATM, MAPP, NGNOP, TTHTK, KYLBO , MACT, MACT2, MACT3, STT, STT2, STTIN," & _
             " SOKK, SODC, SOCL ,LAN_QUET, DANHAN,SOTIENNC,SONGAYNC"
   
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
        
        kykkhai = "'" & TAX_Utilities_Svr_New.ThreeMonths & "/" & TAX_Utilities_Svr_New.Year & "'"
        matkhai = "'01A/TNDN'"
        MATHUE = "'08'"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        
        .GetText .ColLetterToNumber("E"), 32, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 30, kylbo
        If Trim(kylbo) = vbNullString Then
            kylbo = "CTOD('')"
        Else
            If Len(Trim(kylbo)) = 6 Then
                kylbo = "'0" & kylbo & "'"
            Else
                kylbo = "'" & kylbo & "'"
            End If

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
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & LAN_QUET & "," & DANHAN & "," & SOTIENNC & "," & SONGAYNC
                       
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
                MACT2 = "'050'"
            ElseIf MACT2 = "30" Then
                MACT2 = "'031'"
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
            
            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & LAN_QUET & "," & DANHAN & "," & SOTIENNC & "," & SONGAYNC
                           
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
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & LAN_QUET & "," & DANHAN & "," & SOTIENNC & "," & SONGAYNC
                       
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
                '                ElseIf MACT2 = "23" Then
                '                    MACT2 = "'041'"
            ElseIf MACT2 = "24" Then
                MACT2 = "'050'"
            ElseIf MACT2 = "21" Then
                MACT2 = "'039'"
            ElseIf MACT2 = "31" Then
                MACT2 = "'050'"
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
            
            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & LAN_QUET & "," & DANHAN & "," & SOTIENNC & "," & SONGAYNC
                           
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
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & LAN_QUET & "," & DANHAN & "," & SOTIENNC & "," & SONGAYNC
                       
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        .Row = .Row + 4
        mact = "''"
        .GetText .ColLetterToNumber("BE"), .Row, MACT2

        If Trim(MACT2) = vbNullString Then
            MACT2 = "''"
        ElseIf MACT2 = "14" Then
            MACT2 = "'041'"
        ElseIf MACT2 = "16" Then
            MACT2 = "'041'"
            '                ElseIf MACT2 = "23" Then
            '                    MACT2 = "'041'"
        ElseIf MACT2 = "24" Then
            MACT2 = "'050'"
        ElseIf MACT2 = "21" Then
            MACT2 = "'039'"
        ElseIf MACT2 = "31" Then
            MACT2 = "'050'"
        Else
            MACT2 = "'000'"
        End If
                
        MACT3 = "'2'"
        STT = 3
        STT2 = 1
        STTIN = "'1'"
                
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
            
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & LAN_QUET & "," & DANHAN & "," & SOTIENNC & "," & SONGAYNC
                           
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        If MACT2 <> "''" Then
            bln = clsDAO.ExecuteDLL(sSQL)
        End If

    End With
   
    clsDAO.Disconnect
    InsertDTL_KHBS = vbNullString
End Function



Public Function InsertHDR() As String
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
        Dim TENNV      As Variant
    Dim CHUNGCHI   As Variant

    Dim NGAYNOP As Variant
    Dim KYLB As Variant
    Dim bln  As Boolean
    
    Dim QUI As Variant
    Dim SHTEP As Variant
    Dim matkhai As Variant
    Dim madtnt As Variant
    Dim MATHUE As Variant
    Dim MaMuc As Variant
    Dim maTM As Variant
    Dim MAPP As Variant
    Dim kylbo As Variant
    Dim ngnop As Variant
    Dim NGNOP_D As Variant
    Dim NgNhap As Variant
    Dim LoaiTien As Variant
    Dim DoanhSo As Variant
    Dim GTTinhThue As Variant
    Dim ThueTKy As Variant
    Dim DaGhi As Variant
    Dim MaVach As Variant
    Dim HANNOP2 As Variant
    Dim ThueTKy2 As Variant
    Dim NGAYQUET As Variant
    Dim DATHAY As Variant
    Dim CHKGIAHAN  As Variant
    Dim DAGHI1 As Variant
    Dim AUTO_CAL As Variant
    Dim DSTSUAT As Variant
    Dim GHICHU  As Variant
    Dim DANHAN As Variant
    Dim MSTDL As Variant
    Dim LanBS As Variant
    Dim MANNKD As Variant
    Dim GIAHAN As Variant
    Dim MALD    As Variant
    
    sSQLCol = "QUI, SHTEP, MATKHAI, MADTNT, MATHUE, MAMUC, MATM, MAPP, KYLBO, NGNOP, NGNHAP, LOAITIEN, TTHTK, "
    sSQLCol = sSQLCol & "DOANHSO, GTTINHTHUE, THUETKY, DAGHI, DAGHI1, DAGHI2, MAVACH, HANNOP, HANNOP2, THUETKY2, AUTO_CAL, DSTSUAT, GHICHU, LAN_QUET, DANHAN,MADLT,LANBS,MANNKD,GIAHAN,MALD,TENNV,CHUNGCHI"
           
'    If clsConn.Connected = False Then
'        clsConn.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsConn.Connect
'    End If
With fps
        .Sheet = 1
                'lay ten nhan vien dai ly thue, chung chi hanh nghe
        .GetText .ColLetterToNumber("H"), 73, TENNV
        .GetText .ColLetterToNumber("H"), 75, CHUNGCHI

        .GetText .ColLetterToNumber("F"), 12, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 32, kylbo
        If Trim(kylbo) = vbNullString Then
            kylbo = "''"
        Else
            If Len(Trim(kylbo)) = 6 Then
                kylbo = "'0" & kylbo & "'"
            Else
                kylbo = "'" & kylbo & "'"
            End If
        End If

        .GetText .ColLetterToNumber("E"), 34, ngnop
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            NGNOP_D = ngnop
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
        'set DAGHI2 theo checkbox loi dinh dinh
        .Col = .ColLetterToNumber("E")
        .Row = 36
'        chkLoiDDValue = .Value
        If .Value = 1 Then
            DAGHI2 = ".T."
        Else
            DAGHI2 = ".F."
        End If
              
        .GetText .ColLetterToNumber("M"), 34, NgNhap
        If Trim(NgNhap) = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = ToDate(Trim(NgNhap), DDMMYYYY)
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If
        MATHUE = "'08'"
        matkhai = "'01A/TNDN'"
        DANHAN = "''"
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If
        MAPP = "'1'"
        LoaiTien = "''"
                
        'get MaNNKD
        .GetText .ColLetterToNumber("Q"), 80, MANNKD
        If Trim(MANNKD) = vbNullString Then
            MANNKD = "''"
        Else
            MANNKD = "'" & MANNKD & "'"
        End If
                    
        .GetText .ColLetterToNumber("U"), 40, DoanhSo
        If Trim(DoanhSo) = vbNullString Then
            DoanhSo = "0"
        Else
            DoanhSo = Trim(DoanhSo)
        End If
        
        .GetText .ColLetterToNumber("U"), 45, GTTinhThue
        If Trim(GTTinhThue) = vbNullString Then
            GTTinhThue = "0"
        Else
            GTTinhThue = Trim(GTTinhThue)
        End If
        
        .GetText .ColLetterToNumber("U"), 48, ThueTKy
        If Trim(ThueTKy) = vbNullString Then
            ThueTKy = "0"
        Else
            ThueTKy = Trim(ThueTKy)
        End If
            
        DaGhi = ".T."
        
        AUTO_CAL = ".T."
        
        MaVach = ".T."
        
        DSTSUAT = "''"
        GHICHU = "''"
        
        'Check gia han nop
        .GetText .ColLetterToNumber("K"), 62, CHKGIAHAN
'        If Trim(CHKGIAHAN) = "1" Or Trim(CHKGIAHAN) = "x" Then
'            DAGHI1 = ".T."
'            If Trim(TAX_Utilities_Svr_New.Year) = "2009" Then
'               If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
'                HANNOP = "CTOD('" & Format("01/02/2010", "mm/dd/yyyy") & "')"
'               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
'                HANNOP = "CTOD('" & Format("04/05/2010", "mm/dd/yyyy") & "')"
'               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
'                HANNOP = "CTOD('" & Format("30/07/2010", "mm/dd/yyyy") & "')"
'               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
'                HANNOP = "CTOD('" & Format("01/11/2010", "mm/dd/yyyy") & "')"
'               End If
'            ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" Then
'               If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
'                HANNOP = "CTOD('" & Format("30/07/2010", "mm/dd/yyyy") & "')"
'               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
'                HANNOP = "CTOD('" & Format("30/11/2010", "mm/dd/yyyy") & "')"
'               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
'                HANNOP = "CTOD('" & Format("31/01/2011", "mm/dd/yyyy") & "')"
'               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
'                HANNOP = "CTOD('" & Format("03/05/2011", "mm/dd/yyyy") & "')"
'               End If
'            Else
'                If Trim(HANNOP) = vbNullString Then
'                    HANNOP = "CTOD('')"
'                Else
'                    HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
'                End If
'            End If
'        Else
'            DAGHI1 = ".F."
'            If Trim(HANNOP) = vbNullString Then
'                HANNOP = "CTOD('')"
'            Else
'                HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
'            End If
'        End If

        ' sua lai format dinh dang han nop
        If Trim(CHKGIAHAN) = "1" Or Trim(CHKGIAHAN) = "x" Then
            DAGHI1 = ".T."
            If Trim(TAX_Utilities_Svr_New.Year) = "2009" Then
               If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
                HANNOP = "01/02/2010"
               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
                HANNOP = "04/05/2010"
               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
                HANNOP = "30/07/2010"
               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
                HANNOP = "01/11/2010"
               End If
            ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2010" Then
               If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
                HANNOP = "30/07/2010"
               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
                HANNOP = "01/11/2010"
               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
                HANNOP = "08/02/2011"
               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
                'HANNOP = "03/05/2011"
                HANNOP = "30/01/2012"
               End If
            ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2011" Then
                If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
                HANNOP = "02/05/2012"
               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
                HANNOP = "30/07/2012"
               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
                HANNOP = "30/10/2012"
               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 4 Then
                HANNOP = "01/04/2013"
               End If
            ElseIf Trim(TAX_Utilities_Svr_New.Year) = "2013" Then
                If Val(TAX_Utilities_Svr_New.ThreeMonths) = 1 Then
                HANNOP = "30/10/2013"
               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 2 Then
                HANNOP = "30/10/2013"
               ElseIf Val(TAX_Utilities_Svr_New.ThreeMonths) = 3 Then
                HANNOP = "30/01/2014"
               End If
            End If
            If Trim(HANNOP) = vbNullString Then
                    HANNOP = "CTOD('')"
            Else
                    HANNOP = DateSerial(Int(Mid$(HANNOP, 7, 4)), Int(Mid$(HANNOP, 4, 2)), Int(Mid$(HANNOP, 1, 2)))
                    If tthtk = "'2'" And DateDiff("d", HANNOP, NGNOP_D) > 0 And (Trim(TAX_Utilities_Svr_New.Year) = "2011" Or (Trim(TAX_Utilities_Svr_New.Year) = "2010" And Val(TAX_Utilities_Svr_New.ThreeMonths) = 4)) Then
                    HANNOP = NGNOP_D
                    End If
                    HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
            End If
        Else
            DAGHI1 = ".F."
            If Trim(HANNOP) = vbNullString Then
                HANNOP = "CTOD('')"
            Else
                ' nvanhai sua sai dinh dang format ngay thang
                HANNOP = DateSerial(Int(Mid$(HANNOP, 7, 4)), Int(Mid$(HANNOP, 4, 2)), Int(Mid$(HANNOP, 1, 2)))
                If tthtk = "'2'" And DateDiff("d", HANNOP, NGNOP_D) > 0 And (Trim(TAX_Utilities_Svr_New.Year) = "2011" Or (Trim(TAX_Utilities_Svr_New.Year) = "2010" And Val(TAX_Utilities_Svr_New.ThreeMonths) = 4)) Then
                    HANNOP = NGNOP_D
                End If
                HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
            End If
        End If

        ' end
        
        HANNOP2 = HANNOP
        ThueTKy2 = "0"
        QUI = "'" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
        
'        khong con phai luu shtep
'        .GetText .ColLetterToNumber("M"), 10, SHTEP
        If Trim(SHTEP) = vbNullString Then
            SHTEP = "''"
        Else
            SHTEP = "'" & SHTEP & "'"
        End If
'       lay mstdl
        .GetText .ColLetterToNumber("F"), 20, MSTDL
        If Trim(MSTDL) = vbNullString Then
            MSTDL = "''"
        Else
            MSTDL = "'" & MSTDL & "'"
        End If
'       lay so lan bs
        .GetText .ColLetterToNumber("T"), 5, LanBS
        If Trim(LanBS) = vbNullString Then
            LanBS = "0"
        End If
'        .GetText .ColLetterToNumber("M"), 30, LAN_QUET
'        LAN_QUET = TinhSoLanQuet
        If Trim(LAN_QUET) = vbNullString Then
            LAN_QUET = "1"
        End If
        DANHAN = "''"
        
        '       luu check gia han va ly do
        'get check gia han
        .GetText .ColLetterToNumber("K"), 62, GIAHAN
        'get ly do gia han
        .GetText .ColLetterToNumber("K"), 60, MALD
        If Trim(GIAHAN) = "1" Or UCase(Trim(GIAHAN)) = "X" Then
            GIAHAN = ".T."
'            If Trim(MALD) = "01" Then
'                MALD = "'01'"
'            ElseIf Trim(MALD) = "02" Then
'                MALD = "'02'"
'            Else
'                MALD = "Null"
'            End If
             MALD = "'" & MALD & "'"
        Else
            GIAHAN = ".F."
            MALD = "Null"
        End If
End With

    If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
    End If
    '---------------------------------------------------
    sSQLVal = QUI & "," & SHTEP & "," & matkhai & "," & madtnt & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & ", " & kylbo & ", " & ngnop & "," & NgNhap
    sSQLVal = sSQLVal & "," & LoaiTien & "," & tthtk & "," & DoanhSo & "," & GTTinhThue & "," & ThueTKy & "," & DaGhi & "," & DAGHI1 & "," & DAGHI2 & ","
    sSQLVal = sSQLVal & MaVach & "," & HANNOP & "," & HANNOP2 & "," & ThueTKy2 & "," & AUTO_CAL & ", " & DSTSUAT & "," & GHICHU & "," & LAN_QUET & "," & DANHAN & "," & MSTDL & "," & LanBS & "," & MANNKD & "," & GIAHAN & "," & MALD & ",'" & TENNV & "','" & CHUNGCHI & "'"
    
    sSQL = "INSERT INTO TMP_TKDN" & Right(sKyKeKhai, 4) & _
     "( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
     
    bln = clsConn.ExecuteDLL(sSQL)
    clsConn.Disconnect
    InsertHDR = vbNullString
    
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
'                   checkNgayNop vNgayNop, varTemp
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
'dhdang
'ham tinh so lan quet cua to khai (Ghi vao DB trung gian)
'ngay 24/09/2010

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
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TKDN" & Right(sKyKeKhai, 4) & ".DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
        End If
        
            sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_TKDN" & Right(sKyKeKhai, 4) & _
                  " WHERE MATKHAI = '01A/TNDN' AND MAPP='1' AND MADTNT= " & "'" & strMST & "'" & _
                  " AND MATKHAI='01A/TNDN' AND QUI = '" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'"
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 LAN_QUET1 = rs.Fields("LAN_QUET")
                 TinhSoLanQuet = LAN_QUET1
            Else
                 TinhSoLanQuet = 0
            End If
            clsConn.Disconnect
    Else
    TinhSoLanQuet = 0
    End If
'------------------------------------------
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
    vKeKhai = DateSerial(Int(Val(TAX_Utilities_Svr_New.Year)), Int(Val(TAX_Utilities_Svr_New.ThreeMonths)), 1)
    vKeKhai_temp = sKyKeKhai
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   Do While flag
        strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TKDN" & Right(sKyKeKhai, 4) & ".DBF"
        'ktra neu ton tai db thi moi thuc hien truy van
        If fso.FileExists(strTestFileName) Then
            sSQL = "SELECT MADTNT, NGNOP FROM TMP_TKDN" & Right(vKeKhai_temp, 4) & _
              " WHERE MATKHAI = '01A/TNDN' AND MAPP='1' AND MADTNT= " & "'" & strMST & "'" & _
              " AND MATKHAI='01A/TNDN' AND QUI = '" & Right(TAX_Utilities_Svr_New.ThreeMonths, 1) & "/" & TAX_Utilities_Svr_New.Year & "'" & _
              " AND (TTHTK = '1' or TTHTK = '3' or TTHTK = '4')"
            
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

'ham lay ten DM NNKD
Private Function getTenLyDoGD(maGH As Variant) As String
    Dim maDM1 As String
    Dim strDuongDanFile As String
    Dim xmlDomData      As New MSXML.DOMDocument
    Dim xmlNodelist     As MSXML.IXMLDOMNodeList
    Dim xmlNode         As MSXML.IXMLDOMNode
    Dim arrStrName()    As String
    
    getTenLyDoGD = ""
    maDM1 = CStr(Trim(maGH))
    'strDuongDanFile = GetCatalogueFileName
    'lay file chua tenDM
    strDuongDanFile = "..\InterfaceTemplates\Catalogue_DM_Lydo_GH_TT16.xml"
    strDuongDanFile = GetAbsolutePath(strDuongDanFile)
    xmlDomData.Load (strDuongDanFile)
    
    Set xmlNodelist = xmlDomData.getElementsByTagName("Item")
        For Each xmlNode In xmlNodelist
            If Left(GetAttribute(xmlNode, "Value"), 2) = maDM1 Then
                arrStrName = Split(GetAttribute(xmlNode, "Value"), "###")
                getTenLyDoGD = arrStrName(1)
            End If
        Next
End Function

Public Function InsertDTL_PL5A() As String
    Dim sSQL      As String
    Dim sSQLCol   As String
    Dim sSQLVal   As String
    Dim bln       As Boolean
    Dim maTM As Variant
    Dim MaMuc As Variant
    Dim MAPP As Variant
    Dim MATHUE As Variant
    Dim madtnt    As Variant
    Dim ngnop     As Variant
    Dim kylbo     As Variant
    Dim kykkhai   As Variant
    Dim matkhai   As Variant
    Dim SOCT      As Variant
    Dim NGNOPTHUE As Variant
    Dim NOINOP    As Variant
    Dim CQUANQL   As Variant
    Dim THUEDN    As Variant
    Dim DANHAN    As Variant
   
    sSQLCol = "MADTNT, NGNOP, TTHTK, MATKHAI, QUI, MATM, MAMUC, MATHUE, MAPP, NGUOINH, DCHILO, DIENTICH, GIADAT1, GIATRI1, GIADAT2, GIATRI2, LAN_QUET, DANHAN"
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 12, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 34, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
        .GetText .ColLetterToNumber("M"), 32, tthtk

        If Trim(tthtk) = vbNullString Then
            madtnt = "0"
        End If

        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            maTM = "'02'"
            MaMuc = "'002'"
        Else
            MaMuc = "'1050'"
            maTM = "'1052'"
        End If

        MAPP = "'1'"
        MATHUE = "'08'"
        kykkhai = "'" & TAX_Utilities_Svr_New.ThreeMonths & "/" & TAX_Utilities_Svr_New.Year & "'"
        matkhai = "'01A/TNDN'"
        DANHAN = "''"
        .Sheet = 2
        .Col = .ColLetterToNumber("C")
        .Row = 35

        Do
                                       
            .Col = .ColLetterToNumber("E")
            NGNOPTHUE = .Text

            If Trim(NGNOPTHUE) = vbNullString Then
                NGNOPTHUE = "CTOD('')"
            Else
                NGNOPTHUE = ToDate(Trim(NGNOPTHUE), DDMMYYYY)
                NGNOPTHUE = "CTOD('" & Format(NGNOPTHUE, "mm/dd/yyyy") & "')"
            End If
                                                    
            .Col = .ColLetterToNumber("C")
            CQUANQL = .Text

            If Trim(CQUANQL) = vbNullString Then
                CQUANQL = "''"
            Else
                CQUANQL = "'" & TAX_Utilities_Svr_New.Convert(CStr(CQUANQL), UNICODE, TCVN) & "'"
            End If
                       
            .Col = .ColLetterToNumber("F")
            THUEDN = .Value

            If Trim(THUEDN) = vbNullString Then
                THUEDN = "0"
            End If
                        
            If Trim(NOINOP) = vbNullString Then
                NOINOP = "''"
            Else
                NOINOP = "'" & TAX_Utilities_Svr_New.Convert(CStr(NOINOP), UNICODE, TCVN) & "'"
            End If
                        
            If Trim(SOCT) = vbNullString Then
                SOCT = "''"
            Else
                SOCT = "'" & TAX_Utilities_Svr_New.Convert(CStr(SOCT), UNICODE, TCVN) & "'"
            End If
                        
            sSQLVal = madtnt & "," & ngnop & "," & tthtk & "," & kylbo & "," & kykkhai & "," & matkhai & "," & SOCT & "," & NGNOPTHUE & "," & NOINOP & "," & CQUANQL & "," & THUEDN & "," & LAN_QUET & "," & DANHAN
                       
            sSQL = "INSERT INTO TMP_P6" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)

            .Col = .ColLetterToNumber("C")
            .Row = .Row + 1
        Loop Until .Text = "aa"
             
    End With

    clsDAO.Disconnect
   
    InsertDTL_PL5A = vbNullString
   
End Function

