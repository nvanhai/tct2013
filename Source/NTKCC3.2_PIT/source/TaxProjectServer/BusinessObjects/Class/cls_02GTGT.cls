VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_02GTGT"
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
        .Col = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X)
        .Row = HEADER_SO_TT_TRONG_TEP_Y
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetNumeric
        .TypeMaxEditLen = 10
        
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
        .GetText .ColLetterToNumber("J"), 64, LoaiKyKK

        If LoaiKyKK = "0" Then
            LOAI_TO_KHAI = "02/GTGT"
            isTkQuy = False
        Else
            LOAI_TO_KHAI = "02Q/GTGT"
            isTkQuy = True
        End If
    End With
    
    Prepared2 = True
End Function

Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

Public Function Prepared3() As Boolean
    Dim formula As Variant
    With fps
        'set demo formula
        .Sheet = .SheetCount
        .Row = 13
        .Col = .ColLetterToNumber("B")
        .formula = "IF(VALUE(RIGHT('01'!E30,4))<VALUE(T2),0,IF(VALUE(LEFT('01'!E30,2))<VALUE(T3),0,1))"
        
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
        
        'so thu tu to khai
        .Col = .ColLetterToNumber("M")
        .Row = 30
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        'Set ma BPQL
        
''       .Col = .ColLetterToNumber(MA_BPQL_X)
''        .Row = MA_BPQL_Y
''        If strMaBPQL <> "" Then
''            .Text = strMaBPQL
''            UpdateCell fps, .Col, .Row, .Text
''        End If
''
''        .Col = .ColLetterToNumber(TEN_BPQL_X)
''        .Row = TEN_BPQL_Y
''        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
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
    Dim strKyLBO As Variant
    
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 36, varCheckValue
        .GetText .ColLetterToNumber("M"), 36, varNoteValue
        If CStr(varCheckValue) = "1" And Trim(CStr(varNoteValue)) = "" Then
            .Sheet = .SheetCount
            .SetText 2, 14, "0"
            .Col = .ColLetterToNumber("E")
            .Row = 36
            .CellNote = "err"
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
    Dim iRowID       As Long, strSQL As String, strTempSQL As String
    Dim dDate        As Date, strDate() As String
    Dim vTIN, vTEN_DTNT, vDIA_CHI, vLOAI_TKHAI, vNGAY_NOP, vKyLB
    Dim vKyKK, vNGAY_CAP_NHAT, vNGUOI_CAP_NHAT, vCO_LOI_DDANH
    Dim vSO_HIEU_TEP, vSO_TT_TK, vDA_NHAN, vGHI_CHU_LOI, vCO_GTRINH_02A
    Dim vCO_GTRINH_02B, vCO_GTRINH_02C
    Dim vPHONG_XU_LY
    Dim i              As Long, j As Long
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
    
    vCO_GTRINH_02A = TAX_Utilities_Svr_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue
    vCO_GTRINH_02B = TAX_Utilities_Svr_New.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue
    vCO_GTRINH_02C = TAX_Utilities_Svr_New.NodeValidity.childNodes(3).Attributes.getNamedItem("Active").nodeValue
    strSQL = strSQL & "'" & IIf(Val(vCO_GTRINH_02A) = 1, "Y", "") & "',"
    strSQL = strSQL & "'" & IIf(Val(vCO_GTRINH_02B) = 1, "Y", "") & "',"
    strSQL = strSQL & "'" & IIf(Val(vCO_GTRINH_02C) = 1, "Y", "") & "',"
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
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("E_30"), "Value"), "dd/mm/yyyy")
    
    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function

Public Function InsertDTL() As String
    Dim sSQL      As String
    Dim sSQLCol   As String
    Dim sSQLVal   As String
   
    Dim madtnt    As Variant
    Dim MATHUE    As Variant
    Dim matkhai   As Variant
    Dim kykkhai   As Variant
    Dim maTM      As Variant
    Dim MAPP      As Variant
    Dim KYLBO     As Variant
    Dim LoaiTien  As Variant
    Dim ngnop     As Variant
    Dim CTIEUTNDN As Variant
    Dim DSO       As Variant
    Dim THUEGTGT  As Variant
    Dim DSO2      As Variant
    Dim THUEGTGT2 As Variant
    Dim tt        As Variant
    Dim GTRIBR    As Variant
    Dim NGANH     As Variant  ' vttoan
    Dim DANHAN    As Variant
   
    Dim MaMuc     As Variant
    Dim CT30      As Variant
    Dim CT31      As Variant
    Dim CT32      As Variant
    Dim CT33      As Variant
    Dim SCT31     As Boolean
    Dim SCT33     As Boolean
   
    SCT31 = True
    SCT33 = True
   
    Dim bln As Boolean
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
    
    sSQLCol = "MADTNT , MATHUE, MATM, MAPP, KYLBO, LOAITIEN, NGNOP, "
    sSQLCol = sSQLCol & "Ctieutndn , DSO, THUEGTGT, DSO2, THUEGTGT2, TT, MAMUC,TTHTK,KYKKHAI,MATKHAI,NGANH, LAN_QUET, DANHAN,THUECL,DSOCL"
                              
    '-------------
    With fps
   
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        DANHAN = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'210'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 42, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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
        
        NGANH = "'" & TinhMaNganh & "'"
        
    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
   
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
    '-------------
    With fps
   
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        DANHAN = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'211'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 43, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
        tt = "''"

        ' sua MAMUC & MATM tu nam 2009
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        NGANH = "'" & TinhMaNganh & "'"
        
    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
   
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '-------------
    With fps
   
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        DANHAN = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'220'"
        DSO = "0"
        THUEGTGT = "0"
        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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
        
        NGANH = "'" & TinhMaNganh & "'"
        
    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
   
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '----------------------

    '---------------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'230'"
        .GetText .ColLetterToNumber("Q"), 45, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If

        .GetText .ColLetterToNumber("W"), 45, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
   
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
    '-------------
    With fps
   
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        DANHAN = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'240'"
        DSO = "0"
        THUEGTGT = "0"
        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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

        NGANH = "'" & TinhMaNganh & "'"
        
    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
   
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '----------------------
   
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'250'"
        .GetText .ColLetterToNumber("Q"), 47, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If
        
        .GetText .ColLetterToNumber("W"), 47, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
   
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'270'"
        .GetText .ColLetterToNumber("Q"), 48, DSO

        If Trim(DSO) = vbNullString Then
            DSO = "0"
        End If
        
        .GetText .ColLetterToNumber("W"), 48, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
   
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '----------------------
   
    '----------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'280'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 49, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
   
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        CTIEUTNDN = "'281'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 50, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
   
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'290'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 51, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
   
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
        With fps
   
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        DANHAN = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'291'"
        DSO = "0"
        THUEGTGT = "0"
        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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

        NGANH = "'" & TinhMaNganh & "'"
        
    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
   
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '----------------------
    
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'301'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 53, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'300'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 54, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'310'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 55, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)

    '----------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        LoaiTien = "''"
        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        'Chi tieu
        '
        CTIEUTNDN = "'320'"
        DSO = "0"
        .GetText .ColLetterToNumber("W"), 56, THUEGTGT

        If Trim(THUEGTGT) = vbNullString Then
            THUEGTGT = "0"
        End If

        DSO2 = DSO
        THUEGTGT2 = THUEGTGT
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

    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & LoaiTien & "," & ngnop & ","
    sSQLVal = sSQLVal & CTIEUTNDN & "," & DSO & "," & THUEGTGT & "," & DSO2 & "," & THUEGTGT2 & "," & tt & "," & MaMuc & "," & TTHTK & "," & kykkhai & "," & matkhai & "," & NGANH & "," & LAN_QUET & "," & DANHAN & ",0,0"
    sSQL = "INSERT INTO TMP_DA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsDAO.ExecuteDLL(sSQL)
    '----------------------
    clsDAO.Disconnect
    
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2
    End If

    If TAX_Utilities_Svr_New.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2_1
    End If

    If TAX_Utilities_Svr_New.NodeValidity.childNodes(3).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2_2
    End If

    If TAX_Utilities_Svr_New.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2_3
    End If

    If TAX_Utilities_Svr_New.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2_4
    End If

    If TAX_Utilities_Svr_New.NodeValidity.childNodes(6).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2_5
    End If

    If TAX_Utilities_Svr_New.NodeValidity.childNodes(7).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2_6
    End If

    If TAX_Utilities_Svr_New.NodeValidity.childNodes(8).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2_7
    End If

    If TAX_Utilities_Svr_New.NodeValidity.childNodes(9).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2_8
    End If

    If TAX_Utilities_Svr_New.NodeValidity.childNodes(10).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2_9
    End If

    InsertDTL = vbNullString
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
        .Row = .Row + 3
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
            
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                           
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        If MACT2 <> "''" Then
            bln = clsDAO.ExecuteDLL(sSQL)
        End If

    End With
   
    clsDAO.Disconnect
    InsertDTL_KHBS = vbNullString
   
End Function

Public Function InsertDTL_PL2() As String
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
   
    Dim strFileNameP2 As String
    Dim fso           As New FileSystemObject
    Dim DANHAN        As Variant
    Dim MAUHD         As Variant
    
    'kiem tra ton tai tep *.dbf chua
    '    strFileNameP2 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P2" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileNameP2) = False Then
    '    Dim strTepmauP2 As String
    '        strTepmauP2 = spathVat & "\tepmau\P2.DBF"
    '        fso.CopyFile strTepmauP2, strFileNameP2, False
    '    End If
   
    sSQLCol = "MATKHAI, KYLBO, KYKKHAI, MADTNT, TTHTK, NGNOP, LOAIHD, KYHIEU, SOHDON, NGAYPH, HANG, NGMUA," & "MST, DSOCHUA, THUESUAT, THUEGTGT, MACT, STT, STT2, STTIN, GHICHU, MAPL, LAN_QUET, DANHAN, MAUHD"
   
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
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        .Sheet = 2
            
        'insert HHDV dung cho du an dau tu
        MAUHD = "''"
        LOAIHD = "''"
        KYHIEU = "''"
        SOHDON = "''"
        NGAYPH = "CTOD('')"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0101" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        'Muc I
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2001'"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        GhiChu = "''"
        MAPL = "'PL02GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc II
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0102" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2002'"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        GhiChu = "''"
        MAPL = "'PL02GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc III
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0103" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2003'"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        GhiChu = "''"
        MAPL = "'PL02GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc IV
        
        'Muc IV HTKK la muc V cua VAT
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0138" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2004'"
        STT = 4
        STT2 = 0
        STTIN = "'IV'"
        GhiChu = "''"
        MAPL = "'PL02GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc V
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0139" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2005'"
        STT = 5
        STT2 = 0
        STTIN = "'V'"
        GhiChu = "''"
        MAPL = "'PL02GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = 25

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If
            
            .GetText .ColLetterToNumber("C"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("H"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("K"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 5
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("M"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
            sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"
             
    End With
   
    clsDAO.Disconnect
    InsertDTL_PL2 = vbNullString
   
End Function

Public Function InsertDTL_PL2_1()
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
    Dim MAUHD         As Variant
    Dim strFileNameP2 As String
    Dim fso           As New FileSystemObject
    Dim DANHAN        As Variant
    'kiem tra ton tai tep *.dbf chua
    '    strFileNameP2 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P2" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileNameP2) = False Then
    '    Dim strTepmauP2 As String
    '        strTepmauP2 = spathVat & "\tepmau\P2.DBF"
    '        fso.CopyFile strTepmauP2, strFileNameP2, False
    '    End If
   
    sSQLCol = "MATKHAI, KYLBO, KYKKHAI, MADTNT, TTHTK, NGNOP, LOAIHD, KYHIEU, SOHDON, NGAYPH, HANG, NGMUA," & "MST, DSOCHUA, THUESUAT, THUEGTGT, MACT, STT, STT2, STTIN, GHICHU, MAPL, LAN_QUET, DANHAN, MAUHD"
   
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
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        .Sheet = 3
            
        'insert HHDV dung cho du an dau tu
            
        LOAIHD = "''"
        KYHIEU = "''"
        SOHDON = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0101" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        'Muc I
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2001'"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        GhiChu = "''"
        MAPL = "'PL02_1GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc II
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0102" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2002'"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        GhiChu = "''"
        MAPL = "'PL02_1GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc III
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0103" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2003'"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        GhiChu = "''"
        MAPL = "'PL02_1GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc IV
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0138" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2004'"
        STT = 4
        STT2 = 0
        STTIN = "'IV'"
        GhiChu = "''"
        MAPL = "'PL02_1GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc V
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0139" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2005'"
        STT = 5
        STT2 = 0
        STTIN = "'V'"
        GhiChu = "''"
        MAPL = "'PL02_1GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = 25

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If
            
            .GetText .ColLetterToNumber("C"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("H"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("K"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 5
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("M"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02_1GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
            sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"
             
    End With
   
    clsDAO.Disconnect
    InsertDTL_PL2_1 = vbNullString
   
End Function

Public Function InsertDTL_PL2_2()
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
    Dim MAUHD         As Variant
    Dim strFileNameP2 As String
    Dim fso           As New FileSystemObject
    Dim DANHAN        As Variant
    'kiem tra ton tai tep *.dbf chua
    '    strFileNameP2 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P2" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileNameP2) = False Then
    '    Dim strTepmauP2 As String
    '        strTepmauP2 = spathVat & "\tepmau\P2.DBF"
    '        fso.CopyFile strTepmauP2, strFileNameP2, False
    '    End If
   
    sSQLCol = "MATKHAI, KYLBO, KYKKHAI, MADTNT, TTHTK, NGNOP, LOAIHD, KYHIEU, SOHDON, NGAYPH, HANG, NGMUA," & "MST, DSOCHUA, THUESUAT, THUEGTGT, MACT, STT, STT2, STTIN, GHICHU, MAPL, LAN_QUET, DANHAN, MAUHD"
   
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
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        .Sheet = 4
        'insert HHDV dung cho du an dau tu
            
        LOAIHD = "''"
        KYHIEU = "''"
        SOHDON = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0101" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        'Muc I
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2001'"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        GhiChu = "''"
        MAPL = "'PL02_2GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc II
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0102" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2002'"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        GhiChu = "''"
        MAPL = "'PL02_2GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc III
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0103" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2003'"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        GhiChu = "''"
        MAPL = "'PL02_2GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc IV
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0138" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2004'"
        STT = 4
        STT2 = 0
        STTIN = "'IV'"
        GhiChu = "''"
        MAPL = "'PL02_2GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc V
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0139" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2005'"
        STT = 5
        STT2 = 0
        STTIN = "'V'"
        GhiChu = "''"
        MAPL = "'PL02_2GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = 25

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If
            
            .GetText .ColLetterToNumber("C"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("H"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("K"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 5
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("M"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02_2GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
            sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"
             
    End With
   
    clsDAO.Disconnect
    InsertDTL_PL2_2 = vbNullString
End Function

Public Function InsertDTL_PL2_3()
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
    Dim MAUHD         As Variant
    Dim strFileNameP2 As String
    Dim fso           As New FileSystemObject
    Dim DANHAN        As Variant
    'kiem tra ton tai tep *.dbf chua
    '    strFileNameP2 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P2" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileNameP2) = False Then
    '    Dim strTepmauP2 As String
    '        strTepmauP2 = spathVat & "\tepmau\P2.DBF"
    '        fso.CopyFile strTepmauP2, strFileNameP2, False
    '    End If
   
    sSQLCol = "MATKHAI, KYLBO, KYKKHAI, MADTNT, TTHTK, NGNOP, LOAIHD, KYHIEU, SOHDON, NGAYPH, HANG, NGMUA," & "MST, DSOCHUA, THUESUAT, THUEGTGT, MACT, STT, STT2, STTIN, GHICHU, MAPL, LAN_QUET, DANHAN, MAUHD"
   
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
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        .Sheet = 5
            
        'insert HHDV dung cho du an dau tu
            
        LOAIHD = "''"
        KYHIEU = "''"
        SOHDON = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0101" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        'Muc I
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2001'"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        GhiChu = "''"
        MAPL = "'PL02_3GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc II
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0102" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2002'"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        GhiChu = "''"
        MAPL = "'PL02_3GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc III
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0103" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2003'"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        GhiChu = "''"
        MAPL = "'PL02_3GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc IV
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0138" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2004'"
        STT = 4
        STT2 = 0
        STTIN = "'IV'"
        GhiChu = "''"
        MAPL = "'PL02_3GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc V
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0139" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2005'"
        STT = 5
        STT2 = 0
        STTIN = "'V'"
        GhiChu = "''"
        MAPL = "'PL02_3GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = 25

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If
            
            .GetText .ColLetterToNumber("C"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("H"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("K"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 5
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("M"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02_3GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
            sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"
             
    End With
   
    clsDAO.Disconnect
    InsertDTL_PL2_3 = vbNullString
End Function

Public Function InsertDTL_PL2_4()
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
    Dim MAUHD         As Variant
    Dim strFileNameP2 As String
    Dim fso           As New FileSystemObject
    Dim DANHAN        As Variant
    'kiem tra ton tai tep *.dbf chua
    '    strFileNameP2 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P2" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileNameP2) = False Then
    '    Dim strTepmauP2 As String
    '        strTepmauP2 = spathVat & "\tepmau\P2.DBF"
    '        fso.CopyFile strTepmauP2, strFileNameP2, False
    '    End If
   
    sSQLCol = "MATKHAI, KYLBO, KYKKHAI, MADTNT, TTHTK, NGNOP, LOAIHD, KYHIEU, SOHDON, NGAYPH, HANG, NGMUA," & "MST, DSOCHUA, THUESUAT, THUEGTGT, MACT, STT, STT2, STTIN, GHICHU, MAPL, LAN_QUET, DANHAN, MAUHD"
   
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
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        .Sheet = 6
            
        'insert HHDV dung cho du an dau tu
            
        LOAIHD = "''"
        KYHIEU = "''"
        SOHDON = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0101" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        'Muc I
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2001'"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        GhiChu = "''"
        MAPL = "'PL02_4GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc II
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0102" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2002'"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        GhiChu = "''"
        MAPL = "'PL02_4GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc III
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0103" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2003'"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        GhiChu = "''"
        MAPL = "'PL02_4GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc IV
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0138" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2004'"
        STT = 4
        STT2 = 0
        STTIN = "'IV'"
        GhiChu = "''"
        MAPL = "'PL02_4GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc V
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0139" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2005'"
        STT = 5
        STT2 = 0
        STTIN = "'V'"
        GhiChu = "''"
        MAPL = "'PL02_4GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = 25

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If
            
            .GetText .ColLetterToNumber("C"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("H"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("K"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 5
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("M"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02_4GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
            sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"
             
    End With
   
    clsDAO.Disconnect
    InsertDTL_PL2_4 = vbNullString
End Function

Public Function InsertDTL_PL2_5()
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
    Dim MAUHD         As Variant
    Dim strFileNameP2 As String
    Dim fso           As New FileSystemObject
    Dim DANHAN        As Variant
    'kiem tra ton tai tep *.dbf chua
    '    strFileNameP2 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P2" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileNameP2) = False Then
    '    Dim strTepmauP2 As String
    '        strTepmauP2 = spathVat & "\tepmau\P2.DBF"
    '        fso.CopyFile strTepmauP2, strFileNameP2, False
    '    End If
   
    sSQLCol = "MATKHAI, KYLBO, KYKKHAI, MADTNT, TTHTK, NGNOP, LOAIHD, KYHIEU, SOHDON, NGAYPH, HANG, NGMUA," & "MST, DSOCHUA, THUESUAT, THUEGTGT, MACT, STT, STT2, STTIN, GHICHU, MAPL, LAN_QUET, DANHAN, MAUHD"
   
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
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        .Sheet = 7
            
        'insert HHDV dung cho du an dau tu
            
        LOAIHD = "''"
        KYHIEU = "''"
        SOHDON = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0101" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        'Muc I
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2001'"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        GhiChu = "''"
        MAPL = "'PL02_5GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc II
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0102" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2002'"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        GhiChu = "''"
        MAPL = "'PL02_5GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc III
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0103" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2003'"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        GhiChu = "''"
        MAPL = "'PL02_5GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc IV
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0138" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2004'"
        STT = 4
        STT2 = 0
        STTIN = "'IV'"
        GhiChu = "''"
        MAPL = "'PL02_5GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc V
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0139" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2005'"
        STT = 5
        STT2 = 0
        STTIN = "'V'"
        GhiChu = "''"
        MAPL = "'PL02_5GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = 25

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If
            
            .GetText .ColLetterToNumber("C"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("H"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("K"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 5
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("M"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02_5GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
            sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"
             
    End With
   
    clsDAO.Disconnect
    InsertDTL_PL2_5 = vbNullString
End Function

Public Function InsertDTL_PL2_6()
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
    Dim MAUHD         As Variant
    Dim strFileNameP2 As String
    Dim fso           As New FileSystemObject
    Dim DANHAN        As Variant
    'kiem tra ton tai tep *.dbf chua
    '    strFileNameP2 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P2" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileNameP2) = False Then
    '    Dim strTepmauP2 As String
    '        strTepmauP2 = spathVat & "\tepmau\P2.DBF"
    '        fso.CopyFile strTepmauP2, strFileNameP2, False
    '    End If
   
    sSQLCol = "MATKHAI, KYLBO, KYKKHAI, MADTNT, TTHTK, NGNOP, LOAIHD, KYHIEU, SOHDON, NGAYPH, HANG, NGMUA," & "MST, DSOCHUA, THUESUAT, THUEGTGT, MACT, STT, STT2, STTIN, GHICHU, MAPL, LAN_QUET, DANHAN, MAUHD"
   
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
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        .Sheet = 8
            
        'insert HHDV dung cho du an dau tu
            
        LOAIHD = "''"
        KYHIEU = "''"
        SOHDON = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0101" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        'Muc I
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2001'"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        GhiChu = "''"
        MAPL = "'PL02_6GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc II
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0102" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2002'"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        GhiChu = "''"
        MAPL = "'PL02_6GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc III
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0103" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2003'"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        GhiChu = "''"
        MAPL = "'PL02_6GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc IV
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0138" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2004'"
        STT = 4
        STT2 = 0
        STTIN = "'IV'"
        GhiChu = "''"
        MAPL = "'PL02_6GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc V
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0139" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2005'"
        STT = 5
        STT2 = 0
        STTIN = "'V'"
        GhiChu = "''"
        MAPL = "'PL02_6GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = 25

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If
            
            .GetText .ColLetterToNumber("C"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("H"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("K"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 5
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("M"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02_6GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
            sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"
             
    End With
   
    clsDAO.Disconnect
    InsertDTL_PL2_6 = vbNullString
End Function
Public Function InsertDTL_PL2_7()
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim bln  As Boolean
    Dim madtnt As Variant
    Dim ngnop As Variant
    Dim KYLBO As Variant
    Dim kykkhai As Variant
    Dim KYTTBS As Variant
    Dim matkhai As Variant
    Dim LOAIHD As Variant
    Dim KYHIEU As Variant
    Dim SOHDON As Variant
    Dim NGAYPH As Variant
    Dim HANG As Variant
    Dim NGMUA As Variant
    Dim MST As Variant
    Dim DSOCHUA As Variant
    Dim THUESUAT As Variant
    Dim THUEGTGT As Variant
    Dim mact As Variant
    Dim STT As Variant
    Dim STT2 As Variant
    Dim STTIN As Variant
    Dim GhiChu As Variant
    Dim MAPL As Variant
    Dim i As Integer, iCol As Long, iRow As Long
    Dim xmlNode As MSXML.IXMLDOMNode
   Dim MAUHD         As Variant
    Dim strFileNameP2 As String
    Dim fso As New FileSystemObject
    Dim DANHAN As Variant
   'kiem tra ton tai tep *.dbf chua
'    strFileNameP2 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P2" & sKyKeKhai & ".DBF"
'    If fso.FileExists(strFileNameP2) = False Then
'    Dim strTepmauP2 As String
'        strTepmauP2 = spathVat & "\tepmau\P2.DBF"
'        fso.CopyFile strTepmauP2, strFileNameP2, False
'    End If
   
   
   sSQLCol = "MATKHAI, KYLBO, KYKKHAI, MADTNT, TTHTK, NGNOP, LOAIHD, KYHIEU, SOHDON, NGAYPH, HANG, NGMUA," & _
             "MST, DSOCHUA, THUESUAT, THUEGTGT, MACT, STT, STT2, STTIN, GHICHU, MAPL, LAN_QUET, DANHAN, MAUHD"
   
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
         If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        .Sheet = 9
            
'insert HHDV dung cho du an dau tu
            
            LOAIHD = "''"
            KYHIEU = "''"
            SOHDON = "''"
            NGAYPH = "CTOD('')"
            MAUHD = "''"
            For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage
                If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0101" Then
                    HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                    Exit For
                End If
            Next
            'Muc I
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2001'"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        GhiChu = "''"
        MAPL = "'PL02_7GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc II
        
                For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0102" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2002'"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        GhiChu = "''"
        MAPL = "'PL02_7GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc III
        
                For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0103" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2003'"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        GhiChu = "''"
        MAPL = "'PL02_7GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc IV
        
                For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0138" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2004'"
        STT = 4
        STT2 = 0
        STTIN = "'IV'"
        GhiChu = "''"
        MAPL = "'PL02_7GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc V
        
                For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0139" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2005'"
        STT = 5
        STT2 = 0
        STTIN = "'V'"
        GhiChu = "''"
        MAPL = "'PL02_7GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = 25

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If
            
            .GetText .ColLetterToNumber("C"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("H"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("K"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 5
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("M"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02_7GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
            sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"
             
    End With
   
   clsDAO.Disconnect
   InsertDTL_PL2_7 = vbNullString
End Function

Public Function InsertDTL_PL2_8()

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
    Dim MAUHD         As Variant
    Dim strFileNameP2 As String
    Dim fso           As New FileSystemObject
    Dim DANHAN        As Variant
    'kiem tra ton tai tep *.dbf chua
    '    strFileNameP2 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P2" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileNameP2) = False Then
    '    Dim strTepmauP2 As String
    '        strTepmauP2 = spathVat & "\tepmau\P2.DBF"
    '        fso.CopyFile strTepmauP2, strFileNameP2, False
    '    End If
   
    sSQLCol = "MATKHAI, KYLBO, KYKKHAI, MADTNT, TTHTK, NGNOP, LOAIHD, KYHIEU, SOHDON, NGAYPH, HANG, NGMUA," & "MST, DSOCHUA, THUESUAT, THUEGTGT, MACT, STT, STT2, STTIN, GHICHU, MAPL, LAN_QUET, DANHAN, MAUHD"
   
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
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        .Sheet = 10
            
        'insert HHDV dung cho du an dau tu
            
        LOAIHD = "''"
        KYHIEU = "''"
        SOHDON = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0101" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        'Muc I
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2001'"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        GhiChu = "''"
        MAPL = "'PL02_8GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc II
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0102" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2002'"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        GhiChu = "''"
        MAPL = "'PL02_8GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc III
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0103" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2003'"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        GhiChu = "''"
        MAPL = "'PL02_8GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc IV
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0138" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2004'"
        STT = 4
        STT2 = 0
        STTIN = "'IV'"
        GhiChu = "''"
        MAPL = "'PL02_8GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc V
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0139" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2005'"
        STT = 5
        STT2 = 0
        STTIN = "'V'"
        GhiChu = "''"
        MAPL = "'PL02_8GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = 25

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If
            
            .GetText .ColLetterToNumber("C"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("H"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("K"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 5
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("M"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02_8GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
            sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"
             
    End With
   
    clsDAO.Disconnect
    InsertDTL_PL2_8 = vbNullString
End Function

Public Function InsertDTL_PL2_9()

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
    Dim MAUHD         As Variant
    Dim strFileNameP2 As String
    Dim fso           As New FileSystemObject
    Dim DANHAN        As Variant
    'kiem tra ton tai tep *.dbf chua
    '    strFileNameP2 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P2" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileNameP2) = False Then
    '    Dim strTepmauP2 As String
    '        strTepmauP2 = spathVat & "\tepmau\P2.DBF"
    '        fso.CopyFile strTepmauP2, strFileNameP2, False
    '    End If
   
    sSQLCol = "MATKHAI, KYLBO, KYKKHAI, MADTNT, TTHTK, NGNOP, LOAIHD, KYHIEU, SOHDON, NGAYPH, HANG, NGMUA," & "MST, DSOCHUA, THUESUAT, THUEGTGT, MACT, STT, STT2, STTIN, GHICHU, MAPL, LAN_QUET, DANHAN, MAUHD"
   
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
        If isTkQuy = True Then
            kykkhai = "'" & Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year & "'"
        Else
            kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        End If

        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        .Sheet = 11
            
        'insert HHDV dung cho du an dau tu
            
        LOAIHD = "''"
        KYHIEU = "''"
        SOHDON = "''"
        NGAYPH = "CTOD('')"
        MAUHD = "''"

        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0101" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next

        'Muc I
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2001'"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        GhiChu = "''"
        MAPL = "'PL02_9GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc II
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0102" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2002'"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        GhiChu = "''"
        MAPL = "'PL02_9GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc III
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0103" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2003'"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        GhiChu = "''"
        MAPL = "'PL02_9GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc IV
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0138" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2004'"
        STT = 4
        STT2 = 0
        STTIN = "'IV'"
        GhiChu = "''"
        MAPL = "'PL02_9GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Muc V
        
        For Each xmlNode In TAX_Utilities_Svr_New.NodeMessage

            If xmlNode.Attributes.getNamedItem("ID").nodeValue = "0139" Then
                HANG = "'" & TAX_Utilities_Svr_New.Convert(xmlNode.Attributes.getNamedItem("Msg").nodeValue, UNICODE, TCVN) & "'"
                Exit For
            End If

        Next
        
        NGMUA = "''"
        MST = "''"
        DSOCHUA = 0
        THUESUAT = 0
        THUEGTGT = 0
        mact = "'2005'"
        STT = 5
        STT2 = 0
        STTIN = "'V'"
        GhiChu = "''"
        MAPL = "'PL02_9GTKT'"
            
        sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
        sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        i = 1
        .Row = 25

        Do
            LOAIHD = "''"
            .GetText .ColLetterToNumber("D"), .Row, KYHIEU

            If Trim(KYHIEU) = vbNullString Then
                KYHIEU = "''"
            Else
                KYHIEU = "'" & KYHIEU & "'"
            End If
            
            .GetText .ColLetterToNumber("C"), .Row, MAUHD

            If Trim(MAUHD) = vbNullString Then
                MAUHD = "''"
            Else
                MAUHD = "'" & MAUHD & "'"
            End If

            .GetText .ColLetterToNumber("E"), .Row, SOHDON

            If Trim(SOHDON) = vbNullString Then
                SOHDON = "''"
            Else
                SOHDON = "'" & SOHDON & "'"
            End If

            .GetText .ColLetterToNumber("F"), .Row, NGAYPH

            If Trim(NGAYPH) = vbNullString Then
                NGAYPH = "CTOD('')"
            Else
                NGAYPH = ToDate(Trim(NGAYPH), DDMMYYYY)
                NGAYPH = "CTOD('" & Format(NGAYPH, "mm/dd/yyyy") & "')"
            End If

            .GetText .ColLetterToNumber("G"), .Row, NGMUA

            If Trim(NGMUA) = vbNullString Then
                NGMUA = "''"
            Else
                NGMUA = "'" & TAX_Utilities_Svr_New.Convert(CStr(NGMUA), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("H"), .Row, MST

            If Trim(MST) = vbNullString Then
                MST = "''"
            Else
                MST = "'" & MST & "'"
            End If

            .GetText .ColLetterToNumber("I"), .Row, HANG

            If Trim(HANG) = vbNullString Then
                HANG = "''"
            Else
                HANG = "'" & TAX_Utilities_Svr_New.Convert(CStr(HANG), UNICODE, TCVN) & "'"
            End If

            .GetText .ColLetterToNumber("J"), .Row, DSOCHUA

            If Trim(DSOCHUA) = vbNullString Then
                DSOCHUA = "0"
            Else
                DSOCHUA = DSOCHUA
            End If
                
            .GetText .ColLetterToNumber("K"), .Row, THUESUAT

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            Else
                THUESUAT = IIf(InStr(1, THUESUAT, "%") > 0, Left(THUESUAT, Len(THUESUAT) - 1), THUESUAT)
            End If
                
            .GetText .ColLetterToNumber("L"), .Row, THUEGTGT

            If Trim(THUEGTGT) = vbNullString Then
                THUEGTGT = "0"
            Else
                THUEGTGT = THUEGTGT
            End If

            mact = "''"
            STT = 5
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("M"), .Row, GhiChu

            If Trim(GhiChu) = vbNullString Then
                GhiChu = "''"
            Else
                GhiChu = "'" & TAX_Utilities_Svr_New.Convert(CStr(GhiChu), UNICODE, TCVN) & "'"
            End If
                
            MAPL = "'PL02_9GTKT'"
                
            sSQLVal = matkhai & "," & KYLBO & "," & kykkhai & "," & madtnt & "," & TTHTK & "," & ngnop & "," & LOAIHD & "," & KYHIEU & "," & SOHDON & "," & NGAYPH & "," & HANG & "," & NGMUA & "," & MST & "," & DSOCHUA & "," & THUESUAT & "," & THUEGTGT & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & GhiChu & "," & MAPL & "," & LAN_QUET & "," & DANHAN & "," & MAUHD
                       
            sSQL = "INSERT INTO TMP_P2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If NGMUA <> "''" And NGAYPH <> "CTOD('')" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
                
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"
             
    End With
   
    clsDAO.Disconnect
    InsertDTL_PL2_9 = vbNullString
End Function
Public Function InsertHDR() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    Dim bln        As Boolean
    
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
    Dim kykkhai As Variant
    Dim NgNhap     As Variant
    Dim ThueTKy    As Variant
    Dim GTTinhThue As Variant
    Dim DaGhi      As Variant
    Dim KTTRUOC    As Variant
    Dim DoanhSo    As Variant
    Dim MAPP2      As Variant
    Dim MaMuc      As Variant
    Dim MaVach     As Variant
    Dim TENDA      As Variant
    Dim CT27       As Variant
    Dim ct34       As Variant
    Dim CT36       As Variant
    Dim matkhai As Variant
    Dim ThueTKy2   As Variant
    Dim HANNOP2    As Variant
    Dim NGANH      As Variant
    Dim DAGHI1     As Variant
    Dim AUTO_CAL   As Variant
    Dim DANHAN     As Variant
    Dim LoaiKyKK As Variant
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
    sSQLCol = sSQLCol & "KTTRUOC, DOANHSO, MAPP2, MAMUC, MAVACH,TENDA, MATKHAI, HANNOP, THUETKY2, HANNOP2, NGANH, DAGHI1, DAGHI2, LAN_QUET, AUTO_CAL, DANHAN,MADLT,LANBS,THUECL,DSOCL,GTTTCL,NGAYPS,MANNKD,TENNV,CHUNGCHI"
    NGANH = TinhMaNganh

    With fps
        .Sheet = 1
        'lay ten nhan vien dai ly thue, chung chi hanh nghe
        .GetText .ColLetterToNumber("H"), 60, TENNV
        'nshung- bo sung
        If (Trim(TENNV) <> vbNullString) Then
            TENNV = TAX_Utilities_Svr_New.Convert(Trim(TENNV), UNICODE, TCVN)
        End If
        
        .GetText .ColLetterToNumber("H"), 62, CHUNGCHI
        'nshung- bo sung
        If (Trim(CHUNGCHI) <> vbNullString) Then
            CHUNGCHI = TAX_Utilities_Svr_New.Convert(Trim(CHUNGCHI), UNICODE, TCVN)
        End If
        
        'set cac chi so
        DSoCL = "0"
        GTTTCL = "0"
        NgayPs = "CTOD('')"
        MANNKD = "''"
        '
        
        'set DAGHI2 theo chk loi dinh danh
        .Col = .ColLetterToNumber("E")
        .Row = 36

        If .Value = 1 Then
            DAGHI2 = ".T."
        Else
            DAGHI2 = ".F."
        End If
        
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        MATHUE = "'01'"
        MAPP = "'5'"
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

        ' Da bo truong TENDA
        '        .GetText .ColLetterToNumber("G"), 18, TENDA
        If Trim(TENDA) = vbNullString Then
            TENDA = "''"
        Else
            TENDA = "'" & TAX_Utilities_Svr_New.Convert(CStr(TENDA), UNICODE, TCVN) & "'"
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
        
        '   Da bo truong SHTEP
        '        .GetText .ColLetterToNumber("M"), 10, SHTEP
        If Trim(SHTEP) = vbNullString Then
            SHTEP = "''"
        Else
            SHTEP = "'" & SHTEP & "'"
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

        .GetText .ColLetterToNumber("W"), 56, ThueTKy

        If Trim(ThueTKy) = vbNullString Then
            ThueTKy = "0"
        Else
            ThueTKy = Trim(ThueTKy)
        End If
        
        ' Lay loai ky ke khai(thang/quy)
        .GetText .ColLetterToNumber("AB"), 42, LoaiKyKK
        If LoaiKyKK = "0" Then
            LoaiKyKK = ".T."
        Else
            LoaiKyKK = ".F."
        End If
       
        .GetText .ColLetterToNumber("Q"), 42, GTTinhThue

        If Trim(GTTinhThue) = vbNullString Then
            GTTinhThue = "0"
        Else
            GTTinhThue = Trim(GTTinhThue)
        End If

        DaGhi = ".F."
        DAGHI1 = ".F."
        KTTRUOC = "0"
        AUTO_CAL = ".F."
        DANHAN = "''"
        'DOANHSO = "0"
        'DOANHSO = CT27 + CT34 - CT36
        .GetText .ColLetterToNumber("Q"), 42, DoanhSo
        
        If Trim(DoanhSo) = vbNullString Then

            DoanhSo = "0"
        Else

            DoanhSo = Trim(DoanhSo)
        End If
        
        MAPP2 = "'5'"

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

        ' lay so lan quet
        '       .GetText .ColLetterToNumber("M"), 30, LAN_QUET
        '        If Trim(LAN_QUET) = vbNullString Then
        '            LAN_QUET = "1"
        '        End If
        DANHAN = "''"
        THUECL = "0"
    End With
    
    sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & ngnop & ", "
    sSQLVal = sSQLVal & LoaiTien & "," & MABPQL & "," & SHTEP & "," & TTHTK & "," & kykkhai & ", "
    sSQLVal = sSQLVal & NgNhap & "," & ThueTKy & "," & GTTinhThue & "," & DaGhi & ", "
    sSQLVal = sSQLVal & KTTRUOC & "," & DoanhSo & "," & MAPP2 & "," & MaMuc & "," & MaVach & "," & TENDA & ","
    sSQLVal = sSQLVal & matkhai & "," & HANNOP & "," & ThueTKy2 & "," & HANNOP2 & ",'" & NGANH & "'," & DAGHI1 _
    & "," & DAGHI2 & "," & LAN_QUET & "," & AUTO_CAL & "," & DANHAN & "," & madlt & "," & LanBS & "," & THUECL _
    & "," & DSoCL & "," & GTTTCL & "," & NgayPs & "," & MANNKD & ",'" & TENNV & "','" & CHUNGCHI & "'"

    sSQL = "INSERT INTO TMP_TK" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
    End If
    
    bln = clsConn.ExecuteDLL(sSQL)
    clsConn.Disconnect
    
    InsertHDR = vbNullString
    
End Function

Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim strFileNameIDX As String
    
    'kiem tra ton tai tep *.dbf chua
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
'fsdfsdfdsfdsf
    
        strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
        strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_DA" & sKyKeKhai & ".DBF"
        Dim strFileNamePL1 As String
        strFileNamePL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_P2" & sKyKeKhai & ".DBF"
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
            strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_DA.DBF"
            fso.CopyFile strTepmauDTL, strFileNameDTL, False
        End If
        'file phu luc
        If fso.FileExists(strFileNamePL1) = False Then
            Dim strTepmauPL1 As String
            strTepmauPL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_P2.DBF"
            fso.CopyFile strTepmauPL1, strFileNamePL1, False
        End If
        
        'file bo sung
        If fso.FileExists(strFileNameBS) = False Then
            Dim strTepmauBS As String
            strTepmauBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_BS.DBF"
            fso.CopyFile strTepmauBS, strFileNameBS, False
        End If

        isExistFile = True


'
'    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "DA" & sKyKeKhai & ".DBF"
'    If fso.FileExists(strFileNameDTL) = False Then
'    Dim strTepmauDTL As String
'        strTepmauDTL = spathVat & "\tepmau\DA.DBF"
'        fso.CopyFile strTepmauDTL, strFileNameDTL, False
'    End If
    
'    'Xoa file *.IDX
'    strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TK" & sKyKeKhai & ".IDX"
'    If fso.FileExists(strFileNameIDX) = True Then
'        fso.DeleteFile strFileNameIDX, True
'    End If
'
'        strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "DA" & sKyKeKhai & ".IDX"
'    If fso.FileExists(strFileNameIDX) = True Then
'        fso.DeleteFile strFileNameIDX, True
'    End If
'
'        strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "BS" & sKyKeKhai & ".IDX"
'    If fso.FileExists(strFileNameIDX) = True Then
'        fso.DeleteFile strFileNameIDX, True
'    End If
'
'    strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P2" & sKyKeKhai & ".IDX"
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
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyKeKhai & _
        " WHERE MATHUE = '01' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & vKyKK & "'" & _
        " AND TTHTK = '1' AND MATKHAI = '" & LOAI_TO_KHAI & "'"
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
        " WHERE MATHUE = '01' AND MAPP = '5' AND MADTNT = " & "'" & strMST & "'" & _
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
   
   Dim vKyKK As String
    If isTkQuy = True Then
        vKyKK = Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year
    Else
        vKyKK = TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
    End If
   
   sSQL = "DELETE FROM TMP_TK" & sKyKeKhai & _
        " WHERE MATHUE = '01' AND MAPP = '5' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & vKyKK & "'" & " AND MATKHAI = '" & LOAI_TO_KHAI & "'"

    
   clsDAO.ExecuteDLL sSQL
   
   sSQL = "DELETE FROM TMP_DA" & sKyKeKhai & _
        " WHERE MATHUE = '01' AND MAPP = '5' AND MADTNT = " & "'" & strMST & "'" & _
       " AND KYLBO = " & KYLBO
    clsDAO.ExecuteDLL sSQL
   
'   With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("E"), 10, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "''"
'        Else
'            If Len(Trim(KYLBO)) = 6 Then
'                KYLBO = "'0" & KYLBO & "'"
'            Else
'                KYLBO = "'" & KYLBO & "'"
'            End If
'        End If
'   End With
'   sSQL = "DELETE FROM KT" & sKyKeKhai & _
'        " WHERE MATHUE = '01' AND MATM = '01' AND MAPP = '5' AND MADTNT = " & "'" & strMST & "' AND NGNOP = " & dNgayNopDB
'
'   clsDAO.ExecuteDLL sSQL
   
   XoaTKTT = True
   
   clsDAO.Disconnect

End Function

Public Function TKRB() As Boolean

    Dim sSQL As String
    Dim THUEGTGT_KT As Double
    Dim rs As ADODB.Recordset
    Dim sKyTruoc As String
    Dim sKyKKTruoc As String
    Dim fso As New FileSystemObject
    Dim THUEGTGT As Variant
    Dim dNgayNopDBKT As Variant
    
    If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
   If CInt(Month(Now())) = 1 Then
        sKyTruoc = "12" & Trim(str(CInt(Year(Now())) - 1))
   Else
        sKyTruoc = IIf(Len(Trim(str(CInt(Month(Now())) - 1))) = 1, "0" & Trim(str(CInt(Month(Now())) - 1)), Trim(str(CInt(Month(Now())) - 1))) & Year(Now())
   End If
   
   If CInt(TAX_Utilities_Svr_New.Month) = 1 Then
        sKyKKTruoc = "12" & TAX_Utilities_Svr_New.Year
   Else
        sKyKKTruoc = IIf(Len(Trim(str(CInt(TAX_Utilities_Svr_New.Month) - 1))) = 1, "0" & Trim(str(CInt(TAX_Utilities_Svr_New.Month) - 1)), Trim(str(CInt(TAX_Utilities_Svr_New.Month) - 1))) & TAX_Utilities_Svr_New.Year
   End If
   
   If Not fso.FileExists(spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyTruoc & ".dbf") Then
        TKRB = True
        clsDAO.Disconnect
        Exit Function
   End If
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyTruoc & _
        " WHERE MATHUE = '01' AND MAPP = '5' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & Left(sKyKKTruoc, 2) & "/" & Right(sKyKKTruoc, 4) & "'"
    
   Set rs = clsDAO.Execute(sSQL)
   If Not rs Is Nothing Then
        dNgayNopDBKT = rs.Fields("NGNOP")
        If Trim(dNgayNopDBKT) = vbNullString Then
            dNgayNopDBKT = "CTOD('')"
        Else
            dNgayNopDBKT = "CTOD('" & Format(dNgayNopDBKT, "mm/dd/yyyy") & "')"
        End If
   Else
        TKRB = True
        clsDAO.Disconnect
        Exit Function
   End If
   If Not fso.FileExists(spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "KT" & sKyTruoc & ".dbf") Then
        TKRB = True
        clsDAO.Disconnect
        Exit Function
   End If
   
   sSQL = "SELECT thuegtgt FROM KT" & sKyTruoc & _
        " WHERE CT_KT = '085' AND MADTNT = " & "'" & strMST & "' AND NGNOP = " & dNgayNopDBKT
    
   Set rs = clsDAO.Execute(sSQL)
   If Not rs Is Nothing Then
        THUEGTGT_KT = rs.Fields("thuegtgt")
        If THUEGTGT_KT = 0 Then
            TKRB = True
            clsDAO.Disconnect
            Exit Function
        Else
            With fps
                .Sheet = 1
                .GetText .ColLetterToNumber("W"), 39, THUEGTGT
                If Trim(THUEGTGT) = vbNullString Then
                    THUEGTGT = "0"
                End If
                If THUEGTGT_KT = THUEGTGT Then
                     TKRB = True
                     clsDAO.Disconnect
                     Exit Function
                Else
                      TKRB = False
                      clsDAO.Disconnect
                      Exit Function
                End If
            End With
        End If
   Else
        TKRB = True
        clsDAO.Disconnect
        Exit Function
   End If
   
   TKRB = True
   clsDAO.Disconnect

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
        " WHERE MATHUE = '01' AND MAPP = '5' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI <> '" & vKyKK & "'" & _
        " AND MATKHAI = '" & LOAI_TO_KHAI & "'"
    
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
 Dim vNgayQyet As Variant
 With fps
    If .ActiveSheet = 1 Then
        If Col = .ColLetterToNumber("E") And Row = 32 Then
            .GetText Col, Row, varTemp
            .GetText .ColLetterToNumber("M"), Row, vNgayQyet
            If varTemp <> "" And varTemp <> "../../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                    checkNgayNop varTemp, vNgayQyet
                Else
                    .SetActiveCell Col, Row
                End If
            Else
                .SetText Col, Row, ""
                Exit Sub
            End If
'           UpdateCell fps, .Col, .Row, .Text
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
    'vttoan sua
    'ngay 24/09/2010
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
                    " WHERE MATHUE = '01' AND MADTNT = " & "'" & strMST & "'" & _
                    " AND KYKKHAI = '" & vKyKK & "'" & _
                    " AND MATKHAI = '" & LOAI_TO_KHAI & "' "
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
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim vKeKhai As Variant
    Dim vKeKhai_temp As Variant
    Dim vKyLB As Variant
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
        
        Dim vKyKK As String
        If isTkQuy = True Then
            vKyKK = Val(TAX_Utilities_Svr_New.Month) & "/" & TAX_Utilities_Svr_New.Year
        Else
            vKyKK = TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year
        End If
   
        strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & vKeKhai_temp & ".DBF"
        'kiem tra neu co db thi moi thuc hien truy van
        If fso.FileExists(strTestFileName) Then
            sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & vKeKhai_temp & _
            " WHERE MATHUE = '01' AND MADTNT = " & "'" & strMST & "'" & _
            " AND KYKKHAI = '" & vKyKK & "'" & _
            " AND (TTHTK = '1' or TTHTK = '3' or TTHTK = '4') AND MATKHAI = '" & LOAI_TO_KHAI & "' "
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
