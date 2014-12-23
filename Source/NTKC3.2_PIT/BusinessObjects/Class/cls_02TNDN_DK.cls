VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_02TNDN_DK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_Y = 27
Const KY_LAP_BO_X = "F"
Const MA_SO_TEP_Y = 43
Const MA_SO_TEP_X = "T"
Const NGAY_NHAN_TO_KHAI_Y = 29
Const NGAY_NHAN_TO_KHAI_X = "F"
Const HEADER_SO_TT_TRONG_TEP_ROW = 27 'Dong 16 da dc Hidden
Const HEADER_SO_TT_TRONG_TEP_COL = "S"

Const NGUOI_SU_DUNG_Y = 29
Const NGUOI_SU_DUNG_X = "AG"

Const HEADER_KY_LAP_BO_ROW = 13
Const HEADER_KY_LAP_BO_COL = "B"

Const PHONG_XU_LY_ROW = 27
Const PHONG_XU_LY_COL = "AG"

' thong tin dai ly thue
Const TEN_DL_ROW = 17
Const TEN_DL_COL = "I"

Const DIA_CHI_DL_ROW = 21
Const DIA_CHI_DL_COL = "I"

Const DIEN_THOAI_DL_ROW = 23
Const DIEN_THOAI_DL_COL = "I"

Const FAX_DL_ROW = 23
Const FAX_DL_COL = "S"

Const EMAIL_DL_ROW = 23
Const EMAIL_DL_COL = "AG"

Const SO_HOP_DONG_DL_ROW = 25
Const SO_HOP_DONG_DL_COL = "I"

Const NGAY_HOP_DONG_DL_ROW = 25
Const NGAY_HOP_DONG_DL_COL = "S"

'Longvh
Const MA_SO_THUE_X = "I"
Const MA_SO_THUE_Y = 11
Const NGAY_NOP_Y = 29
Const NGAY_NOP_X = "F"
Const NGAY_QUET_Y = 29
Const NGAY_QUET_X = "S"

Const TEN_GOI_X = "I"
Const TEN_GOI_Y = 9
Const DIA_CHI_X = "I"
Const DIA_CHI_Y = 13
Const MA_BPQL_X = "AG"
Const MA_BPQL_Y = 31
Const DIEN_THOAI_X = "I"
Const DIEN_THOAI_Y = 15
Const FAX_X = "S"
Const FAX_Y = 15

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
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
Public sKyKeKhai As String

Public strTenDL As String
Public strDiaChiDL As String
Public strDienThoaiDL As String
Public strFaxDL As String
Public strEmailDL As String
Public strSoHopDongDL As String
Public strNgayHopDongDL As String


Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
        'Ghi chu
        .Col = .ColLetterToNumber("N")
        .Row = 33
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100
        
        'Phong xu ly
        .Col = .ColLetterToNumber("AG")
        .Row = 24
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60

        'Ky lap bo
        SetDateFormat fps, 1, 27, .ColLetterToNumber("F"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 29, .ColLetterToNumber("F"), DDMMYYYY
        .Row = 29
        .Col = .ColLetterToNumber("F")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 29, .ColLetterToNumber("S"), DDMMYYYY
        .Row = 29
        .Col = .ColLetterToNumber("S")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft

    End With
    
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
    Dim i                As Integer, intIndexCombo As Integer
    Dim strLTN           As Variant, Col7 As Variant, strLTNCu As Variant, strId As Variant
    Dim xmlCellNode      As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim varID            As Variant, varDVT As Variant, varThueSuat As Variant
    Dim iTenTN           As String, strDVT As String, strThueSuat As String
    Dim iCol             As Long, iRow As Long
    Dim strIdQLT         As String, blnSuaThueSuat As Boolean
    Dim strMucThueAnDinh As String
    Dim vKHBS

    i = 0
    ReDim Preserve larrPhongXuLy(0)
    larrPhongXuLy(0) = "00"
    ReDim Preserve larrid(0)
    larrid(0) = "00"
    
    With fps
        .Sheet = 1
        .Col = .ColLetterToNumber(PHONG_XU_LY_COL)
        .Row = PHONG_XU_LY_ROW

        If rsPXL.Fields.count > 0 Then

            Do While Not rsPXL.EOF
                .TypeComboBoxIndex = -1
                .TypeComboBoxString = TAX_Utilities_Srv_New.Convert(rsPXL.Fields(1).Value, TCVN, UNICODE)
                i = i + 1
                ReDim Preserve larrPhongXuLy(i)
                larrPhongXuLy(i) = TAX_Utilities_Srv_New.Convert(rsPXL.Fields(1).Value, TCVN, UNICODE)
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

Private Sub DataDM(ByVal Id As String, Optional ByRef strIdQLT As String, Optional ByRef TenTN As String, Optional ByRef DVT As String, Optional ByRef strThueSuat As String, Optional ByRef blnSuaThueSuat As Boolean)
Dim arrDanhsach() As String
Dim strDataFileName As String
Dim xmlDOMdata As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = GetCatalogueFileName
    
       If xmlDOMdata.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDOMdata.getElementsByTagName("Cell")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If Id = arrDanhsach(1) Then
                            strIdQLT = arrDanhsach(0)
                            TenTN = arrDanhsach(2)
                            DVT = arrDanhsach(3)
                            strThueSuat = arrDanhsach(4)
                            blnSuaThueSuat = IIf(arrDanhsach(5) = "1", True, False)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub

Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String) 'ByRef strPMaSoTep, ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
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

        'Set NguoiSuDung
        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_X)
        .Row = NGUOI_SU_DUNG_Y
        If strNguoiSuDung <> "" Then
            .Text = strNguoiSuDung
            UpdateCell fps, .Col, .Row, .Text
        End If

         'set Phong xu ly
        .Col = .ColLetterToNumber(PHONG_XU_LY_COL)
        .Row = PHONG_XU_LY_ROW
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
        
'        .Col = .ColLetterToNumber(NGAY_NOP_X)
'        .Row = NGAY_NOP_Y
'        .Value = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
'         UpdateCell fps, .Col, .Row, .Value
'        .Lock = False
        
'        .Col = .ColLetterToNumber(NGAY_QUET_X)
'        .Row = NGAY_QUET_Y
'        .Value = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
'        UpdateCell fps, .Col, .Row, .Value
'
        .Col = .ColLetterToNumber(TEN_GOI_X)
        .Row = TEN_GOI_Y
        If strTenGoi <> "" Then
            .Text = Trim(TAX_Utilities_Srv_New.Convert(strTenGoi, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(DIA_CHI_X)
        .Row = DIA_CHI_Y
        If strDchi <> "" Then
            .Text = Trim(TAX_Utilities_Srv_New.Convert(strDchi, TCVN, UNICODE))
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
        
        .Col = .ColLetterToNumber(MA_BPQL_X)
        .Row = MA_BPQL_Y
        If strMaBPQL <> "" Then
            .Text = strMaBPQL
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        ' set thong tin dai ly thue
        ' ten dai ly thue
        .Col = .ColLetterToNumber(TEN_DL_COL)
        .Row = TEN_DL_ROW
        If strTenDL <> "" Then
            .Text = strTenDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' dia chi dai ly thue
        .Col = .ColLetterToNumber(DIA_CHI_DL_COL)
        .Row = DIA_CHI_DL_ROW
        If strDiaChiDL <> "" Then
            .Text = strDiaChiDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' dien thoai dai ly
        .Col = .ColLetterToNumber(DIEN_THOAI_DL_COL)
        .Row = DIEN_THOAI_DL_ROW
        If strDienThoaiDL <> "" Then
            .Text = strDienThoaiDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' fax dai ly thue
        .Col = .ColLetterToNumber(FAX_DL_COL)
        .Row = FAX_DL_ROW
        If strFaxDL <> "" Then
            .Text = strFaxDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' Email dai ly
        .Col = .ColLetterToNumber(EMAIL_DL_COL)
        .Row = EMAIL_DL_ROW
        If strEmailDL <> "" Then
            .Text = strEmailDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' so hop dong
        .Col = .ColLetterToNumber(SO_HOP_DONG_DL_COL)
        .Row = SO_HOP_DONG_DL_ROW
        If strSoHopDongDL <> "" Then
            .Text = strSoHopDongDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' ngay hop dong
        .Col = .ColLetterToNumber(NGAY_HOP_DONG_DL_COL)
        .Row = NGAY_HOP_DONG_DL_ROW
        If strNgayHopDongDL <> "" Then
            .Text = strNgayHopDongDL
            UpdateCell fps, .Col, .Row, .Text
        End If
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
        
        ' Get MaSoTep
'        .Col = .ColLetterToNumber(MA_SO_TEP_X)
'        .Row = MA_SO_TEP_Y
'        strMaSoTep = .Text
        
        ' Get Phong xu ly
        .Col = .ColLetterToNumber(PHONG_XU_LY_COL)
        .Row = PHONG_XU_LY_ROW
        strPhongXuLy = .Text
        
        ' Get NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
        .Row = NGAY_NHAN_TO_KHAI_Y
        strNgayNhanToKhai = .Text
        
        'Go to last sheet (header sheet)
        .Sheet = .SheetCount
        
        If Trim(strKyLapBo) = "" Or Trim(strKyLapBo) = "../...." Then
            blnValid = False
            .Col = .ColLetterToNumber("B")
            .Row = 19
            .Formula = IIf(blnValid, "1", "0")
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 19
            .Formula = IIf(blnValid, "1", "0")
            
            If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
                blnValid = False
            Else
                blnValid = True
            End If
            .Col = .ColLetterToNumber(HEADER_KY_LAP_BO_COL)
            .Row = HEADER_KY_LAP_BO_ROW
            
            If Not blnValid Then
                .Formula = "0"
            Else
                .Formula = "1"
            End If
            
            
            ' Kiem tra xem ky ke khai trong cung ky hay khac ky
'            If (DateAdd("M", 1, dNgayDauKy) = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Or (dNgayDauKy = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Then
'                blnValid = True
'            Else
'                blnValid = False
'            End If
'
'            .Col = .ColLetterToNumber("B")
'            .Row = 17
'            .Text = "1"
'
'            If Not blnValid Then
'                .Formula = "0"
'            Else
'                .Formula = "1"
'            End If
            ' kiem tra ky lap bo khong duoc lon hon ky hien tai
            If DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) > Date Then
                blnValid = False
            Else
                blnValid = True
            End If
            .Col = .ColLetterToNumber("B")
            .Row = 18
            
            If Not blnValid Then
                .Formula = "0"
            Else
                .Formula = "1"
            End If
        End If
        
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

Public Function GenerateSQL_Header(xmlDOMdata As MSXML.DOMDocument, strSQL_HDR As String, vHdrID As Variant, vKKBS As Byte, ByVal dNgayDauKy As Date) As String
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
    Dim sDATE_FROM, sDATE_TO As Variant
    Dim vTIN_DL, vSO_HOP_DONG, vNGAY_HOP_DONG, vLAN_BS, vCHINH_THUC As Variant
    Dim vDAU_THO, vCONDENSATE, vKHI_THIEN_NHIEN, vLOAI_TK_DK As Variant
On Error GoTo ErrHandle
    strSQL = strSQL_HDR
    strSQL = Replace(strSQL, ")", ",LOAI_TK_DK)")
    Set xmlList = xmlDOMdata.getElementsByTagName("Cell")
    For Each xmlNode In xmlList
        With xmlNode.Attributes
        
        If Trim(GetAttribute(xmlNode, "MCT")) = vbNullString Then
            Select Case Trim(GetAttribute(xmlNode, "CellID"))
                Case "I_11"
                    vTIN = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "I_9"
                    vTEN_DTNT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "I_13"
                    vDIA_CHI = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "F_29"
                    vNGAY_NOP = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "F_27"
                    vKyLB = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "S_29"
                    vNGAY_CAP_NHAT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "AG_29"
                    vNGUOI_CAP_NHAT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "G_33"
                    vCO_LOI_DDANH = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "T_43"
                    ' vSO_HIEU_TEP = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                    ' Tu phien ban 1.3.0 song tep se thuc hien bang tay=> So hieu tep luon bang ""
                    vSO_HIEU_TEP = ""
                Case HEADER_SO_TT_TRONG_TEP_COL & "_" & HEADER_SO_TT_TRONG_TEP_ROW
                    vSO_TT_TK = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "N_33"
                    vGHI_CHU_LOI = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case PHONG_XU_LY_COL & "_" & PHONG_XU_LY_ROW
                    vPHONG_XU_LY = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
            End Select
        End If
        End With
    Next
    
    strSQL = strSQL & "'" & vHdrID & "',"
    strSQL = strSQL & vKKBS & ","
    strSQL = strSQL & "'" & vTIN & "',"
    strSQL = strSQL & "'" & vTEN_DTNT & "',"
    strSQL = strSQL & "'" & vDIA_CHI & "',"
    'vLOAI_TKHAI = TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue
    vLOAI_TKHAI = "02_TNDN_DK"
    strSQL = strSQL & "'" & vLOAI_TKHAI & "',"
    strSQL = strSQL & "To_date('" & vNGAY_NOP & "','dd/mm/yyyy'),"
    
    'Ky/Quy LB
    'Ngay dau ky lap bo va ngay cuoi ky lap bo
    strDate = Split(vKyLB, "/")
    dDate = DateSerial(Val(strDate(1)), Val(strDate(0)), 1)
    strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    dDate = DateAdd("m", 1, dDate)
    dDate = DateAdd("d", -1, dDate)
    strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    
    'Ky/ Quy KK
    'Ngay dau ky ke khai va ngay cuoi ky ke khai
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("R"), 37, sDATE_FROM
        .GetText .ColLetterToNumber("S"), 37, sDATE_TO
        
        .GetText .ColLetterToNumber("D"), 38, vDAU_THO
        .GetText .ColLetterToNumber("G"), 38, vCONDENSATE
        .GetText .ColLetterToNumber("J"), 38, vKHI_THIEN_NHIEN
    End With
    'Giá tr?  =”0’ t? khai tháng, =”1’ t? khai l?n phát sinh, =”2” t? khai theo l?n xu?t bán
    If (vDAU_THO = "1") Then
        vLOAI_TK_DK = "DT"
    ElseIf (vCONDENSATE = "1") Then
        vLOAI_TK_DK = "CD"
    ElseIf (vKHI_THIEN_NHIEN = "1") Then
        vLOAI_TK_DK = "KTN"
    End If
    
    'SET KY KE KHAI
    'tu ngay
    dDate = DateSerial(CInt(TAX_Utilities_Srv_New.Year), 1, 1)
    strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    'den ngay
    dDate = DateAdd("m", 12, dDate)
    dDate = DateAdd("d", -1, dDate)
    strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    
    strSQL = strSQL & "To_date('" & vNGAY_CAP_NHAT & "','dd/mm/yyyy'),"
    strSQL = strSQL & "'" & vNGUOI_CAP_NHAT & "',"
    strSQL = strSQL & "'" & vCO_LOI_DDANH & "',"
    strSQL = strSQL & "'" & vSO_HIEU_TEP & "',"
    strSQL = strSQL & "'" & vSO_TT_TK & "',"
    
    strSQL = strSQL & "'" & vDA_NHAN & "',"
    strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
    
    'vCO_GTRINH_02A = TAX_Utilities_Srv_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue
    vCO_GTRINH_02A = 0
    vCO_GTRINH_02B = 0
    vCO_GTRINH_02C = 0
    strSQL = strSQL & "'" & IIf(Val(vCO_GTRINH_02A) = 1, "Y", "") & "',"
    strSQL = strSQL & "'" & IIf(Val(vCO_GTRINH_02B) = 1, "Y", "") & "',"
    strSQL = strSQL & "'" & IIf(Val(vCO_GTRINH_02C) = 1, "Y", "") & "',"
    ' Them danh dau 3 phu luc
    strSQL = strSQL & "'','','',"
    ' Ket thuc
    'strSQL = strSQL & "null,null,"
    
    '//SET KY QUYET TOAN
    strSQL = strSQL & "To_date('" & "01/" & sDATE_FROM & "','dd/mm/yyyy'),"
    dDate = Format(CDate("01/" & sDATE_TO), "dd/mm/yyyy")
    dDate = DateAdd("m", 1, dDate)
    dDate = DateAdd("d", -1, dDate)
    strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    
    With fps
        For i = 1 To lSoPhongXL
            If vPHONG_XU_LY = TAX_Utilities_Srv_New.Convert(larrPhongXuLy(i), UNICODE, TCVN) Then
                strMaPhongXuLy = larrid(i)
                Exit For
            End If
        Next
        
        
        .Sheet = 1
       ' Ma dai ly thue
        .GetText .ColLetterToNumber("I"), 19, vTIN_DL
        .GetText .ColLetterToNumber(SO_HOP_DONG_DL_COL), SO_HOP_DONG_DL_ROW, vSO_HOP_DONG
        .GetText .ColLetterToNumber(NGAY_HOP_DONG_DL_COL), NGAY_HOP_DONG_DL_ROW, vNGAY_HOP_DONG
        .GetText .ColLetterToNumber("J"), 37, vLAN_BS
        .GetText .ColLetterToNumber("D"), 37, vCHINH_THUC
        If Trim(vCHINH_THUC) = "1" Then
            vLAN_BS = "0"
        End If
        
        If Len(Trim(vTIN_DL)) = 13 Then
            vTIN_DL = Left(vTIN_DL, 10) & "-" & Right(vTIN_DL, 3)
        End If
        
        
        If Trim(vNGAY_HOP_DONG) = "" Or vNGAY_HOP_DONG = vbNullString Then
            vNGAY_HOP_DONG = "null"
        Else
            vNGAY_HOP_DONG = "'" & vNGAY_HOP_DONG & "'"
            vNGAY_HOP_DONG = "to_date(" & vNGAY_HOP_DONG & ",'dd/mm/yyyy')"
        End If
    End With
    strSQL = strSQL & "'" & strMaPhongXuLy & "', null,"
    'strSQL = strSQL & "null)"
    strSQL = strSQL & "null,"
    strSQL = strSQL & "'" & Trim(vTIN_DL) & "','" & Trim(vSO_HOP_DONG) & "'," & vNGAY_HOP_DONG & "," & Trim(vLAN_BS) & ",'" & vLOAI_TK_DK & "')"
    GenerateSQL_Header = strSQL
    'Debug.Print GenerateSQL_Header
    
    Exit Function
ErrHandle:
    SaveErrorLog "cls_02TAIN_DK", "GenerateSQL_Header", Err.Number, Err.Description
End Function

'Private Sub fps_KeyPress(KeyAscii As Integer)
''    If fps.ActiveCol = fps.ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) And fps.ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then _
''        CellEditFormatNumber fps, fps.ActiveSheet, fps.ActiveCol, fps.ActiveRow, KeyAscii
'    With fps
'        If .ActiveSheet = 1 And .ActiveCol = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) _
'           And .ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then
'            If KeyAscii = 46 Or KeyAscii = 45 Then ' KeyAscii of '.' character
'                KeyAscii = 0
'            End If
'        End If
'    End With
'End Sub

Public Function CheckValidData() As Boolean
    Dim varCheckValue As Variant, varNoteValue As Variant
    Dim varNgayNop As Variant
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("G"), 33, varCheckValue
        .GetText .ColLetterToNumber("N"), 33, varNoteValue
        '
        .GetText .ColLetterToNumber("F"), 29, varNgayNop
        If CStr(varCheckValue) = "1" And Trim(CStr(varNoteValue)) = "" Then
            .Sheet = .SheetCount
            .SetText 2, 14, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText 2, 14, "1"
        End If
        '
        If CStr(Trim(varNgayNop)) = "" Or CStr(Trim(varNgayNop)) = "../../...." Then
            .Sheet = .SheetCount
            .SetText 2, 12, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText 2, 12, "1"
        End If
        
    End With
End Function

Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
    KiemTraKhoaSo = True
    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Srv_New.Data(0).nodeFromID("F_27"), "Value"), "dd/mm/yyyy")
    
    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    With fps
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("F") And Row = 29 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    Else
                        .SetActiveCell Col, Row
                    End If
'                Else
'                 .SetActiveCell Col, Row
                End If
               .Col = Col
               .Row = Row
               UpdateCell fps, .Col, .Row, .Text
            End If
            'ky lap bo
            If Col = .ColLetterToNumber("F") And Row = 27 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../...." Then
                    If Format_mmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_mmyyyy(CStr(varTemp))
                    Else
                        .SetActiveCell Col, Row
                    End If
'                Else
'                 .SetActiveCell Col, Row
                End If
               .Col = Col
               .Row = Row
               UpdateCell fps, .Col, .Row, .Text
            End If
        End If
    End With
End Sub
