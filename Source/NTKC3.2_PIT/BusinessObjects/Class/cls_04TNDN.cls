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
Const KY_LAP_BO_ROW = 30
Const KY_LAP_BO_COL = "E"
Const NGAY_NHAN_TO_KHAI_ROW = 32
Const NGAY_NHAN_TO_KHAI_COL = "E"
Const NGUOI_SU_DUNG_ROW = 32
Const NGUOI_SU_DUNG_COL = "V"
Const HEADER_KY_LAP_BO_ROW = 13
Const HEADER_KY_LAP_BO_COL = "B"

Const HEADER_SO_TT_TRONG_TEP_ROW = 39 'Dong 16 da dc Hidden
Const HEADER_SO_TT_TRONG_TEP_COL = "V"

Const PHONG_XU_LY_ROW = 30
Const PHONG_XU_LY_COL = "V"
' thong tin dai ly thue
Const TEN_DL_ROW = 18
Const TEN_DL_COL = "H"

Const DIA_CHI_DL_ROW = 22
Const DIA_CHI_DL_COL = "F"

Const DIEN_THOAI_DL_ROW = 26
Const DIEN_THOAI_DL_COL = "F"

Const FAX_DL_ROW = 26
Const FAX_DL_COL = "N"

Const EMAIL_DL_ROW = 26
Const EMAIL_DL_COL = "V"

Const SO_HOP_DONG_DL_ROW = 28
Const SO_HOP_DONG_DL_COL = "H"

Const NGAY_HOP_DONG_DL_ROW = 28
Const NGAY_HOP_DONG_DL_COL = "N"


' Ngay phat sinh
Const NGAY_PS_ROW = 28
Const NGAY_PS_COL = "F"
'Lan phat sinh
Const LAN_PS_ROW = 37
Const LAN_PS_COL = "R"


' Phong quan ly
Const PHONG_QUAN_LY_ROW = 34
Const PHONG_QUAN_LY_COL = "V"



Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public strNgayNhanToKhai As String
Public strNguoiSuDung As String
Public strPhongXuLy As String
Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long

Public strTenDL As String
Public strDiaChiDL As String
Public strDienThoaiDL As String
Public strFaxDL As String
Public strEmailDL As String
Public strSoHopDongDL As String
Public strNgayHopDongDL As String

Public strMaCQT As String
Public strMaPQL As String
Public strTenPQL As String

Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
        .Col = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_COL)
        .Row = HEADER_SO_TT_TRONG_TEP_ROW
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetNumeric
        .TypeMaxEditLen = 10
        
        
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
                
        'Ky lap bo
        SetDateFormat fps, 1, 30, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 32, .ColLetterToNumber("E"), DDMMYYYY
        .Row = 32
        .Col = .ColLetterToNumber("E")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 32, .ColLetterToNumber("M"), DDMMYYYY
        .Row = 32
        .Col = .ColLetterToNumber("M")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
    End With
    
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
    Dim i       As Integer
    Dim varMaDM As Variant
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
                
'        'get ten chi tieu trong danh muc NNKD
'        .Sheet = 1
'        .Col = .ColLetterToNumber("C")
'        .Row = 44
'        i = 0
'
'        Do
'            '.GetText .ColLetterToNumber("AV"), .Row, varMaDM
'            .GetText .ColLetterToNumber("C"), .Row, varMaDM
'
'            If Trim(varMaDM) <> vbNullString Then
'                '.SetText .ColLetterToNumber("C"), .Row, getTenDMNNKD(varMaDM)
'                .Col = .ColLetterToNumber("C")
'                If (Trim(getTenDMNNKD(varMaDM)) <> "") Then
'                    .Text = getTenDMNNKD(varMaDM)
'                Else
'                    .Text = varMaDM
'                End If
'
'            End If
'
'            i = i + 1
'            .Col = .ColLetterToNumber("B")
'            .Row = 44 + i
'        Loop Until .Text = "aa"

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
        '.EventEnabled(EventAllEvents) = False
        'Set MaSoTep
        .Sheet = 1
'        .Col = .ColLetterToNumber(MA_SO_TEP_X)
'        .Row = MA_SO_TEP_Y
'        .Text = strMaSoTep
'        UpdateCell fps, .Col, .Row, .Text
        
        'Set NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_COL)
        .Row = NGAY_NHAN_TO_KHAI_ROW
        If strNgayNhanToKhai <> "" Then
            .Text = strNgayNhanToKhai
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        'Set NguoiSuDung
        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_COL)
        .Row = NGUOI_SU_DUNG_ROW
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
        
        ' set Phong quan ly
        .Col = .ColLetterToNumber(PHONG_QUAN_LY_COL)
        .Row = PHONG_QUAN_LY_ROW
        .Text = TAX_Utilities_Srv_New.Convert(strTenPQL, TCVN, UNICODE)
        UpdateCell fps, .Col, .Row, .Text
        
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
        
        ' Get ky lap bo
        .Col = .ColLetterToNumber(KY_LAP_BO_COL)
        .Row = KY_LAP_BO_ROW
        strKyLapBo = .Text
                        
        ' Get Phong xu ly
        .Col = .ColLetterToNumber(PHONG_XU_LY_COL)
        .Row = PHONG_XU_LY_ROW
        strPhongXuLy = .Text
        
        ' Get NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_COL)
        .Row = NGAY_NHAN_TO_KHAI_ROW
        strNgayNhanToKhai = .Text
        
        'Go to last sheet (header sheet)
        .Sheet = .SheetCount
'        ' Kiem tra ky lap bo va ky ke khai
'        If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
'            blnValid = False
'        End If
'
'        .Col = .ColLetterToNumber(HEADER_KY_LAP_BO_COL)
'        .Row = HEADER_KY_LAP_BO_ROW
'
'        If Not blnValid Then
'            .Formula = "0"
'        Else
'            .Formula = "1"
'        End If
'
'        ' Kiem tra xem ky ke khai trong cung ky hay khac ky
'        If (DateAdd("M", 1, dNgayDauKy) = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Or (dNgayDauKy = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Then
'            blnValid = True
'        Else
'            blnValid = False
'        End If
'
'        .Col = .ColLetterToNumber("B")
'        .Row = 17
'
'        If Not blnValid Then
'            .Formula = "0"
'        Else
'            .Formula = "1"
'        End If
        ' Kiem tra neu ky lap bo trang se bao loi
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
            
            
'            ' Kiem tra xem ky ke khai trong cung ky hay khac ky
'            If (DateAdd("M", 1, dNgayDauKy) = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Or (dNgayDauKy = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Then
'                blnValid = True
'            Else
'                blnValid = False
'            End If
'
'            .Col = .ColLetterToNumber("B")
'            .Row = 17
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

Private Sub fps_KeyPress(KeyAscii As Integer)
'    If fps.ActiveCol = fps.ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_COL) And fps.ActiveRow = HEADER_SO_TT_TRONG_TEP_ROW Then _
'        CellEditFormatNumber fps, fps.ActiveSheet, fps.ActiveCol, fps.ActiveRow, KeyAscii
'    With fps
'        If .ActiveSheet = 1 And .ActiveCol = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_COL) _
'           And .ActiveRow = HEADER_SO_TT_TRONG_TEP_ROW Then
'            If KeyAscii = 46 Or KeyAscii = 45 Then ' KeyAscii of '.' character
'                KeyAscii = 0
'            End If
'        End If
'    End With
End Sub

Public Function CheckValidData() As Boolean
    Dim varCheckValue As Variant, varNoteValue As Variant
    Dim varNgayNop As Variant
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 36, varCheckValue
        .GetText .ColLetterToNumber("M"), 36, varNoteValue
        '
        .GetText .ColLetterToNumber("E"), 32, varNgayNop
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

Public Function GenerateSQL_Header(xmlDOMdata As MSXML.DOMDocument, strSQL_HDR As String, vHdrID As Variant, vKKBS As Byte, ByVal dNgayDauKy As Date, Optional chkQuetBangKe As Boolean) As String
    Dim xmlList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlAttribute As MSXML.IXMLDOMAttribute
    Dim iRowID As Long, strSQL As String, strTempSQL As String
    Dim dDate As Date, strDate() As String
    Dim vTIN, vTEN_DTNT, vDIA_CHI, vLOAI_TKHAI, vNGAY_NOP, vKyLB
    Dim vKYKK, vNGAY_CAP_NHAT, vNGUOI_CAP_NHAT, vCO_LOI_DDANH
    Dim vSO_HIEU_TEP, vSO_TT_TK, vDA_NHAN, vGHI_CHU_LOI
    Dim vCO_GTRINH_02A, vCO_GTRINH_02B, vCO_GTRINH_02C
    Dim vCO_PLUC_GTGT_01, vCO_PLUC_GTGT_02, vCO_PLUC_GTGT_03
    Dim vPHONG_XU_LY
    Dim i As Long, j As Long
    Dim strMaPhongXuLy As String, isPhatSinh As Boolean
    
    Dim vTIN_DL, vSO_HOP_DONG, vNGAY_HOP_DONG, vLAN_BS, vCHINH_THUC, vNGANH_NGHE_KD, vTO_KHAI_LAN_PS, vNGAY_PS As Variant
    
On Error GoTo ErrHandle
    strSQL = strSQL_HDR
    Set xmlList = xmlDOMdata.getElementsByTagName("Cell")
    For Each xmlNode In xmlList
        With xmlNode.Attributes
        
        If Trim(GetAttribute(xmlNode, "MCT")) = vbNullString Then
            Select Case Trim(GetAttribute(xmlNode, "CellID"))
                Case "F_10"
                    vTIN = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "H_8"
                    vTEN_DTNT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "F_12"
                    vDIA_CHI = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_32"
                    vNGAY_NOP = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_30"
                    vKyLB = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "M_32"
                    vNGAY_CAP_NHAT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "V_32"
                    vNGUOI_CAP_NHAT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_36"
                    vCO_LOI_DDANH = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "M_13"
                    ' vSO_HIEU_TEP = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                    ' Tu phien ban 1.3.0 song tep se thuc hien bang tay=> So hieu tep luon bang ""
                    vSO_HIEU_TEP = ""
                Case HEADER_SO_TT_TRONG_TEP_COL & "_" & HEADER_SO_TT_TRONG_TEP_ROW
                    vSO_TT_TK = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "M_36"
                    vGHI_CHU_LOI = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case PHONG_XU_LY_COL & "_" & PHONG_XU_LY_ROW
                    vPHONG_XU_LY = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
            End Select
        End If
        End With
    Next
    
    isPhatSinh = False
    ' ngay phat sinh
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber(NGAY_PS_COL), NGAY_PS_ROW, vNGAY_PS
        If (Trim(vNGAY_PS) <> "") Then
            isPhatSinh = True
        End If
    End With
    
    strSQL = strSQL & "'" & vHdrID & "',"
    strSQL = strSQL & vKKBS & ","
    strSQL = strSQL & "'" & vTIN & "',"
    strSQL = strSQL & "'" & vTEN_DTNT & "',"
    strSQL = strSQL & "'" & vDIA_CHI & "',"
    'vLOAI_TKHAI = TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue
    vLOAI_TKHAI = "04_TNDN"
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
    If Trim(TAX_Utilities_Srv_New.ThreeMonths) <> "" And isPhatSinh = False Then
        'Ngay dau ky ke khai va ngay cuoi ky ke khai
        dDate = DateSerial(CInt(TAX_Utilities_Srv_New.Year), 1, 1)
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
        dDate = DateAdd("m", 12, dDate)
        dDate = DateAdd("d", -1, dDate)
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    ElseIf isPhatSinh = True Then
            strSQL = strSQL & "To_date('" & vNGAY_PS & "','dd/mm/yyyy'),"
            strSQL = strSQL & "To_date('" & vNGAY_PS & "','dd/mm/yyyy'),"
    End If
    
    strSQL = strSQL & "To_date('" & vNGAY_CAP_NHAT & "','dd/mm/yyyy'),"
    strSQL = strSQL & "'" & vNGUOI_CAP_NHAT & "',"
    strSQL = strSQL & "'" & vCO_LOI_DDANH & "',"
    strSQL = strSQL & "'" & vSO_HIEU_TEP & "',"
    strSQL = strSQL & "'" & vSO_TT_TK & "',"
    
    strSQL = strSQL & "'" & vDA_NHAN & "',"
    strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
    
    'vCO_GTRINH_02A = TAX_Utilities_Srv_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue
    
    vCO_PLUC_GTGT_01 = 0
    vCO_PLUC_GTGT_02 = 0
    vCO_PLUC_GTGT_03 = 0
    vCO_GTRINH_02A = 0
    vCO_GTRINH_02B = 0
    vCO_GTRINH_02C = 0
    strSQL = strSQL & "'" & IIf(Val(vCO_GTRINH_02A) = 1, "Y", "") & "',"
    strSQL = strSQL & "'" & IIf(Val(vCO_GTRINH_02B) = 1, "Y", "") & "',"
    strSQL = strSQL & "'" & IIf(Val(vCO_GTRINH_02C) = 1, "Y", "") & "',"
    strSQL = strSQL & "'" & IIf(Val(vCO_PLUC_GTGT_01) = 1, "Y", "") & "',"
    strSQL = strSQL & "'" & IIf(Val(vCO_PLUC_GTGT_02) = 1, "Y", "") & "',"
    strSQL = strSQL & "'" & IIf(Val(vCO_PLUC_GTGT_03) = 1, "Y", "") & "',"
    strSQL = strSQL & "null,null,"
    
    With fps
        For i = 1 To lSoPhongXL
            If vPHONG_XU_LY = TAX_Utilities_Srv_New.Convert(larrPhongXuLy(i), UNICODE, TCVN) Then
                strMaPhongXuLy = larrid(i)
                Exit For
            End If
        Next
        
        .Sheet = 1
       ' Ma dai ly thue
        .GetText .ColLetterToNumber("F"), 20, vTIN_DL
        .GetText .ColLetterToNumber(SO_HOP_DONG_DL_COL), SO_HOP_DONG_DL_ROW, vSO_HOP_DONG
        .GetText .ColLetterToNumber(NGAY_HOP_DONG_DL_COL), NGAY_HOP_DONG_DL_ROW, vNGAY_HOP_DONG
        .GetText .ColLetterToNumber("O"), 37, vLAN_BS
        .GetText .ColLetterToNumber("M"), 37, vCHINH_THUC
        If Trim(vCHINH_THUC) = "1" Then
            vLAN_BS = "0"
        End If
        
        If Len(Trim(vTIN_DL)) = 13 Then
            vTIN_DL = Left(vTIN_DL, 10) & "-" & Right(vTIN_DL, 3)
        End If
        
  
        vNGANH_NGHE_KD = "null"
        ' ngay phat sinh
        '.GetText .ColLetterToNumber(NGAY_PS_COL), NGAY_PS_ROW, vNGAY_PS
        If Trim(vNGAY_PS) = "" Or vNGAY_PS = vbNullString Then
            vNGAY_PS = "null"
            vTO_KHAI_LAN_PS = "null"
        Else
            vNGAY_PS = "'" & vNGAY_PS & "'"
            vNGAY_PS = "to_date(" & vNGAY_PS & ",'dd/mm/yyyy')"
            vTO_KHAI_LAN_PS = "1"
        End If
        
        
        If Trim(vNGAY_HOP_DONG) = "" Or vNGAY_HOP_DONG = vbNullString Then
            vNGAY_HOP_DONG = "null"
        Else
            vNGAY_HOP_DONG = "'" & vNGAY_HOP_DONG & "'"
            vNGAY_HOP_DONG = "to_date(" & vNGAY_HOP_DONG & ",'dd/mm/yyyy')"
        End If
    End With
    strSQL = strSQL & "'" & strMaPhongXuLy & "',"
    ' Truong hop co quet bang ke thi danh dau truong CO_BANG_KE de xu ly khi dua vao QLT
    If chkQuetBangKe = True Then
        strSQL = strSQL & "1,"
    Else
        strSQL = strSQL & "null,"
    End If
    'strSQL = strSQL & "null)"
    strSQL = strSQL & "null,"
    strSQL = strSQL & "'" & Trim(vTIN_DL) & "','" & Trim(vSO_HOP_DONG) & "'," & vNGAY_HOP_DONG & "," & Trim(vLAN_BS) & "," & vNGANH_NGHE_KD & "," & vTO_KHAI_LAN_PS & "," & vNGAY_PS & ")"
    GenerateSQL_Header = strSQL
    
    Exit Function
ErrHandle:
    SaveErrorLog "cls_04TNDN", "GenerateSQL_Header", Err.Number, Err.Description
End Function

Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
    KiemTraKhoaSo = True
    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Srv_New.Data(0).nodeFromID("E_30"), "Value"), "dd/mm/yyyy")
    
    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
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
            If Col = .ColLetterToNumber("E") And Row = 30 Then
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

