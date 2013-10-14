VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_BC26AC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_ROW = 10
Const KY_LAP_BO_COL = "E"
Const MA_SO_TEP_ROW = 10
Const MA_SO_TEP_COL = "K"
Const NGAY_NHAN_TO_KHAI_ROW = 12
Const NGAY_NHAN_TO_KHAI_COL = "E"
Const NGUOI_SU_DUNG_ROW = 12
Const NGUOI_SU_DUNG_COL = "Q"
Const HEADER_KY_LAP_BO_ROW = 13
Const HEADER_KY_LAP_BO_COL = "B"

Const HEADER_SO_TT_TRONG_TEP_ROW = 16 'Dong 16 da dc Hidden
Const HEADER_SO_TT_TRONG_TEP_COL = "V"

Const PHONG_XU_LY_ROW = 10
Const PHONG_XU_LY_COL = "Q"
Const PHONG_QUAN_LY_ROW = 24
Const PHONG_QUAN_LY_COL = "Q"

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public strNgayNhanToKhai As String
Public strNguoiSuDung As String
Public strPhongXuLy As String
Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long

Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
        .Col = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_COL)
        .Row = HEADER_SO_TT_TRONG_TEP_ROW
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetNumeric
        .TypeMaxEditLen = 10
        
        
        'Ma so tep
        .Col = .ColLetterToNumber("K")
        .Row = 10
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetAlphanumeric
        .TypeMaxEditLen = 20
        
        'Ghi chu
        .Col = .ColLetterToNumber("K")
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
                
        'Ky lap bo
        SetDateFormat fps, 1, 10, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 12, .ColLetterToNumber("E"), DDMMYYYY
        .Row = 12
        .Col = .ColLetterToNumber("E")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 12, .ColLetterToNumber("K"), DDMMYYYY
        .Row = 12
        .Col = .ColLetterToNumber("K")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
    End With
    
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
    Dim tungay As Variant
    Dim denngay As Variant
    Dim ngay As Variant
    Dim i As Integer
    i = 0
    ReDim Preserve larrPhongXuLy(0)
    larrPhongXuLy(0) = "00"
    ReDim Preserve larrid(0)
    larrid(0) = "00"
    
    With fps
        .Sheet = 1
        .Col = .ColLetterToNumber(PHONG_XU_LY_COL)
        .Row = PHONG_XU_LY_ROW
        If rsPXL.Fields.Count > 0 Then
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

Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

Public Function Prepared3() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim strSoXoa, strSoMat, strSoHuy As Variant
    Dim strTongSo As Variant
    Dim ct20 As Variant
    Dim ct21 As Variant
    Dim ct22 As Variant
    Dim ct05 As Variant
    Dim ct06 As Variant
    Dim ct07 As Variant
    Dim ct08 As Variant
    Dim ct09 As Variant
    Dim tuso As Variant
    Dim denso As Variant
    Dim soluong As Variant
    
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
        '.EventEnabled(EventAllEvents) = True
 
    End With
    Prepared3 = True
End Function

Public Sub Prepared4(ByVal dNgayDauKy As Date)
Dim strKyLapBo As String
Dim Col As Variant
Dim Row As Variant
Dim varTemp As Variant
Dim TongSo As Variant
Dim blnValid As Boolean
    blnValid = True
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        
'        ' Get ky lap bo
'        .Col = .ColLetterToNumber(KY_LAP_BO_COL)
'        .Row = KY_LAP_BO_ROW
'        strKyLapBo = .Text
'
        ' Get Phong xu ly
        .Col = .ColLetterToNumber(PHONG_XU_LY_COL)
        .Row = PHONG_XU_LY_ROW
        strPhongXuLy = .Text
        
        ' Get NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_COL)
        .Row = NGAY_NHAN_TO_KHAI_ROW
        strNgayNhanToKhai = .Text
        
'        'Go to last sheet (header sheet)
'        .Sheet = .SheetCount
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
        
        ' Kiem tra xem ky ke khai trong cung ky hay khac ky
'        .Sheet = .SheetCount
'        If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
'            blnValid = False
'        End If
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
    
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 14, varCheckValue
        .GetText .ColLetterToNumber("K"), 14, varNoteValue
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
    Dim strMaPhongXuLy As String
    
On Error GoTo ErrHandle
    strSQL = strSQL_HDR
    Set xmlList = xmlDOMdata.getElementsByTagName("Cell")
    For Each xmlNode In xmlList
        With xmlNode.Attributes
        
        If Trim(GetAttribute(xmlNode, "MCT")) = vbNullString Then
            Select Case Trim(GetAttribute(xmlNode, "CellID"))
                Case "E_4"
                    vTIN = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_5"
                    vTEN_DTNT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_6"
                    vDIA_CHI = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_12"
                    vNGAY_NOP = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_10"
                    vKyLB = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "K_12"
                    vNGAY_CAP_NHAT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "Q_12"
                    vNGUOI_CAP_NHAT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_14"
                    vCO_LOI_DDANH = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "K_10"
                    ' vSO_HIEU_TEP = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                    ' Tu phien ban 1.3.0 song tep se thuc hien bang tay=> So hieu tep luon bang ""
                    vSO_HIEU_TEP = ""
                Case HEADER_SO_TT_TRONG_TEP_COL & "_" & HEADER_SO_TT_TRONG_TEP_ROW
                    vSO_TT_TK = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "K_14"
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
    vLOAI_TKHAI = "BC26AC"
    strSQL = strSQL & "'" & vLOAI_TKHAI & "',"
    strSQL = strSQL & "To_date('" & vNGAY_NOP & "','dd/mm/yyyy'),"
    
    'Ky/Quy LB
    If Trim(TAX_Utilities_Srv_New.Month) <> "" Then
        'Ngay dau ky lap bo va ngay cuoi ky lap bo
        strDate = Split(vKyLB, "/")
        dDate = DateSerial(Val(strDate(1)), Val(strDate(0)), 1)
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
        dDate = DateAdd("m", 1, dDate)
        dDate = DateAdd("d", -1, dDate)
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    End If
    
    'Ky/ Quy KK
    If Trim(TAX_Utilities_Srv_New.Month) <> "" Then
        'Ngay dau ky ke khai va ngay cuoi ky ke khai
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
    ' Danh dau da nhan la "N" trong truong hop nguoi dung chi can quet bang ke
    If chkQuetBangKe = True Then
        vDA_NHAN = "N"
    Else
        vDA_NHAN = vbNullString
    End If
    strSQL = strSQL & "'" & vDA_NHAN & "',"
    strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
    
    'vCO_GTRINH_02A = TAX_Utilities_Srv_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue
    
    vCO_PLUC_GTGT_01 = TAX_Utilities_Srv_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue
    vCO_PLUC_GTGT_02 = TAX_Utilities_Srv_New.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue
    vCO_PLUC_GTGT_03 = TAX_Utilities_Srv_New.NodeValidity.childNodes(3).Attributes.getNamedItem("Active").nodeValue
    vCO_GTRINH_02A = 0
    vCO_GTRINH_02B = TAX_Utilities_Srv_New.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue
    vCO_GTRINH_02C = TAX_Utilities_Srv_New.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue
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
    End With
    strSQL = strSQL & "'" & strMaPhongXuLy & "',"
    ' Truong hop co quet bang ke thi danh dau truong CO_BANG_KE de xu ly khi dua vao QLT
    If chkQuetBangKe = True Then
        strSQL = strSQL & "1,"
    Else
        strSQL = strSQL & "null,"
    End If
    strSQL = strSQL & "null)"
    GenerateSQL_Header = strSQL
    
    Exit Function
ErrHandle:
    SaveErrorLog "BC26AC", "GenerateSQL_Header", Err.Number, Err.Description
End Function

Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
    KiemTraKhoaSo = True
    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Srv_New.Data(0).nodeFromID("E_10"), "Value"), "dd/mm/yyyy")
    
    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim TongSo As Variant
    Dim strCt10 As Variant
    Dim strCt11 As Variant
    Dim strCt12 As Variant
    Dim FormatSoHD As Variant
    Dim strCt5 As Variant
    Dim soluongTonCK As Variant
    
    With fps
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("E") And Row = 10 Then
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







