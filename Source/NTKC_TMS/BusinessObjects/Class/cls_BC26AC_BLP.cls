VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_BC26AC_BLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Const KY_LAP_BO_ROW = 10
'Const KY_LAP_BO_COL = "E"
'Const MA_SO_TEP_ROW = 10
'Const MA_SO_TEP_COL = "M"
Const NGAY_NHAN_TO_KHAI_ROW = 10
Const NGAY_NHAN_TO_KHAI_COL = "E"
Const NGUOI_SU_DUNG_ROW = 12
Const NGUOI_SU_DUNG_COL = "Q"

Const HEADER_SO_TT_TRONG_TEP_ROW = 10 'Dong 24 da dc Hidden
Const HEADER_SO_TT_TRONG_TEP_COL = "K"

Const PHONG_XU_LY_ROW = 10
Const PHONG_XU_LY_COL = "Q"

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public strNgayNhanToKhai As String
Public strNguoiSuDung As String
Public strPhongXuLy As String
Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long
Public strMaCQT As String
Public strMaPQL As String
Public strTenPQL As String

Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
'        .Col = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_COL)
'        .Row = HEADER_SO_TT_TRONG_TEP_ROW
'        .CellType = CellTypeEdit
'        .TypeEditCharSet = TypeEditCharSetNumeric
'        .TypeMaxEditLen = 10
        
        'Ma so tep
'        .Col = .ColLetterToNumber("M")
'        .Row = 10
'        .CellType = CellTypeEdit
'        .TypeEditCharSet = TypeEditCharSetAlphanumeric
'        .TypeMaxEditLen = 20
        
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
    
        'Ky lap bo
'        SetDateFormat fps, 1, 10, .ColLetterToNumber("E"), MMYYYY
'        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 10, .ColLetterToNumber("E"), DDMMYYYY
        .Row = 10
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

Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

Public Function Prepared3() As Boolean
    Dim strDataFileName As String
    Dim strId, strMauso, strTuSo, strDenSo As Variant
    Dim strTen As String
    Dim i As Integer
    With fps
        '.EventEnabled(EventAllEvents) = False
        'Set MaSoTep
        .Sheet = 1
        
        
'        strDataFileName = GetAbsolutePath("..\InterfaceTemplates\xml\Catalogue_loai_HD.xml")
   
    ' Lay danh muc loai hoa don
'        i = 42
'        Do
'
'             .Row = i
'             .GetText .ColLetterToNumber("AH"), .Row, strId
'             .GetText .ColLetterToNumber("O"), .Row, strMauso
'             .GetText .ColLetterToNumber("V"), .Row, strTuSo
'             .GetText .ColLetterToNumber("Y"), .Row, strDenSo
'
'             DataDM Trim(CStr(strId)), Trim(CStr(strMauso)), strTen
'             .SetText .ColLetterToNumber("C"), .Row, strTen
'
'             UpdateCell fps, .ColLetterToNumber("C"), .Row, .Text
'
'             .SetText .ColLetterToNumber("AA"), .Row, Val(strDenSo) - Val(strTuSo) + 1
'             UpdateCell fps, .ColLetterToNumber("AA"), .Row, .Text
'
'             .Col = .ColLetterToNumber("B")
'             i = i + 1
'        Loop Until .Text = "aa"
        
        ' set Phong quan ly
        .Col = .ColLetterToNumber("Q")
        .Row = 14
        .Text = TAX_Utilities_Srv_New.Convert(strTenPQL, TCVN, UNICODE)
        
        UpdateCell fps, .Col, .Row, .Text
        
        
        
        
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
    'UpdateSheets
    Prepared3 = True
End Function

Public Sub Prepared4(ByVal dNgayDauKy As Date)
Dim strKyLapBo As String
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
'
        ' Get NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_COL)
        .Row = NGAY_NHAN_TO_KHAI_ROW
        strNgayNhanToKhai = .Text
'
'        'Go to last sheet (header sheet)
'        .Sheet = .SheetCount
'
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
    Dim varNgayNop As Variant
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 12, varCheckValue
        .GetText .ColLetterToNumber("E"), 14, varNoteValue
        
        .GetText .ColLetterToNumber("E"), 10, varNgayNop
        
        If CStr(varCheckValue) = "1" And Trim(CStr(varNoteValue)) = "" Then
            .Sheet = .SheetCount
            .SetText 2, 13, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText 2, 13, "1"
        End If
        
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

Public Function GenerateSQL_Header(xmlDOMdata As MSXML.DOMDocument, strSQL_HDR As String, vHdrID As Variant, vKKBS As Byte, ByVal dNgayDauKy As Date) As String
    Dim xmlList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlAttribute As MSXML.IXMLDOMAttribute
    Dim iRowID As Long, strSQL As String, strTempSQL As String
    Dim dDate As Date, strDate() As String
    Dim vTIN, vTEN_DTNT, vDIA_CHI, vLOAI_TKHAI, vNGAY_NOP, vKyLB, vPHONG_QUAN_LY, vNGUOI_DAI_DIEN, vTEN_CQ_TIEP_NHAN, vGHI_CHU, vNGAY_BC, vLY_DO
    Dim vNGAY_CAP_NHAT, vNGUOI_CAP_NHAT, vCO_LOI_DDANH, vNGAY_CHAY_HONG, vMA_CQT, vNGUOI_LB, vKY_BC_TUNGAY, vKY_BC_DENNGAY, vQUY_BC, vLoai_BC26
    Dim vSO_HIEU_TEP, vSO_TT_TK, vDA_NHAN, vGHI_CHU_LOI, vCO_GTRINH_02A
    Dim vCO_BANG_KE1, vCO_BANG_KE2, vCO_BANG_KE
    Dim vPHONG_XU_LY
    Dim i As Long, j As Long
    Dim strMaPhongXuLy As String
    Dim vKyBCCuoi, vChuyenDiaDiem
    
On Error GoTo ErrHandle
    strSQL = SqlHdr_AC
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
                Case "E_10"
                    vNGAY_NOP = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "K_12"
                    vNGAY_CAP_NHAT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "Q_12"
                    vNGUOI_CAP_NHAT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_12"
                    vCO_LOI_DDANH = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case HEADER_SO_TT_TRONG_TEP_COL & "_" & HEADER_SO_TT_TRONG_TEP_ROW
                    vSO_TT_TK = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "E_14"
                    vGHI_CHU_LOI = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case PHONG_XU_LY_COL & "_" & PHONG_XU_LY_ROW
                    vPHONG_XU_LY = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
            End Select
        End If
        End With
    Next
    
    vCO_BANG_KE1 = TAX_Utilities_Srv_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue
    'vCO_BANG_KE2 = TAX_Utilities_Srv_New.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue
    If vCO_BANG_KE1 = 1 Then
        vCO_BANG_KE = 1
'        fps.Sheet = 2
'        fps.GetText fps.ColLetterToNumber("M"), 19, vTEN_CQ_TIEP_NHAN
    ElseIf vCO_BANG_KE2 = 1 Then
        vCO_BANG_KE = 1
        vTEN_CQ_TIEP_NHAN = ""
    Else
        vCO_BANG_KE = 0
        vTEN_CQ_TIEP_NHAN = ""
    End If
    
    'Lay thong tin ly do, nguoi dai dien, ngay BC
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        .GetText .ColLetterToNumber("I"), 19, vNGUOI_LB
        .GetText .ColLetterToNumber("J"), 19, vNGUOI_DAI_DIEN
        .GetText .ColLetterToNumber("K"), 19, vNGAY_BC
        
        .GetText .ColLetterToNumber("E"), 23, vKY_BC_TUNGAY
        .GetText .ColLetterToNumber("F"), 23, vKY_BC_DENNGAY
        
        .GetText .ColLetterToNumber("E"), 16, vKyBCCuoi
        .GetText .ColLetterToNumber("K"), 16, vChuyenDiaDiem
                
        .EventEnabled(EventAllEvents) = False
    End With
    vNGUOI_LB = Trim(TAX_Utilities_Srv_New.Convert(CStr(vNGUOI_LB), UNICODE, TCVN))
    vNGUOI_DAI_DIEN = Trim(TAX_Utilities_Srv_New.Convert(CStr(vNGUOI_DAI_DIEN), UNICODE, TCVN))
    If Trim(vTEN_CQ_TIEP_NHAN) <> "" Then
        vTEN_CQ_TIEP_NHAN = Trim(TAX_Utilities_Srv_New.Convert(CStr(vTEN_CQ_TIEP_NHAN), UNICODE, TCVN))
        vTEN_CQ_TIEP_NHAN = Mid(vTEN_CQ_TIEP_NHAN, InStr(vTEN_CQ_TIEP_NHAN, "{") + 1, InStr(vTEN_CQ_TIEP_NHAN, "}") - InStr(vTEN_CQ_TIEP_NHAN, "{") - 1)
    End If
    vPHONG_QUAN_LY = strMaPQL
    ' Lay Quy BC
    If Trim(TAX_Utilities_Srv_New.ThreeMonths) <> "" Then
        dDate = dNgayDauKy
        vQUY_BC = Format(dDate, "dd/mm/yyyy")
    End If
    ' lay gia tri loai BC 26
'    If Trim(CStr(vKyBCCuoi)) = "0" And Trim(CStr(vChuyenDiaDiem)) = "0" Then
'        ' bao cao quy
'        vLoai_BC26 = "01"
'    ElseIf Trim(CStr(vKyBCCuoi)) = "0" And Trim(CStr(vChuyenDiaDiem)) = "1" Then
'        ' bao cao chuyen dia diem
'        vLoai_BC26 = "03"
'    ElseIf Trim(CStr(vKyBCCuoi)) = "1" And Trim(CStr(vChuyenDiaDiem)) = "0" Then
'        ' bao cao quyet toan cuoi tung
'        vLoai_BC26 = "02"
'    Else
'        '
'        vLoai_BC26 = "04"
'    End If
    vLoai_BC26 = "01"
    
    
    ' Tach ma so thue thanh ma 14 ky tu
    If Len(Trim(CStr(vTIN))) = 13 Then
        vTIN = Left(CStr(vTIN), 10) & "-" & Right(CStr(vTIN), 3)
    End If
    
    strSQL = strSQL & "'" & vHdrID & "',"
    strSQL = strSQL & "'" & vTIN & "',"
'    strSQL = strSQL & "'" & vTEN_DTNT & "',"
'    strSQL = strSQL & "'" & vDIA_CHI & "',"
    
    vLOAI_TKHAI = "BC26_AC_BLP"
    vMA_CQT = strMaCQT
    
    strSQL = strSQL & "'" & vLOAI_TKHAI & "',"
    strSQL = strSQL & "To_date('" & vNGAY_NOP & "','dd/mm/yyyy'),"
    strSQL = strSQL & "To_date('" & vKY_BC_TUNGAY & "','dd/mm/yyyy'),"
    strSQL = strSQL & "To_date('" & vKY_BC_DENNGAY & "','dd/mm/yyyy'),"
    strSQL = strSQL & "To_date('" & vNGAY_CAP_NHAT & "','dd/mm/yyyy'),"
    strSQL = strSQL & "'" & vNGUOI_CAP_NHAT & "',"
'    strSQL = strSQL & "'" & vCO_LOI_DDANH & "',"
    strSQL = strSQL & "'" & vSO_TT_TK & "',"
    
    strSQL = strSQL & "'" & vDA_NHAN & "',"
    
    'strMaPhongXuLy = larrid(fps.TypeComboBoxCurSel)
'    strSQL = strSQL & "'','','','','',"
'    strSQL = strSQL & "'',null,null,"
    With fps
        For i = 1 To lSoPhongXL
            If vPHONG_XU_LY = TAX_Utilities_Srv_New.Convert(larrPhongXuLy(i), UNICODE, TCVN) Then
                strMaPhongXuLy = larrid(i)
                Exit For
            End If
        Next
    End With
    
    strSQL = strSQL & "'" & strMaPhongXuLy & "',"
    strSQL = strSQL & "'" & vPHONG_QUAN_LY & "',"
    strSQL = strSQL & vCO_BANG_KE & ",null,null,null,null,"
    strSQL = strSQL & "To_date('" & vNGAY_BC & "','dd/mm/yyyy'),"
    strSQL = strSQL & "'" & vNGUOI_DAI_DIEN & "',"
    strSQL = strSQL & "'" & vTEN_CQ_TIEP_NHAN & "',"
    strSQL = strSQL & "'" & "" & "',"
    strSQL = strSQL & "null,"
    strSQL = strSQL & "null,null,"
    strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
    strSQL = strSQL & "'" & vMA_CQT & "',"
    strSQL = strSQL & "'" & vLoai_BC26 & "',"
    'strSQL = strSQL & "null,"
    strSQL = strSQL & "'" & vNGUOI_LB & "',"
    strSQL = strSQL & "To_date('" & vQUY_BC & "','dd/mm/yyyy'))"
    GenerateSQL_Header = strSQL
    Exit Function
ErrHandle:
    SaveErrorLog "BC26_AC_BLP", "GenerateSQL_Header", Err.Number, Err.Description
End Function

Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
    KiemTraKhoaSo = False
'    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
'    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Srv_New.Data(0).nodeFromID("E_10"), "Value"), "dd/mm/yyyy")
'
'    If dNgayKhoaSo < dKyLapBo Then
'        KiemTraKhoaSo = False
'    End If
End Function

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
 Dim varTemp As Variant
 With fps
    If .ActiveSheet = 1 Then
        .Sheet = 1
        If Col = .ColLetterToNumber("E") And Row = 10 Then
            .GetText Col, Row, varTemp
            If varTemp <> "" And varTemp <> "../../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                Else
                    .SetActiveCell Col, Row
                End If
'            Else
'             .SetActiveCell Col, Row
            End If
            .Col = .ColLetterToNumber("E")
            .Row = 10
           UpdateCell fps, .Col, .Row, .Text
        End If
    End If
End With
End Sub


Private Sub DataDM(ByVal Id As String, ByVal strMauso As String, Optional ByRef strTen As String)
Dim arrDanhsach() As String
Dim strDataFileName As String
Dim xmlDOMdata As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = GetAbsolutePath("..\InterfaceTemplates\xml\Catalogue_loai_HD.xml")
    
       If xmlDOMdata.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDOMdata.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If Id = "0" Then
                            strMauso = Mid(strMauso, 1, 6)
                        Else
                            strMauso = Mid(strMauso, 1, 3)
                        End If
                        
                        If Id = arrDanhsach(0) And strMauso = arrDanhsach(1) Then
                            strTen = arrDanhsach(2)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub

