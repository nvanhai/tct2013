VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_BC26AC_BK01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Const KY_LAP_BO_ROW = 10
'Const KY_LAP_BO_COL = "E"
'Const MA_SO_TEP_ROW = 10
'Const MA_SO_TEP_COL = "K"
Const NGAY_NHAN_TO_KHAI_ROW = 10
Const NGAY_NHAN_TO_KHAI_COL = "E"
Const NGUOI_SU_DUNG_ROW = 12
Const NGUOI_SU_DUNG_COL = "Q"
Const HEADER_SO_TT_TRONG_TEP_ROW = 10 'Dong 16 da dc Hidden
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
       .Sheet = 1
        ' lay danh muc loai hoa don
'         strDataFileName = GetAbsolutePath("..\InterfaceTemplates\xml\Catalogue_loai_HD.xml")
'          i = 29
'        Do
'             .Row = i
'             .GetText .ColLetterToNumber("AF"), .Row, strId
'             .GetText .ColLetterToNumber("I"), .Row, strMauso
'             .GetText .ColLetterToNumber("M"), .Row, strTuSo
'             .GetText .ColLetterToNumber("O"), .Row, strDenSo
'
'             DataDM Trim(CStr(strId)), Trim(CStr(strMauso)), strTen
'             .SetText .ColLetterToNumber("D"), .Row, strTen
'
'             UpdateCell fps, .ColLetterToNumber("D"), .Row, .Text
'             .SetText .ColLetterToNumber("K"), .Row, Val(strDenSo) - Val(strTuSo) + 1
'             UpdateCell fps, .ColLetterToNumber("K"), .Row, .Text
'
'             .Col = .ColLetterToNumber("B")
'             i = i + 1
'        Loop Until .Text = "aa"

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
                 
        ' set Phong quan ly
        .Col = .ColLetterToNumber("Q")
        .Row = 14
        .Text = TAX_Utilities_Srv_New.Convert(strTenPQL, TCVN, UNICODE)
        UpdateCell fps, .Col, .Row, .Text
              
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
        .Col = .ColLetterToNumber(PHONG_XU_LY_COL)
        .Row = PHONG_XU_LY_ROW
        strPhongXuLy = .Text
        
        ' get nguoi su dung
        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_COL)
        .Row = NGUOI_SU_DUNG_ROW
        strNguoiSuDung = .Text
        
        ' Get NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_COL)
        .Row = NGAY_NHAN_TO_KHAI_ROW
        strNgayNhanToKhai = .Text
      
        
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
    Dim vTIN, vTEN_DTNT, vLOAI_TKHAI, vNGAY_NOP, vPHONG_QUAN_LY, vNGUOI_DAI_DIEN, vGHI_CHU, vNGAY_BC, vLY_DO
    Dim vNGAY_CAP_NHAT, vDIA_CHI, vNGUOI_CAP_NHAT, vCO_LOI_DDANH, vNGAY_CHAY_HONG, vMA_CQT, vNGUOI_LB, vKY_BC_TUNGAY, vKY_BC_DENNGAY
    Dim vSO_HIEU_TEP, vSO_TT_TK, vDA_NHAN, vGHI_CHU_LOI, vHTHUC_NOP, vTEN_DV_CQ, vTIN_DV_CQ, vTEN_CQ_TIEP_NHAN
    Dim vDUNG_DN_CQ As Variant
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
    'Lay thong tin don vi chu quan, nguoi dai dien, ngay BC
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        .GetText .ColLetterToNumber("R"), 22, vNGUOI_DAI_DIEN
        .GetText .ColLetterToNumber("O"), 22, vNGAY_BC
        .GetText .ColLetterToNumber("E"), 22, vTEN_DV_CQ
        .GetText .ColLetterToNumber("I"), 22, vTIN_DV_CQ
        .GetText .ColLetterToNumber("L"), 22, vTEN_CQ_TIEP_NHAN
        .EventEnabled(EventAllEvents) = False
    End With

    vNGUOI_DAI_DIEN = TAX_Utilities_Srv_New.Convert(CStr(vNGUOI_DAI_DIEN), UNICODE, TCVN)
    vNGUOI_CAP_NHAT = TAX_Utilities_Srv_New.Convert(CStr(vNGUOI_CAP_NHAT), UNICODE, TCVN)
    vGHI_CHU_LOI = TAX_Utilities_Srv_New.Convert(CStr(vGHI_CHU_LOI), UNICODE, TCVN)
    vTEN_DV_CQ = TAX_Utilities_Srv_New.Convert(CStr(vTEN_DV_CQ), UNICODE, TCVN)
    vTEN_CQ_TIEP_NHAN = TAX_Utilities_Srv_New.Convert(CStr(vTEN_CQ_TIEP_NHAN), UNICODE, TCVN)
    
    
    ' Tach ma so thue thanh ma 14 ky tu
    If Len(Trim(CStr(vTIN))) = 13 Then
        vTIN = Left(CStr(vTIN), 10) & "-" & Right(CStr(vTIN), 3)
    End If
     ' Tach ma so thue thanh ma 14 ky tu
    If Len(Trim(CStr(vTIN_DV_CQ))) = 13 Then
        vTIN_DV_CQ = Left(CStr(vTIN_DV_CQ), 10) & "-" & Right(CStr(vTIN_DV_CQ), 3)
    End If
    
    
    strSQL = strSQL & "'" & vHdrID & "',"
    strSQL = strSQL & "'" & vTIN & "',"
'    strSQL = strSQL & "'" & vTEN_DTNT & "',"
'    strSQL = strSQL & "'" & vDIA_CHI & "',"

    vLOAI_TKHAI = "01_BK_BC26_AC"
    vMA_CQT = strMaCQT
    vPHONG_QUAN_LY = strMaPQL
    ' Lay thong tin don vi chu quan dat in, tu in
    If Trim(vTIN) = Trim(vTIN_DV_CQ) Or Trim(vTIN_DV_CQ) = "" Then
        vDUNG_DN_CQ = "N"
    Else
        vDUNG_DN_CQ = "Y"
    End If
    
    
    strSQL = strSQL & "'" & vLOAI_TKHAI & "',"
    strSQL = strSQL & "To_date('" & vNGAY_NOP & "','dd/mm/yyyy'),"
    strSQL = strSQL & "null,null,"
    strSQL = strSQL & "To_date('" & vNGAY_CAP_NHAT & "','dd/mm/yyyy'),"
    strSQL = strSQL & "'" & vNGUOI_CAP_NHAT & "',"
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
    strSQL = strSQL & "'" & vPHONG_QUAN_LY & "',null,null,null,"
    strSQL = strSQL & "'" & vTEN_DV_CQ & "',"
    strSQL = strSQL & "'" & vTIN_DV_CQ & "',"
    strSQL = strSQL & "To_date('" & vNGAY_BC & "','dd/mm/yyyy'),"
    strSQL = strSQL & "'" & vNGUOI_DAI_DIEN & "',"
    strSQL = strSQL & "'" & vTEN_CQ_TIEP_NHAN & "',null,null,null,"
    strSQL = strSQL & "'" & vDUNG_DN_CQ & "',"
    strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
    strSQL = strSQL & "'" & vMA_CQT & "',null,null,null)"
    GenerateSQL_Header = strSQL
    Exit Function
ErrHandle:
    SaveErrorLog "01_BK_BC26_AC", "GenerateSQL_Header", Err.Number, Err.Description
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
        If Col = .ColLetterToNumber("E") And Row = 10 Then
            .Sheet = 1
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
