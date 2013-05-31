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
Private clsConn As New TAX_Utilities_Svr_New.clsADO

Const KY_LAP_BO_Y = 10
Const KY_LAP_BO_X = "E"
Const MA_SO_TEP_Y = 10
Const MA_SO_TEP_X = "M"
Const NGAY_NHAN_TO_KHAI_Y = 12
Const NGAY_NHAN_TO_KHAI_X = "E"
Const NGUOI_SU_DUNG_Y = 12
Const NGUOI_SU_DUNG_X = "V"
Const HEADER_KY_LAP_BO_Y = 13
Const HEADER_KY_LAP_BO_X = "B"
Const HEADER_SO_TT_TRONG_TEP_X = "V"
Const HEADER_SO_TT_TRONG_TEP_Y = 16
Const PHONG_XU_LY_Y = 10
Const PHONG_XU_LY_X = "V"
'Longvh
Const MA_SO_THUE_X = "G"
Const MA_SO_THUE_Y = 4
Const NGAY_NOP_Y = 12
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 12
Const NGAY_QUET_X = "M"

Const TEN_GOI_X = "G"
Const TEN_GOI_Y = 5
Const DIA_CHI_X = "G"
Const DIA_CHI_Y = 6
Const MA_BPQL_X = "V"
Const MA_BPQL_Y = 12
Const DIEN_THOAI_X = "E"
Const DIEN_THOAI_Y = 8
Const FAX_X = "M"
Const FAX_Y = 8


'Const HEADER_SO_TT_TRONG_TEP_X = "V"
'Const HEADER_SO_TT_TRONG_TEP_Y = 10
'Const PHONG_XU_LY_Y = 16
'Const PHONG_XU_LY_X = "E"

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'Public strMaSoTep As String
Public strNgayNhanToKhai As String
Public strNguoiSuDung As String
Public strMST As String
Public strPhongXuLy As String
Public strTenGoi As String
Public strDchi As String
Public strNganh As String
Public strMaBPQL As String
Public strDThoai As String
Public strFax As String
Public spathVat As String

Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long
Public sSaiCT11 As String
Public sKyKeKhai As String
Public dNgayNopDB As Variant
Public HANNOP As String
Public TTHTK As String


Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
'        .Col = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X)
'        .Row = HEADER_SO_TT_TRONG_TEP_Y
'        .CellType = CellTypeEdit
'        .TypeEditCharSet = TypeEditCharSetNumeric
'        .TypeMaxEditLen = 10
        
        'Ma so tep
        .Col = .ColLetterToNumber("M")
        .Row = 10
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetAlphanumeric
        .TypeMaxEditLen = 20
        
        'Ghi chu
        .Col = .ColLetterToNumber("M")
        .Row = 14
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100
        
        'Phong xu ly
        .Col = .ColLetterToNumber("E")
        .Row = 16
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
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
        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y
        If rsPXL.Fields.Count > 0 Then
            Do While Not rsPXL.EOF
                .TypeComboBoxIndex = -1
                .TypeComboBoxString = TAX_Utilities_Svr_New.Convert(rsPXL.Fields(1).Value, TCVN, UNICODE)
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
        .Value = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        UpdateCell fps, .Col, .Row, .Value
        
        .Col = .ColLetterToNumber(NGAY_QUET_X)
        .Row = NGAY_QUET_Y
        .Value = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
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
        
        .Col = .ColLetterToNumber(MA_BPQL_X)
        .Row = MA_BPQL_Y
        If strMaBPQL <> "" Then
            .Text = strMaBPQL
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
    With fps
        If .ActiveSheet = 1 And .ActiveCol = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) _
           And .ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then
            If KeyAscii = 46 Or KeyAscii = 45 Then ' KeyAscii of '.' character
                KeyAscii = 0
            End If
        End If
    End With
End Sub

Public Function CheckValidData() As Boolean
    Dim varCheckValue As Variant, varNoteValue As Variant
    
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 14, varCheckValue
        .GetText .ColLetterToNumber("M"), 14, varNoteValue
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

Public Function GenerateSQL_Header(xmlDOMdata As MSXML.DOMDocument, strSQL_HDR As String, vHdrID As Variant, ByVal dNgayDauKy As Date) As String
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
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'',"
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
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("E_10"), "Value"), "dd/mm/yyyy")
    
    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function

Public Function InsertDTL() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim MST As Variant
   Dim KYKK As Variant
   Dim IDHDR As Variant
   Dim LOAITK As Variant
   Dim MACT As Variant
   Dim GIATRI As Variant
   Dim bln  As Boolean
   
   sSQLCol = "IDHDR, LOAITK, MACT, GIATRI"
   
   If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & "\TK_NEW\" 'GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
   End If
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("G"), 4, MST
        .GetText .ColLetterToNumber("E"), 10, KYKK
        IDHDR = MST & Mid(KYKK, 1, 2) & Mid(KYKK, 4, 4)
        
        If Trim(IDHDR) = vbNullString Then
            IDHDR = "''"
        Else
            IDHDR = "'" & IDHDR & "'"
        End If
        LOAITK = "'04TNDN'"
   End With
   
   ' Ma chi tieu [10]
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("S"), 17, MACT
        
        If Trim(MACT) = vbNullString Then
            MACT = "''"
        Else
            MACT = "'" & MACT & "'"
        End If
        
        .GetText .ColLetterToNumber("T"), 17, GIATRI
        If Trim(GIATRI) = vbNullString Then
            GIATRI = "0"
        End If
   End With
    
   sSQLVal = IDHDR & "," & LOAITK & "," & MACT & "," & GIATRI
   sSQL = "INSERT INTO TKDTL( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
   
   ' Ma chi tieu [11]
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("S"), 18, MACT
        
        If Trim(MACT) = vbNullString Then
            MACT = "''"
        Else
            MACT = "'" & MACT & "'"
        End If
        
        .GetText .ColLetterToNumber("T"), 18, GIATRI
        If Trim(GIATRI) = vbNullString Then
            GIATRI = "0"
        End If
   End With
   
   sSQLVal = IDHDR & "," & LOAITK & "," & MACT & "," & GIATRI
   sSQL = "INSERT INTO TKDTL( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)
   
   ' Ma chi tieu [12]
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("S"), 19, MACT
        
        If Trim(MACT) = vbNullString Then
            MACT = "''"
        Else
            MACT = "'" & MACT & "'"
        End If
        
        .GetText .ColLetterToNumber("T"), 19, GIATRI
        If Trim(GIATRI) = vbNullString Then
            GIATRI = "0"
        End If
   End With
    
   sSQLVal = IDHDR & "," & LOAITK & "," & MACT & "," & GIATRI
   sSQL = "INSERT INTO TKDTL( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)

   ' Ma chi tieu [13]
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("S"), 20, MACT
        
        If Trim(MACT) = vbNullString Then
            MACT = "''"
        Else
            MACT = "'" & MACT & "'"
        End If
        
        .GetText .ColLetterToNumber("T"), 20, GIATRI
        If Trim(GIATRI) = vbNullString Then
            GIATRI = "0"
        End If
   End With
    
   sSQLVal = IDHDR & "," & LOAITK & "," & MACT & "," & GIATRI
   sSQL = "INSERT INTO TKDTL( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)

   ' Ma chi tieu [14]
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("S"), 21, MACT
        
        If Trim(MACT) = vbNullString Then
            MACT = "''"
        Else
            MACT = "'" & MACT & "'"
        End If
        
        .GetText .ColLetterToNumber("T"), 21, GIATRI
        If Trim(GIATRI) = vbNullString Then
            GIATRI = "0"
        End If
   End With
    
   sSQLVal = IDHDR & "," & LOAITK & "," & MACT & "," & GIATRI
   sSQL = "INSERT INTO TKDTL( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)

   ' Ma chi tieu [15]
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("S"), 22, MACT
        
        If Trim(MACT) = vbNullString Then
            MACT = "''"
        Else
            MACT = "'" & MACT & "'"
        End If
        
        .GetText .ColLetterToNumber("T"), 22, GIATRI
        If Trim(GIATRI) = vbNullString Then
            GIATRI = "0"
        End If
   End With
    
   sSQLVal = IDHDR & "," & LOAITK & "," & MACT & "," & GIATRI
   sSQL = "INSERT INTO TKDTL( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)

   ' Ma chi tieu [16]
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("S"), 23, MACT
        
        If Trim(MACT) = vbNullString Then
            MACT = "''"
        Else
            MACT = "'" & MACT & "'"
        End If
        
        .GetText .ColLetterToNumber("T"), 23, GIATRI
        If Trim(GIATRI) = vbNullString Then
            GIATRI = "0"
        End If
   End With
    
   sSQLVal = IDHDR & "," & LOAITK & "," & MACT & "," & GIATRI
   sSQL = "INSERT INTO TKDTL( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
   bln = clsConn.ExecuteDLL(sSQL)

   clsConn.Disconnect
   
   InsertDTL = vbNullString
   
End Function

Public Function InsertHDR() As String
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    
    Dim Id As Variant
    Dim MST As Variant
    Dim LOAITK  As Variant
    Dim NGAYNOP As Variant
    Dim KYLB As Variant
    Dim PHONGXL As Variant
    Dim KYKK As Variant
    Dim NGAYQUET As Variant
    Dim GHICHU As Variant
    Dim LOIDD As Variant
    Dim bln  As Boolean
    
    sSQLCol = "ID, MST, LOAITK, NGAYNOP, KYLB, PHONGXL, KYKK, NGAYQUET, GHICHU, LOIDD"
           
    If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & "\TK_NEW\" 'GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
    End If
   
    With fps
        .Sheet = 1
        
        .GetText .ColLetterToNumber("G"), 4, MST
        If Trim(MST) = vbNullString Then
            MST = "''"
        Else
            MST = "'" & MST & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 10, KYLB
        If Trim(KYLB) = vbNullString Then
            KYLB = "''"
        Else
            KYLB = "'" & KYLB & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 12, NGAYNOP
        If Trim(NGAYNOP) = vbNullString Then
            NGAYNOP = "CTOD('')"
        Else
            NGAYNOP = ToDate(Trim(NGAYNOP), DDMMYYYY)
            NGAYNOP = "CTOD('" & Format(NGAYNOP, "mm/dd/yyyy") & "')"
        End If
        
        .GetText .ColLetterToNumber("V"), 12, PHONGXL
        If Trim(PHONGXL) = vbNullString Then
            PHONGXL = "''"
        Else
            PHONGXL = "'" & PHONGXL & "'"
        End If
                
        .GetText .ColLetterToNumber("E"), 10, KYKK
        If Trim(KYKK) = vbNullString Then
            KYKK = "''"
        Else
            KYKK = "'" & KYKK & "'"
        End If
        
        .GetText .ColLetterToNumber("M"), 12, NGAYQUET
        If Trim(NGAYQUET) = vbNullString Then
            NGAYQUET = "CTOD('')"
        Else
            NGAYQUET = "CTOD('" & Format(NGAYQUET, "mm/dd/yyyy") & "')"
        End If
        
        .GetText .ColLetterToNumber("M"), 14, GHICHU
        If Trim(GHICHU) = vbNullString Then
            GHICHU = "''"
        Else
            GHICHU = "'& GHICHU &'"
        End If
        
        .GetText .ColLetterToNumber("E"), 14, LOIDD
        If Trim(LOIDD) = 0 Then
            LOIDD = ".F."
        ElseIf Trim(LOIDD) = 1 Then
            LOIDD = ".T."
        End If
        
        Id = "'" & Mid(MST, 2, Len(MST) - 2) & Mid(KYKK, 2, 2) & Mid(KYKK, 5, 4) & "'"
        LOAITK = "'04TNDN'"
        
    End With
    
    sSQLVal = Id & "," & MST & "," & LOAITK & "," & NGAYNOP & "," & KYLB & "," & PHONGXL & "," & KYKK & "," & NGAYQUET & "," & GHICHU & "," & LOIDD
    
    sSQL = "INSERT INTO TKHDR " & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsConn.ExecuteDLL(sSQL)
    
    clsConn.Disconnect
    InsertHDR = vbNullString
End Function
Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    If clsConn.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsConn.CreateConnectionString spathVat & "\TK_NEW\" 'GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
        
   sSQL = "SELECT MST, NGAYNOP FROM TKHDR WHERE LOAITK = '04TNDN' AND ID = '" & strMST & sKyKeKhai & "'"
   Set rs = clsConn.Execute(sSQL)
   If Not rs Is Nothing Then
        dNgayNopDB = rs.Fields("NGAYNOP")
        If Trim(dNgayNopDB) = vbNullString Then
            dNgayNopDB = "CTOD('')"
        Else
            dNgayNopDB = "CTOD('" & Format(dNgayNopDB, "mm/dd/yyyy") & "')"
        End If
        TKTT = True
   Else
        TKTT = False
   End If
   clsConn.Disconnect

End Function

Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim KYLBO As Variant
    Dim rs As ADODB.Recordset
    Dim idTKHDR As String 'Tam thoi ID cua to khai nay la MST & KYKK
    If clsConn.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsConn.CreateConnectionString spathVat & "\TK_NEW\" 'GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   idTKHDR = strMST & sKyKeKhai
   sSQL = "DELETE FROM TKHDR WHERE LOAITK = '04TNDN' AND ID = '" & idTKHDR & "'"
    
   clsConn.ExecuteDLL sSQL
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
   sSQL = vbNullString
   sSQL = "DELETE FROM TKDTL WHERE LOAITK = '04TNDN' IDHDR = '" & idTKHDR & "'"
   
   clsConn.ExecuteDLL sSQL
   
   sSQL = vbNullString
   sSQL = "DELETE FROM PLDTL WHERE LOAITK = '04TNDN' IDHDR = '" & idTKHDR & "'"
   
   clsConn.ExecuteDLL sSQL
   
   XoaTKTT = True
   
   clsConn.Disconnect

End Function

Public Function TKRB() As Boolean

   TKRB = True

End Function

Public Function TKTTNGNOP() As Boolean
   Dim NGNOP As Variant
   Dim NGNOPDB As Variant
   Dim sSQL As String
   Dim rs As ADODB.Recordset
   Dim i As Integer
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 12, NGNOP
        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = ""
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
        End If
   End With
   
   If clsConn.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsConn.CreateConnectionString spathVat & "\TK_NEW\" 'GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
   sSQL = "SELECT MST, NGAYNOP FROM TKHDR WHERE LOAITK = '04TNDN' AND ID = '" & strMST & sKyKeKhai & "'"
    
   Set rs = clsConn.Execute(sSQL)
   If Not rs Is Nothing Then
        For i = 1 To rs.RecordCount
            If Not rs.EOF Then
                NGNOPDB = rs.Fields("NGAYNOP")
                If Trim(NGNOPDB) = vbNullString Then
                    NGNOPDB = ""
                Else
                    NGNOPDB = "CTOD('" & Format(NGNOPDB, "mm/dd/yyyy") & "')"
                End If
                If Trim(NGNOPDB) <> vbNullString And Trim(NGNOPDB) = Trim(NGNOP) Then
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

