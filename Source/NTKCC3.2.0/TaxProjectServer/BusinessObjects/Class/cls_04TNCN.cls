VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_04TNCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_Y = 10
Const KY_LAP_BO_X = "E"
Const MA_SO_TEP_Y = 10
Const MA_SO_TEP_X = "L"
Const NGAY_NHAN_TO_KHAI_Y = 12
Const NGAY_NHAN_TO_KHAI_X = "E"
Const NGUOI_SU_DUNG_Y = 12
Const NGUOI_SU_DUNG_X = "R"
Const HEADER_KY_LAP_BO_Y = 13
Const HEADER_KY_LAP_BO_X = "B"
Const MA_BPQL_X = "V"
Const MA_BPQL_Y = 16 'Dong 16 da dc Hidden
Const PHONG_XU_LY_Y = 10
Const PHONG_XU_LY_X = "R"

'Longvh
Const MA_SO_THUE_X = "G"
Const MA_SO_THUE_Y = 4
Const NGAY_NOP_Y = 12
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 12
Const NGAY_QUET_X = "L"

Const TEN_GOI_X = "G"
Const TEN_GOI_Y = 5
Const DIA_CHI_X = "G"
Const DIA_CHI_Y = 6
Const TEN_BPQL_X = "R"
Const TEN_BPQL_Y = 12
Const DIEN_THOAI_X = "E"
Const DIEN_THOAI_Y = 8
Const FAX_X = "L"
Const FAX_Y = 8
Private LAN_QUET As Variant
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

'Longvh
Public strMST As String
Public strTenGoi As String
Public strDchi As String
Public strNganh As String
Public strMaBPQL As String
Public strDThoai As String
Public strFax As String
Public spathVat As String
Public sSaiCT11 As String
Public sKyKeKhai As String
Public dNgayNopDB As Variant
Public HANNOP As String
Public strTenBpql As String
Public TTHTK As String


Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
        
        
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
        
        ' Ky lap bo
        SetDateFormat fps, 1, 10, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 12, .ColLetterToNumber("E"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 12
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 12, .ColLetterToNumber("M"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 12
        .Col = .ColLetterToNumber("M")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
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
        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y
        PHONG_XU_LY_X1 = PHONG_XU_LY_X
        PHONG_XU_LY_Y1 = PHONG_XU_LY_Y
        If rsPXL.Fields.Count > 0 Then
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
        
        .Col = .ColLetterToNumber(MA_BPQL_X)
        .Row = MA_BPQL_Y
        If strMaBPQL <> "" Then
            .Text = strMaBPQL
            UpdateCell fps, .Col, .Row, .Text
        End If
        .Col = .ColLetterToNumber(TEN_BPQL_X)
        .Row = TEN_BPQL_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
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


Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
    KiemTraKhoaSo = True
    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("E_10"), "Value"), "dd/mm/yyyy")
    
    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function

Public Function InsertDTL_PL1() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim bln  As Boolean
   
   Dim MADTNT As Variant
   Dim NGNOP As Variant
   Dim QT As Variant
   Dim SOTT As Variant
   Dim TENDTNT2 As Variant
   Dim MADTNT2 As Variant
   Dim SOTIEN01 As Variant
   Dim SOTIEN02 As Variant
   Dim SOTIEN03 As Variant
   Dim SOTIEN04 As Variant
   Dim SOTIEN05 As Variant
   Dim SOTIEN06 As Variant
   Dim SOTIEN07 As Variant
   Dim SOTIEN08 As Variant
   Dim SOTIEN09 As Variant
   Dim SOTIEN092 As Variant
   Dim GHICHU As Variant
   Dim STT As Integer
   Dim NAMQT As Variant
   Dim DANHAN As Variant
   Dim ROWNUM As Integer
    
   sSQLCol = "MADTNT, NGNOP, TTHTK,QT, SOTT, TENDTNT2, MADTNT2, SOTIEN01, SOTIEN02, SOTIEN03,"
   sSQLCol = sSQLCol + "SOTIEN04, SOTIEN05,SOTIEN06, SOTIEN07,SOTIEN08, SOTIEN09, SOTIEN092, GHICHU, NAMQT, LAN_QUET, DANHAN"
   
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   With fps
   .Sheet = 1
        .GetText .ColLetterToNumber("G"), 4, MADTNT
         If Trim(MADTNT) = vbNullString Then
             MADTNT = "''"
         Else
             MADTNT = "'" & MADTNT & "'"
         End If
         
         .GetText .ColLetterToNumber("E"), 12, NGNOP
         If Trim(NGNOP) = vbNullString Then
             NGNOP = "CTOD('')"
         Else
             NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
             NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
         End If
         
         QT = "''"
         NAMQT = "'" & TAX_Utilities_Svr_New.Year & "'"
        .Sheet = 2
             STT = 0
             .Col = .ColLetterToNumber("B")
             .Row = 26
                 Do
                 
                       .Col = .ColLetterToNumber("C")
                       TENDTNT2 = .Text
                       If Trim(TENDTNT2) = vbNullString Then
                            TENDTNT2 = "''"
                       Else
                            TENDTNT2 = "'" & TAX_Utilities_Svr_New.Convert(CStr(TENDTNT2), UNICODE, TCVN) & "'"
                       End If
                        
                       .Col = .ColLetterToNumber("E")
                       MADTNT2 = .Text
                       If Trim(MADTNT2) = vbNullString Then
                            MADTNT2 = "''"
                       Else
                            MADTNT2 = "'" & MADTNT2 & "'"
                       End If
                       
                       .Col = .ColLetterToNumber("G")
                       SOTIEN01 = .Value
                       If Trim(SOTIEN01) = vbNullString Then
                            SOTIEN01 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("H")
                       SOTIEN02 = .Value
                       If Trim(SOTIEN02) = vbNullString Then
                            SOTIEN02 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("I")
                       SOTIEN03 = .Value
                       If Trim(SOTIEN03) = vbNullString Then
                            SOTIEN03 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("J")
                       SOTIEN04 = .Value
                       If Trim(SOTIEN04) = vbNullString Then
                            SOTIEN04 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("K")
                       SOTIEN05 = .Value
                       If Trim(SOTIEN05) = vbNullString Then
                            SOTIEN05 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("L")
                       SOTIEN06 = .Value
                       If Trim(SOTIEN06) = vbNullString Then
                            SOTIEN06 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("M")
                       SOTIEN07 = .Value
                       If Trim(SOTIEN07) = vbNullString Then
                            SOTIEN07 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("N")
                       SOTIEN08 = .Value
                       If Trim(SOTIEN08) = vbNullString Then
                            SOTIEN08 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("O")
                       SOTIEN09 = .Value
                       If Trim(SOTIEN09) = vbNullString Then
                            SOTIEN09 = "0"
                       End If
                       
                       SOTIEN092 = "0"
                       
                       .Col = .ColLetterToNumber("P")
                       GHICHU = .Text
                       If Trim(GHICHU) = vbNullString Then
                            GHICHU = "''"
                       Else
                            GHICHU = "'" & TAX_Utilities_Svr_New.Convert(CStr(GHICHU), UNICODE, TCVN) & "'"
                       End If


                       SOTT = STT
                       DANHAN = "''"
                 
                       sSQLVal = MADTNT & "," & NGNOP & "," & TTHTK & "," & QT & "," & SOTT & "," & TENDTNT2 & _
                       "," & MADTNT2 & "," & SOTIEN01 & "," & SOTIEN02 & "," & SOTIEN03 & "," & SOTIEN04 & "," & _
                       SOTIEN05 & "," & SOTIEN06 & "," & SOTIEN07 & "," & SOTIEN08 & "," & SOTIEN09 & "," & SOTIEN092 & _
                       "," & GHICHU & "," & NAMQT & "," & LAN_QUET & "," & DANHAN
                       sSQL = "INSERT INTO TMP_Q41" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                       bln = clsDAO.ExecuteDLL(sSQL)
                
                       STT = STT + 1
                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                  Loop Until .Text = "aa"
                  ROWNUM = .Col
                  
                  
                  
                 .Col = .ColLetterToNumber("B")
                 .Row = ROWNUM + 5
                 Do
                 
                       .Col = .ColLetterToNumber("C")
                       TENDTNT2 = .Text
                       If Trim(TENDTNT2) = vbNullString Then
                            TENDTNT2 = "''"
                       Else
                            TENDTNT2 = "'" & TAX_Utilities_Svr_New.Convert(CStr(TENDTNT2), UNICODE, TCVN) & "'"
                       End If
                        
                       .Col = .ColLetterToNumber("E")
                       MADTNT2 = .Text
                       If Trim(MADTNT2) = vbNullString Then
                            MADTNT2 = "''"
                       Else
                            MADTNT2 = "'" & MADTNT2 & "'"
                       End If
                       
                       .Col = .ColLetterToNumber("G")
                       SOTIEN01 = .Value
                       If Trim(SOTIEN01) = vbNullString Then
                            SOTIEN01 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("H")
                       SOTIEN02 = .Value
                       If Trim(SOTIEN02) = vbNullString Then
                            SOTIEN02 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("I")
                       SOTIEN03 = .Value
                       If Trim(SOTIEN03) = vbNullString Then
                            SOTIEN03 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("J")
                       SOTIEN04 = .Value
                       If Trim(SOTIEN04) = vbNullString Then
                            SOTIEN04 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("K")
                       SOTIEN05 = .Value
                       If Trim(SOTIEN05) = vbNullString Then
                            SOTIEN05 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("L")
                       SOTIEN06 = .Value
                       If Trim(SOTIEN06) = vbNullString Then
                            SOTIEN06 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("M")
                       SOTIEN07 = .Value
                       If Trim(SOTIEN07) = vbNullString Then
                            SOTIEN07 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("N")
                       SOTIEN08 = .Value
                       If Trim(SOTIEN08) = vbNullString Then
                            SOTIEN08 = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("O")
                       SOTIEN09 = .Value
                       If Trim(SOTIEN09) = vbNullString Then
                            SOTIEN09 = "0"
                       End If
                       
                       SOTIEN092 = "0"
                       
                       .Col = .ColLetterToNumber("P")
                       GHICHU = .Text
                       If Trim(GHICHU) = vbNullString Then
                            GHICHU = "''"
                       Else
                            GHICHU = "'" & TAX_Utilities_Svr_New.Convert(CStr(GHICHU), UNICODE, TCVN) & "'"
                       End If


                       SOTT = STT
                        
                 
                       sSQLVal = MADTNT & "," & NGNOP & "," & TTHTK & "," & QT & "," & SOTT & "," & TENDTNT2 & _
                       "," & MADTNT2 & "," & SOTIEN01 & "," & SOTIEN02 & "," & SOTIEN03 & "," & SOTIEN04 & "," & _
                       SOTIEN05 & "," & SOTIEN06 & "," & SOTIEN07 & "," & SOTIEN08 & "," & SOTIEN09 & "," & SOTIEN092 & _
                       "," & GHICHU & "," & NAMQT & "," & LAN_QUET & "," & DANHAN
                       sSQL = "INSERT INTO TMP_Q41" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                       bln = clsDAO.ExecuteDLL(sSQL)
                
                       STT = STT + 1
                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                  Loop Until .Text = "bb"
                  
        
   End With
   
   
   
   clsDAO.Disconnect
   InsertDTL_PL1 = vbNullString
End Function


Public Function InsertDTL() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim bln  As Boolean

  Dim NAMQT As Variant
  Dim MADTNT As Variant
  Dim NGNOP As Variant
  Dim CTTN As Variant
  Dim SOTIEN As Variant
  Dim SOTIEN2 As Variant
  Dim LOAITK As Variant
  Dim SONGUOI As Variant
  Dim SONGUOI2 As Variant
  Dim DANHAN As Variant
   
   
   


   
   sSQLCol = "NAMQT, MADTNT, NGNOP, TTHTK, CTTN, SOTIEN, SONGUOI, SOTIEN2, SONGUOI2, LOAITK, LAN_QUET,DANHAN "
  
   
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("G"), 4, MADTNT
        If Trim(MADTNT) = vbNullString Then
            MADTNT = "''"
        Else
            MADTNT = "'" & MADTNT & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 12, NGNOP
        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
        End If
                    
        NAMQT = TAX_Utilities_Svr_New.Year
        
        
        LOAITK = "'04/TNCN'"
        DANHAN = "''"
        ' Chi tieu
        CTTN = "'002'"
        .GetText .ColLetterToNumber("M"), 20, SONGUOI
        If Trim(SONGUOI) = vbNullString Then
                SONGUOI = "0"
        End If
        SONGUOI2 = SONGUOI
        SOTIEN = "0"
        SOTIEN2 = "0"
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

                              
   End With
        
        '--------------
         
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'003'"
        .GetText .ColLetterToNumber("M"), 21, SONGUOI
        If Trim(SONGUOI) = vbNullString Then
                SONGUOI = "0"
        End If
        SONGUOI2 = SONGUOI
        SOTIEN = "0"
        SOTIEN2 = "0"
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   '--------------------
    With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'004'"
        .GetText .ColLetterToNumber("M"), 22, SONGUOI
        If Trim(SONGUOI) = vbNullString Then
                SONGUOI = "0"
        End If
        SONGUOI2 = SONGUOI
        SOTIEN = "0"
        SOTIEN2 = "0"
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
    
   '----------------
    With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'005'"
        .GetText .ColLetterToNumber("M"), 23, SONGUOI
        If Trim(SONGUOI) = vbNullString Then
                SONGUOI = "0"
        End If
        SONGUOI2 = SONGUOI
        SOTIEN = "0"
        SOTIEN2 = "0"
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   '--------------
    With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'006'"
        .GetText .ColLetterToNumber("M"), 24, SONGUOI
        If Trim(SONGUOI) = vbNullString Then
                SONGUOI = "0"
        End If
        SONGUOI2 = SONGUOI
        SOTIEN = "0"
        SOTIEN2 = "0"
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   '--------------
    With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'007'"
        .GetText .ColLetterToNumber("M"), 25, SONGUOI
        If Trim(SONGUOI) = vbNullString Then
                SONGUOI = "0"
        End If
        SONGUOI2 = SONGUOI
        SOTIEN = "0"
        SOTIEN2 = "0"
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   '--------------
    With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'008'"
        .GetText .ColLetterToNumber("P"), 26, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   '---------------
       With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'009'"
        .GetText .ColLetterToNumber("P"), 27, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   '---------
   
       With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'010'"
        .GetText .ColLetterToNumber("P"), 28, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   '------------
       With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'011'"
        .GetText .ColLetterToNumber("P"), 29, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   '--------------
   
       With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'012'"
        .GetText .ColLetterToNumber("P"), 30, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-----------
       With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'013'"
        .GetText .ColLetterToNumber("P"), 31, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-----------
       With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'014'"
        .GetText .ColLetterToNumber("P"), 32, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '------------
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'015'"
        .GetText .ColLetterToNumber("P"), 33, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-------------
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'016'"
        .GetText .ColLetterToNumber("P"), 34, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-------------
   
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'018'"
        .GetText .ColLetterToNumber("M"), 36, SONGUOI
        If Trim(SONGUOI) = vbNullString Then
                SONGUOI = "0"
        End If
        SONGUOI2 = SONGUOI
        SOTIEN = "0"
        SOTIEN2 = "0"
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '----------
   
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'019'"
        .GetText .ColLetterToNumber("M"), 37, SONGUOI
        If Trim(SONGUOI) = vbNullString Then
                SONGUOI = "0"
        End If
        SONGUOI2 = SONGUOI
        SOTIEN = "0"
        SOTIEN2 = "0"
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-----------------
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'020'"
        .GetText .ColLetterToNumber("M"), 38, SONGUOI
        If Trim(SONGUOI) = vbNullString Then
                SONGUOI = "0"
        End If
        SONGUOI2 = SONGUOI
        SOTIEN = "0"
        SOTIEN2 = "0"
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '---------
   
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'021'"
        .GetText .ColLetterToNumber("P"), 39, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '------------
   
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'022'"
        .GetText .ColLetterToNumber("P"), 40, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '----------
   
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'023'"
        .GetText .ColLetterToNumber("P"), 41, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   
   '-----------
   
    With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'024'"
        .GetText .ColLetterToNumber("P"), 42, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
 
   '-----------
   
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'025'"
        .GetText .ColLetterToNumber("P"), 43, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '----------------
   
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'026'"
        .GetText .ColLetterToNumber("P"), 44, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '------------
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'027'"
        .GetText .ColLetterToNumber("P"), 45, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '--------------
   
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'028'"
        .GetText .ColLetterToNumber("P"), 46, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '------------
   
    With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'029'"
        .GetText .ColLetterToNumber("P"), 47, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '------------
   
    With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'031'"
        .GetText .ColLetterToNumber("P"), 49, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '---------------
   
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'032'"
        .GetText .ColLetterToNumber("P"), 50, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-------------
   
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'033'"
        .GetText .ColLetterToNumber("P"), 51, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-----------
   
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'034'"
        .GetText .ColLetterToNumber("P"), 52, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '---------------
   
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'035'"
        .GetText .ColLetterToNumber("P"), 53, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '---------------
   
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'036'"
        .GetText .ColLetterToNumber("P"), 54, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-----------
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'037'"
        .GetText .ColLetterToNumber("P"), 55, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-------------
   
   With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'038'"
        .GetText .ColLetterToNumber("P"), 56, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '------------
   
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'039'"
        .GetText .ColLetterToNumber("P"), 57, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '------------
   
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'041'"
        .GetText .ColLetterToNumber("P"), 59, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-------------
   
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'042'"
        .GetText .ColLetterToNumber("P"), 60, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   '------------
   
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'043'"
        .GetText .ColLetterToNumber("P"), 61, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-------------
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'044'"
        .GetText .ColLetterToNumber("P"), 62, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '------------
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'045'"
        .GetText .ColLetterToNumber("P"), 63, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-----------
   
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'046'"
        .GetText .ColLetterToNumber("P"), 64, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-----------
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'047'"
        .GetText .ColLetterToNumber("P"), 65, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '--------------
   
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'048'"
        .GetText .ColLetterToNumber("P"), 66, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '------------
   
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'049'"
        .GetText .ColLetterToNumber("P"), 67, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-------------
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'050'"
        .GetText .ColLetterToNumber("P"), 68, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-------------
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'051'"
        .GetText .ColLetterToNumber("P"), 69, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   '-------------
   
      With fps
        .Sheet = 1
        ' Chi tieu 2
        CTTN = "'052'"
        .GetText .ColLetterToNumber("P"), 70, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
                SOTIEN = "0"
        End If
        SONGUOI2 = "0"
        SONGUOI = "0"
        SOTIEN2 = SOTIEN
        
 
        sSQLVal = NAMQT & "," & MADTNT & "," & NGNOP & "," & TTHTK & "," & CTTN & "," & _
        SOTIEN & "," & SONGUOI & "," & SOTIEN2 & "," & SONGUOI2 & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
        sSQL = "INSERT INTO TMP_QT4D" & Right(sKyKeKhai, 4) & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
   End With
   
   clsDAO.Disconnect
   
   If TAX_Utilities_Svr_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL1
   End If
   
   InsertDTL = vbNullString
End Function

Public Function InsertDTL_KHBS() As String
'    Dim sSQL As String
'    Dim sSQLCol As String
'    Dim sSQLVal As String
'    Dim bln  As Boolean
'    Dim MADTNT As Variant
'    Dim NGNOP As Variant
'    Dim KYLBO As Variant
'    Dim KYKKHAI As Variant
'    Dim MATKHAI As Variant
'    Dim MATHUE As Variant
'    Dim MAMUC As Variant
'    Dim MATM As Variant
'    Dim MAPP As Variant
'    Dim MACT As Variant
'    Dim MACT2 As Variant
'    Dim MACT3 As Variant
'    Dim STT As Variant
'    Dim STT2 As Variant
'    Dim STTIN As Variant
'    Dim SOKK As Variant
'    Dim SODC As Variant
'    Dim SOCL As Variant
'    Dim i As Integer, iCol As Long, iRow As Long
'    Dim xmlNode As MSXML.IXMLDOMNode
'
'    Dim strFileName As String
'    Dim fso As New FileSystemObject
'
'   'kiem tra ton tai tep *.dbf chua
'    strFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "BS" & sKyKeKhai & ".DBF"
'    If fso.FileExists(strFileName) = False Then
'    Dim strTepmau As String
'        strTepmau = spathVat & "\tepmau\BS.DBF"
'        fso.CopyFile strTepmau, strFileName, False
'    End If
'
'
'   sSQLCol = "MADTNT, KYKKHAI, MATKHAI, MATHUE, MAMUC, MATM, MAPP, NGNOP, TTHTK, KYLBO , MACT, MACT2, MACT3, STT, STT2, STTIN," & _
'             " SOKK, SODC, SOCL"
'
'   If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'   End If
'
'
'   With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("G"), 4, MADTNT
'        If Trim(MADTNT) = vbNullString Then
'            MADTNT = "''"
'        Else
'            MADTNT = "'" & MADTNT & "'"
'        End If
'
'        KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'        MATKHAI = "'03/TNDN'"
'        MATHUE = "'01'"
'        MAMUC = "'014'"
'        MATM = "'01'"
'        MAPP = "'2'"
'
'        .GetText .ColLetterToNumber("E"), 12, NGNOP
'        'NGNOP = Date
'        If Trim(NGNOP) = vbNullString Then
'            NGNOP = "CTOD('')"
'        Else
'            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
'            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
'        End If
'
'        .GetText .ColLetterToNumber("E"), 10, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "CTOD('')"
'        Else
'            KYLBO = "'" & KYLBO & "'"
'        End If
'
'' Insert so dieu chinh tang
'
'        MACT = "'1000'"
'        MACT2 = "''"
'        MACT3 = "''"
'        STT = 1
'        STT2 = 0
'        STTIN = "'I'"
'        SOKK = "0"
'        SODC = "0"
'        SOCL = "0"
'
'            sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                      NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                      SOKK & "," & SODC & "," & SOCL
'
'
'            sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'            bln = clsDAO.ExecuteDLL(sSQL)
'
'        .Sheet = .SheetCount - 1
'        i = 1
'        .Row = 24
'        Do
'                MACT = "''"
'                .GetText .ColLetterToNumber("E"), .Row, MACT2
'                If Trim(MACT2) = vbNullString Then
'                    MACT2 = "''"
'                ElseIf MACT2 = "20" Then
'                    MACT2 = "'062'"
'                ElseIf MACT2 = "21" Then
'                    MACT2 = "'076'"
'                End If
'
'                MACT3 = "'2'"
'                STT = 1
'                STT2 = i
'                STTIN = "'" & CStr(i) & "'"
'
'                .GetText .ColLetterToNumber("F"), .Row, SOKK
'                If Trim(SOKK) = vbNullString Then
'                    SOKK = "0"
'                Else
'                    SOKK = SOKK
'                End If
'
'                .GetText .ColLetterToNumber("G"), .Row, SODC
'                If Trim(SODC) = vbNullString Then
'                    SODC = "0"
'                Else
'                    SODC = SODC
'                End If
'
'                .GetText .ColLetterToNumber("H"), .Row, SOCL
'                If Trim(SOCL) = vbNullString Then
'                    SOCL = "0"
'                Else
'                    SOCL = SOCL
'                End If
'
'                sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                          NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                          SOKK & "," & SODC & "," & SOCL
'
'                sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'
'                If MACT2 <> "''" Then
'                    bln = clsDAO.ExecuteDLL(sSQL)
'                End If
'
'            .Col = .ColLetterToNumber("B")
'            .Row = 1 + .Row
'            i = i + 1
'        Loop Until .Text = "aa"
'' Insert so dieu chinh giam
'
'        MACT = "'2000'"
'        MACT2 = "''"
'        MACT3 = "''"
'        STT = 2
'        STT2 = 0
'        STTIN = "'II'"
'        SOKK = "0"
'        SODC = "0"
'        SOCL = "0"
'
'            sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                      NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                      SOKK & "," & SODC & "," & SOCL
'
'
'            sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'            bln = clsDAO.ExecuteDLL(sSQL)
'
'        i = 1
'        .Row = .Row + 3
'         Do
'                MACT = "''"
'                .GetText .ColLetterToNumber("E"), .Row, MACT2
'                If Trim(MACT2) = vbNullString Then
'                    MACT2 = "''"
'                ElseIf MACT2 = "20" Then
'                    MACT2 = "'062'"
'                ElseIf MACT2 = "21" Then
'                    MACT2 = "'076'"
'                End If
'
'                MACT3 = "'2'"
'                STT = 2
'                STT2 = i
'                STTIN = "'" & CStr(i) & "'"
'
'                .GetText .ColLetterToNumber("F"), .Row, SOKK
'                If Trim(SOKK) = vbNullString Then
'                    SOKK = "0"
'                Else
'                    SOKK = SOKK
'                End If
'
'                .GetText .ColLetterToNumber("G"), .Row, SODC
'                If Trim(SODC) = vbNullString Then
'                    SODC = "0"
'                Else
'                    SODC = SODC
'                End If
'
'                .GetText .ColLetterToNumber("H"), .Row, SOCL
'                If Trim(SOCL) = vbNullString Then
'                    SOCL = "0"
'                Else
'                    SOCL = SOCL
'                End If
'
'                sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                          NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                          SOKK & "," & SODC & "," & SOCL
'
'                sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'                If MACT2 <> "''" Then
'                    bln = clsDAO.ExecuteDLL(sSQL)
'                End If
'
'            .Col = .ColLetterToNumber("B")
'            .Row = 1 + .Row
'            i = i + 1
'        Loop Until .Text = "bb"
'
'' Insert Tong hop dieu chinh
'
'        MACT = "'3000'"
'        MACT2 = "''"
'        MACT3 = "''"
'        STT = 3
'        STT2 = 0
'        STTIN = "'III'"
'        SOKK = "0"
'        SODC = "0"
'        SOCL = "0"
'
'            sSQLVal = MADTNT & "," & KYKKHAI & "," & MATKHAI & "," & MATHUE & "," & MAMUC & "," & MATM & "," & MAPP & "," & _
'                      NGNOP & "," & TTHTK & "," & KYLBO & "," & MACT & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & _
'                      SOKK & "," & SODC & "," & SOCL
'
'            sSQL = "INSERT INTO BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'            bln = clsDAO.ExecuteDLL(sSQL)
'
'   End With
'
'   clsDAO.Disconnect
   InsertDTL_KHBS = vbNullString
   
End Function



Public Function InsertHDR() As String
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    
    Dim MADTNT As Variant
    Dim NGNOP As Variant
    Dim KYTTTU As Variant
    Dim KYTTDEN As Variant
    Dim NAMQT As Variant
    Dim LOAITK As Variant
    Dim DANHAN As Variant
    
    
    
    sSQLCol = "MADTNT, NGNOP, TTHTK, KYTTTU, KYTTDEN, NAMQT, LOAITK, LAN_QUET,DANHAN "
    
    
       
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("G"), 4, MADTNT
        If Trim(MADTNT) = vbNullString Then
            MADTNT = "''"
        Else
            MADTNT = "'" & MADTNT & "'"
        End If
      
        
        .GetText .ColLetterToNumber("E"), 12, NGNOP
        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
        End If
        
        KYTTTU = "'01/" & TAX_Utilities_Svr_New.Year & "'"
        KYTTDEN = "'12/" & TAX_Utilities_Svr_New.Year & "'"
        NAMQT = TAX_Utilities_Svr_New.Year
        LOAITK = "'04/TNCN'"
        DANHAN = "''"
       
         
         
        
    End With
    
  
    sSQLCol = "MADTNT, NGNOP, TTHTK, KYTTTU, KYTTDEN, NAMQT, LOAITK"
    sSQLVal = MADTNT & "," & NGNOP & "," & TTHTK & "," & KYTTTU & "," & KYTTDEN & "," & NAMQT & "," & LOAITK & "," & LAN_QUET & "," & DANHAN
    
       
    sSQL = "INSERT INTO TMP_QT4" & Right(sKyKeKhai, 4) & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    InsertHDR = sSQL
    
End Function




Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
   
   
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QT4" & Right(sKyKeKhai, 4) & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QT4D" & Right(sKyKeKhai, 4) & ".DBF"
    Dim strFileNamePL1 As String
    strFileNamePL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_Q41" & Right(sKyKeKhai, 4) & ".DBF"
    Dim strFileNameBS As String
    strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & Right(sKyKeKhai, 4) & ".DBF"
    'file HDR
    If fso.FileExists(strFileNameHDR) = False Then
        Dim strTepmauHDR As String
        strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_QT4.DBF"
        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    End If
    'file dtl
    If fso.FileExists(strFileNameDTL) = False Then
        Dim strTepmauDTL As String
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_QT4D.DBF"
        fso.CopyFile strTepmauDTL, strFileNameDTL, False
    End If
    'file phu luc
    If fso.FileExists(strFileNamePL1) = False Then
        Dim strTepmauPL1 As String
        strTepmauPL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_Q41.DBF"
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
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_QT4" & Right(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND NAMQT = " & TAX_Utilities_Svr_New.Year & _
        " AND LOAITK='04/TNCN'"
    
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
   
'   sSQL = "UPDATE TMP_QT4" & Right(sKyKeKhai, 4) & _
'        " SET SODC=1" & _
'        " WHERE MADTNT = " & "'" & strMST & "'" & _
'        " AND NAMQT = " & TAX_Utilities_Svr_New.Year & _
'        " AND LOAITK='04/TNCN'"
'
'   clsDAO.ExecuteDLL sSQL
   
   UpdateTHUETKY2 = True
   
   clsDAO.Disconnect

End Function

Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim KYLBO As Variant
    Dim rs As ADODB.Recordset
    
    If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
  sSQL = "DELETE FROM TMP_QT4" & Right(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND NAMQT = " & TAX_Utilities_Svr_New.Year & " AND LOAITK='04/TNCN'"
        
    
   clsDAO.ExecuteDLL sSQL
   
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
   sSQL = "DELETE FROM TMP_QT4D" & Right(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT = " & TAX_Utilities_Svr_New.Year & _
        " AND LOAITK='04/TNCN'"
    
   clsDAO.ExecuteDLL sSQL
   'Xoa phu luc 1
   
      sSQL = "DELETE FROM TMP_Q41" & Right(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "' AND NAMQT = '" & TAX_Utilities_Svr_New.Year & _
        "'"
    
   clsDAO.ExecuteDLL sSQL
   
   XoaTKTT = True
   
   clsDAO.Disconnect

End Function

Public Function TKRB() As Boolean

'    Dim sSQL As String
'    Dim THUEGTGT_KT As Double
'    Dim rs As ADODB.Recordset
'    Dim sKyTruoc As String
'    Dim sKyKKTruoc As String
'    Dim fso As New FileSystemObject
'    Dim THUEGTGT As Variant
'    Dim dNgayNopDBKT As Variant
'
'    If clsDAO.Connected = False Then
''        Me.MousePointer = vbHourglass
''        frmSystem.MousePointer = vbHourglass
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
''        frmSystem.MousePointer = vbDefault
''        Me.MousePointer = vbDefault
'   End If
'
'   If CInt(Month(Now())) = 1 Then
'        sKyTruoc = "12" & Trim(Str(CInt(Year(Now())) - 1))
'   Else
'        sKyTruoc = IIf(Len(Trim(Str(CInt(Month(Now())) - 1))) = 1, "0" & Trim(Str(CInt(Month(Now())) - 1)), Trim(Str(CInt(Month(Now())) - 1))) & Year(Now())
'   End If
'
'   If CInt(TAX_Utilities_Svr_New.Month) = 1 Then
'        sKyKKTruoc = "12" & TAX_Utilities_Svr_New.Year
'   Else
'        sKyKKTruoc = IIf(Len(Trim(Str(CInt(TAX_Utilities_Svr_New.Month) - 1))) = 1, "0" & Trim(Str(CInt(TAX_Utilities_Svr_New.Month) - 1)), Trim(Str(CInt(TAX_Utilities_Svr_New.Month) - 1))) & TAX_Utilities_Svr_New.Year
'   End If
'
'   If Not fso.FileExists(spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TK" & sKyTruoc & ".dbf") Then
'        TKRB = True
'        clsDAO.Disconnect
'        Exit Function
'   End If
'
'   sSQL = "SELECT MADTNT, NGNOP FROM TK" & sKyTruoc & _
'        " WHERE MATHUE = '01' AND MATM = '01' AND MAPP = '5' AND MADTNT = " & "'" & strMST & "'" & _
'        " AND KYKKHAI = '" & Left(sKyKKTruoc, 2) & "/" & Right(sKyKKTruoc, 4) & "'"
'
'   Set rs = clsDAO.Execute(sSQL)
'   If Not rs Is Nothing Then
'        dNgayNopDBKT = rs.Fields("NGNOP")
'        If Trim(dNgayNopDBKT) = vbNullString Then
'            dNgayNopDBKT = "CTOD('')"
'        Else
'            dNgayNopDBKT = "CTOD('" & Format(dNgayNopDBKT, "mm/dd/yyyy") & "')"
'        End If
'   Else
'        TKRB = True
'        clsDAO.Disconnect
'        Exit Function
'   End If
'   If Not fso.FileExists(spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "KT" & sKyTruoc & ".dbf") Then
'        TKRB = True
'        clsDAO.Disconnect
'        Exit Function
'   End If
'
'   sSQL = "SELECT thuegtgt FROM KT" & sKyTruoc & _
'        " WHERE CT_KT = '085' AND MADTNT = " & "'" & strMST & "' AND NGNOP = " & dNgayNopDBKT
'
'   Set rs = clsDAO.Execute(sSQL)
'   If Not rs Is Nothing Then
'        THUEGTGT_KT = rs.Fields("thuegtgt")
'        If THUEGTGT_KT = 0 Then
'            TKRB = True
'            clsDAO.Disconnect
'            Exit Function
'        Else
'            With fps
'                .Sheet = 1
'                .GetText .ColLetterToNumber("W"), 39, THUEGTGT
'                If Trim(THUEGTGT) = vbNullString Then
'                    THUEGTGT = "0"
'                End If
'                If THUEGTGT_KT = THUEGTGT Then
'                     TKRB = True
'                     clsDAO.Disconnect
'                     Exit Function
'                Else
'                      TKRB = False
'                      clsDAO.Disconnect
'                      Exit Function
'                End If
'            End With
'        End If
'   Else
'        TKRB = True
'        clsDAO.Disconnect
'        Exit Function
'   End If
'
   TKRB = True
'   clsDAO.Disconnect

End Function

Public Function TKTTNGNOP() As Boolean
   Dim NGNOP As Variant
   Dim NGNOPDB As Variant
   Dim sSQL As String
   Dim THUEGTGT_KT As Double
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
   
   If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_QT4" & Right(sKyKeKhai, 4) & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND NAMQT = " & TAX_Utilities_Svr_New.Year & _
        " AND LOAITK='04/TNCN'"
    
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
                If Trim(NGNOPDB) <> vbNullString And Trim(NGNOPDB) = Trim(NGNOP) Then
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
        If Col = .ColLetterToNumber("E") And Row = 12 Then
            .GetText Col, Row, varTemp
            If varTemp <> "" And varTemp <> "../../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                Else
                    .SetActiveCell Col, Row
                End If
            Else
             .SetText Col, Row, ""
            End If
           UpdateCell fps, .Col, .Row, .Text
        End If
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
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_QT4" & Right(sKyKeKhai, 4) & ".DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        End If
        
         sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_QT4" & Right(sKyKeKhai, 4) & _
                " WHERE MADTNT = " & "'" & strMST & "'" & _
                " AND NAMQT = " & TAX_Utilities_Svr_New.Year & _
                " AND LOAITK='04/TNCN'" & " AND (TTHTK = '1' OR TTHTK = '3' OR TTHTK = '4')"
            
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


