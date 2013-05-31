VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_07TNCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
Const MA_BPQL_X = "V"
Const MA_BPQL_Y = 16 'Dong 16 da dc Hidden
Const PHONG_XU_LY_Y = 10
Const PHONG_XU_LY_X = "V"

Private rs As ADODB.Recordset

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
Const TEN_BPQL_X = "V"
Const TEN_BPQL_Y = 12
Const DIEN_THOAI_X = "E"
Const DIEN_THOAI_Y = 8
Const FAX_X = "M"
Const FAX_Y = 8

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
Public TTHTK As String
Public strTenBpql As String


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
        
        
        ' Dat lai check E_36 Duoc gian thue theo TT27/TT-BTC
        ' Thoi gian gian thue tu thang 1 den thang 6 nam 2009 cua to khai thang
        If Val(TAX_Utilities_Svr_New.Year) = 2009 And Val(TAX_Utilities_Svr_New.Month) <= 6 Then
            ' Neu trong hieu luc cua gian thue thi default la checked
            .Row = 36
            .Col = .ColLetterToNumber("E")
            .Text = "1"
            UpdateCell fps, .Col, .Row, "x"
        Else
            ' Neu ngoai hieu luc cua gian thue thi visible cac dieu kien check nay di
            .Row = 36
            .Col = .ColLetterToNumber("B")
            .CellType = CellTypeEdit
            .Lock = True
            .Text = ""
            
            .Row = 36
            .Col = .ColLetterToNumber("E")
            .CellType = CellTypeEdit
            .Lock = True
            .Text = ""
            UpdateCell fps, .Col, .Row, ""
        End If
        
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

Public Function InsertDTL() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   
   Dim MATKHAI As Variant
   Dim MADTNT As Variant
   Dim KYLBO As Variant
   Dim NGNOP As Variant
   Dim KYKKHAI As Variant
   
   Dim CTTN As Variant
   Dim THUETKY As Variant
   Dim THUETKY2 As Variant
   Dim MAMUC As Variant
   Dim MATM As Variant
   Dim TT As Variant
   
   
   Dim bln  As Boolean
   
   If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
    
   sSQLCol = "MATKHAI, MADTNT, KYLBO, KYKKHAI, TTHTK, NGNOP, CTTN, THUETKY, THUETKY2, MAMUC, MATM, TT"
  
   '-------------
   With fps
   
        .Sheet = 1
        MATKHAI = "'07/KK-TNCN'"
        .GetText .ColLetterToNumber("G"), 4, MADTNT
        If Trim(MADTNT) = vbNullString Then
            MADTNT = "''"
        Else
            MADTNT = "'" & MADTNT & "'"
        End If
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
      
        .GetText .ColLetterToNumber("E"), 12, NGNOP
        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
        End If
        KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        
        MAMUC = "'1000'"
        MATM = "'1001'"
        
        THUETKY2 = "0"
        TT = "''"
                
        'Chi tieu "'08'"
        CTTN = "'008'"
        .GetText .ColLetterToNumber("R"), 40, THUETKY
        If Trim(THUETKY) = vbNullString Then
           THUETKY = "0"
        End If
        THUETKY2 = THUETKY
        
        sSQLVal = MATKHAI & "," & MADTNT & "," & KYLBO & "," & KYKKHAI & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & CTTN & "," & THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & ", " & TT
    
        sSQL = "INSERT INTO C7" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Chi tieu 09
        CTTN = "'009'"
        .GetText .ColLetterToNumber("R"), 41, THUETKY
        If Trim(THUETKY) = vbNullString Then
                THUETKY = "0"
        End If
        THUETKY2 = THUETKY
        sSQLVal = MATKHAI & "," & MADTNT & "," & KYLBO & "," & KYKKHAI & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & CTTN & "," & THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & "," & TT
    
        sSQL = "INSERT INTO C7" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Chi tieu 10
        CTTN = "'010'"
        .GetText .ColLetterToNumber("R"), 42, THUETKY
        If Trim(THUETKY) = vbNullString Then
                THUETKY = "0"
        End If
        THUETKY2 = THUETKY
        sSQLVal = MATKHAI & "," & MADTNT & "," & KYLBO & "," & KYKKHAI & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & CTTN & "," & THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & "," & TT
    
        sSQL = "INSERT INTO C7" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

       'Chi tieu 11
        CTTN = "'011'"
        .GetText .ColLetterToNumber("R"), 43, THUETKY
        If Trim(THUETKY) = vbNullString Then
                THUETKY = "0"
        End If
        THUETKY2 = THUETKY
        sSQLVal = MATKHAI & "," & MADTNT & "," & KYLBO & "," & KYKKHAI & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & CTTN & "," & THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & "," & TT
    
        sSQL = "INSERT INTO C7" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Chi tieu 12
        CTTN = "'012'"
        .GetText .ColLetterToNumber("R"), 44, THUETKY
        If Trim(THUETKY) = vbNullString Then
                THUETKY = "0"
        End If
        THUETKY2 = THUETKY
        sSQLVal = MATKHAI & "," & MADTNT & "," & KYLBO & "," & KYKKHAI & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & CTTN & "," & THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & "," & TT
    
        sSQL = "INSERT INTO C7" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        'Chi tieu 13
        CTTN = "'013'"
        .GetText .ColLetterToNumber("R"), 45, THUETKY
        If Trim(THUETKY) = vbNullString Then
                THUETKY = "0"
        End If
        THUETKY2 = THUETKY
        sSQLVal = MATKHAI & "," & MADTNT & "," & KYLBO & "," & KYKKHAI & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & CTTN & "," & THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & "," & TT
    
        sSQL = "INSERT INTO C7" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

        'Chi tieu 14
        CTTN = "'014'"
        .GetText .ColLetterToNumber("R"), 46, THUETKY
        If Trim(THUETKY) = vbNullString Then
                THUETKY = "0"
        End If
        THUETKY2 = THUETKY
        sSQLVal = MATKHAI & "," & MADTNT & "," & KYLBO & "," & KYKKHAI & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & CTTN & "," & THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & "," & TT
    
        sSQL = "INSERT INTO C7" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

        'Chi tieu 15
        CTTN = "'015'"
        .GetText .ColLetterToNumber("R"), 47, THUETKY
        If Trim(THUETKY) = vbNullString Then
                THUETKY = "0"
        End If
        THUETKY2 = THUETKY
        sSQLVal = MATKHAI & "," & MADTNT & "," & KYLBO & "," & KYKKHAI & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & CTTN & "," & THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & "," & TT
    
        sSQL = "INSERT INTO C7" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

        'Chi tieu 16
        CTTN = "'016'"
        .GetText .ColLetterToNumber("R"), 48, THUETKY
        If Trim(THUETKY) = vbNullString Then
                THUETKY = "0"
        End If
        THUETKY2 = THUETKY
        sSQLVal = MATKHAI & "," & MADTNT & "," & KYLBO & "," & KYKKHAI & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & CTTN & "," & THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & "," & TT
    
        sSQL = "INSERT INTO C7" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

        'Chi tieu 17
        CTTN = "'017'"
        .GetText .ColLetterToNumber("R"), 49, THUETKY
        If Trim(THUETKY) = vbNullString Then
                THUETKY = "0"
        End If
        THUETKY2 = THUETKY
        sSQLVal = MATKHAI & "," & MADTNT & "," & KYLBO & "," & KYKKHAI & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & CTTN & "," & THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & "," & TT
    
        sSQL = "INSERT INTO C7" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

        'Chi tieu 18
        CTTN = "'018'"
        .GetText .ColLetterToNumber("R"), 50, THUETKY
        If Trim(THUETKY) = vbNullString Then
                THUETKY = "0"
        End If
        THUETKY2 = THUETKY
        sSQLVal = MATKHAI & "," & MADTNT & "," & KYLBO & "," & KYKKHAI & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & CTTN & "," & THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & "," & TT
    
        sSQL = "INSERT INTO C7" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

        'Chi tieu 19
        CTTN = "'019'"
        .GetText .ColLetterToNumber("R"), 51, THUETKY
        If Trim(THUETKY) = vbNullString Then
                THUETKY = "0"
        End If
        THUETKY2 = THUETKY
        sSQLVal = MATKHAI & "," & MADTNT & "," & KYLBO & "," & KYKKHAI & "," & TTHTK & "," & NGNOP & ","
        sSQLVal = sSQLVal & CTTN & "," & THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & "," & TT
    
        sSQL = "INSERT INTO C7" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

     End With
   
     clsDAO.Disconnect
     
     With fps
        .Sheet = 1
        .Row = 55
        .Col = .ColLetterToNumber("C")
        If Trim(.Text) <> vbNullString Then
            InsertDTL_PL1
        End If
     End With
   
     InsertDTL = vbNullString
End Function

Public Function InsertDTL_PL1() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim bln  As Boolean
   
  Dim MADTNT As Variant
  Dim MATKHAI As Variant
  Dim KYLBO As Variant
  Dim KYKKHAI As Variant
  Dim NGNOP As Variant
  
  Dim TENGOI As Variant
  Dim MADTNT2 As Variant
  Dim NGAYSINH As Variant
  Dim SOCMT As Variant
  Dim SOHC As Variant
  Dim QHEDTNT As Variant
   
   
  Dim strFileName As String
  Dim fso As New FileSystemObject
   'kiem tra ton tai tep *.dbf chua
'    strFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "P3" & sKyKeKhai & ".DBF"
'    If fso.FileExists(strFileName) = False Then
'    Dim strTepmau As String
'        strTepmau = spathVat & "\tepmau\P3.DBF"
'        fso.CopyFile strTepmau, strFileName, False
'    End If
    
   
   sSQLCol = "MADTNT, MATKHAI, KYLBO, KYKKHAI, NGNOP, TENGOI, MADTNT2, NGAYSINH, SOCMT, SOHC, QHEDTNT"
   
   
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
                    
        .GetText .ColLetterToNumber("E"), 10, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
                
        
        KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        MATKHAI = "'07/KK-TNCN'"
        .Sheet = 1
             .Col = .ColLetterToNumber("B")
             .Row = 55
                 Do
                        .Col = .ColLetterToNumber("C")
                        TENGOI = .Text
                        If Trim(TENGOI) = vbNullString Then
                            TENGOI = "''"
                        Else
                            TENGOI = "'" & Trim(TENGOI) & "'"
                        End If
                        
                        .Col = .ColLetterToNumber("H")
                        MADTNT2 = .Text
                        If Trim(MADTNT2) = vbNullString Then
                            MADTNT2 = "''"
                        Else
                            MADTNT2 = "'" & Trim(MADTNT2) & "'"
                        End If
                        
                        .Col = .ColLetterToNumber("L")
                        NGAYSINH = .Text
                        If Trim(NGAYSINH) = vbNullString Then
                            NGAYSINH = "CTOD('')"
                        Else
                            NGAYSINH = ToDate(Trim(NGAYSINH), DDMMYYYY)
                            NGAYSINH = "CTOD('" & Format(NGAYSINH, "mm/dd/yyyy") & "')"
                        End If
                        
                        .Col = .ColLetterToNumber("P")
                        SOCMT = .Text
                        If Trim(SOCMT) = vbNullString Then
                            SOCMT = "''"
                        Else
                            SOCMT = "'" & Trim(SOCMT) & "'"
                        End If
                        
                        .Col = .ColLetterToNumber("U")
                        SOHC = .Text
                        If Trim(SOHC) = vbNullString Then
                            SOHC = "''"
                        Else
                            SOHC = "'" & Trim(SOHC) & "'"
                        End If
                        
                        .Col = .ColLetterToNumber("Y")
                        QHEDTNT = .Text
                        If Trim(QHEDTNT) = vbNullString Then
                            QHEDTNT = "''"
                        Else
                            QHEDTNT = "'" & Trim(QHEDTNT) & "'"
                        End If
                       
                       sSQLVal = MADTNT & "," & MATKHAI & "," & KYLBO & "," & KYKKHAI & "," & NGNOP & "," & TENGOI & "," & MADTNT2 & "," & NGAYSINH & "," & SOCMT & "," & SOHC & "," & QHEDTNT
                       
                       sSQL = "INSERT INTO T7" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                       bln = clsDAO.ExecuteDLL(sSQL)
                       
                       .Col = .ColLetterToNumber("B")
                       .Row = .Row + 1
                  Loop Until .Text = "aa"
        End With
   
   clsDAO.Disconnect
   InsertDTL_PL1 = vbNullString
   
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
    
   Dim MATKHAI As Variant
   Dim KYLBO As Variant
   Dim NGNOP As Variant
   Dim NGNHAP As Variant
   Dim KYKKHAI As Variant
   Dim MADTNT As Variant
   Dim DAGHI As Variant
   Dim LOAITIEN As Variant
   
   Dim MAVACH As Variant
   Dim SHTEP As Variant
   Dim HANNOP2 As Variant
   Dim THUETKY As Variant
   Dim THUETKY2 As Variant
   
   Dim MAMUC As Variant
   Dim MATM As Variant
   Dim MATHUE As Variant
   Dim MAPP As Variant
   
   Dim MATS As Variant
   Dim AUTOCAL As Variant
   
   
   Dim CHKGIAHAN As Variant
    
    
    sSQLCol = "MATKHAI, KYLBO, NGNOP,NGNHAP, KYKKHAI, MADTNT, TTHTK,DAGHI,LOAITIEN, MAVACH,"
    sSQLCol = sSQLCol + "HANNOP, HANNOP2, THUETKY, THUETKY2, MAMUC,MATM, SHTEP, MATS, AUTO_CAL, MATHUE, MAPP "
    
    
       
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("G"), 4, MADTNT
        If Trim(MADTNT) = vbNullString Then
            MADTNT = "''"
        Else
            MADTNT = "'" & MADTNT & "'"
        End If
      
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
        
        .GetText .ColLetterToNumber("E"), 12, NGNOP
        'NGNOP = Date
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
        End If
        
        .GetText .ColLetterToNumber("M"), 12, NGNHAP
        'NGNHAP = Date
        If Trim(NGNHAP) = vbNullString Then
            NGNHAP = "CTOD('')"
        Else
            NGNHAP = ToDate(Trim(NGNHAP), DDMMYYYY)
            NGNHAP = "CTOD('" & Format(NGNHAP, "mm/dd/yyyy") & "')"
        End If
            
        
        KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        
        
'        .GetText .ColLetterToNumber(MA_BPQL_X), MA_BPQL_Y, MABPQL
'
'        If Trim(MABPQL) = vbNullString Then
'            MABPQL = "''"
'        Else
'            MABPQL = "'" & MABPQL & "'"
'        End If
        
        .GetText .ColLetterToNumber("M"), 10, SHTEP
        If Trim(SHTEP) = vbNullString Then
            SHTEP = "''"
        Else
            SHTEP = "'" & SHTEP & "'"
        End If
        
'Check gia han nop
        .GetText .ColLetterToNumber("E"), 36, CHKGIAHAN
        If Trim(CHKGIAHAN) = "1" Then
                HANNOP = "CTOD('" & Format("20/06/2009", "mm/dd/yyyy") & "')"
        Else
            If Trim(HANNOP) = vbNullString Then
                HANNOP = "CTOD('')"
            Else
                HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
            End If
        End If
         
         
         HANNOP2 = HANNOP
         
         MATKHAI = "'07/KK-TNCN'"
         DAGHI = ".F."
         LOAITIEN = "''"
         
         MAVACH = ".T."
         
         MAMUC = "'1000'"
         MATM = "'1001'"

         MATS = "''"
         AUTOCAL = ".T."
         
         MATHUE = "'09'"
         MAPP = "'07'"
         
         .GetText .ColLetterToNumber("R"), 51, THUETKY
         If Trim(THUETKY) = vbNullString Then
                THUETKY = "0"
         End If
         THUETKY2 = "0"
        
        
    End With
    
    
    sSQLVal = MATKHAI & "," & KYLBO & "," & NGNOP & "," & NGNHAP & "," & KYKKHAI & "," & MADTNT & "," & _
    TTHTK & "," & DAGHI & "," & LOAITIEN & "," & MAVACH & "," & HANNOP & "," & HANNOP2 & "," & _
    THUETKY & "," & THUETKY2 & "," & MAMUC & "," & MATM & "," & SHTEP & "," & MATS & "," & AUTOCAL & "," & MATHUE & "," & MAPP
    
       
    sSQL = "INSERT INTO TZ" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    InsertHDR = sSQL
    
End Function

Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim strFileNameIDX As String
    
    'kiem tra ton tai tep *.dbf chua
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TZ" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileNameHDR) = False Then
'    Dim strTepmauHDR As String
'        strTepmauHDR = spathVat & "\tepmau\TZ.DBF"
'        fso.CopyFile strTepmauHDR, strFileNameHDR, False
        isExistFile = False
        DisplayMessage "0109", msOKOnly, miCriticalError
        Exit Function
    Else
        isExistFile = True
    End If
'
'    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "T2" & sKyKeKhai & ".DBF"
'    If fso.FileExists(strFileNameDTL) = False Then
'    Dim strTepmauDTL As String
'        strTepmauDTL = spathVat & "\tepmau\T2.DBF"
'        fso.CopyFile strTepmauDTL, strFileNameDTL, False
'    End If
'
'    'Xoa file *.IDX
    strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TZ" & sKyKeKhai & ".IDX"
    If fso.FileExists(strFileNameIDX) = True Then
        fso.DeleteFile strFileNameIDX, True
    End If
    
        strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "T2" & sKyKeKhai & ".IDX"
    If fso.FileExists(strFileNameIDX) = True Then
        fso.DeleteFile strFileNameIDX, True
    End If
    
        strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "BS" & sKyKeKhai & ".IDX"
    If fso.FileExists(strFileNameIDX) = True Then
        fso.DeleteFile strFileNameIDX, True
    End If
    
     strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "T7" & sKyKeKhai & ".IDX"
    If fso.FileExists(strFileNameIDX) = True Then
        fso.DeleteFile strFileNameIDX, True
    End If
    
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   sSQL = "SELECT MADTNT, NGNOP FROM TZ" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='07/KK-TNCN' "
        
    
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
  
   clsDAO.Disconnect

End Function

Public Function UpdateTHUETKY2() As Boolean
    Dim sSQL As String

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   sSQL = "UPDATE TZ" & sKyKeKhai & _
        " SET THUETKY2 = 1 " & _
        " WHERE MATHUE = '09' and MAPP = '07' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='07/KK-TNCN'"

   clsDAO.ExecuteDLL sSQL

   UpdateTHUETKY2 = True

   clsDAO.Disconnect
End Function

Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL As String
    

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   sSQL = "SELECT * FROM TZ" & sKyKeKhai & _
        " WHERE TTHTK = '1' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='07/KK-TNCN'"
   
   Set rs = clsDAO.Execute(sSQL)
   If rs Is Nothing Then
        isToKhaiChinhThuc = False
   Else
        isToKhaiChinhThuc = True
   End If
   
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
   
   sSQL = "DELETE FROM TZ" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='07/KK-TNCN'"
     
    
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
   sSQL = "DELETE FROM T2" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "' AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' " & _
        " AND MATKHAI='07/KK-TNCN'"
    
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
   
   sSQL = "SELECT MADTNT, NGNOP FROM TZ" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI <> '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
    
    
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


