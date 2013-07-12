VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_05TNDN"
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
Const MA_BPQL_X = "V"
Const MA_BPQL_Y = 16 'Dong 16 da dc Hidden
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
Const TEN_BPQL_X = "V"
Const TEN_BPQL_Y = 12
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
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
Public isExistFile As Boolean

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
Private LAN_QUET As Variant

Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long
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


Public Function InsertDTL() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim bln  As Boolean
   
   Dim KYLBO As Variant
   Dim MADTNT As Variant
   Dim KYKKHAI  As Variant
   Dim NGNOP As Variant
   Dim STT As Variant
   Dim TENGOI As Variant
   Dim HOPDONG As Variant
   Dim MADTNT2 As Variant
   Dim NGAYTT As Variant
   Dim SOTIENHH As Variant
   Dim TYLE As Variant
   Dim THUEPN As Variant
   Dim GHICHU As Variant
   Dim MATKHAI As String
   Dim DANHAN As Variant
   
   
   
   sSQLCol = "KYLBO, MADTNT, TTHTK, KYKKHAI, NGNOP, STT,MADTNT2 ,TENGOI, HOPDONG, NGAYTT, SOTIENHH, TYLE, THUEPN, GHICHU,Matkhai, LAN_QUET,DANHAN "
      
   'tinh so lan quet
    'vttoan
    'ngay 28/09/10
    '------------------------------------------------
     
     
    
     
     
   
   
   If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
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
             .Col = .ColLetterToNumber("B")
             .Row = 24
                 Do
                 
                      .Col = .ColLetterToNumber("B")
                      STT = Trim(.Text)
                      If Trim(STT) = vbNullString Then
                            STT = "0"
                      End If
'                      If Trim(STT) = vbNullString Then
'                            STT = "''"
'                      Else
'                            STT = "'" & STT & "'"
'                      End If
                        
                     .Col = .ColLetterToNumber("C")
                     TENGOI = Trim(.Text)
                     If Trim(TENGOI) = vbNullString Then
                        TENGOI = "''"
                     Else
                        TENGOI = "'" & TAX_Utilities_Svr_New.Convert(CStr(TENGOI), UNICODE, TCVN) & "'"
                     End If
                     
                     .Col = .ColLetterToNumber("J")
                      MADTNT2 = Trim(.Text)
                      If Trim(MADTNT2) = vbNullString Then
                            MADTNT2 = "''"
                      Else
                            MADTNT2 = "'" & MADTNT2 & "'"
                      End If
                      
                     .Col = .ColLetterToNumber("N")
                     HOPDONG = Trim(.Text)
                     If Trim(HOPDONG) = vbNullString Then
                        HOPDONG = "''"
                     Else
                        HOPDONG = "'" & HOPDONG & "'"
                     End If
                     
                     .Col = .ColLetterToNumber("R")
                      NGAYTT = Trim(.Text)
                      If Trim(NGAYTT) = vbNullString Then
                            NGAYTT = "CTOD('')"
                      Else
                            NGAYTT = ToDate(Trim(NGAYTT), DDMMYYYY)
                            NGAYTT = "CTOD('" & Format(NGAYTT, "mm/dd/yyyy") & "')"
                      End If
                      
                      .Col = .ColLetterToNumber("V")
                      SOTIENHH = .Value
                      If Trim(SOTIENHH) = vbNullString Then
                            SOTIENHH = "0"
                      End If
                      
                      .Col = .ColLetterToNumber("Z")
                      TYLE = .Value
                      If Trim(TYLE) = vbNullString Then
                            TYLE = "0"
                      End If
                      
                      .Col = .ColLetterToNumber("AD")
                      THUEPN = .Value
                      If Trim(THUEPN) = vbNullString Then
                            THUEPN = "0"
                      End If
                        
                      .Col = .ColLetterToNumber("AH")
                      GHICHU = .Text
                      If Trim(GHICHU) = vbNullString Then
                            GHICHU = "''"
                      Else
                            GHICHU = "'" & TAX_Utilities_Svr_New.Convert(CStr(GHICHU), UNICODE, TCVN) & "'"
                      End If
                      
                     
                    
                    MATKHAI = "'05/TNDN'"
                    DANHAN = "''"
                        
                       sSQLVal = KYLBO & "," & MADTNT & "," & TTHTK & "," & KYKKHAI & "," & NGNOP & "," & STT & "," & MADTNT2 & "," & TENGOI & "," & HOPDONG & "," & NGAYTT & "," & SOTIENHH & "," & TYLE & "," & THUEPN & "," & GHICHU & "," & MATKHAI & "," & LAN_QUET & "," & DANHAN
                       sSQLVal = Replace(sSQLVal, Chr(10), "")
                       
                       
                       sSQL = "INSERT INTO TMP_HH" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                                  
                       bln = clsConn.ExecuteDLL(sSQL)

                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                  Loop Until .Text = "aa"
                 
             
        End With
   
   


   clsConn.Disconnect
   
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
    Dim MATHUE As Variant
    Dim MATM As Variant
    Dim MAPP As Variant
    Dim KYLBO As Variant
    Dim NGNOP As Variant
    Dim LOAITIEN As Variant
    Dim MAPBQL As Variant
    Dim SHSTEP As Variant
    Dim KYKKHAI As Variant
    Dim NGNHAP As Variant
    Dim THUETKY As Variant
    Dim GTTINHTHUE As Variant
    Dim DAGHI As Variant
    Dim KTTRUOC As Variant
    Dim DOANHSO As Variant
    Dim MAMUC As Variant
    Dim MAVACH As Variant
    Dim MATKHAI As Variant
    Dim THUETKY2 As Variant
    Dim HANNOP2 As Variant
    Dim TENDA As Variant
    Dim MAPP2 As Variant
    Dim TEMP As Variant
    Dim TEMP2 As Variant
    Dim NGANH As Variant
    Dim DAGHI1 As Variant
    Dim DAGHI2 As Variant
    Dim AUTO_CAL As Variant
    Dim DANHAN As Variant
   
    Dim bln  As Boolean
    
    sSQLCol = "MADTNT, MATHUE, MATM, KYLBO, NGNOP, LOAITIEN, MABPQL, SHTEP, TTHTK, KYKKHAI, NGNHAP, "
    sSQLCol = sSQLCol + " THUETKY, GTTINHTHUE, DAGHI, KTTRUOC, DOANHSO, MAPP, MAMUC, MAVACH, MATKHAI, HANNOP, THUETKY2, HANNOP2, TENDA, MAPP2,NGANH, DAGHI1,DAGHI2,AUTO_CAL, LAN_QUET,DANHAN "
           
    If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
    End If
   
    With fps
        .Sheet = 1
        MATHUE = "'08'"
        MAPP = "'5'"
        LOAITIEN = "''"
        DAGHI = ".T."
        MAVACH = ".T."
        MATKHAI = "'05/TNDN'"
        TENDA = "''"
        HANNOP2 = "CTOD('')"
        DANHAN = "''"
        ' sua MAMUC & MATM tu nam 2009
        If CLng(Right(sKyKeKhai, 4)) < 2009 Then
            MATM = "'99'"
            MAMUC = "'002'"
        Else
            MAMUC = "'1050'"
            MATM = "'1099'"
        End If
        
        
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
            KYLBO = "'" & KYLBO & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 12, NGNOP
        If Trim(NGNOP) = vbNullString Then
            NGNOP = "CTOD('')"
        Else
            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
        End If
        
        .GetText .ColLetterToNumber(MA_BPQL_X), MA_BPQL_Y, MAPBQL
        If Trim(MAPBQL) = vbNullString Then
            MAPBQL = "''"
        Else
            MAPBQL = "'" & MAPBQL & "'"
            
        End If
        
        .GetText .ColLetterToNumber("M"), 10, SHSTEP
        If Trim(SHSTEP) = vbNullString Then
            SHSTEP = "''"
        Else
            SHSTEP = "'" & SHSTEP & "'"
        End If
        
        KYKKHAI = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        
        .GetText .ColLetterToNumber("M"), 12, NGNHAP
        If Trim(NGNHAP) = vbNullString Then
            NGNHAP = "CTOD('')"
        Else
            NGNHAP = ToDate(Trim(NGNHAP), DDMMYYYY)
            NGNHAP = "CTOD('" & Format(NGNHAP, "mm/dd/yyyy") & "')"
        End If
        
         .Col = .ColLetterToNumber("B")
             .Row = 24
                 Do
                        
                      .Col = .ColLetterToNumber("V")
                      TEMP = .Value
                      If Trim(TEMP) = vbNullString Then
                            TEMP = "0"
                      End If
                      .Col = .ColLetterToNumber("AD")
                      TEMP2 = .Value
                      If Trim(TEMP2) = vbNullString Then
                            TEMP2 = "0"
                      End If
                      THUETKY = THUETKY + TEMP
                      GTTINHTHUE = GTTINHTHUE + TEMP2
                      
                      
                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                  Loop Until .Text = "aa"
        
     
        THUETKY = "0"
       ' GTTINHTHUE = "0"
        KTTRUOC = "0"
        DOANHSO = "0"
        THUETKY2 = "0"
        MAPP2 = "''"
        
        If Trim(HANNOP) = vbNullString Then
            HANNOP = "CTOD('')"
        Else
            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
            HANNOP2 = HANNOP
        End If
        
        
        
    End With
'---------------------------
' vttoan sua 04/10/2010
' them truong DAGHI1, DAGHI2, AUTO_CAL, NGANH
    NGANH = TinhMaNganh
    DAGHI1 = ".F."
    DAGHI2 = ".F."
    AUTO_CAL = ".F."
    
    sSQLVal = MADTNT & "," & MATHUE & "," & MATM & "," & KYLBO & "," & NGNOP & "," & _
              LOAITIEN & "," & MAPBQL & "," & SHSTEP & "," & TTHTK & "," & KYKKHAI & "," & NGNHAP & "," & _
              THUETKY & "," & GTTINHTHUE & "," & DAGHI & "," & KTTRUOC & "," & DOANHSO & "," & MAPP & "," & _
              MAMUC & "," & MAVACH & "," & MATKHAI & "," & HANNOP & "," & THUETKY2 & "," & HANNOP2 & "," & TENDA & "," & MAPP2 & ",'" & NGANH & "'," & DAGHI1 & "," & DAGHI2 & "," & AUTO_CAL & "," & LAN_QUET & "," & DANHAN
                
    sSQL = "INSERT INTO TMP_TK" & sKyKeKhai & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
    bln = clsConn.ExecuteDLL(sSQL)
    
    clsConn.Disconnect
    InsertHDR = vbNullString
End Function
Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim KYLBO As Variant
  Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim strFileNameIDX As String
    
'    'kiem tra ton tai tep *.dbf chua
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
'
'    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "HH" & sKyKeKhai & ".DBF"
'    If fso.FileExists(strFileNameDTL) = False Then
'    Dim strTepmauDTL As String
'        strTepmauDTL = spathVat & "\tepmau\HH.DBF"
'        fso.CopyFile strTepmauDTL, strFileNameDTL, False
'    End If
    
    'Xoa file *.IDX
'    strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TK" & sKyKeKhai & ".IDX"
'    If fso.FileExists(strFileNameIDX) = True Then
'        fso.DeleteFile strFileNameIDX, True
'    End If
'
'        strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "HH" & sKyKeKhai & ".IDX"
'    If fso.FileExists(strFileNameIDX) = True Then
'        fso.DeleteFile strFileNameIDX, True
'    End If
'
'        strFileNameIDX = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "BS" & sKyKeKhai & ".IDX"
'    If fso.FileExists(strFileNameIDX) = True Then
'        fso.DeleteFile strFileNameIDX, True
'    End If
    'Ghi vao DB trung gian kiem tra file .dbf
    'vttoan  sua
    'ngay 28/09/2010
    '-----------------------------
      strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
        strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_HH" & sKyKeKhai & ".DBF"
        'file HDR
        If fso.FileExists(strFileNameHDR) = False Then
            Dim strTepmauHDR As String
            strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_TK.DBF"
            fso.CopyFile strTepmauHDR, strFileNameHDR, False
        End If
        'file dtl
        If fso.FileExists(strFileNameDTL) = False Then
            Dim strTepmauDTL As String
            strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_HH.DBF"
            fso.CopyFile strTepmauDTL, strFileNameDTL, False
        End If
        
        isExistFile = True
    
    '----------------------------
    
    With fps
        .Sheet = 1
        
        .GetText .ColLetterToNumber("E"), 10, KYLBO
        If Trim(KYLBO) = vbNullString Then
                  KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
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
        
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyKeKhai & " WHERE MATKHAI = '05/TNDN' AND MADTNT = '" & strMST & "'" & _
   " AND KYLBO= " & KYLBO & " AND MATHUE='08' AND MAPP='5'"
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

Public Function UpdateTHUETKY2() As Boolean
    Dim sSQL As String
    Dim KYLBO As Variant
    
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 10, KYLBO
        If Trim(KYLBO) = vbNullString Then
                  KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
    End With
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   sSQL = "UPDATE TMP_TK" & sKyKeKhai & _
        " SET THUETKY2=1" & _
        " WHERE MATKHAI = '05/TNDN' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYLBO= " & KYLBO & " AND MATHUE='08' AND MAPP='5'"
    
   clsDAO.ExecuteDLL sSQL
   
   UpdateTHUETKY2 = True
   
   clsDAO.Disconnect

End Function


Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim KYLBO As Variant
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
   
    sSQL = "DELETE FROM TMP_TK" & sKyKeKhai & " WHERE MATKHAI = '05/TNDN' AND MADTNT = '" & strMST & "'" & _
   " AND KYLBO= " & KYLBO & " AND MATHUE='08' AND MAPP='5'"
   clsConn.ExecuteDLL sSQL
   
   ' xoa bang dtl
   
  sSQL = "DELETE FROM TMP_HH" & sKyKeKhai & " WHERE MADTNT = '" & strMST & "'" & _
   " AND KYLBO= " & KYLBO
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
   Dim KYLBO As Variant
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
        
        .GetText .ColLetterToNumber("E"), 10, KYLBO
        If Trim(KYLBO) = vbNullString Then
                  KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
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
   
    sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyKeKhai & " WHERE MATKHAI = '05/TNDN' AND MADTNT = '" & strMST & "'" & _
   " AND KYLBO= " & KYLBO & " AND MATHUE='08' AND MAPP='5'"
    
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
    Dim LAN_QUET As Variant
    Dim KYLBO As Variant
    
'-----------------------------
    'kiem tra ton tai tep *.dbf chua
    'Tinh so lan quyet de ghi vao DB trung gian
    'vttoan sua
    'ngay 28/09/2010
'-----------------------------

    With fps
        .Sheet = 1
        
        .GetText .ColLetterToNumber("E"), 10, KYLBO
        If Trim(KYLBO) = vbNullString Then
                  KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
    End With

    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsConn.Connected = False Then
        clsConn.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsConn.Connect
        End If
        
            sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_TK" & sKyKeKhai & _
                    " WHERE MATKHAI = '05/TNDN' AND MADTNT = '" & strMST & "'" & _
                    " AND KYLBO= " & KYLBO & " AND MATHUE='08' AND MAPP='5'" & _
                    " AND (TTHTK = '1' OR TTHTK = '3' OR TTHTK = '4')"
            Set rs = clsConn.Execute(sSQL)
            If Not rs Is Nothing Then
                 LAN_QUET = rs.Fields("LAN_QUET")
                 TinhSoLanQuet = LAN_QUET
            Else
                 TinhSoLanQuet = 0
            End If
            clsConn.Disconnect
    Else
    TinhSoLanQuet = 0
    End If
'------------------------------------------
End Function
'dhdang tinh ma nganh
'ngay 13/1/11
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

