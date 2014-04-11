VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01_AC_BLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_Y = 19
Const KY_LAP_BO_X = "E"
Const MA_SO_TEP_Y = 10
Const MA_SO_TEP_X = "M"
Const NGAY_NHAN_TO_KHAI_Y = 10
Const NGAY_NHAN_TO_KHAI_X = "E"
Const NGUOI_SU_DUNG_Y = 12
Const NGUOI_SU_DUNG_X = "V"
Const HEADER_KY_LAP_BO_Y = 13
Const HEADER_KY_LAP_BO_X = "B"
Const MA_BPQL_X = "V"
Const MA_BPQL_Y = 18 'Dong 16 da dc Hidden
Const PHONG_XU_LY_Y = 10
Const PHONG_XU_LY_X = "Q"

'Longvh
Const MA_SO_THUE_X = "E"
Const MA_SO_THUE_Y = 5
Const NGAY_NOP_Y = 10
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 12
Const NGAY_QUET_X = "K"

Const TEN_GOI_X = "E"
Const TEN_GOI_Y = 4
Const DIA_CHI_X = "E"
Const DIA_CHI_Y = 6
Const TEN_BPQL_X = "Q"
Const TEN_BPQL_Y = 12
Const DIEN_THOAI_X = "E"
Const DIEN_THOAI_Y = 8
Const FAX_X = "K"
Const FAX_Y = 8
'dntai
Const STT_X = "K"
Const STT_Y = 10
Const PHONG_QL_X = "Q"
Const PHONG_QL_Y = 14

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
Private rs As ADODB.Recordset
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
Private tempKyLb As Variant
Private tempNgayNop As Variant
Private NgayDauKy As Date

Public strMaCQT As String
Public strMaPQL As String
Public strTenPQL As String
Public strSoTTTKhai As String
Public isTKTonTai As Boolean
Public vHDR_ID As String




Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
        
        'dien thoai
        .Col = .ColLetterToNumber("E")
        .Row = 8
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetAlphanumeric
        .TypeMaxEditLen = 20
        
        'Ghi chu
        .Col = .ColLetterToNumber("E")
        .Row = 14
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 1000
        
        'Fax
        .Col = .ColLetterToNumber("K")
        .Row = 8
        .TypeEditCharSet = TypeEditCharSetNumeric
        .TypeMaxEditLen = 20
    
        'STT to khai
        .Col = .ColLetterToNumber("K")
        .Row = 10
        .TypeEditCharSet = TypeEditCharSetNumeric
        .TypeMaxEditLen = 20
        
        'Phong quan ly
        .Col = .ColLetterToNumber("Q")
        .Row = 14
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
        
        'Phong xu ly
        .Col = .ColLetterToNumber("Q")
        .Row = 10
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 80
        
        ' Nguoi su dung
        .Col = .ColLetterToNumber("Q")
        .Row = 12
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
        
        'Email
        .Col = .ColLetterToNumber("Q")
        .Row = 8
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
                
        ' Ngay nop
        SetDateFormat fps, 1, 10, .ColLetterToNumber("E"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 10
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        '.Text = Day(Date) & "/" & Month(Date) & "/" & Year(Date)
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 12, .ColLetterToNumber("K"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 12
        .Col = .ColLetterToNumber("K")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        '.Text = Day(Date) & "/" & Month(Date) & "/" & Year(Date)
        .TypeHAlign = TypeHAlignLeft
        
        'set ky lapbo
        SetDateFormat fps, 1, 19, .ColLetterToNumber("E"), MMYYYY
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
        If rsPXL.Fields.count > 0 Then
            Do While Not rsPXL.EOF
                'dhdang sua combox ngay xu ly
                'ngay 05-08
                .TypeComboBoxIndex = -1
                .TypeComboBoxString = TAX_Utilities_Svr_New.Convert(rsPXL.Fields(1).Value, TCVN, UNICODE) + "                {" + TAX_Utilities_Svr_New.Convert(Trim(rsPXL.Fields(0).Value), TCVN, UNICODE) + "}"
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
        .TypeComboBoxCurSel = 2
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
        
        'set kylapbo
        
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
        '.Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
         UpdateCell fps, .Col, .Row, .Text
        .Lock = False
        
        .Col = .ColLetterToNumber(NGAY_QUET_X)
        .Row = NGAY_QUET_Y
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
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
        
        'set nguoi su dung
        
        .Col = .ColLetterToNumber(TEN_BPQL_X)
        .Row = TEN_BPQL_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strNguoiSuDung, TCVN, UNICODE))
        '.EventEnabled(EventAllEvents) = True
        
        'set phong quan ly
        .Col = .ColLetterToNumber(PHONG_QL_X)
        .Row = PHONG_QL_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
       
        ' set STT to khai
        .Col = .ColLetterToNumber(STT_X)
        .Row = STT_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        
    End With
    'get hdr_id
    vHDR_ID = GetHdrIdAC(spathVat)
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
        
        NgayDauKy = dNgayDauKy
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
    Dim vNgayNop, vNgayQuet As Variant
    
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 12, varCheckValue
        .GetText .ColLetterToNumber("E"), 14, varNoteValue
        
        .GetText .ColLetterToNumber("E"), 10, vNgayNop
        .GetText .ColLetterToNumber("K"), 12, vNgayQuet
        ' Kiem tra ngay nop khong duoc lon hon ngay quet
        If ToDate(CStr(vNgayNop), "dd/mm/yyyy") > ToDate(CStr(vNgayQuet), "dd/mm/yyyy") Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 12, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 12, "1"
        End If
'        If CStr(varCheckValue) = "1" And Trim(CStr(varNoteValue)) = "" Then
'            .Sheet = .SheetCount
'            .SetText 2, 14, "0"
'            CheckValidData = False
'        Else
'            .Sheet = .SheetCount
'            .SetText 2, 14, "1"
'        End If
        
    End With
End Function


Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
   ' Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
   ' KiemTraKhoaSo = True
   ' dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
   ' dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("E_10"), "Value"), "dd/mm/yyyy")
    
   ' If dNgayKhoaSo < dKyLapBo Then
   KiemTraKhoaSo = False
   ' End If
End Function

Public Function InsertDTL() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   
   Dim MST As Variant
   Dim hdr_ID1 As Variant
   Dim Id1 As Variant
   
   Dim tenToChuc As Variant
   Dim diaChi As Variant
   Dim soHopDong As Variant
   Dim ngayHopDong As Variant
   Dim tenLoaiBienLai As Variant
   Dim kyHieuMauBienLai As Variant
   Dim kyHieuBienLai As Variant
   Dim tuSO As Variant
   Dim denSo As Variant
   Dim SOLUONG As Variant
   Dim STT As Variant
   Dim LOAIHD As Variant
   
   Dim bln  As Boolean
   

   If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
    
   sSQLCol = "id, hdr_id, so_tt,mst_datin,ten_datin,dia_chi,so_bl,ngay_bl, ten_bl , mau_so, ky_hieu_bl, tu_so, den_so, so_luong,loai_bl"
  
   '-------------
   With fps
   
   .Sheet = 1
        'set hdrID
        If vHDR_ID = vbNullString Then
            vHDR_ID = "0"
        End If
        .Row = 29
        Do
            Id1 = GetDtlIdAC(spathVat)
            'set sTT
            .GetText .ColLetterToNumber("B"), .Row, STT
            If Trim(STT) = vbNullString Then
                STT = "0"
            End If
            'set MaSoThue
            .GetText .ColLetterToNumber("D"), .Row, MST
            If MST = vbNullString Then
                MST = "''"
            Else
                MST = Trim(CStr(MST))
                If Len(MST) = 13 Then
                    MST = Left(CStr(MST), 10) & "-" & Right(CStr(MST), 3)
                End If
                MST = "'" & MST & "'"
            End If
            'set tentochuc
            .GetText .ColLetterToNumber("E"), .Row, tenToChuc
            If tenToChuc = vbNullString Then
                tenToChuc = "''"
            Else
                tenToChuc = TAX_Utilities_Svr_New.Convert(CStr(tenToChuc), UNICODE, TCVN)
                tenToChuc = "'" & tenToChuc & "'"
            End If
            'set diachi
            .GetText .ColLetterToNumber("I"), .Row, diaChi
            If diaChi = vbNullString Then
                diaChi = "''"
            Else
                diaChi = TAX_Utilities_Svr_New.Convert(CStr(diaChi), UNICODE, TCVN)
                diaChi = "'" & diaChi & "'"
            End If
            'set dohopdong
            .GetText .ColLetterToNumber("K"), .Row, soHopDong
            If soHopDong = vbNullString Then
                soHopDong = "''"
            Else
                soHopDong = "'" & soHopDong & "'"
            End If
            'set ngayhopdong
            .GetText .ColLetterToNumber("M"), .Row, ngayHopDong
            If ngayHopDong = vbNullString Then
                ngayHopDong = "CTOD('')"
            Else
                ngayHopDong = ToDate(Trim(ngayHopDong), DDMMYYYY)
                ngayHopDong = "CTOD('" & Format(ngayHopDong, "mm/dd/yyyy") & "')"
            End If
            'set tenloaiBienLai
            .GetText .ColLetterToNumber("O"), .Row, tenLoaiBienLai
            If tenLoaiBienLai = vbNullString Then
                tenLoaiBienLai = "''"
            Else
                tenLoaiBienLai = TAX_Utilities_Svr_New.Convert(CStr(tenLoaiBienLai), UNICODE, TCVN)
                If Len(tenLoaiBienLai) > 250 Then
                    tenLoaiBienLai = Left(CStr(tenLoaiBienLai), 250)
                End If
                tenLoaiBienLai = "'" & tenLoaiBienLai & "'"
            End If
            'set mauso
            .GetText .ColLetterToNumber("Q"), .Row, kyHieuMauBienLai
            If kyHieuMauBienLai = vbNullString Then
                kyHieuMauBienLai = "''"
            Else
                kyHieuMauBienLai = "'" & kyHieuMauBienLai & "'"
            End If
            'set kyhieu bien lai
            .GetText .ColLetterToNumber("T"), .Row, kyHieuBienLai
            If kyHieuBienLai = vbNullString Then
                kyHieuBienLai = "''"
            Else
                kyHieuBienLai = "'" & kyHieuBienLai & "'"
            End If
            'set tuso
            .GetText .ColLetterToNumber("V"), .Row, tuSO
            If tuSO = vbNullString Then
                tuSO = "'0'"
            Else
                tuSO = "'" & tuSO & "'"
            End If
            'set denso
            .GetText .ColLetterToNumber("X"), .Row, denSo
            If denSo = vbNullString Then
                denSo = "'0'"
            Else
                denSo = "'" & denSo & "'"
            End If
            'set soluong
            .GetText .ColLetterToNumber("Z"), .Row, SOLUONG
            If SOLUONG = vbNullString Then
                SOLUONG = "0"
            End If
            'set loaiBienLai
            .GetText .ColLetterToNumber("AB"), .Row, LOAIHD
            If LOAIHD = vbNullString Then
                LOAIHD = "''"
            Else
                LOAIHD = "'" & LOAIHD & "'"
            End If
            sSQLVal = Id1 & "," & vHDR_ID & "," & STT & "," & MST & "," & tenToChuc & "," & diaChi & "," & soHopDong & "," & ngayHopDong & "," & tenLoaiBienLai & "," & kyHieuMauBienLai & "," & kyHieuBienLai & "," & tuSO & "," & denSo & "," & SOLUONG & "," & LOAIHD
            
            sSQL = "INSERT INTO tmp_dtl_01_AC_BLP( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        
     End With
   
    clsDAO.Disconnect
   
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
'    Dim MACT As Varian
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
    Dim strComboBox As Variant
    Dim combovalue() As String
    
    Dim maST As Variant
    Dim kyBCTuNgay As Variant
    Dim kyBCDenNgay As Variant
    Dim ngayLapBC As Variant
    Dim nguoiDDPL As Variant
    Dim loaiBC As String
    Dim NGAYNOP As Variant
    Dim DANHAN As Variant
    Dim phongXuLy As Variant
    Dim phongQuanLy As Variant
    Dim NGAYQUET As Variant
    Dim nguoiCapNhat As Variant
    Dim kyBCCuoiCung As Variant
    Dim STT As Variant
    Dim maCQT As Variant
    Dim quyBC As Variant
    Dim dDate As Date
    Dim vQUY_BC
    Dim Ghi_chu As Variant
    
    
   Dim bln  As Boolean
   ' end
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
    
  
    sSQLCol = "id,tin,loai_bc,ngay_nop,tu_ngay,den_ngay,ngay_cn,nguoi_cn,da_nhan,phong_xly,phong_qly,ngay_bc,nguoi_dd,loai_bc26,so_tt_tk,ma_cqt,quy_bc,ghi_chu"
    'set hdrID
    If vHDR_ID = vbNullString Then
        vHDR_ID = "0"
    End If
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber(MA_SO_THUE_X), MA_SO_THUE_Y, maST
        If Trim(maST) = vbNullString Then
            maST = "''"
        Else
            maST = Trim(CStr(maST))
            If Len(maST) = 13 Then
                maST = Left(CStr(maST), 10) & "-" & Right(CStr(maST), 3)
            End If
            maST = "'" & maST & "'"
        End If

        .GetText .ColLetterToNumber("E"), 10, NGAYNOP
        If Trim(NGAYNOP) = vbNullString Then
            NGAYNOP = "CTOD('')"
        Else
            NGAYNOP = ToDate(Trim(NGAYNOP), DDMMYYYY)
            NGAYNOP = "CTOD('" & Format(NGAYNOP, "mm/dd/yyyy") & "')"
        End If
        
        ' set ngayquet
        .GetText .ColLetterToNumber("K"), 12, NGAYQUET
'        NGAYQUET = Date
        If Trim(NGAYQUET) = vbNullString Then
            NGAYQUET = "CTOD('')"
        Else
            NGAYQUET = ToDate(Trim(NGAYQUET), DDMMYYYY)
            NGAYQUET = "CTOD('" & Format(NGAYQUET, "mm/dd/yyyy") & "')"
        End If
        'get nguoi cap nhat
'        .GetText .ColLetterToNumber("Q"), 12, nguoi_cn
'        If Trim(nguoi_cn) = vbNullString Then
'            nguoi_cn = "''"
'        Else
'            nguoi_cn = TAX_Utilities_Svr_New.Convert(CStr(nguoi_cn), UNICODE, TCVN)
'            nguoi_cn = "'" & nguoi_cn & "'"
'        End If
        
        'set ngay baocaotungay
        .GetText .ColLetterToNumber("E"), 23, kyBCTuNgay
        If Trim(kyBCTuNgay) = vbNullString Then
            kyBCTuNgay = "CTOD('')"
        Else
            kyBCTuNgay = ToDate(Trim(kyBCTuNgay), DDMMYYYY)
            kyBCTuNgay = "CTOD('" & Format(kyBCTuNgay, "mm/dd/yyyy") & "')"
        End If
        
        'set ngaybaocaodenngay
        .GetText .ColLetterToNumber("F"), 23, kyBCDenNgay
        If Trim(kyBCDenNgay) = vbNullString Then
            kyBCDenNgay = "CTOD('')"
        Else
            kyBCDenNgay = ToDate(Trim(kyBCDenNgay), DDMMYYYY)
            kyBCDenNgay = "CTOD('" & Format(kyBCDenNgay, "mm/dd/yyyy") & "')"
        End If
    
        'set quy bao cao
        
        If Trim(TAX_Utilities_Svr_New.ThreeMonths) <> "" Then
            vQUY_BC = "CTOD('" & Format(NgayDauKy, "mm/dd/yyyy") & "')"
        End If
                  
        'set ngayLapBC
        .GetText .ColLetterToNumber("L"), 23, ngayLapBC
        If Trim(ngayLapBC) = vbNullString Then
            ngayLapBC = "CTOD('')"
        Else
            ngayLapBC = ToDate(Trim(ngayLapBC), DDMMYYYY)
            ngayLapBC = "CTOD('" & Format(ngayLapBC, "mm/dd/yyyy") & "')"
        End If
        
        'set nguoi dai dien phap luat
        .GetText .ColLetterToNumber("J"), 23, nguoiDDPL
        If Trim(nguoiDDPL) = vbNullString Then
            nguoiDDPL = "''"
        Else
            nguoiDDPL = TAX_Utilities_Svr_New.Convert(CStr(nguoiDDPL), UNICODE, TCVN)
            nguoiDDPL = "'" & nguoiDDPL & "'"
        End If
        
        'set loaiBC
        loaiBC = "'01_AC_BLP'"
        
        'set DaNhan
        DANHAN = "''"
        
        'set PhongXuLy
        .Col = .ColLetterToNumber("Q")
        .Row = 10
        
        .GetText .ColLetterToNumber("Q"), 10, strComboBox
        If strComboBox = vbNullString Then
            phongXuLy = "''"
        Else
            combovalue = Split(strComboBox, "{")
            phongXuLy = Mid(combovalue(1), 1, Len(combovalue(1)) - 1)
            phongXuLy = "'" & phongXuLy & "'"
        End If
        
        'set phongQuanLY
        If Trim(strMaBPQL) = vbNullString Then
            phongQuanLy = "''"
        Else
            phongQuanLy = "'" & Trim(strMaBPQL) & "'"
        End If
        'set nguoi capnhat
        .GetText .ColLetterToNumber("Q"), 12, nguoiCapNhat
        If nguoiCapNhat = vbNullString Then
            nguoiCapNhat = "''"
        Else
            nguoiCapNhat = TAX_Utilities_Svr_New.Convert(CStr(nguoiCapNhat), UNICODE, TCVN)
            nguoiCapNhat = "'" & nguoiCapNhat & "'"
        End If
        
        'set ky bao cao cuoi cung
        .GetText .ColLetterToNumber("E"), 16, kyBCCuoiCung
        If kyBCCuoiCung = 1 Then
            kyBCCuoiCung = "'Y'"
            Else
            kyBCCuoiCung = "'N'"
        End If
        
        'set sTT
        STT = strSoTTTKhai
        If STT = vbNullString Then
            STT = "0"
        End If
        'set maCQT
        maCQT = GetMaCqt

        'set ghichu
        
        .GetText .ColLetterToNumber("E"), 14, Ghi_chu
        If Ghi_chu = vbNullString Then
            Ghi_chu = "''"
        Else
            Ghi_chu = TAX_Utilities_Svr_New.Convert(CStr(Ghi_chu), UNICODE, TCVN)
            Ghi_chu = "'" & Ghi_chu & "'"
        End If
'sSQLCol = "id,tin,loai_bc,ngay_nop,tu_ngay,den_ngay,ngay_cn,nguoi_cn,da_nhan,phong_xly,phong_qly,ngay_bc,nguoi_dd,loai_bc26,so_tt_tk,ma_cqt,quy_bc,ghi_chu"

        sSQLVal = vHDR_ID & "," & maST & "," & loaiBC & "," & NGAYNOP & "," & kyBCTuNgay + "," & kyBCDenNgay & "," & NGAYQUET & "," & nguoiCapNhat & "," & DANHAN & "," & phongXuLy & "," & phongQuanLy & "," & ngayLapBC & "," & nguoiDDPL & "," & kyBCCuoiCung & "," & STT & "," & maCQT & "," & vQUY_BC & "," & Ghi_chu
        sSQL = "INSERT INTO tmp_bcao_hdr_ac( " & sSQLCol & " ) VALUES ( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

    End With


    InsertHDR = vbNullString
    clsDAO.Disconnect
End Function

Public Function TKTT() As Boolean
    isExistFile = True
    TKTT = isTKTonTai
End Function

Public Function UpdateTHUETKY2() As Boolean
    Dim sSQL As String

'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'   End If
'
'   sSQL = "UPDATE TMP_TZ" & sKyKeKhai & _
'        " SET THUETKY2 = 1 " & _
'        " WHERE MATHUE = '09' and MAPP = '03' AND MADTNT = " & "'" & strMST & "'" & _
'        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='03T/KK-TNCN'"
'
'   clsDAO.ExecuteDLL sSQL

   UpdateTHUETKY2 = True

   clsDAO.Disconnect


End Function

Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL As String
    

'    If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'   End If
'
'   sSQL = "SELECT * FROM TMP_TZ" & sKyKeKhai & _
'        " WHERE TTHTK = '1' AND MADTNT = " & "'" & strMST & "'" & _
'        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='03T/KK-TNCN'"
'
'   Set rs = clsDAO.Execute(sSQL)
'   If rs Is Nothing Then
'        isToKhaiChinhThuc = False
'   Else
'        isToKhaiChinhThuc = True
'   End If
   isToKhaiChinhThuc = True
   
   clsDAO.Disconnect
End Function

Public Function XoaTKTT() As Boolean

'    Dim sSQL As String
'    Dim KYLBO As Variant
'    Dim rs As ADODB.Recordset
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
'   sSQL = "DELETE FROM TMP_TZ" & sKyKeKhai & _
'        " WHERE MADTNT = " & "'" & strMST & "'" & _
'        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='03T/KK-TNCN'"
'
'
'   clsDAO.ExecuteDLL sSQL
'
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
'   sSQL = "DELETE FROM T2" & sKyKeKhai & _
'        " WHERE MADTNT = " & "'" & strMST & "' AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' " & _
'        " AND MATKHAI='03T/KK-TNCN'"
'
'   clsDAO.ExecuteDLL sSQL
'
   XoaTKTT = True
'
'   clsDAO.Disconnect

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
'   Dim NGNOP As Variant
'   Dim NGNOPDB As Variant
'   Dim sSQL As String
'   Dim THUEGTGT_KT As Double
'   Dim rs As ADODB.Recordset
'   Dim i As Integer
'
'   With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("E"), 10, NGNOP
'        'NGNOP = Date
'        If Trim(NGNOP) = vbNullString Then
'            NGNOP = ""
'        Else
'            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
'            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
'        End If
'   End With
'
'   If clsDAO.Connected = False Then
''        Me.MousePointer = vbHourglass
''        frmSystem.MousePointer = vbHourglass
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
''        frmSystem.MousePointer = vbDefault
''        Me.MousePointer = vbDefault
'   End If
'
'   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TZ" & sKyKeKhai & _
'        " WHERE MADTNT = " & "'" & strMST & "'" & _
'        " AND KYKKHAI <> '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'
'
'   Set rs = clsDAO.Execute(sSQL)
'   If Not rs Is Nothing Then
'        For i = 1 To rs.RecordCount
'            If Not rs.EOF Then
'                NGNOPDB = rs.Fields("NGNOP")
'                If Trim(NGNOPDB) = vbNullString Then
'                    NGNOPDB = ""
'                Else
'                    NGNOPDB = "CTOD('" & Format(NGNOPDB, "mm/dd/yyyy") & "')"
'                End If
'                If Trim(NGNOPDB) <> vbNullString And Trim(NGNOPDB) = Trim(NGNOP) Then
'                    clsDAO.Disconnect
'                    TKTTNGNOP = False
'                    Exit Function
'                Else
'                    TKTTNGNOP = True
'                End If
'            End If
'            rs.MoveNext
'        Next
'   Else
'        clsDAO.Disconnect
    TKTTNGNOP = True
'   End If

End Function
Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
 Dim varTemp As Variant
 With fps
    If .ActiveSheet = 1 Then
        If Col = .ColLetterToNumber("E") And Row = 10 Then
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
'        If Col = .ColLetterToNumber("E") And Row = 10 Then
'            .GetText Col, Row, varTemp
'            If varTemp <> "" And varTemp <> "../...." Then
'                If Format_mmyyyy(CStr(varTemp)) <> "" Then
'                    .SetText Col, Row, Format_mmyyyy(CStr(varTemp))
'                Else
'                    .SetActiveCell Col, Row
'                End If
'            Else
'             .SetActiveCell Col, Row
'            End If
'           UpdateCell fps, .Col, .Row, .Text
'        End If
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
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TZ" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        End If
        
            sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_TZ" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='03T/KK-TNCN' AND (TTHTK = '1' OR TTHTK = '3' OR TTHTK = '4' )"
        
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


Public Function GetMaCqt() As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim maCQT As Variant
    Dim TEMP As String
    Dim clsConn As New TAX_Utilities_Svr_New.clsADO

        If clsConn.Connected = False Then
            clsConn.CreateConnectionString spathVat & "\dtnt\"
            clsConn.Connect
         End If
         sSQL = "SELECT madbhc,tengoi,dchi FROM dtnt2 WHERE MADTNT =  '" & strMST & "'"
        
        Set rs = clsConn.Execute(sSQL)
        If Not rs Is Nothing Then
            maCQT = rs.Fields("madbhc")
            TEMP = "'" & Mid$(Trim(maCQT), 1, 5) & "'"
        Else
            TEMP = "''"
        End If
        clsConn.Disconnect
        GetMaCqt = TEMP
End Function
