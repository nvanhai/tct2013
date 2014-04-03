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
Const MA_SO_THUE_X = "E"
Const MA_SO_THUE_Y = 4
Const NGAY_NOP_Y = 10
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 12
Const NGAY_QUET_X = "K"
Const TEN_GOI_X = "E"
Const TEN_GOI_Y = 5
Const DIA_CHI_X = "E"
Const DIA_CHI_Y = 6
Const TEN_BPQL_X = "Q"
Const TEN_BPQL_Y = 14
Const MA_BPQL_Y = 18
Const MA_BPQL_X = "I"


Const DIEN_THOAI_X = "E"
Const DIEN_THOAI_Y = 8
Const FAX_X = "K"
Const FAX_Y = 8


Const KY_LAP_BO_Y = 10
Const KY_LAP_BO_X = "I"
Const MA_SO_TEP_Y = 10
Const MA_SO_TEP_X = "T"
Const NGAY_NHAN_TO_KHAI_Y = 12
Const NGAY_NHAN_TO_KHAI_X = "K"
Const NGUOI_SU_DUNG_Y = 12
Const NGUOI_SU_DUNG_X = "Q"
Const HEADER_KY_LAP_BO_Y = 13
Const HEADER_KY_LAP_BO_X = "B"
Const HEADER_SO_TT_TRONG_TEP_X = "AC"
Const HEADER_SO_TT_TRONG_TEP_Y = 16
Const PHONG_XU_LY_Y = 10
Const PHONG_XU_LY_X = "Q"

Const CHECK_COL = "M"
Const CHECK_ROW = 27

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
Public isExistFile As Boolean
Private LAN_QUET As Variant
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

Public vHDR_ID As String
Public vNgayDauKy As Date

Public strMaCQT As String
Public strMaPQL As String
Public strTenPQL As String
Public strSoTTTKhai As String
Public isTKTonTai As Boolean



Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
'        .Col = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X)
'        .Row = HEADER_SO_TT_TRONG_TEP_Y
'        .CellType = CellTypeEdit
'        .TypeEditCharSet = TypeEditCharSetNumeric
'        .TypeMaxEditLen = 10
'
'        'Ma so tep
'        .Col = .ColLetterToNumber(MA_SO_TEP_X)
'        .Row = MA_SO_TEP_Y
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
        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
        
        ' Ky lap bo
'        SetDateFormat fps, 1, 10, .ColLetterToNumber("I"), MMYYYY
'        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 10, .ColLetterToNumber("E"), DDMMYYYY
        .Sheet = 1
        .Row = 10
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 12, .ColLetterToNumber("K"), DDMMYYYY
        .Sheet = 1
        .Row = 12
        .Col = .ColLetterToNumber("K")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
        
    End With
    vHDR_ID = GetHdrIdAC(spathVat)
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
Dim i As Integer, intIndexCombo As Integer
Dim strLTN As Variant, Col7 As Variant, strLTNCu As Variant, strId As Variant
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim varID As Variant, varDVT As Variant, varThueSuat As Variant
Dim iTenHHDV As String, strCheck As String
Dim strThueSuat As String, strDVT As String
Dim strIdQLT As String
Dim iCol As Long, iRow As Long, lCurrSheet As Long
Dim blnLoadThueSuat As Boolean, blnSuaDVT As Boolean
Dim MATM As String
    With fps
        lCurrSheet = .Sheet
        
        .Sheet = 1
'            i = 0
'            blnLoadThueSuat = True
'            strCheck = GetAttribute(TAX_Utilities_Svr_New.Data(.Sheet - 1).nodeFromID( _
'                     GetCellID(fps, .ColLetterToNumber(CHECK_COL), CHECK_ROW)), "Value")
'            Do
'
'                varID = Empty
'               .GetText .ColLetterToNumber("D"), i + 31, varID
'               If varID <> "" Then
'                    .Row = i + 31
'                    varThueSuat = .Value
'                    DataDM varID, strIdQLT, iTenHHDV, strDVT, strThueSuat, blnSuaDVT, MATM
'                    If strIdQLT <> vbNullString Then
'                        .Col = .ColLetterToNumber("F")
'                        .Row = i + 31
'                        .Text = iTenHHDV
'                        .Col = .ColLetterToNumber("C")
'                        .Text = MATM
'                       ' UpdateCell fps, .ColLetterToNumber("F"), i + 31, .Text
'                    ElseIf CStr(varID) <> vbNullString Then
'                        DisplayMessage "0081", msOKOnly, miCriticalError
'                        Exit Function
'                    End If
'               End If
'                    .RowHeight(i + 31) = .MaxTextRowHeight(i + 31)
'                    'UpdateCell fps, .Col, .Row, .Text
'                    i = i + 1
'                    .Col = .ColLetterToNumber("B")
'                    .Row = i + 31
'                    If Trim(.Text) = "III" Then blnLoadThueSuat = False
'
'            Loop Until .Row >= .MaxRows
'        .Sheet = lCurrSheet
'    End With
'
'    i = 0
'    ReDim Preserve larrPhongXuLy(0)
'    larrPhongXuLy(0) = "00"
'    ReDim Preserve larrid(0)
'    larrid(0) = "00"

        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y
        PHONG_XU_LY_X1 = PHONG_XU_LY_X
        PHONG_XU_LY_Y1 = PHONG_XU_LY_Y
        If rsPXL.Fields.count > 0 Then
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



Public Function TKRB() As Boolean

   TKRB = True

End Function

Public Function TKTT() As Boolean
    isExistFile = True
    TKTT = isTKTonTai
End Function

Public Function UpdateTHUETKY2() As Boolean
    Dim sSQL As String
    Dim MATKHAI As Variant
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
'   With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("B"), 3, MATKHAI
'        If Trim(MATKHAI) = vbNullString Then
'                MATKHAI = "''"
'        Else
'                MATKHAI = "'" & TAX_Utilities_Svr_New.Convert(CStr(MATKHAI), UNICODE, TCVN) & "'"
'        End If
'   End With
'
'
'   sSQL = "UPDATE TMP_TK" & sKyKeKhai & _
'        " SET THUETKY2=1" & _
'        " WHERE MATKHAI = " & MATKHAI & " AND  MAPP='1' AND MADTNT= " & "'" & strMST & "'" & _
'        " AND KYKKHAI ='" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'
'   clsDAO.ExecuteDLL sSQL
   
   UpdateTHUETKY2 = True
   
   clsDAO.Disconnect

End Function


Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim KYLBO As Variant
    Dim MATKHAI As Variant
    Dim rs As ADODB.Recordset
    Dim idTKHDR As String 'Tam thoi ID cua to khai nay la MST & KYKK
    If clsDAO.Connected = False Then
'        Me.MousePointer = vbHourglass
'        frmSystem.MousePointer = vbHourglass
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
'        frmSystem.MousePointer = vbDefault
'        Me.MousePointer = vbDefault
   End If
  ' idTKHDR = strMST & sKyKeKhai
    
'    With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("I"), 10, KYLBO
'        If Trim(KYLBO) = vbNullString Then
'            KYLBO = "''"
'        Else
'            If Len(Trim(KYLBO)) = 6 Then
'                KYLBO = "'0" & KYLBO & "'"
'            Else
'                KYLBO = "'" & KYLBO & "'"
'            End If
'        End If
'
'        .GetText .ColLetterToNumber("B"), 3, MATKHAI
'        If Trim(MATKHAI) = vbNullString Then
'                MATKHAI = "''"
'        Else
'                MATKHAI = "'" & TAX_Utilities_Svr_New.Convert(CStr(MATKHAI), UNICODE, TCVN) & "'"
'        End If
'   End With
'
'
'   sSQL = "DELETE FROM TMP_TK" & sKyKeKhai & _
'   " WHERE MATKHAI = " & MATKHAI & " AND  MAPP='1' AND MADTNT= " & "'" & strMST & "'" & _
'   " AND KYKKHAI ='" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
'
'   clsDAO.ExecuteDLL sSQL
'
'   sSQL = vbNullString
'   sSQL = "DELETE FROM TMP_TB" & sKyKeKhai & " WHERE MATKHAI = " & MATKHAI & " AND MADTNT = '" & strMST & "'" & _
'   " AND KYLBO=" & KYLBO
'
'
'   clsDAO.ExecuteDLL sSQL
'   ' Xoa phu luc 1
'
'    sSQL = "DELETE FROM TMP_D1" & sKyKeKhai & " WHERE MADTNT = '" & strMST & "'" & _
'   " AND KYLBO=" & KYLBO
'   clsDAO.ExecuteDLL sSQL
   XoaTKTT = True
   
   clsDAO.Disconnect

End Function



Public Function TKTTNGNOP() As Boolean
'   Dim NGNOP As Variant
'   Dim NGNOPDB As Variant
'   Dim sSQL As String
'   Dim rs As ADODB.Recordset
'   Dim i As Integer
'
'   With fps
'        .Sheet = 1
'        .GetText .ColLetterToNumber("E"), 12, NGNOP
'        'NGNOP = Date
'        If Trim(NGNOP) = vbNullString Then
'            NGNOP = ""
'        Else
'            NGNOP = ToDate(Trim(NGNOP), DDMMYYYY)
'            NGNOP = "CTOD('" & Format(NGNOP, "mm/dd/yyyy") & "')"
'        End If
'   End With
'
'   If clsConn.Connected = False Then
''        Me.MousePointer = vbHourglass
''        frmSystem.MousePointer = vbHourglass
'        clsConn.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsConn.Connect
''        frmSystem.MousePointer = vbDefault
''        Me.MousePointer = vbDefault
'   End If
'
'   'sSQL = "SELECT MST, NGAYNOP FROM TKHDR WHERE LOAITK = '01BTNDN' AND ID = '" & strMST & sKyKeKhai & "'"
'
'    sSQL = "SELECT MADTNT, NGNOP FROM TKDN" & Right(sKyKeKhai, 4) & " WHERE MATKHAI = '01B/TNDN' AND MADTNT = '" & strMST & "'"
'
'   Set rs = clsConn.Execute(sSQL)
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
'                    clsConn.Disconnect
'                    TKTTNGNOP = False
'                    Exit Function
'                Else
'                    TKTTNGNOP = True
'                End If
'            End If
'            rs.MoveNext
'        Next
'   Else
'        clsConn.Disconnect
 TKTTNGNOP = True
'   End If
'
'
End Function
Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String) 'ByRef strPMaSoTep, ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub



Public Function InsertHDR() As String
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim bln  As Boolean
    
    Dim Id As Variant
    Dim Tin As Variant
    Dim Loai_bc  As Variant
    Dim Ngay_nop As Variant
    Dim TU_NGAY As Variant
    Dim DEN_NGAY As Variant
    Dim Ngay_cn As Variant
    Dim nguoi_cn As Variant
    Dim So_tt_tk As Variant
    Dim DA_NHAN As Variant
    Dim PHONG_XL As Variant
    Dim PHONG_QL As Variant
    Dim CO_BANG_KE, CO_BANG_KE1, CO_BANG_KE2 As Variant
    Dim Ten_dv_cq As Variant
    Dim Tin_dv_cq As Variant
    Dim Ngay_bc As Variant
    Dim Nguoi_dd As Variant
    Dim Cq_nhan As Variant
    Dim LY_DO_MAT As Variant
    Dim Ngay_mat_h As Variant
    Dim Pp_huy As Variant
    Dim DUNG_DN_CQ As Variant
    Dim Ghi_chu As Variant
    Dim MA_CQT As Variant
    Dim LOAI_BC26 As Variant
    Dim Nguoi_lb As Variant
    Dim QUY_BC As Variant
    
    Dim vPHONG_XU_LY As Variant
    Dim arrPXL()   As String
    
    Dim vKyBCCuoi, vChuyenDiaDiem As Variant
      
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
    
    
    sSQLCol = "ID, TIN, LOAI_BC, NGAY_NOP, TU_NGAY, DEN_NGAY, "
    sSQLCol = sSQLCol & "NGAY_CN, NGUOI_CN, SO_TT_TK, DA_NHAN, PHONG_XLY, "
    sSQLCol = sSQLCol & "PHONG_QLY, CO_BANG_KE, TEN_DV_CQ, TIN_DV_CQ, "
    sSQLCol = sSQLCol & "NGAY_BC, NGUOI_DD, CQ_NHAN, LY_DO_MAT, NGAY_MAT_H,PP_HUY,DUNG_DN_CQ, GHI_CHU,MA_CQT,LOAI_BC26,NGUOI_LB,QUY_BC"
    Id = vHDR_ID
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        ' Thong tin ma so thue
        .GetText .ColLetterToNumber("E"), 4, Tin
        If Trim(Tin) = vbNullString Then
            Tin = "''"
        Else
            Tin = Trim(CStr(Tin))
            If Len(Tin) = 13 Then
                Tin = Left(CStr(Tin), 10) & "-" & Right(CStr(Tin), 3)
            End If
            Tin = "'" & Trim(Tin) & "'"
        End If
        ' Ma loai bao cao
        Loai_bc = "'BC26_AC_BLP'"
        ' Thong tin ngay nop bao cao an chi
        .GetText .ColLetterToNumber("E"), 10, Ngay_nop
        If Trim(Ngay_nop) = vbNullString Then
            Ngay_nop = "CTOD('')"
        Else
            Ngay_nop = ToDate(Trim(Ngay_nop), DDMMYYYY)
            Ngay_nop = "CTOD('" & Format(Ngay_nop, "mm/dd/yyyy") & "')"
        End If
        ' Ky bao cao tu ngay
        .GetText .ColLetterToNumber("E"), 23, TU_NGAY
        If Trim(TU_NGAY) = vbNullString Then
            TU_NGAY = "CTOD('')"
        Else
            TU_NGAY = ToDate(Trim(TU_NGAY), DDMMYYYY)
            TU_NGAY = "CTOD('" & Format(TU_NGAY, "mm/dd/yyyy") & "')"
        End If
        ' Ky bao cao den ngay
        .GetText .ColLetterToNumber("F"), 23, DEN_NGAY
        If Trim(DEN_NGAY) = vbNullString Then
            DEN_NGAY = "CTOD('')"
        Else
            DEN_NGAY = ToDate(Trim(DEN_NGAY), DDMMYYYY)
            DEN_NGAY = "CTOD('" & Format(DEN_NGAY, "mm/dd/yyyy") & "')"
        End If
        ' Thong tin ngay cap nhat
        .GetText .ColLetterToNumber("K"), 12, Ngay_cn
        If Trim(Ngay_cn) = vbNullString Then
            Ngay_cn = "CTOD('')"
        Else
            Ngay_cn = ToDate(Trim(Ngay_cn), DDMMYYYY)
            Ngay_cn = "CTOD('" & Format(Ngay_cn, "mm/dd/yyyy") & "')"
        End If
        ' Thong tin nguoi cap nhat
        .GetText .ColLetterToNumber("Q"), 12, nguoi_cn
        If Trim(nguoi_cn) = vbNullString Then
            nguoi_cn = "''"
        Else
            nguoi_cn = TAX_Utilities_Svr_New.Convert(CStr(nguoi_cn), UNICODE, TCVN)
            nguoi_cn = "'" & Trim(nguoi_cn) & "'"
        End If
        ' So thu tu quet to khai
        .GetText .ColLetterToNumber("K"), 10, So_tt_tk
        If Trim(So_tt_tk) = vbNullString Then
            So_tt_tk = 0
        End If
        ' trang thai chuyen sang an chi
        DA_NHAN = "''"
        ' Thong tin ma phong xu ly
        .GetText .ColLetterToNumber("Q"), 10, vPHONG_XU_LY
        arrPXL = Split(vPHONG_XU_LY, "{")
        PHONG_XL = arrPXL(1)
        PHONG_XL = Left$(PHONG_XL, Len(PHONG_XL) - 1)
        If Trim(PHONG_XL) = vbNullString Then
            PHONG_XL = "''"
        Else
            PHONG_XL = "'" & Trim(PHONG_XL) & "'"
        End If
        
        ' Thong tin phong quan ly
        If Trim(strMaBPQL) = vbNullString Then
            PHONG_QL = "''"
        Else
            PHONG_QL = "'" & Trim(strMaBPQL) & "'"
        End If
        
        ' Ten don vi chu quan
        Ten_dv_cq = "''"
        Tin_dv_cq = "''"
        ' Thong tin ngay bao cao
        .GetText .ColLetterToNumber("K"), 19, Ngay_bc
        If Trim(Ngay_bc) = vbNullString Then
            Ngay_bc = "CTOD('')"
        Else
            Ngay_bc = ToDate(Trim(Ngay_bc), DDMMYYYY)
            Ngay_bc = "CTOD('" & Format(Ngay_bc, "mm/dd/yyyy") & "')"
        End If
        ' thong tin nguoi dai dien
        .GetText .ColLetterToNumber("J"), 19, Nguoi_dd
        If Trim(Nguoi_dd) = vbNullString Then
            Nguoi_dd = "''"
        Else
            Nguoi_dd = TAX_Utilities_Svr_New.Convert(CStr(Nguoi_dd), UNICODE, TCVN)
            Nguoi_dd = "'" & Trim(Nguoi_dd) & "'"
        End If
       
        LY_DO_MAT = "''"
        Ngay_mat_h = "CTOD('')"
        Pp_huy = "''"
        DUNG_DN_CQ = "''"
        ' thong tin ghi chu
        .GetText .ColLetterToNumber("E"), 14, Ghi_chu
        If Trim(Ghi_chu) = vbNullString Then
            Ghi_chu = "''"
        Else
            Ghi_chu = TAX_Utilities_Svr_New.Convert(CStr(Ghi_chu), UNICODE, TCVN)
            Ghi_chu = "'" & Trim(Ghi_chu) & "'"
        End If
        ' thong tin ma co quan thue
        MA_CQT = GetMaCqt
        
        ' Loai bao cao BC26
        .GetText .ColLetterToNumber("E"), 16, vKyBCCuoi
        .GetText .ColLetterToNumber("K"), 16, vChuyenDiaDiem
        If Trim(CStr(vKyBCCuoi)) = "0" And Trim(CStr(vChuyenDiaDiem)) = "0" Then
        ' bao cao quy
            LOAI_BC26 = "'01'"
        ElseIf Trim(CStr(vKyBCCuoi)) = "0" And Trim(CStr(vChuyenDiaDiem)) = "1" Then
            ' bao cao chuyen dia diem
            LOAI_BC26 = "'03'"
        ElseIf Trim(CStr(vKyBCCuoi)) = "1" And Trim(CStr(vChuyenDiaDiem)) = "0" Then
            ' bao cao quyet toan cuoi tung
            LOAI_BC26 = "'02'"
        Else
            '
            LOAI_BC26 = "'04'"
        End If
        ' thong tin nguoi lap bieu
        .GetText .ColLetterToNumber("I"), 19, Nguoi_lb
        If Trim(Nguoi_lb) = vbNullString Then
            Nguoi_lb = "''"
        Else
            Nguoi_lb = TAX_Utilities_Svr_New.Convert(CStr(Nguoi_lb), UNICODE, TCVN)
            Nguoi_lb = "'" & Trim(Nguoi_lb) & "'"
        End If
        ' Ky bao cao
        If Trim(TAX_Utilities_Svr_New.ThreeMonths) <> "" Then
            QUY_BC = "CTOD('" & Format(vNgayDauKy, "mm/dd/yyyy") & "')"
         End If
        
        Cq_nhan = "''"
        CO_BANG_KE = "0"
        
        .EventEnabled(EventAllEvents) = True
    End With
    
    sSQLVal = Id & "," & Tin & "," & Loai_bc & "," & Ngay_nop & "," & TU_NGAY & "," & DEN_NGAY & ","
    sSQLVal = sSQLVal & Ngay_cn & "," & nguoi_cn & "," & So_tt_tk & "," & DA_NHAN & "," & PHONG_XL
    sSQLVal = sSQLVal & "," & PHONG_QL & "," & CO_BANG_KE & "," & Ten_dv_cq & "," & Tin_dv_cq & ","
    sSQLVal = sSQLVal & Ngay_bc & "," & Nguoi_dd & "," & Cq_nhan & "," & LY_DO_MAT & "," & Ngay_mat_h & ","
    sSQLVal = sSQLVal & Pp_huy & "," & DUNG_DN_CQ & "," & Ghi_chu & "," & MA_CQT & "," & LOAI_BC26 & "," & Nguoi_lb & "," & QUY_BC

    sSQL = "INSERT INTO tmp_bcao_hdr_ac" & _
     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

     bln = clsDAO.ExecuteDLL(sSQL)
' '----------------
 
   clsDAO.Disconnect
   
   InsertHDR = vbNullString
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

Public Function InsertDTL() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim bln  As Boolean
   
 Dim Id As Variant
 Dim Hdr_id As Variant
 Dim So_tt As Variant
 Dim Ten_bl As Variant
 Dim Mau_so As Variant
 Dim Ky_hieu_bl As Variant
 Dim TONG_SO As Variant
 Dim TU_SO_TON As Variant
 Dim DEN_SO_TON As Variant
 Dim TU_SO_PS As Variant
 Dim DEN_SO_PS As Variant
 Dim TSO_SD_MAT As Variant
 Dim DSO_SD_MAT As Variant
 Dim SUM_SD_MAT As Variant
 Dim SL_DA_SD As Variant
 Dim SL_XOA As Variant
 Dim SO_XOA As Variant
 Dim SL_MAT As Variant
 Dim SO_MAT As Variant
 Dim SL_HUY As Variant
 Dim SO_HUY As Variant
 Dim TSO_TON_CK As Variant
 Dim DSO_TON_CK As Variant
 Dim SL_TON_CK As Variant
 Dim Loai_bl As Variant
   
    
 
   
   sSQLCol = "ID, HDR_ID, SO_TT, TEN_bl, MAU_SO, KY_HIEU_bl, TONG_SO, TU_SO_TON, DEN_SO_TON, TU_SO_PS, DEN_SO_PS, TSO_SD_MAT, DSO_SD_MAT, SUM_SD_MAT, "
   sSQLCol = sSQLCol & "SL_DA_SD, SL_XOA, SO_XOA, SL_MAT, SO_MAT, SL_HUY, SO_HUY, TSO_TON_CK, DSO_TON_CK, SL_TON_CK,LOAI_bl"
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If


 With fps
        .Sheet = 1
        .Col = .ColLetterToNumber("B")
        .Row = 32
        Do
            Id = GetDtlIdAC(spathVat)
            Hdr_id = vHDR_ID
            .GetText .ColLetterToNumber("B"), .Row, So_tt
            ' ten hoa don
            .GetText .ColLetterToNumber("D"), .Row, Ten_bl
            If Trim(Ten_bl) = vbNullString Then
                Ten_bl = "''"
            Else
                Ten_bl = TAX_Utilities_Svr_New.Convert(CStr(Ten_bl), UNICODE, TCVN)
                Ten_bl = "'" & Trim(Ten_bl) & "'"
            End If
            ' mau so hoa don
            .GetText .ColLetterToNumber("J"), .Row, Mau_so
            If Trim(Mau_so) = vbNullString Then
                Mau_so = "''"
            Else
                Mau_so = "'" & Trim(Mau_so) & "'"
            End If
            ' ky hieu hoa don
            .GetText .ColLetterToNumber("K"), .Row, Ky_hieu_bl
            If Trim(Ky_hieu_bl) = vbNullString Then
                Ky_hieu_bl = "''"
            Else
                Ky_hieu_bl = "'" & Trim(Ky_hieu_bl) & "'"
            End If
            ' tong so hoa don
            .GetText .ColLetterToNumber("L"), .Row, TONG_SO
            If Trim(TONG_SO) = vbNullString Then
                TONG_SO = "0"
            End If
            ' tu so ton dau ky
            .GetText .ColLetterToNumber("M"), .Row, TU_SO_TON
            If Trim(TU_SO_TON) = vbNullString Then
                TU_SO_TON = "''"
            Else
                TU_SO_TON = "'" & Trim(TU_SO_TON) & "'"
            End If
            ' den so ton dau ky
            .GetText .ColLetterToNumber("N"), .Row, DEN_SO_TON
            If Trim(DEN_SO_TON) = vbNullString Then
                DEN_SO_TON = "''"
            Else
                DEN_SO_TON = "'" & Trim(DEN_SO_TON) & "'"
            End If
            ' tu so phat sinh
            .GetText .ColLetterToNumber("O"), .Row, TU_SO_PS
            If Trim(TU_SO_PS) = vbNullString Then
                TU_SO_PS = "''"
            Else
                TU_SO_PS = "'" & Trim(TU_SO_PS) & "'"
            End If
            ' den so phat sinh
            .GetText .ColLetterToNumber("P"), .Row, DEN_SO_PS
            If Trim(DEN_SO_PS) = vbNullString Then
                DEN_SO_PS = "''"
            Else
                DEN_SO_PS = "'" & Trim(DEN_SO_PS) & "'"
            End If
            ' tu so su dung, mat, xoa , huy
            .GetText .ColLetterToNumber("Q"), .Row, TSO_SD_MAT
            If Trim(TSO_SD_MAT) = vbNullString Then
                TSO_SD_MAT = "''"
            Else
                TSO_SD_MAT = "'" & Trim(TSO_SD_MAT) & "'"
            End If
            ' den so su dung, mat, xoa, huy
            .GetText .ColLetterToNumber("R"), .Row, DSO_SD_MAT
            If Trim(DSO_SD_MAT) = vbNullString Then
                DSO_SD_MAT = "''"
            Else
                DSO_SD_MAT = "'" & Trim(DSO_SD_MAT) & "'"
            End If
            'so luong su dung, mat, xoa , huy
            .GetText .ColLetterToNumber("S"), .Row, SUM_SD_MAT
            If Trim(SUM_SD_MAT) = vbNullString Then
                SUM_SD_MAT = "0"
            End If
            ' So luong da su dung
            .GetText .ColLetterToNumber("U"), .Row, SL_DA_SD
            If Trim(SL_DA_SD) = vbNullString Then
                SL_DA_SD = "0"
            End If
            ' so luong xoa
            .GetText .ColLetterToNumber("V"), .Row, SL_XOA
            If Trim(SL_XOA) = vbNullString Then
                SL_XOA = "0"
            End If
            ' so xoa
            .GetText .ColLetterToNumber("W"), .Row, SO_XOA
            If Trim(SO_XOA) = vbNullString Then
                SO_XOA = "''"
            Else
                SO_XOA = "'" & Trim(SO_XOA) & "'"
            End If
            ' so luong mat
            .GetText .ColLetterToNumber("X"), .Row, SL_MAT
            If Trim(SL_MAT) = vbNullString Then
                SL_MAT = "0"
            End If
            ' So mat
            .GetText .ColLetterToNumber("Y"), .Row, SO_MAT
            If Trim(SO_MAT) = vbNullString Then
                SO_MAT = "''"
            Else
                SO_MAT = "'" & Trim(SO_MAT) & "'"
            End If
            ' so luong huy
            .GetText .ColLetterToNumber("Z"), .Row, SL_HUY
            If Trim(SL_HUY) = vbNullString Then
                SL_HUY = "0"
            End If
            ' so huy
            .GetText .ColLetterToNumber("AA"), .Row, SO_HUY
            If Trim(SO_HUY) = vbNullString Then
                SO_HUY = "''"
            Else
                SO_HUY = "'" & Trim(SO_HUY) & "'"
            End If
            ' tu so ton cuoi ky
            .GetText .ColLetterToNumber("AC"), .Row, TSO_TON_CK
            If Trim(TSO_TON_CK) = vbNullString Then
                TSO_TON_CK = "''"
            Else
                TSO_TON_CK = "'" & Trim(TSO_TON_CK) & "'"
            End If
            ' den so ton cuoi ky
            .GetText .ColLetterToNumber("AD"), .Row, DSO_TON_CK
            If Trim(DSO_TON_CK) = vbNullString Then
                DSO_TON_CK = "''"
            Else
                DSO_TON_CK = "'" & Trim(DSO_TON_CK) & "'"
            End If
            ' so luong ton cuoi ky
            .GetText .ColLetterToNumber("AE"), .Row, SL_TON_CK
            If Trim(SL_TON_CK) = vbNullString Then
                SL_TON_CK = "0"
            End If
            ' loai hoa don
            .GetText .ColLetterToNumber("AF"), .Row, Loai_bl
            If Trim(Loai_bl) = vbNullString Then
                Loai_bl = "''"
            Else
                Loai_bl = "'" & Trim(Loai_bl) & "'"
            End If
            
             sSQLVal = Id & "," & Hdr_id & "," & So_tt & "," & Ten_bl & "," & Mau_so & "," & _
            Ky_hieu_bl & "," & TONG_SO & "," & TU_SO_TON & "," & DEN_SO_TON & "," & TU_SO_PS & "," & DEN_SO_PS & "," & _
            TSO_SD_MAT & "," & DSO_SD_MAT & "," & SUM_SD_MAT & "," & SL_DA_SD & "," & SL_XOA & "," & SO_XOA & _
            "," & SL_MAT & "," & SO_MAT & "," & SL_HUY & "," & SO_HUY & "," & TSO_TON_CK & "," & _
            DSO_TON_CK & "," & SL_TON_CK & "," & Loai_bl
        
            sSQL = "INSERT INTO tmp_dtl_BC26_AC_BLP( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
            
        .Col = .ColLetterToNumber("B")
        .Row = .Row + 1
        Loop Until .Text = "aa"
   End With
   
 
   
   
   clsDAO.Disconnect
   
   InsertDTL = vbNullString
   
End Function

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
'**************************************************************
'       ThanhDX remove
'       Date: 14/09/2006
'       Reason: Doi Nguoi su dung -> Phong quan ly.

        'Set NguoiSuDung
'        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_X)
'        .Row = NGUOI_SU_DUNG_Y
'        If strNguoiSuDung <> "" Then
'            .Text = strNguoiSuDung
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'**************************************************************
        ' so thu tu lan quet ma vach
        .Col = .ColLetterToNumber("K")
        .Row = 10
        If strSoTTTKhai <> "" Then
            .Text = strSoTTTKhai
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
      
              
        .Col = .ColLetterToNumber("Q")
        .Row = 12
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strNguoiSuDung, TCVN, UNICODE))
      
      
      
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
'        .Col = .ColLetterToNumber(KY_LAP_BO_X)
'        .Row = KY_LAP_BO_Y
'        strKyLapBo = .Text
'
'        ' Get MaSoTep
''        .Col = .ColLetterToNumber(MA_SO_TEP_X)
''        .Row = MA_SO_TEP_Y
''        strMaSoTep = .Text
'
'        ' Get Phong xu ly
'        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
'        .Row = PHONG_XU_LY_Y
'        strPhongXuLy = .Text
'
'        ' Get NgayNhanToKhai
'        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
'        .Row = NGAY_NHAN_TO_KHAI_Y
'        strNgayNhanToKhai = .Text
'
'        'Go to last sheet (header sheet)
'        .Sheet = .SheetCount
'
'        If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
'            blnValid = False
'        End If
'
''        .Col = .ColLetterToNumber(HEADER_KY_LAP_BO_X)
''        .Row = HEADER_KY_LAP_BO_Y
'
''        If Not blnValid Then
''            .Formula = "0"
''        Else
''            .Formula = "1"
''        End If
'        .EventEnabled(EventAllEvents) = True
    vNgayDauKy = dNgayDauKy
    End With
End Sub

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
    
    Dim vNgayNop, vNgayQuet As Variant
    
    CheckValidData = True
    With fps
         .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        .GetText .ColLetterToNumber("I"), 14, varCheckValue
        .GetText .ColLetterToNumber("T"), 14, varNoteValue
        
        .GetText .ColLetterToNumber("E"), 10, vNgayNop
        .GetText .ColLetterToNumber("K"), 12, vNgayQuet
        
'        If CStr(varCheckValue) = "1" And Trim(CStr(varNoteValue)) = "" Then
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 14, "0"
'            CheckValidData = False
'        Else
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 14, "1"
'        End If
        ' Kiem tra ngay nop khong duoc lon hon ngay quet
        If ToDate(CStr(vNgayNop), "dd/mm/yyyy") > ToDate(CStr(vNgayQuet), "dd/mm/yyyy") Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "1"
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Function


Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
'    KiemTraKhoaSo = True
'    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
'    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("I_10"), "Value"), "dd/mm/yyyy")
'
'    If dNgayKhoaSo < dKyLapBo Then
'        KiemTraKhoaSo = False
'    End If
    KiemTraKhoaSo = False
End Function



Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
 Dim varTemp As Variant
 With fps
    If .ActiveSheet = 1 Then
        If Col = .ColLetterToNumber("I") And Row = 12 Then
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
        If Col = .ColLetterToNumber("I") And Row = 10 Then
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
    Dim MATKHAI As Variant
        
    'kiem tra ton tai tep *.dbf chua
    'Tinh so lan quyet de ghi vao DB trung gian
    'dhdang sua
    'ngay 24/09/2010
'-----------------------------
'    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tmp_bcao_hdr_ac.DBF"
'    If fso.FileExists(strFileNameHDR) = True Then
'        If clsDAO.Connected = False Then
'        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
'        clsDAO.Connect
'        End If
'        sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM tmp_bcao_hdr_ac" & _
'            " WHERE MATKHAI = " & MATKHAI & " AND  MAPP='1' AND MADTNT= " & "'" & strMST & "'" & _
'            " AND KYKKHAI ='" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND (TTHTK = '1' OR TTHTK = '3' OR TTHTK = '4')"
'
'            Set rs = clsDAO.Execute(sSQL)
'            If Not rs Is Nothing Then
'                 LAN_QUET1 = rs.Fields("LAN_QUET")
'
'                 TinhSoLanQuet = LAN_QUET1
'            Else
'                 TinhSoLanQuet = 0
'            End If
'            clsDAO.Disconnect
'    Else
'    TinhSoLanQuet = 0
'    End If
    TinhSoLanQuet = 0
End Function


'dhdang tinh ma nganh
'ngay 25/10/10
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

'Get ve ma CQT
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
