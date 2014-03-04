VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01TTDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_Y = 23
Const KY_LAP_BO_X = "E"
Const NGAY_NHAN_TO_KHAI_Y = 25
Const NGAY_NHAN_TO_KHAI_X = "R"
Const NGUOI_SU_DUNG_Y = 25
Const NGUOI_SU_DUNG_X = "AF"
Const PHONG_XU_LY_Y = 23
Const PHONG_XU_LY_X = "AF"
Const NGAY_NOP_Y = 25
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 25
Const NGAY_QUET_X = "R"
Const TEN_BPQL_X = "AF"
Const TEN_BPQL_Y = 30
Const MA_BPQL_X = "AF"
Const MA_BPQL_Y = 27
Private rs As ADODB.Recordset
' thong tin din vi
Const MA_SO_THUE_X = "H"
Const MA_SO_THUE_Y = 7
Const TEN_GOI_X = "H"
Const TEN_GOI_Y = 5
Const DIA_CHI_X = "H"
Const DIA_CHI_Y = 9
'Const QUAN_HUYEN_X = "F"
'Const QUAN_HUYEN_Y = 11
'Const TINH_TPHO_X = "T"
'Const TINH_TPHO_Y = 11
Const DIEN_THOAI_X = "H"
Const DIEN_THOAI_Y = 11
Const FAX_X = "R"
Const FAX_Y = 11
Const MAIL_X = "AF"
Const MAIL_Y = 11
'thong tin Dai ly thue
Const MA_SO_THUE_DL_X = "H"
Const MA_SO_THUE_DL_Y = 15
Const TEN_GOI_DL_X = "H"
Const TEN_GOI_DL_Y = 13
Const DIA_CHI_DL_X = "H"
Const DIA_CHI_DL_Y = 17
'Const QUAN_HUYEN_DL_X = "F"
'Const QUAN_HUYEN_DL_Y = 17
'Const TINH_TPHO_DL_X = "T"
'Const TINH_TPHO_DL_Y = 17
Const DIEN_THOAI_DL_X = "H"
Const DIEN_THOAI_DL_Y = 19
Const FAX_DL_X = "R"
Const FAX_DL_Y = 19
Const MAIL_DL_X = "AF"
Const MAIL_DL_Y = 19
Const SO_HDDL_DL_X = "H"
Const SO_HDDL_DL_Y = 21
Const NGAY_HDDL_DL_X = "R"
Const NGAY_HDDL_DL_Y = 21

Const CHECK_COL = "M"
Const CHECK_ROW = 49

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
Public strMail As String
Public strQHuyen As String
Public strTPho As String
'vttoan them thong tin dai ly thue
Public strMST_DLT As String
Public strTen_DLT As String
Public strDchi_DLT As String
Public strQHuyen_DLT As String
Public strTTPho_DLT As String
Public strDthoai_DLT As String
Public strFax_DLT As String
Public strMail_DLT As String
Public strSoHD_DLT As String
Public strNgayHD_DLT As String
'end
Public spathVat As String
Public sSaiCT11 As String
Public sKyKeKhai As String
Public dNgayNopDB As Variant
Public HANNOP As String
Public TTHTK As String
Public strTenBpql As String
Private tempKyLb As Variant
Private tempNgayNop As Variant
Public hdrId As String
Public strMaCQT As String
Public strMaPQL As String
Public strTenPQL As String
Public strSoTTTKhai As String
Public isTKTonTai As Boolean
Public checkDLTTonTai As Boolean

Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
        'Ghi chu
        .Col = .ColLetterToNumber("E")
        .Row = 29
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
        SetDateFormat fps, 1, 23, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 25, .ColLetterToNumber("E"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 25
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 25, .ColLetterToNumber("R"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 25
        .Col = .ColLetterToNumber("R")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
    End With
    
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
Dim i As Integer, intIndexCombo As Integer
Dim strLTN As Variant, Col7 As Variant, strLTNCu As Variant, strID As Variant
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim varID As Variant, varDVT As Variant, varThueSuat As Variant
Dim iTenHHDV As String, strCheck As String
Dim strThueSuat As String, strDVT As String
Dim strIdQLT As String
Dim iCol As Long, iRow As Long, lCurrSheet As Long
Dim blnLoadThueSuat As Boolean, blnSuaDVT As Boolean
Dim maTM As String
    With fps
        lCurrSheet = .Sheet
        
        .Sheet = 1
            i = 0
            blnLoadThueSuat = True
            strCheck = GetAttribute(TAX_Utilities_Svr_New.Data(.Sheet - 1).nodeFromID( _
                     GetCellID(fps, .ColLetterToNumber(CHECK_COL), CHECK_ROW)), "Value")
            Do
                varID = Empty
               .GetText .ColLetterToNumber("AP"), i + 53, varID
               If varID <> "" Then
                    .Row = i + 53
                    varThueSuat = .Value
                    DataDM varID, strIdQLT, iTenHHDV, strDVT, strThueSuat, blnSuaDVT, maTM
                    If strIdQLT <> vbNullString Then
                        .Col = .ColLetterToNumber("C")
                        .Row = i + 53
                        .Text = iTenHHDV
                        .Col = .ColLetterToNumber("AO")
                        .Text = maTM
                       ' UpdateCell fps, .ColLetterToNumber("C"), i + 53, .Text
                    ElseIf CStr(varID) <> vbNullString Then
                        DisplayMessage "0081", msOKOnly, miCriticalError
                        Exit Function
                    End If
               End If
                    .RowHeight(i + 53) = .MaxTextRowHeight(i + 53)
                    'UpdateCell fps, .Col, .Row, .Text
                    i = i + 1
                    .Col = .ColLetterToNumber("B")
                    .Row = i + 53
                    If Trim(.Text) = "III" Then blnLoadThueSuat = False
                    
            Loop Until .Row >= .MaxRows
         
            .Sheet = 2
            i = 0
            Do
                strIdQLT = ""
                .GetText .ColLetterToNumber("I"), i + 25, varID
            If varID <> "" Then
                    .Row = i + 25
                    DataDM varID, strIdQLT, iTenHHDV, strDVT, strThueSuat, blnSuaDVT, maTM
                    If strIdQLT <> vbNullString Then
                        .Col = .ColLetterToNumber("L")
                        .Row = i + 25
                        .Text = iTenHHDV
'                        .Col = .ColLetterToNumber("AO")
'                        .Text = maTM
                    End If
               End If
                i = i + 1
                .Col = .ColLetterToNumber("B")
                .Row = i + 25
            Loop Until .Row >= .MaxRows
        .Sheet = lCurrSheet
    End With

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

Private Sub DataDM(ByVal Id As String, Optional ByRef strIdQLT As String, Optional ByRef tenTN As String, Optional ByRef strDVT As String, Optional ByRef strThueSuat As String, Optional ByRef blnSuaThueSuat As Boolean, Optional ByRef maTM As String)
Dim arrDanhsach() As String
Dim strDataFileName As String
Dim xmlDomData As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = GetCatalogueFileName
    
       If xmlDomData.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If Id = arrDanhsach(1) Then
                            strIdQLT = arrDanhsach(0)
                            tenTN = arrDanhsach(2)
                            strDVT = arrDanhsach(3)
                            strThueSuat = arrDanhsach(4)
                            blnSuaThueSuat = IIf(arrDanhsach(5) = "1", True, False)
                            If Val(TAX_Utilities_Svr_New.Year) > 2010 Then
                                maTM = arrDanhsach(8)
                            ElseIf Val(TAX_Utilities_Svr_New.Month) >= 4 And Val(TAX_Utilities_Svr_New.Year) = 2010 Then
                                maTM = arrDanhsach(8)
                            Else
                                maTM = arrDanhsach(7)
                            End If
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub

Public Function TKRB() As Boolean

   TKRB = True

End Function

Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim matkhai As Variant
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TB" & sKyKeKhai & ".DBF"
    Dim strFileNamePL1 As String
    strFileNamePL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_D1" & sKyKeKhai & ".DBF"
    Dim strFileNamePL2 As String
    strFileNamePL2 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_D2" & sKyKeKhai & ".DBF"
    
    Dim strFileNameBS As String
    strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"
    'file HDR
    If fso.FileExists(strFileNameHDR) = False Then
        Dim strTepmauHDR As String
        strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_TK.DBF"
        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    End If
    'file dtl
    If fso.FileExists(strFileNameDTL) = False Then
        Dim strTepmauDTL As String
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_TB.DBF"
        fso.CopyFile strTepmauDTL, strFileNameDTL, False
    End If
    'file phu luc 1
    If fso.FileExists(strFileNamePL1) = False Then
        Dim strTepmauPL1 As String
        strTepmauPL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_D1.DBF"
        fso.CopyFile strTepmauPL1, strFileNamePL1, False
    End If
    'file phu luc 2
    If fso.FileExists(strFileNamePL2) = False Then
        Dim strTepmauPL2 As String
        strTepmauPL2 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_D2.DBF"
        fso.CopyFile strTepmauPL2, strFileNamePL2, False
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
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("AO"), 3, matkhai
        If Trim(matkhai) = vbNullString Then
                matkhai = "''"
        Else
                matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        End If
   End With
        
   'sSQL = "SELECT MST, NGAYNOP FROM TKHDR WHERE LOAITK = '01BTNDN' AND ID = '" & strMST & sKyKeKhai & "'"
    sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyKeKhai & _
          " WHERE MATKHAI = " & matkhai & " AND  MAPP='1' AND MADTNT= " & "'" & strMST & "'" & _
          " AND KYKKHAI ='" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
          
   
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
    Dim matkhai As Variant
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("AO"), 3, matkhai
        If Trim(matkhai) = vbNullString Then
                matkhai = "''"
        Else
                matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        End If
   End With
   
   
   sSQL = "UPDATE TMP_TK" & sKyKeKhai & _
        " SET THUETKY2=1" & _
        " WHERE MATKHAI = " & matkhai & " AND  MAPP='1' AND MADTNT= " & "'" & strMST & "'" & _
        " AND KYKKHAI ='" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
    
   clsDAO.ExecuteDLL sSQL
   
   UpdateTHUETKY2 = True
   
   clsDAO.Disconnect

End Function


Public Function XoaTKTT() As Boolean

    Dim sSQL As String
    Dim KYLBO As Variant
    Dim matkhai As Variant
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
    
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 23, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If
        
        .GetText .ColLetterToNumber("AO"), 3, matkhai
        If Trim(matkhai) = vbNullString Then
                matkhai = "''"
        Else
                matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        End If
   End With
    
  
   sSQL = "DELETE FROM TMP_TK" & sKyKeKhai & _
   " WHERE MATKHAI = " & matkhai & " AND  MAPP='1' AND MADTNT= " & "'" & strMST & "'" & _
   " AND KYKKHAI ='" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
          
   clsDAO.ExecuteDLL sSQL
  
   sSQL = vbNullString
   sSQL = "DELETE FROM TMP_TB" & sKyKeKhai & " WHERE MATKHAI = " & matkhai & " AND MADTNT = '" & strMST & "'" & _
   " AND KYLBO=" & KYLBO
   
   
   clsDAO.ExecuteDLL sSQL
   ' Xoa phu luc 1
   
    sSQL = "DELETE FROM TMP_D1" & sKyKeKhai & " WHERE MADTNT = '" & strMST & "'" & _
   " AND KYLBO=" & KYLBO
   clsDAO.ExecuteDLL sSQL
   ' Xoa phu luc 2
   sSQL = "DELETE FROM TMP_D2" & sKyKeKhai & " WHERE MADTNT = '" & strMST & "'" & _
   " AND KYLBO=" & KYLBO
   clsDAO.ExecuteDLL sSQL
   
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
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    Dim bln        As Boolean
    
    Dim madtnt     As Variant
    Dim MATHUE     As Variant
    Dim maTM       As Variant
    Dim MAPP       As Variant
    Dim KYLBO      As Variant
    Dim ngnop      As Variant
    Dim LoaiTien   As Variant
    Dim MABPQL     As Variant
    Dim SHTEP      As Variant
    Dim kykkhai    As Variant
    Dim NgNhap     As Variant
    Dim ThueTKy    As Variant
    Dim GTTinhThue As Variant
    Dim DaGhi      As Variant
    Dim DAGHI1     As Variant
    Dim KTTRUOC    As Variant
    Dim DoanhSo    As Variant
    Dim MAPP2      As Variant
    Dim MaMuc      As Variant
    Dim MaVach     As Variant
    Dim CT27       As Variant
    Dim ct34       As Variant
    Dim CT36       As Variant
    Dim matkhai    As Variant
    Dim HANNOP2    As Variant
    Dim ThueTKy2   As Variant
    Dim TENDA      As Variant
    Dim ROWNUM     As Integer
    Dim THUECL     As Variant
    Dim NGANH      As Variant
    Dim DAGHI2     As Variant
    Dim AUTO_CAL   As Variant
    Dim TEMP       As Variant
    Dim MADLT      As Variant
    Dim LanBS      As Variant
    Dim NGAYPS     As Variant
    Dim TTHTK1     As Variant
    Dim solanBS    As Variant
    Dim DANHAN     As Variant
    Dim MANNKD     As Variant
    Dim DSoCL      As Variant
    Dim GTTTCL     As Variant
    Dim TENNV      As Variant
    Dim CHUNGCHI   As Variant

    'test
    Dim arrMTM()   As String, arrMTMSave() As String
    Dim arrVitri() As Integer
    Dim idx, idx1, i, countMTM, Section As Integer
    Dim vTemp As String, vVitriTrung As String
    'end test
    NGANH = TinhMaNganh
      
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
    sSQLCol = "MADTNT, MATHUE, MATM, MAPP, KYLBO, NGNOP, "
    sSQLCol = sSQLCol & "LOAITIEN, MABPQL, SHTEP, TTHTK, KYKKHAI, "
    sSQLCol = sSQLCol & "NGNHAP, THUETKY, GTTINHTHUE, DAGHI, "
    sSQLCol = sSQLCol & "KTTRUOC, DOANHSO, MAPP2, MAMUC, MAVACH,MATKHAI,HANNOP, THUETKY2,HANNOP2,TENDA,DAGHI1,DAGHI2, AUTO_CAL, NGANH, LAN_QUET, MADLT, LANBS, NgayPS, THUECL, MANNKD, DSOCL, GTTTCL, DANHAN, TENNV, CHUNGCHI"
    
    ' Ghi HDR muc I
    With fps
        .Sheet = 1
        Section = 0
        
        .GetText .ColLetterToNumber("T"), .MaxRows - 4, TENNV
        .GetText .ColLetterToNumber("T"), .MaxRows - 2, CHUNGCHI

        'dhdang phan lay vao mang
        'ngay 14/09/2010
        .Col = .ColLetterToNumber("B")
        .Row = 53

        Do
            ' Lay ma tieu muc tren cel
            .Col = .ColLetterToNumber("AO")

            If .Text <> vbNullString And .Text <> "" Then vTemp = vTemp & .Text & "~"
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1

            If .Text = "aa" Then Section = Section + 1
        Loop Until .Text = "aa" And Section = 5
         
        If vTemp <> "" Then
            arrMTM = Split(vTemp, "~")
            'vttoan comment: vi khong gi duoc MATM khi chi nhap mot dong du lieu
            '         vTemp = ""
            ReDim arrMTMSave(1)
            arrMTMSave(0) = arrMTM(0)

            For idx = 1 To UBound(arrMTM) - 1
                countMTM = 0
                vTemp = arrMTM(idx)

                For i = 0 To UBound(arrMTMSave)

                    If vTemp = arrMTMSave(i) Then
                        countMTM = countMTM + 1
                    End If

                Next

                If countMTM = 0 Then
                    ReDim Preserve arrMTMSave(UBound(arrMTMSave) + 1)
                    arrMTMSave(UBound(arrMTMSave) - 1) = vTemp
                End If

            Next

        End If

        ' end lay ma tieu muc
        
        .GetText .ColLetterToNumber("AO"), 3, matkhai

        If Trim(matkhai) = vbNullString Then
            matkhai = "''"
        Else
            matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        End If

        'mst
        .GetText .ColLetterToNumber("H"), 7, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        ' ma so thue dai ly thue
        .GetText .ColLetterToNumber("H"), 15, MADLT

        If Trim(MADLT) = vbNullString Then
            MADLT = "''"
        Else
            MADLT = "'" & MADLT & "'"
        End If

        'trang thai to khai
        .GetText .ColLetterToNumber("U"), 3, TTHTK1

        If Trim(TTHTK1) = vbNullString Then
            TTHTK = "'" & 2 & "'"
        End If

        ' so lan bo sung
        If TTHTK = "'" & 2 & "'" Then
            .GetText .ColLetterToNumber("AB"), 3, solanBS
            LanBS = solanBS
        Else
            LanBS = "0"
        End If
        
        ' ngay nop
        .GetText .ColLetterToNumber("E"), 25, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        ' ngay phat sinh
        .GetText .ColLetterToNumber("AA"), 44, NGAYPS

        If (NGAYPS) = vbNullString Then
            NGAYPS = "CTOD('')"
        Else
            NGAYPS = ToDate(Trim(NGAYPS), DDMMYYYY)
            NGAYPS = "CTOD('" & Format(NGAYPS, "mm/dd/yyyy") & "')"
            HANNOP = DateAdd("d", 10, ngnop)
        End If

        ' ma nghanh nghe kinh doanh
        .GetText .ColLetterToNumber("AE"), 44, MANNKD

        If (MANNKD) = vbNullString Then
            MANNKD = "''"
        Else
            MANNKD = "'" & MANNKD & "'"
        End If
                
        MATHUE = "'02'"
        MAPP = "'1'"
        LoaiTien = "''"
        DaGhi = ".F."
        DAGHI1 = ".F."
        MAPP2 = "'1'"
        THUECL = "0"
        MaVach = ".T."

        If checkDLTTonTai = False Then
            DAGHI2 = ".T."
        Else
            DAGHI2 = ".F."
        End If

        AUTO_CAL = ".F."
        SHTEP = "''"
        DANHAN = "''"
        DSoCL = "0"
        GTTTCL = "0"
        'ky lap bo
        .GetText .ColLetterToNumber("E"), 23, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        LoaiTien = "''"
        .GetText .ColLetterToNumber(MA_BPQL_X), MA_BPQL_Y, MABPQL

        If Trim(MABPQL) = vbNullString Then
            MABPQL = "''"
        Else
            MABPQL = "'" & MABPQL & "'"
        End If

        ' ky ke khai
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        
        '.GetText .ColLetterToNumber("M"), 12, NGNHAP
        NgNhap = Date

        If Trim(NgNhap) = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If
       
        If Trim(HANNOP) = vbNullString Then
            HANNOP = "CTOD('')"
        Else
            HANNOP = ToDate(Trim(HANNOP), DDMMYYYY)
            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
        End If
        
        ThueTKy2 = "0"
        HANNOP2 = HANNOP
        TENDA = "''"
        
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'06'"
            MaMuc = "'015'"
        Else
            MaMuc = "'1750'"
            maTM = "''"
        End If
        
        '        .GetText .ColLetterToNumber("W"), 51, DOANHSO
        '        If Trim(DOANHSO) = vbNullString Then
        '            DOANHSO = "0"
        '        End If
        '
        '        .GetText .ColLetterToNumber("AG"), 51, KTTRUOC
        '        If Trim(KTTRUOC) = vbNullString Then
        '            KTTRUOC = "0"
        '        End If
        '
        '        .GetText .ColLetterToNumber("AB"), 51, GTTINHTHUE
        '        If Trim(GTTINHTHUE) = vbNullString Then
        '            GTTINHTHUE = "0"
        '        End If
        '
        '        .GetText .ColLetterToNumber("AN"), 51, THUETKY
        '        If Trim(THUETKY) = vbNullString Then
        '            THUETKY = "0"
        '        End If
        '
        '    End With
        '    sSQLVal = MADTNT & "," & MATHUE & "," & MATM & "," & MAPP & "," & KYLBO & "," & NGNOP & ", "
        '    sSQLVal = sSQLVal & LOAITIEN & "," & MABPQL & "," & SHTEP & "," & TTHTK & "," & KYKKHAI & ", "
        '    sSQLVal = sSQLVal & NGNHAP & "," & THUETKY & "," & GTTINHTHUE & "," & DAGHI & ", "
        '    sSQLVal = sSQLVal & KTTRUOC & "," & DOANHSO & "," & MAPP2 & "," & MAMUC & "," & MAVACH & ", "
        '    sSQLVal = sSQLVal & MATKHAI & "," & HANNOP & "," & THUETKY2 & "," & HANNOP2 & "," & TENDA & "," & DAGHI1 & "," & DAGHI2 & "," & AUTO_CAL & ",'" & NGANH & "'," & LAN_QUET & "," & MADLT & "," & LANBS & "," & NgayPS & "," & THUECL & "," & MANNKD & "," & DSOCL & "," & GTTTCL & "," & DANHAN
        '
        '    sSQL = "INSERT INTO TMP_TK" & sKyKeKhai & _
        '     "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        '      bln = clsDAO.ExecuteDLL(sSQL)
        '    'bo cau insert ngay 15/09/2010

        '--------------
        'Ghi du lieu HDR
        .Sheet = 1

        If vTemp <> "" Then

            For i = 0 To UBound(arrMTMSave) - 1
                KTTRUOC = 0

                DoanhSo = 0
                ThueTKy = 0
                GTTinhThue = 0
                vTemp = arrMTMSave(i)
                maTM = "'" & vTemp & "'"
                .Col = .ColLetterToNumber("B")
                .Row = 53
                Section = 0

                Do
                    .Col = .ColLetterToNumber("AO")

                    If vTemp = Trim(.Text) Then
                        .Col = .ColLetterToNumber("W")
                        TEMP = .Value

                        If Trim(TEMP) = vbNullString Then
                            TEMP = "0"
                        End If

                        DoanhSo = DoanhSo + CDbl(TEMP)
                             
                        .Col = .ColLetterToNumber("AG")
                        TEMP = .Value

                        If Trim(TEMP) = vbNullString Then
                            TEMP = "0"
                        End If

                        KTTRUOC = KTTRUOC + CDbl(TEMP)
                             
                        .Col = .ColLetterToNumber("AB")
                        TEMP = .Value

                        If Trim(TEMP) = vbNullString Then
                            TEMP = "0"
                        End If

                        GTTinhThue = GTTinhThue + CDbl(TEMP)
                            
                        .Col = .ColLetterToNumber("AN")
                        TEMP = .Value

                        If Trim(TEMP) = vbNullString Then
                            TEMP = "0"
                        End If
                            
                        ThueTKy = ThueTKy + CDbl(TEMP)
                    End If

                    .Col = .ColLetterToNumber("B")
                    .Row = .Row + 1

                    If .Text = "aa" Then Section = Section + 1
                Loop Until .Text = "aa" And Section = 5

                sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & ngnop & ", "
                sSQLVal = sSQLVal & LoaiTien & "," & MABPQL & "," & SHTEP & "," & TTHTK & "," & kykkhai & ", "
                sSQLVal = sSQLVal & NgNhap & "," & ThueTKy & "," & GTTinhThue & "," & DaGhi & ", "
                sSQLVal = sSQLVal & KTTRUOC & "," & DoanhSo & "," & MAPP2 & "," & MaMuc & "," & MaVach & ", "
                sSQLVal = sSQLVal & matkhai & "," & HANNOP & "," & ThueTKy2 & "," & HANNOP2 & "," & TENDA & "," & DAGHI1 & "," & DAGHI2 & "," & AUTO_CAL & ",'" & NGANH & "'," & LAN_QUET & "," & MADLT & "," & LanBS & "," & NGAYPS & "," & THUECL & "," & MANNKD & "," & DSoCL & "," & GTTTCL & "," & DANHAN & ",'" & TENNV & "','" & CHUNGCHI & "'"
                
                sSQL = "INSERT INTO TMP_TK" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                bln = clsDAO.ExecuteDLL(sSQL)
            Next

        Else
            'dntai sua ngay 10/08/2011
            KTTRUOC = 0

            DoanhSo = 0
            ThueTKy = 0
            GTTinhThue = 0
            'end
            sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & ngnop & ", "
            sSQLVal = sSQLVal & LoaiTien & "," & MABPQL & "," & SHTEP & "," & TTHTK & "," & kykkhai & ", "
            sSQLVal = sSQLVal & NgNhap & "," & ThueTKy & "," & GTTinhThue & "," & DaGhi & ", "
            sSQLVal = sSQLVal & KTTRUOC & "," & DoanhSo & "," & MAPP2 & "," & MaMuc & "," & MaVach & ", "
            sSQLVal = sSQLVal & matkhai & "," & HANNOP & "," & ThueTKy2 & "," & HANNOP2 & "," & TENDA & "," & DAGHI1 & "," & DAGHI2 & "," & AUTO_CAL & ",'" & NGANH & "'," & LAN_QUET & "," & MADLT & "," & LanBS & "," & NGAYPS & "," & THUECL & "," & MANNKD & "," & DSoCL & "," & GTTTCL & "," & DANHAN & ",'" & TENNV & "','" & CHUNGCHI & "'"
                
            sSQL = "INSERT INTO TMP_TK" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
        End If

    End With
                
    clsDAO.Disconnect
   
    InsertHDR = vbNullString
End Function
Public Function InsertDTL_KHBS() As String
    Dim sSQL        As String
    Dim sSQLCol     As String
    Dim sSQLVal     As String
    Dim bln         As Boolean
    Dim madtnt      As Variant
    Dim ngnop       As Variant
    Dim KYLBO       As Variant
    Dim kykkhai     As Variant
    Dim matkhai     As Variant
    Dim MATHUE      As Variant
    Dim MaMuc       As Variant
    Dim maTM        As Variant
    Dim MAPP        As Variant
    Dim mact        As Variant
    Dim MACT2       As Variant
    Dim MACT3       As Variant
    Dim STT         As Variant
    Dim STT2        As Variant
    Dim STTIN       As Variant
    Dim SOKK        As Variant
    Dim SODC        As Variant
    Dim SOCL        As Variant
    Dim i           As Integer, iCol As Long, iRow As Long
    Dim xmlNode     As MSXML.IXMLDOMNode
    
    Dim strFileName As String
    Dim fso         As New FileSystemObject
    Dim DANHAN      As Variant
    'them bien So Tien nop cham va so ngay nop cham trong TT38
    Dim SOTIENNC    As Variant
    Dim SONGAYNC    As Variant
    
    'kiem tra ton tai tep *.dbf chua
    '    strFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "BS" & sKyKeKhai & ".DBF"
    '    If fso.FileExists(strFileName) = False Then
    '    Dim strTepmau As String
    '        strTepmau = spathVat & "\tepmau\BS.DBF"
    '        fso.CopyFile strTepmau, strFileName, False
    '    End If
   
    sSQLCol = "MADTNT, KYKKHAI, MATKHAI, MATHUE, MAMUC, MATM, MAPP, NGNOP, TTHTK, KYLBO , MACT, MACT2, MACT3, STT, STT2, STTIN," & " SOKK, SODC, SOCL,LAN_QUET, DANHAN,SOTIENNC,SONGAYNC"
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("H"), 7, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"


        .GetText .ColLetterToNumber("AO"), 3, matkhai
        If Trim(matkhai) = vbNullString Then
            matkhai = "''"
        Else
            matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        End If
        MATHUE = "'01'"
        DANHAN = "''"
        .GetText .ColLetterToNumber("E"), 23, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If

        End If
        
        ' sua MAMUC & MATM tu nam 2009
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If

        MAPP = "'5'"
        
        .GetText .ColLetterToNumber("E"), 25, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                
        ' Insert so dieu chinh tang
        
        mact = "'1000'"
        MACT2 = "''"
        MACT3 = "''"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        SOKK = "0"
        SODC = "0"
        SOCL = "0"
        SOTIENNC = "0"
        SONGAYNC = "0"
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & LAN_QUET & "," & DANHAN & "," & SOTIENNC & "," & SONGAYNC
                       
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        .Sheet = .SheetCount - 1
        
        'get so tien nop cham va so ngay nop cham
        .GetText .ColLetterToNumber("BE"), .MaxRows - 27, SOTIENNC

        If SOTIENNC = vbNullString Then
            SOTIENNC = "0"
        Else
            SOTIENNC = Replace(SOTIENNC, ".", "")
        End If

        .GetText .ColLetterToNumber("BE"), .MaxRows - 28, SONGAYNC

        If SONGAYNC = vbNullString Then
            SONGAYNC = "0"
        Else
            SONGAYNC = Replace(SONGAYNC, ".", "")
        End If

        'end
        
        i = 1
        .Row = 9

        Do
            mact = "''"
            .GetText .ColLetterToNumber("BE"), .Row, MACT2

            If Trim(MACT2) = vbNullString Then
                MACT2 = "''"
            ElseIf MACT2 = "14" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "16" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "23" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "24" Then
                MACT2 = "'050'"
            ElseIf MACT2 = "21" Then
                MACT2 = "'039'"
            ElseIf MACT2 = "31" Then
                MACT2 = "'050'"
            ElseIf MACT2 = "30" Then
                MACT2 = "'031'"
            Else
                MACT2 = "'000'"
            End If
                
            MACT3 = "'2'"
            STT = 1
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("BF"), .Row, SOKK

            If Trim(SOKK) = vbNullString Then
                SOKK = "0"
            Else
                SOKK = SOKK
            End If
                
            .GetText .ColLetterToNumber("BG"), .Row, SODC

            If Trim(SODC) = vbNullString Then
                SODC = "0"
            Else
                SODC = SODC
            End If
                
            .GetText .ColLetterToNumber("BH"), .Row, SOCL

            If Trim(SOCL) = vbNullString Then
                SOCL = "0"
            Else
                SOCL = SOCL
            End If
            
            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & LAN_QUET & "," & DANHAN & "," & SOTIENNC & "," & SONGAYNC
                           
            sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                
            If MACT2 <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
        
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"

        ' Insert so dieu chinh giam

        mact = "'2000'"
        MACT2 = "''"
        MACT3 = "''"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        SOKK = "0"
        SODC = "0"
        SOCL = "0"
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & LAN_QUET & "," & DANHAN & "," & SOTIENNC & "," & SONGAYNC
                       
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

        i = 1
        .Row = .Row + 3

        Do
            mact = "''"
            .GetText .ColLetterToNumber("BE"), .Row, MACT2

            If Trim(MACT2) = vbNullString Then
                MACT2 = "''"
            ElseIf MACT2 = "14" Then
                MACT2 = "'041'"
            ElseIf MACT2 = "16" Then
                MACT2 = "'041'"
                '                ElseIf MACT2 = "23" Then
                '                    MACT2 = "'041'"
            ElseIf MACT2 = "24" Then
                MACT2 = "'050'"
            ElseIf MACT2 = "21" Then
                MACT2 = "'039'"
            ElseIf MACT2 = "31" Then
                MACT2 = "'050'"
            Else
                MACT2 = "'000'"
            End If
                
            MACT3 = "'2'"
            STT = 2
            STT2 = i
            STTIN = "'" & CStr(i) & "'"
                
            .GetText .ColLetterToNumber("BF"), .Row, SOKK

            If Trim(SOKK) = vbNullString Then
                SOKK = "0"
            Else
                SOKK = SOKK
            End If
                
            .GetText .ColLetterToNumber("BG"), .Row, SODC

            If Trim(SODC) = vbNullString Then
                SODC = "0"
            Else
                SODC = SODC
            End If
                
            .GetText .ColLetterToNumber("BH"), .Row, SOCL

            If Trim(SOCL) = vbNullString Then
                SOCL = "0"
            Else
                SOCL = SOCL
            End If
            
            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & LAN_QUET & "," & DANHAN & "," & SOTIENNC & "," & SONGAYNC
                           
            sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If MACT2 <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
        
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "bb"
        
        ' Insert Tong hop dieu chinh

        mact = "'3000'"
        MACT2 = "''"
        MACT3 = "''"
        STT = 3
        STT2 = 0
        STTIN = "'III'"
        SOKK = "0"
        SODC = "0"
        SOCL = "0"
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & LAN_QUET & "," & DANHAN & "," & SOTIENNC & "," & SONGAYNC
                       
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        .Row = .Row + 4
        mact = "''"
        .GetText .ColLetterToNumber("BE"), .Row, MACT2

        If Trim(MACT2) = vbNullString Then
            MACT2 = "''"
        ElseIf MACT2 = "14" Then
            MACT2 = "'041'"
        ElseIf MACT2 = "16" Then
            MACT2 = "'041'"
            '                ElseIf MACT2 = "23" Then
            '                    MACT2 = "'041'"
        ElseIf MACT2 = "24" Then
            MACT2 = "'050'"
        ElseIf MACT2 = "21" Then
            MACT2 = "'039'"
        ElseIf MACT2 = "31" Then
            MACT2 = "'050'"
        Else
            MACT2 = "'000'"
        End If
                
        MACT3 = "'2'"
        STT = 3
        STT2 = 1
        STTIN = "'1'"
                
        .GetText .ColLetterToNumber("BF"), .Row, SOKK

        If Trim(SOKK) = vbNullString Then
            SOKK = "0"
        Else
            SOKK = SOKK
        End If
                
        .GetText .ColLetterToNumber("BG"), .Row, SODC

        If Trim(SODC) = vbNullString Then
            SODC = "0"
        Else
            SODC = SODC
        End If
                
        .GetText .ColLetterToNumber("BH"), .Row, SOCL

        If Trim(SOCL) = vbNullString Then
            SOCL = "0"
        Else
            SOCL = SOCL
        End If
            
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & LAN_QUET & "," & DANHAN & "," & SOTIENNC & "," & SONGAYNC
                           
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        If MACT2 <> "''" Then
            bln = clsDAO.ExecuteDLL(sSQL)
        End If

    End With
   
    clsDAO.Disconnect
    InsertDTL_KHBS = vbNullString
   
End Function


Public Function InsertDTL_PL1() As String
   Dim sSQL As String
   Dim sSQLCol As String
   Dim sSQLVal As String
   Dim bln  As Boolean
   
  Dim madtnt As Variant
  Dim ngnop As Variant
  Dim KYLBO As Variant
  Dim kykkhai As Variant
  Dim KHHD As Variant
  Dim SOHD As Variant
  Dim NGAYHD As Variant
  Dim TENKH As Variant
  Dim TENHHDV As Variant
  Dim SOLUONG As Variant
  Dim DONGIA As Variant
  Dim DSTTDB As Variant
  Dim DANHAN As Variant
   
   
   
   
   sSQLCol = "MADTNT, NGNOP, TTHTK, KYLBO, KYKKHAI,KHHD, SOHD, NGAYHD, TENKH, TENHHDV,"
   sSQLCol = sSQLCol + "SOLUONG, DONGIA, DSTTDB,LAN_QUET, DANHAN"
   
   DANHAN = "''"
   
   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   
   With fps
        .Sheet = 1
         .GetText .ColLetterToNumber("H"), 7, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 25, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber("E"), 23, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
                
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
       .Sheet = 2
             .Col = .ColLetterToNumber("B")
             .Row = 25
                 Do
                       .Col = .ColLetterToNumber("C")
                       KHHD = .Text
                       If Trim(KHHD) = vbNullString Then
                            KHHD = "''"
                       Else
                            KHHD = "'" & TAX_Utilities_Svr_New.Convert(CStr(KHHD), UNICODE, TCVN) & "'"
                       End If
                       
                       .Col = .ColLetterToNumber("D")
                       SOHD = .Text
                       If Trim(SOHD) = vbNullString Then
                            SOHD = "''"
                       Else
                            SOHD = "'" & TAX_Utilities_Svr_New.Convert(CStr(SOHD), UNICODE, TCVN) & "'"
                       End If
                       
                       .Col = .ColLetterToNumber("E")
                       NGAYHD = .Text
                       If Trim(NGAYHD) = vbNullString Then
                            NGAYHD = "CTOD('')"
                       Else
                            NGAYHD = ToDate(Trim(NGAYHD), DDMMYYYY)
                            NGAYHD = "CTOD('" & Format(NGAYHD, "mm/dd/yyyy") & "')"
                       End If
                       
                       .Col = .ColLetterToNumber("F")
                       TENKH = .Text
                       If Trim(TENKH) = vbNullString Then
                            TENKH = "''"
                       Else
                            TENKH = "'" & TAX_Utilities_Svr_New.Convert(CStr(TENKH), UNICODE, TCVN) & "'"
                       End If
                       
                       .Col = .ColLetterToNumber("J")
                       TENHHDV = .Text
                       If Trim(TENHHDV) = vbNullString Then
                            TENHHDV = "''"
                       Else
                            TENHHDV = "'" & TAX_Utilities_Svr_New.Convert(CStr(TENHHDV), UNICODE, TCVN) & "'"
                       End If
                       
                       
                       
                       .Col = .ColLetterToNumber("N")
                       SOLUONG = .Value
                       If Trim(SOLUONG) = vbNullString Then
                            SOLUONG = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("O")
                       DONGIA = .Value
                       If Trim(DONGIA) = vbNullString Then
                            DONGIA = "0"
                       End If
                       
                       .Col = .ColLetterToNumber("P")
                       DSTTDB = .Value
                       If Trim(DSTTDB) = vbNullString Then
                            DSTTDB = "0"
                       End If
                       
                       sSQLVal = madtnt & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & kykkhai & "," & _
                       KHHD & "," & SOHD & "," & NGAYHD & "," & TENKH & "," & TENHHDV & "," & _
                       SOLUONG & "," & DONGIA & "," & DSTTDB & "," & LAN_QUET & "," & DANHAN
                       
                       
                       sSQL = "INSERT INTO TMP_D1" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                       bln = clsDAO.ExecuteDLL(sSQL)

                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                  Loop Until .Text = "aa"
                 
        End With
   
   clsDAO.Disconnect
   
   InsertDTL_PL1 = vbNullString
   
End Function
Public Function InsertDTL_PL2() As String
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim bln  As Boolean

    Dim madtnt As Variant
    Dim ngnop As Variant
    Dim KYLBO As Variant
    Dim kykkhai As Variant
    Dim KHCT As Variant
    Dim SOCT As Variant
    Dim NGAYCT As Variant
    Dim TENNL As Variant
    Dim SLDVMV As Variant
    Dim THUEDN As Variant
    Dim THUEDNMV As Variant
    Dim THUEKT As Variant
    Dim THUECHUAKT As Variant
    Dim DANHAN As Variant
    Dim ROWNUM As Variant


   sSQLCol = "MADTNT, NGNOP, TTHTK, KYLBO, KYKKHAI, KHCT, SOCT, NGAYCT, TENNL, SLDVMV, THUEDN, THUEDNMV, "
   sSQLCol = sSQLCol + "THUEKT, THUECHUAKT, LAN_QUET, DANHAN"


   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If


   With fps
        .Sheet = 1
         .GetText .ColLetterToNumber("H"), 7, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        .GetText .ColLetterToNumber("E"), 25, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 23, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
        DANHAN = "''"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
       .Sheet = 3
              'I
             .Col = .ColLetterToNumber("B")
             .Row = 24
                 Do
                       .Col = .ColLetterToNumber("C")
                       KHCT = .Text
                       If Trim(KHCT) = vbNullString Then
                            KHCT = "''"
                       Else
                            KHCT = "'" & TAX_Utilities_Svr_New.Convert(CStr(KHCT), UNICODE, TCVN) & "'"
                       End If

                       .Col = .ColLetterToNumber("D")
                       SOCT = .Text
                       If Trim(SOCT) = vbNullString Then
                            SOCT = "''"
                       Else
                            SOCT = "'" & TAX_Utilities_Svr_New.Convert(CStr(SOCT), UNICODE, TCVN) & "'"
                       End If

                       .Col = .ColLetterToNumber("E")
                       NGAYCT = .Text
                       If Trim(NGAYCT) = vbNullString Then
                            NGAYCT = "CTOD('')"
                       Else
                            NGAYCT = ToDate(Trim(NGAYCT), DDMMYYYY)
                            NGAYCT = "CTOD('" & Format(NGAYCT, "mm/dd/yyyy") & "')"
                       End If

                       .Col = .ColLetterToNumber("F")
                       TENNL = .Text
                       If Trim(TENNL) = vbNullString Then
                            TENNL = "''"
                       Else
                            TENNL = "'" & TAX_Utilities_Svr_New.Convert(CStr(TENNL), UNICODE, TCVN) & "'"
                       End If

                       .Col = .ColLetterToNumber("G")
                       SLDVMV = .Value
                       If Trim(SLDVMV) = vbNullString Then
                            SLDVMV = "0"
                       End If

                       .Col = .ColLetterToNumber("H")
                       THUEDN = .Value
                       If Trim(THUEDN) = vbNullString Then
                            THUEDN = "0"
                       End If

                       .Col = .ColLetterToNumber("I")
                       THUEDNMV = .Value
                       If Trim(THUEDNMV) = vbNullString Then
                            THUEDNMV = "0"
                       End If

                       .Col = .ColLetterToNumber("J")
                       THUEKT = .Value
                       If Trim(THUEKT) = vbNullString Then
                            THUEKT = "0"
                       End If

                       .Col = .ColLetterToNumber("K")
                       THUECHUAKT = .Value
                       If Trim(THUECHUAKT) = vbNullString Then
                            THUECHUAKT = "0"
                       End If

                       sSQLVal = madtnt & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & kykkhai & "," & _
                       KHCT & "," & SOCT & "," & NGAYCT & "," & TENNL & "," & SLDVMV & "," & _
                       THUEDN & "," & THUEDNMV & "," & THUEKT & "," & THUECHUAKT & "," & LAN_QUET & "," & DANHAN


                       sSQL = "INSERT INTO TMP_D2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                       bln = clsDAO.ExecuteDLL(sSQL)

                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                  Loop Until .Text = "aa"
                  ROWNUM = .Row

                  ' II
'                  .Col = .ColLetterToNumber("B")
'                  .Row = ROWNUM + 9
'                 Do
'                       .Col = .ColLetterToNumber("C")
'                       KHCT = .Text
'                       If Trim(KHCT) = vbNullString Then
'                            KHCT = "''"
'                       Else
'                            KHCT = "'" & TAX_Utilities_Svr_New.Convert(CStr(KHCT), UNICODE, TCVN) & "'"
'                       End If
'
'                       .Col = .ColLetterToNumber("D")
'                       SOCT = .Text
'                       If Trim(SOCT) = vbNullString Then
'                            SOCT = "''"
'                       Else
'                            SOCT = "'" & TAX_Utilities_Svr_New.Convert(CStr(SOCT), UNICODE, TCVN) & "'"
'                       End If
'
'                       .Col = .ColLetterToNumber("E")
'                       NGAYCT = .Text
'                       If Trim(NGAYCT) = vbNullString Then
'                            NGAYCT = "CTOD('')"
'                       Else
'                            NGAYCT = ToDate(Trim(NGAYCT), DDMMYYYY)
'                            NGAYCT = "CTOD('" & Format(NGAYCT, "mm/dd/yyyy") & "')"
'                       End If
'
'                       .Col = .ColLetterToNumber("F")
'                       TENNL = .Text
'                       If Trim(TENNL) = vbNullString Then
'                            TENNL = "''"
'                       Else
'                            TENNL = "'" & TAX_Utilities_Svr_New.Convert(CStr(TENNL), UNICODE, TCVN) & "'"
'                       End If
'
'                       .Col = .ColLetterToNumber("G")
'                       SLDVMV = .Value
'                       If Trim(SLDVMV) = vbNullString Then
'                            SLDVMV = "0"
'                       End If
'
'                       .Col = .ColLetterToNumber("H")
'                       THUEDN = .Value
'                       If Trim(THUEDN) = vbNullString Then
'                            THUEDN = "0"
'                       End If
'
'                       .Col = .ColLetterToNumber("I")
'                       THUEDNMV = .Value
'                       If Trim(THUEDNMV) = vbNullString Then
'                            THUEDNMV = "0"
'                       End If
'
'                       .Col = .ColLetterToNumber("J")
'                       THUEKT = .Value
'                       If Trim(THUEKT) = vbNullString Then
'                            THUEKT = "0"
'                       End If
'
'                       .Col = .ColLetterToNumber("K")
'                       THUECHUAKT = .Value
'                       If Trim(THUECHUAKT) = vbNullString Then
'                            THUECHUAKT = "0"
'                       End If
'
'                       sSQLVal = MADTNT & "," & NGNOP & "," & TTHTK & "," & KYLBO & "," & KYKKHAI & "," & _
'                       KHCT & "," & SOCT & "," & NGAYCT & "," & TENNL & "," & SLDVMV & "," & _
'                       THUEDN & "," & THUEDNMV & "," & THUEKT & "," & THUECHUAKT & "," & LAN_QUET & "," & DANHAN
'
'
'                       sSQL = "INSERT INTO TMP_D2" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
'                       bln = clsDAO.ExecuteDLL(sSQL)
'
'                     .Col = .ColLetterToNumber("B")
'                     .Row = .Row + 1
'                  Loop Until .Text = "bb"
            End With

   clsDAO.Disconnect

   InsertDTL_PL2 = vbNullString

End Function

Public Function InsertDTL() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    Dim bln        As Boolean
   
    Dim KYLBO      As Variant
    Dim madtnt     As Variant
    Dim ngnop      As Variant
    Dim MATHUE     As Variant
    Dim MAPP       As Variant
    Dim NGANHDB    As Variant
    Dim MaMuc      As Variant
    Dim maTM       As Variant
    Dim mact       As Variant
    Dim STT        As Variant
    Dim STT2       As Variant
    Dim STTIN      As Variant
    Dim DVT        As Variant
    Dim S_LUONG    As Variant
    Dim DOANHTHUBH As Variant
    Dim THUETINH   As Variant
    Dim THUESUAT   As Variant
    Dim GIATRITT   As Variant
    Dim THUEKT     As Variant
    Dim THUEDC     As Variant
    Dim THUETTDB   As Variant
    Dim matkhai    As Variant
    Dim kykkhai    As Variant
    Dim TTHTK1     As Variant
    Dim ROWNUM     As Integer
    Dim DANHAN     As Variant
   
    sSQLCol = "KYLBO, MADTNT, NGNOP, MATHUE, MAPP, NGANHDB, MAMUC, MATM, MACT, STT, STT2, STTIN, DVT, S_LUONG, "
    sSQLCol = sSQLCol + "DOANHTHUBH, THUETINH, THUESUAT, GIATRITT, THUEKT, THUEDC, THUETTDB, TTHTK, MATKHAI, KYKKHAI,LAN_QUET, DANHAN"

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If

    'Du lieu muc I

    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("H"), 7, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        .GetText .ColLetterToNumber("E"), 25, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 23, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If

        MATHUE = "'02'"
        MAPP = "'1'"
        MaMuc = "'1750'"
        DANHAN = "''"
        
        .GetText .ColLetterToNumber("AO"), 3, matkhai
        matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        
        STT = "1"
        NGANHDB = "''"
        maTM = "''"
        mact = "'1000'"
        STT2 = "0"
        STTIN = "'I'"
        DVT = "''"
        
        S_LUONG = "0"

        DOANHTHUBH = "0"
        THUETINH = "0"
        THUESUAT = "0"
        GIATRITT = "0"
        THUEKT = "0"
        THUEDC = "0"
        THUETTDB = "0"
        
    End With
 
    sSQLVal = KYLBO & "," & madtnt & "," & ngnop & "," & MATHUE & "," & MAPP & "," & NGANHDB & "," & MaMuc & "," & maTM & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & DVT & "," & S_LUONG & "," & DOANHTHUBH & "," & THUETINH & "," & THUESUAT & "," & GIATRITT & "," & THUEKT & "," & THUEDC & "," & THUETTDB & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN

    sSQL = "INSERT INTO TMP_TB" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    bln = clsDAO.ExecuteDLL(sSQL)

    'Du lieu muc con cua I

    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("H"), 7, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        .GetText .ColLetterToNumber("E"), 25, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 23, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If

        MATHUE = "'02'"
        MAPP = "'1'"
        MaMuc = "'1750'"
        .GetText .ColLetterToNumber("AO"), 3, matkhai
        matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        
        .Col = .ColLetterToNumber("B")
        .Row = 53
             
        Do
            .Col = .ColLetterToNumber("AO")
            maTM = .Text

            If Trim(maTM) = vbNullString Then
                maTM = "''"
            Else
                maTM = "'" & maTM & "'"
            End If
   
            .Col = .ColLetterToNumber("AP")
            NGANHDB = .Text

            If Trim(NGANHDB) = vbNullString Then
                NGANHDB = "''"
            Else
                NGANHDB = "'" & NGANHDB & "'"
            End If
                    
            mact = "''"
            STT = "1"
            STT2 = STT2 + 1
            STTIN = "'" & STT2 & "'"
                    
            .Col = .ColLetterToNumber("O")
            DVT = .Text

            If Trim(DVT) = vbNullString Then
                DVT = "''"
            Else
                DVT = "'" & TAX_Utilities_Svr_New.Convert(CStr(DVT), UNICODE, TCVN) & "'"
            End If
                    
            .Col = .ColLetterToNumber("R")
            S_LUONG = .Value

            If Trim(S_LUONG) = vbNullString Then
                S_LUONG = "0"
            End If
                    
            .Col = .ColLetterToNumber("W")

            DOANHTHUBH = .Value

            If Trim(DOANHTHUBH) = vbNullString Then

                DOANHTHUBH = "0"
            End If
                    
            .Col = .ColLetterToNumber("AB")
            GIATRITT = .Value

            If Trim(GIATRITT) = vbNullString Then
                GIATRITT = "0"
            End If
                    
            .Col = .ColLetterToNumber("AD")
            THUESUAT = .Value

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            End If
                    
            .Col = .ColLetterToNumber("AG")
            THUEKT = .Value

            If Trim(THUEKT) = vbNullString Then
                THUEKT = "0"
            End If
                    
            .Col = .ColLetterToNumber("AJ")
            THUEDC = .Value

            If Trim(THUEDC) = vbNullString Then
                THUEDC = "0"
            End If
                    
            .Col = .ColLetterToNumber("AN")
            THUETTDB = .Value

            If Trim(THUETTDB) = vbNullString Then
                THUETTDB = "0"
            End If

            THUETINH = CDbl(GIATRITT) * CDbl(THUESUAT) / 100
                    
            sSQLVal = KYLBO & "," & madtnt & "," & ngnop & "," & MATHUE & "," & MAPP & "," & NGANHDB & "," & MaMuc & "," & maTM & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & DVT & "," & S_LUONG & "," & DOANHTHUBH & "," & THUETINH & "," & THUESUAT & "," & GIATRITT & "," & THUEKT & "," & THUEDC & "," & THUETTDB & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
                
            sSQL = "INSERT INTO TMP_TB" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                    
            'ktra neu dong trong thi khong ghi vao db
            If NGANHDB <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

            'end
                    
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"

        ROWNUM = .Row

    End With
   
    'Du lieu ghi cua muc 2

    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("H"), 7, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        .GetText .ColLetterToNumber("E"), 25, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 23, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If

        MATHUE = "'02'"
        MAPP = "'1'"
        MaMuc = "'1750'"
        
        .GetText .ColLetterToNumber("AO"), 3, matkhai
        matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        STT = "2"
        NGANHDB = "''"
        maTM = "''"
        mact = "'2000'"
        STT2 = "0"
        STTIN = "'II'"
        DVT = "''"
        
        S_LUONG = "0"

        DOANHTHUBH = "0"
        THUETINH = "0"
        THUESUAT = "0"
        GIATRITT = "0"
        THUEKT = "0"
        THUEDC = "0"
        THUETTDB = "0"
      
    End With
 
    sSQLVal = KYLBO & "," & madtnt & "," & ngnop & "," & MATHUE & "," & MAPP & "," & NGANHDB & "," & MaMuc & "," & maTM & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & DVT & "," & S_LUONG & "," & DOANHTHUBH & "," & THUETINH & "," & THUESUAT & "," & GIATRITT & "," & THUEKT & "," & THUEDC & "," & THUETTDB & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN

    sSQL = "INSERT INTO TMP_TB" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    bln = clsDAO.ExecuteDLL(sSQL)
  
    ' Ghi du lieu muc con cua II
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("H"), 7, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        .GetText .ColLetterToNumber("E"), 25, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 23, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If

        MATHUE = "'02'"
        MAPP = "'1'"
        MaMuc = "'1750'"
        .GetText .ColLetterToNumber("AO"), 3, matkhai
        matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        
        .Col = .ColLetterToNumber("B")
        .Row = ROWNUM + 3
        STT2 = 0

        Do
            .Col = .ColLetterToNumber("AO")
            maTM = .Text

            If Trim(maTM) = vbNullString Then
                maTM = "''"
            Else
                maTM = "'" & maTM & "'"
            End If
            
            .Col = .ColLetterToNumber("AP")
            NGANHDB = .Text

            If Trim(NGANHDB) = vbNullString Then
                NGANHDB = "''"
            Else
                NGANHDB = "'" & NGANHDB & "'"
            End If
                    
            mact = "''"
            STT = "2"
            STT2 = STT2 + 1
            STTIN = "'" & STT2 & "'"
                    
            .Col = .ColLetterToNumber("O")
            DVT = .Text

            If Trim(DVT) = vbNullString Then
                DVT = "''"
            Else
                DVT = "'" & TAX_Utilities_Svr_New.Convert(CStr(DVT), UNICODE, TCVN) & "'"
            End If
                    
            .Col = .ColLetterToNumber("R")
            S_LUONG = .Value

            If Trim(S_LUONG) = vbNullString Then
                S_LUONG = "0"
            End If
                    
            .Col = .ColLetterToNumber("W")

            DOANHTHUBH = .Value

            If Trim(DOANHTHUBH) = vbNullString Then

                DOANHTHUBH = "0"
            End If
                    
            .Col = .ColLetterToNumber("AB")
            GIATRITT = .Value

            If Trim(GIATRITT) = vbNullString Then
                GIATRITT = "0"
            End If
                    
            .Col = .ColLetterToNumber("AD")
            THUESUAT = .Value

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            End If
                    
            .Col = .ColLetterToNumber("AG")
            THUEKT = .Value

            If Trim(THUEKT) = vbNullString Then
                THUEKT = "0"
            End If
                    
            .Col = .ColLetterToNumber("AJ")
            THUEDC = .Value

            If Trim(THUEDC) = vbNullString Then
                THUEDC = "0"
            End If
                    
            .Col = .ColLetterToNumber("AN")
            THUETTDB = .Value

            If Trim(THUETTDB) = vbNullString Then
                THUETTDB = "0"
            End If

            THUETINH = CDbl(GIATRITT) * CDbl(THUESUAT) / 100
                    
            sSQLVal = KYLBO & "," & madtnt & "," & ngnop & "," & MATHUE & "," & MAPP & "," & NGANHDB & "," & MaMuc & "," & maTM & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & DVT & "," & S_LUONG & "," & DOANHTHUBH & "," & THUETINH & "," & THUESUAT & "," & GIATRITT & "," & THUEKT & "," & THUEDC & "," & THUETTDB & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
                
            sSQL = "INSERT INTO TMP_TB" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                    
            'ktra neu dong trong thi khong ghi vao db
            If NGANHDB <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

            'end

            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"

        ROWNUM = .Row

    End With

    'Ghi du lieu muc III
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("H"), 7, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        .GetText .ColLetterToNumber("E"), 25, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 23, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If

        MATHUE = "'02'"
        MAPP = "'1'"
        MaMuc = "'1750'"
        
        .GetText .ColLetterToNumber("AO"), 3, matkhai
        matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        STT = "3"
        NGANHDB = "''"
        maTM = "''"
        mact = "'3000'"
        STT2 = "0"
        STTIN = "'III'"
        DVT = "''"
        
        S_LUONG = "0"

        DOANHTHUBH = "0"
        THUETINH = "0"
        THUESUAT = "0"
        GIATRITT = "0"
        THUEKT = "0"
        THUEDC = "0"
        THUETTDB = "0"
    End With

    sSQLVal = KYLBO & "," & madtnt & "," & ngnop & "," & MATHUE & "," & MAPP & "," & NGANHDB & "," & MaMuc & "," & maTM & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & DVT & "," & S_LUONG & "," & DOANHTHUBH & "," & THUETINH & "," & THUESUAT & "," & GIATRITT & "," & THUEKT & "," & THUEDC & "," & THUETTDB & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN

    sSQL = "INSERT INTO TMP_TB" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    bln = clsDAO.ExecuteDLL(sSQL)

    'Ghi du lieu muc con cua III 1
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("H"), 7, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        .GetText .ColLetterToNumber("E"), 25, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 23, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If

        MATHUE = "'02'"
        MAPP = "'1'"
        MaMuc = "'1750'"
        
        .GetText .ColLetterToNumber("AO"), 3, matkhai
        matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        'dhdang sua ghi STT
        'ngay 28/10
        STT = "4"
        NGANHDB = "''"
        maTM = "''"
        mact = "'3001'"
        STT2 = "0"
        STTIN = "'1'"
        DVT = "''"
        S_LUONG = "0"

        DOANHTHUBH = "0"
        THUETINH = "0"
        THUESUAT = "0"
        GIATRITT = "0"
        THUEKT = "0"
        THUEDC = "0"
        THUETTDB = "0"
                
        sSQLVal = KYLBO & "," & madtnt & "," & ngnop & "," & MATHUE & "," & MAPP & "," & NGANHDB & "," & MaMuc & "," & maTM & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & DVT & "," & S_LUONG & "," & DOANHTHUBH & "," & THUETINH & "," & THUESUAT & "," & GIATRITT & "," & THUEKT & "," & THUEDC & "," & THUETTDB & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
        
        sSQL = "INSERT INTO TMP_TB" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        .Col = .ColLetterToNumber("B")
        .Row = ROWNUM + 5
        STT2 = 0

        '--------------------------------------------
        Do

            .Col = .ColLetterToNumber("AO")
            maTM = .Text

            If Trim(maTM) = vbNullString Then
                maTM = "''"
            Else
                maTM = "'" & maTM & "'"
            End If
                    
            .Col = .ColLetterToNumber("AP")
            NGANHDB = .Text

            If Trim(NGANHDB) = vbNullString Then
                NGANHDB = "''"
            Else
                NGANHDB = "'" & NGANHDB & "'"
            End If
                    
            mact = "''"
            STT = "4"
            STT2 = STT2 + 1
            STTIN = "''"
                    
            .Col = .ColLetterToNumber("O")
            DVT = .Text

            If Trim(DVT) = vbNullString Then
                DVT = "''"
            Else
                DVT = "'" & TAX_Utilities_Svr_New.Convert(CStr(DVT), UNICODE, TCVN) & "'"
            End If
                    
            .Col = .ColLetterToNumber("R")
            S_LUONG = .Value

            If Trim(S_LUONG) = vbNullString Then
                S_LUONG = "0"
            End If
                    
            .Col = .ColLetterToNumber("W")

            DOANHTHUBH = .Value

            If Trim(DOANHTHUBH) = vbNullString Then

                DOANHTHUBH = "0"
            End If

            GIATRITT = "0"
            THUESUAT = "0"
            THUEKT = "0"
            THUEDC = "0"
            THUETTDB = "0"
                    
            sSQLVal = KYLBO & "," & madtnt & "," & ngnop & "," & MATHUE & "," & MAPP & "," & NGANHDB & "," & MaMuc & "," & maTM & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & DVT & "," & S_LUONG & "," & DOANHTHUBH & "," & THUETINH & "," & THUESUAT & "," & GIATRITT & "," & THUEKT & "," & THUEDC & "," & THUETTDB & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
                
            sSQL = "INSERT INTO TMP_TB" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            'ktra neu dong trong thi khong ghi vao db
            If NGANHDB <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

            'end

            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"

        ROWNUM = .Row

    End With
        
    'Ghi du lieu muc con III 2
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("H"), 7, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        .GetText .ColLetterToNumber("E"), 25, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 23, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If

        MATHUE = "'02'"
        MAPP = "'1'"
        MaMuc = "'1750'"
        .GetText .ColLetterToNumber("AO"), 3, matkhai
        matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        'dhdang sua ghi STT
        'ngay 28/10
        STT = "5"
        NGANHDB = "''"
        maTM = "''"
        mact = "'3002'"
        STT2 = "0"
        STTIN = "'2'"
        DVT = "''"
        S_LUONG = "0"

        DOANHTHUBH = "0"
        THUETINH = "0"
        THUESUAT = "0"
        GIATRITT = "0"
        THUEKT = "0"
        THUEDC = "0"
        THUETTDB = "0"
                
        sSQLVal = KYLBO & "," & madtnt & "," & ngnop & "," & MATHUE & "," & MAPP & "," & NGANHDB & "," & MaMuc & "," & maTM & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & DVT & "," & S_LUONG & "," & DOANHTHUBH & "," & THUETINH & "," & THUESUAT & "," & GIATRITT & "," & THUEKT & "," & THUEDC & "," & THUETTDB & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
        
        sSQL = "INSERT INTO TMP_TB" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
            
        .Col = .ColLetterToNumber("B")
        .Row = ROWNUM + 3
        STT2 = 0

        '--------------------------------------------
        Do
            .Col = .ColLetterToNumber("AO")
            maTM = .Text

            If Trim(maTM) = vbNullString Then
                maTM = "''"
            Else
                maTM = "'" & maTM & "'"
            End If
                    
            .Col = .ColLetterToNumber("AP")
            NGANHDB = .Text

            If Trim(NGANHDB) = vbNullString Then
                NGANHDB = "''"
            Else
                NGANHDB = "'" & NGANHDB & "'"
            End If
                    
            mact = "''"
            STT = "5"
            STT2 = STT2 + 1
            STTIN = "''"
                    
            .Col = .ColLetterToNumber("O")
            DVT = .Text

            If Trim(DVT) = vbNullString Then
                DVT = "''"
            Else
                DVT = "'" & TAX_Utilities_Svr_New.Convert(CStr(DVT), UNICODE, TCVN) & "'"
            End If
                    
            .Col = .ColLetterToNumber("R")
            S_LUONG = .Value

            If Trim(S_LUONG) = vbNullString Then
                S_LUONG = "0"
            End If
                    
            .Col = .ColLetterToNumber("W")

            DOANHTHUBH = .Value

            If Trim(DOANHTHUBH) = vbNullString Then

                DOANHTHUBH = "0"
            End If

            GIATRITT = "0"
            THUESUAT = "0"
            THUEKT = "0"
            THUEDC = "0"
            THUETTDB = "0"
                    
            sSQLVal = KYLBO & "," & madtnt & "," & ngnop & "," & MATHUE & "," & MAPP & "," & NGANHDB & "," & MaMuc & "," & maTM & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & DVT & "," & S_LUONG & "," & DOANHTHUBH & "," & THUETINH & "," & THUESUAT & "," & GIATRITT & "," & THUEKT & "," & THUEDC & "," & THUETTDB & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
                
            sSQL = "INSERT INTO TMP_TB" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            'ktra neu dong trong thi khong ghi vao db
            If NGANHDB <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

            'end

            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"

        ROWNUM = .Row

    End With

    'Ghi du lieu muc con cua III 3
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("H"), 7, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        .GetText .ColLetterToNumber("E"), 25, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 23, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If

        MATHUE = "'02'"
        MAPP = "'1'"
        MaMuc = "'1750'"
        .GetText .ColLetterToNumber("AO"), 3, matkhai
        matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        'dhdang sua ghi STT
        'ngay 28/10
        STT = "6"
        NGANHDB = "''"
        maTM = "''"
        mact = "'3003'"
        STT2 = "0"
        STTIN = "'3'"
        DVT = "''"
        S_LUONG = "0"

        DOANHTHUBH = "0"
        THUETINH = "0"
        THUESUAT = "0"
        GIATRITT = "0"
        THUEKT = "0"
        THUEDC = "0"
        THUETTDB = "0"
        sSQLVal = KYLBO & "," & madtnt & "," & ngnop & "," & MATHUE & "," & MAPP & "," & NGANHDB & "," & MaMuc & "," & maTM & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & DVT & "," & S_LUONG & "," & DOANHTHUBH & "," & THUETINH & "," & THUESUAT & "," & GIATRITT & "," & THUEKT & "," & THUEDC & "," & THUETTDB & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
        
        sSQL = "INSERT INTO TMP_TB" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        '--------------------------------------------
            
        .Col = .ColLetterToNumber("B")
        .Row = ROWNUM + 3
        STT2 = 0

        Do
            .Col = .ColLetterToNumber("AO")
            maTM = .Text

            If Trim(maTM) = vbNullString Then
                maTM = "''"
            Else
                maTM = "'" & maTM & "'"
            End If
   
            .Col = .ColLetterToNumber("AP")
            NGANHDB = .Text

            If Trim(NGANHDB) = vbNullString Then
                NGANHDB = "''"
            Else
                NGANHDB = "'" & NGANHDB & "'"
            End If
                    
            mact = "''"
            STT = "6"
            STT2 = STT2 + 1
            STTIN = "''"
                    
            .Col = .ColLetterToNumber("O")
            DVT = .Text

            If Trim(DVT) = vbNullString Then
                DVT = "''"
            Else
                DVT = "'" & TAX_Utilities_Svr_New.Convert(CStr(DVT), UNICODE, TCVN) & "'"
            End If
                    
            .Col = .ColLetterToNumber("R")
            S_LUONG = .Value

            If Trim(S_LUONG) = vbNullString Then
                S_LUONG = "0"
            End If
                    
            .Col = .ColLetterToNumber("W")

            DOANHTHUBH = .Value

            If Trim(DOANHTHUBH) = vbNullString Then

                DOANHTHUBH = "0"
            End If

            GIATRITT = "0"
            THUESUAT = "0"
            THUEKT = "0"
            THUEDC = "0"
            THUETTDB = "0"
                    
            sSQLVal = KYLBO & "," & madtnt & "," & ngnop & "," & MATHUE & "," & MAPP & "," & NGANHDB & "," & MaMuc & "," & maTM & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & DVT & "," & S_LUONG & "," & DOANHTHUBH & "," & THUETINH & "," & THUESUAT & "," & GIATRITT & "," & THUEKT & "," & THUEDC & "," & THUETTDB & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN
                
            sSQL = "INSERT INTO TMP_TB" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            'ktra neu dong trong thi khong ghi vao db
            If NGANHDB <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

            'end

            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"

        ROWNUM = .Row

    End With

    clsDAO.Disconnect
   
    If TAX_Utilities_Svr_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL1
    End If

    If TAX_Utilities_Svr_New.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_PL2
    End If

    If TAX_Utilities_Svr_New.NodeValidity.childNodes(3).Attributes.getNamedItem("Active").nodeValue = 1 Then
        InsertDTL = InsertDTL_KHBS
    End If

    InsertDTL = vbNullString
   
End Function

Public Function Prepared3() As Boolean
    With fps
        .Sheet = 1
        'set check loi dinh danh
        If checkDLTTonTai = False Then
            .Col = .ColLetterToNumber("F")
            .Row = 27
            .Value = 1
        End If
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
        
        ' vttoan: them thong tin dai ly thue
        'ten
        .Col = .ColLetterToNumber(TEN_GOI_DL_X)
        .Row = TEN_GOI_DL_Y
        If strTen_DLT <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strTen_DLT, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        'mst
        .Col = .ColLetterToNumber(MA_SO_THUE_DL_X)
        .Row = MA_SO_THUE_DL_Y
        If strMST_DLT <> "" Then
            .Text = strMST_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        'dia chi
        .Col = .ColLetterToNumber(DIA_CHI_DL_X)
        .Row = DIA_CHI_DL_Y
        If strDchi_DLT <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strDchi_DLT, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
'        ' quan huyen
'        .Col = .ColLetterToNumber(QUAN_HUYEN_DL_X)
'        .Row = QUAN_HUYEN_DL_Y
'        If strQHuyen_DLT <> "" Then
'            .Text = strQHuyen_DLT
'        End If
'        ' thanh pho
'        .Col = .ColLetterToNumber(TINH_TPHO_DL_X)
'        .Row = TINH_TPHO_DL_Y
'        If strTTPho_DLT <> "" Then
'            .Text = strTTPho_DLT
'        End If
        'dien thoai
        .Col = .ColLetterToNumber(DIEN_THOAI_DL_X)
        .Row = DIEN_THOAI_DL_Y
        If strDthoai_DLT <> vbNullString Then
            .Text = strDthoai_DLT
        End If
        'fax
        .Col = .ColLetterToNumber(FAX_DL_X)
        .Row = FAX_DL_Y
        If strFax_DLT <> "" Then
            .Text = strFax_DLT
        End If
        ' mail
        .Col = .ColLetterToNumber(MAIL_DL_X)
        .Row = MAIL_DL_Y
        If strMail_DLT <> "" Then
            .Text = strMail_DLT
        End If
        ' so hop dong
        .Col = .ColLetterToNumber(SO_HDDL_DL_X)
        .Row = SO_HDDL_DL_Y
        If strSoHD_DLT <> "" Then
            .Text = strSoHD_DLT
        End If
        ' ngay hop dong
        .Col = .ColLetterToNumber(NGAY_HDDL_DL_X)
        .Row = NGAY_HDDL_DL_Y
        If strNgayHD_DLT <> "" Then
            .Text = strNgayHD_DLT
        End If
        'so thu tu to khai
        .Col = .ColLetterToNumber("R")
        .Row = 23
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
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
        
        If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
            blnValid = False
        End If
        
        .EventEnabled(EventAllEvents) = True
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
    
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("I"), 14, varCheckValue
        .GetText .ColLetterToNumber("T"), 14, varNoteValue
        If CStr(varCheckValue) = "1" And Trim(CStr(varNoteValue)) = "" Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 14, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 14, "1"
        End If
        
    End With
End Function

Public Function GenerateSQL_Header(xmlDomData As MSXML.DOMDocument, strSQL_HDR As String, vHdrID As Variant, ByVal dNgayDauKy As Date) As String
'    Dim xmlList As MSXML.IXMLDOMNodeList
'    Dim xmlNode As MSXML.IXMLDOMNode
'    Dim xmlAttribute As MSXML.IXMLDOMAttribute
'    Dim iRowID As Long, strSQL As String, strTempSQL As String
'    Dim dDate As Date, strDate() As String
'    Dim vTIN, vTEN_DTNT, vDIA_CHI, vLOAI_TKHAI, vNGAY_NOP, vKyLB
'    Dim vKYKK, vNGAY_CAP_NHAT, vNGUOI_CAP_NHAT, vCO_LOI_DDANH
'    Dim vSO_HIEU_TEP, vSO_TT_TK, vDA_NHAN, vGHI_CHU_LOI, vCO_GTRINH_02A
'    Dim vCO_GTRINH_02B, vCO_GTRINH_02C
'    Dim vPHONG_XU_LY
'    Dim i As Long, j As Long
'    Dim strMaPhongXuLy As String
'
'On Error GoTo ErrHandle
'    strSQL = strSQL_HDR
'    Set xmlList = xmlDOMdata.getElementsByTagName("Cell")
'    For Each xmlNode In xmlList
'        With xmlNode.Attributes
'
'        If Trim(GetAttribute(xmlNode, "MCT")) = vbNullString Then
'            Select Case Trim(GetAttribute(xmlNode, "CellID"))
'                Case "H_7"
'                    vTIN = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
'                Case "H_5"
'                    vTEN_DTNT = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
'                Case "H_9"
'                    vDIA_CHI = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
'                Case "E_25"
'                    vNGAY_NOP = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
'                Case "E_23"
'                    vKyLB = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
'                Case "R_25"
'                    vNGAY_CAP_NHAT = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
'                Case "AF_23"
'                    vNGUOI_CAP_NHAT = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
'                Case "F_27"
'                    vCO_LOI_DDANH = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
'                Case "E_29"
'                    vGHI_CHU_LOI = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
'                Case PHONG_XU_LY_X & "_" & PHONG_XU_LY_Y
'                    vPHONG_XU_LY = Trim(TAX_Utilities_Svr_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
'            End Select
'        End If
'        End With
'    Next
'
'    strSQL = strSQL & "'" & vHdrID & "',"
'    strSQL = strSQL & "'" & vTIN & "',"
'    strSQL = strSQL & "'" & vTEN_DTNT & "',"
'    strSQL = strSQL & "'" & vDIA_CHI & "',"
'    vLOAI_TKHAI = TAX_Utilities_Svr_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue
'    strSQL = strSQL & "'" & vLOAI_TKHAI & "',"
'    strSQL = strSQL & "To_date('" & vNGAY_NOP & "','dd/mm/yyyy'),"
'
'    'Ky/Quy LB
'    If Trim(TAX_Utilities_Svr_New.Month) <> "" Then
'        'Ngay dau ky lap bo va ngay cuoi ky lap bo
'        strDate = Split(vKyLB, "/")
'        dDate = DateSerial(Val(strDate(1)), Val(strDate(0)), 1)
'        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
'        dDate = DateAdd("m", 1, dDate)
'        dDate = DateAdd("d", -1, dDate)
'        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
'    End If
'
'    'Ky/ Quy KK
'    If Trim(TAX_Utilities_Svr_New.Month) <> "" Then
'        'Ngay dau ky ke khai va ngay cuoi ky ke khai
'        'strDate = Split(TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year, "/")
'        'dDate = DateSerial(Val(strDate(1)), Val(strDate(0)), 1)
'        dDate = dNgayDauKy
'        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
'        dDate = DateAdd("m", 1, dDate)
'        dDate = DateAdd("d", -1, dDate)
'        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
'    End If
'
'    strSQL = strSQL & "To_date('" & vNGAY_CAP_NHAT & "','dd/mm/yyyy'),"
'    strSQL = strSQL & "'" & vNGUOI_CAP_NHAT & "',"
'    strSQL = strSQL & "'" & vCO_LOI_DDANH & "',"
'    strSQL = strSQL & "'" & vSO_HIEU_TEP & "',"
'    strSQL = strSQL & "'" & vSO_TT_TK & "',"
'
'    strSQL = strSQL & "'" & vDA_NHAN & "',"
'    strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
'    strSQL = strSQL & "'',"
'    strSQL = strSQL & "'',"
'    strSQL = strSQL & "'',"
'    strSQL = strSQL & "'','',"
'
'    With fps
'        For i = 1 To lSoPhongXL
'            If vPHONG_XU_LY = TAX_Utilities_Svr_New.Convert(larrPhongXuLy(i), UNICODE, TCVN) Then
'                strMaPhongXuLy = larrid(i)
'                Exit For
'            End If
'        Next
'    End With
'    strSQL = strSQL & "'" & strMaPhongXuLy & "')"
'    GenerateSQL_Header = strSQL
'
'    Exit Function
'ErrHandle:
'    SaveErrorLog "cls001", "GenerateSQL_Header", Err.Number, Err.Description
End Function

Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
    KiemTraKhoaSo = True
    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Svr_New.Data(0).nodeFromID("E_23"), "Value"), "dd/mm/yyyy")

    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function
Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
 Dim varTemp As Variant
 Dim vNgayQyet As Variant
 With fps
    If .ActiveSheet = 1 Then
        If Col = .ColLetterToNumber("E") And Row = 25 Then
            .GetText Col, Row, varTemp
            If varTemp <> "" And varTemp <> "../../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                    checkNgayNop varTemp, vNgayQyet
                Else
                    .SetActiveCell Col, Row
                End If
            Else
             .SetText Col, Row, ""
            End If
           UpdateCell fps, .Col, .Row, .Text
        End If
        If Col = .ColLetterToNumber("R") And Row = 25 Then
            .GetText Col, Row, varTemp
            If varTemp <> "" And varTemp <> "../...." Then
                If Format_mmyyyy(CStr(varTemp)) <> "" Then
                    .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
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
    Dim sSQL           As String
    Dim rs             As ADODB.Recordset
    Dim fso            As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim LAN_QUET1      As Variant
    Dim matkhai        As Variant
        
    'kiem tra ton tai tep *.dbf chua
    'Tinh so lan quyet de ghi vao DB trung gian
    'dhdang sua
    'ngay 24/09/2010
    '-----------------------------
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("AO"), 3, matkhai

        If Trim(matkhai) = vbNullString Then
            matkhai = "''"
        Else
            matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        End If

    End With

    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"

    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
            clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
            clsDAO.Connect
        End If

        sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_TK" & sKyKeKhai & " WHERE MATKHAI = " & matkhai & " AND  MAPP='1' AND MADTNT= " & "'" & strMST & "'" & " AND KYKKHAI ='" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND (TTHTK = '1' OR TTHTK = '3' OR TTHTK = '4' or TTHTK = '2')"
                 
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
Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim vKeKhai As Variant
    Dim vKeKhai_temp As Variant
    Dim vKyLB As Variant
    Dim strTestFileName As String
    Dim fso As New FileSystemObject
    Dim matkhai As Variant
     
    Dim flag As Boolean
    'kiem tra da co to khai chinh thuc chua
    flag = True
    vKyLB = DateSerial(Int(Year(Now())), Int(Month(Now())), 1)
    vKeKhai = DateSerial(Int(Val(TAX_Utilities_Svr_New.Year)), Int(Val(TAX_Utilities_Svr_New.Month)), 1)
    vKeKhai_temp = sKyKeKhai
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   'get MATK
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("AO"), 3, matkhai
        If Trim(matkhai) = vbNullString Then
                matkhai = "''"
        Else
                matkhai = "'" & TAX_Utilities_Svr_New.Convert(CStr(matkhai), UNICODE, TCVN) & "'"
        End If
   End With
   
   Do While flag
   
        strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & vKeKhai_temp & ".DBF"
        'kiem tra neu co db thi moi thuc hien truy van
        If fso.FileExists(strTestFileName) Then
            sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & vKeKhai_temp & _
            " WHERE MATHUE = '02' AND MADTNT = " & "'" & strMST & "'" & _
            " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'" & _
            " AND (TTHTK = '1' or TTHTK = '3' or TTHTK = '4') AND MATKHAI = " & matkhai
            Set rs = clsDAO.Execute(sSQL)
        End If
        
        If rs Is Nothing Then
             isToKhaiChinhThuc = False
        Else
             isToKhaiChinhThuc = True
             Exit Do
        End If
        
        If vKeKhai < vKyLB Then
            vKyLB = DateAdd("M", -1, vKyLB)
            ' ghep ten bang de kiem tra du lieu
            If Len(Month(vKyLB)) = 1 Then
                vKeKhai_temp = "0" & Month(vKyLB) & Year(vKyLB)
            Else
                vKeKhai_temp = Month(vKyLB) & Year(vKyLB)
            End If
            flag = True
        Else
            flag = False
        End If
   Loop
   clsDAO.Disconnect
End Function
Private Function checkNgayNop(ByVal vNgayNop As Variant, ByVal vNgayQuet As Variant)
    Dim varNgayNop As Variant
    Dim varNgayQuet As Variant
    On Error GoTo e
    
    varNgayNop = DateSerial(Int(Right(CStr(vNgayNop), 4)), Int(Mid(CStr(vNgayNop), 4, 2)), Int(Left(CStr(vNgayNop), 2)))
    varNgayQuet = DateSerial(Int(Right(CStr(vNgayQuet), 4)), Int(Mid(CStr(vNgayQuet), 4, 2)), Int(Left(CStr(vNgayQuet), 2)))

    If Year(vNgayNop) > Year(vNgayQuet) Then
        GoTo e
    ElseIf Month(vNgayNop) > Month(vNgayQuet) Then
        GoTo e
    ElseIf Day(vNgayNop) > Day(vNgayQuet) Then
        GoTo e
    End If
    Exit Function
e:
    DisplayMessage "0144", msOKOnly, miCriticalError

End Function

