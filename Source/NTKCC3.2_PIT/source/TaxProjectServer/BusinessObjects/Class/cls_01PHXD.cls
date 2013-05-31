VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01PHXD"
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
Const PHONG_XU_LY_Y = 30
Const PHONG_XU_LY_X = "V"
Const NGAY_NOP_Y = 25
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 25
Const NGAY_QUET_X = "R"
'Const TEN_BPQL_X = "V"
'Const TEN_BPQL_Y = 34
Const MA_BPQL_X = "V"
Const MA_BPQL_Y = 34
Private rs As ADODB.Recordset

' thong tin din vi
Const MA_SO_THUE_X = "F"
Const MA_SO_THUE_Y = 10
Const TEN_GOI_X = "H"
Const TEN_GOI_Y = 8
Const DIA_CHI_X = "F"
Const DIA_CHI_Y = 12
'Const QUAN_HUYEN_X = "F"
'Const QUAN_HUYEN_Y = 11
'Const TINH_TPHO_X = "T"
'Const TINH_TPHO_Y = 11
Const DIEN_THOAI_X = "F"
Const DIEN_THOAI_Y = 16
Const FAX_X = "N"
Const FAX_Y = 16
Const MAIL_X = "V"
Const MAIL_Y = 16

'thong tin Dai ly thue
Const MA_SO_THUE_DL_X = "F"
Const MA_SO_THUE_DL_Y = 20
Const TEN_GOI_DL_X = "H"
Const TEN_GOI_DL_Y = 18
Const DIA_CHI_DL_X = "F"
Const DIA_CHI_DL_Y = 22
Const QUAN_HUYEN_DL_X = "F"
Const QUAN_HUYEN_DL_Y = 24
Const TINH_TPHO_DL_X = "Q"
Const TINH_TPHO_DL_Y = 24
Const DIEN_THOAI_DL_X = "F"
Const DIEN_THOAI_DL_Y = 26
Const FAX_DL_X = "N"
Const FAX_DL_Y = 26
Const MAIL_DL_X = "V"
Const MAIL_DL_Y = 26
Const SO_HDDL_DL_X = "H"
Const SO_HDDL_DL_Y = 28
Const NGAY_HDDL_DL_X = "N"
Const NGAY_HDDL_DL_Y = 28
'end

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
Public isExistFile As Boolean
Public LAN_QUET As Variant
Const STT_X = "M"
Const STT_Y = 30
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
Public tthtk As String
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
        .Col = .ColLetterToNumber("M")
        .Row = 36
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100

        ' Ky lap bo
        SetDateFormat fps, 1, 30, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft

        'Phong xu ly
        .Col = .ColLetterToNumber("V")
        .Row = 30
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60

        ' Ngay nop
        SetDateFormat fps, 1, 32, .ColLetterToNumber("E"), DDMMYYYY
        .Sheet = 1
        .Row = 32
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft

        ' Ngay quet
        SetDateFormat fps, 1, 32, .ColLetterToNumber("M"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 32
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
    Dim iCol As Long, iRow As Long

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
                            If Val(TAX_Utilities_Svr_New.Month) >= 4 And Val(TAX_Utilities_Svr_New.Year) >= 2010 Then
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
    Dim strFileNameBS As String
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_PHXD" & sKyKeKhai & ".DBF"
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
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_PHXD.DBF"
        fso.CopyFile strTepmauDTL, strFileNameDTL, False
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

        
    sSQL = "SELECT MADTNT, NGNOP FROM  TMP_TK" & sKyKeKhai & "  WHERE MADTNT = " & "'" & strMST & "'" & " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'" & " AND MATKHAI = '01/PHXD'"
          
   
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
    Dim kylbo As Variant
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
        .GetText .ColLetterToNumber("E"), 23, kylbo
        If Trim(kylbo) = vbNullString Then
            kylbo = "''"
        Else
            If Len(Trim(kylbo)) = 6 Then
                kylbo = "'0" & kylbo & "'"
            Else
                kylbo = "'" & kylbo & "'"
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
   " AND KYLBO=" & kylbo
   
   
   clsDAO.ExecuteDLL sSQL
   ' Xoa phu luc 1
   
    sSQL = "DELETE FROM TMP_D1" & sKyKeKhai & " WHERE MADTNT = '" & strMST & "'" & _
   " AND KYLBO=" & kylbo
   clsDAO.ExecuteDLL sSQL
   ' Xoa phu luc 2
   sSQL = "DELETE FROM TMP_D2" & sKyKeKhai & " WHERE MADTNT = '" & strMST & "'" & _
   " AND KYLBO=" & kylbo
   clsDAO.ExecuteDLL sSQL
   
   XoaTKTT = True
   
   clsDAO.Disconnect

End Function



Public Function TKTTNGNOP() As Boolean

 TKTTNGNOP = True

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
    Dim bln         As Boolean
    
    Dim madtnt     As Variant
    Dim MATHUE     As Variant
    Dim maTM       As Variant
    Dim MAPP       As Variant
    Dim kylbo      As Variant
    Dim ngnop      As Variant
    Dim LoaiTien   As Variant
    Dim MABPQL     As Variant
    Dim SHTEP      As Variant
    Dim kykkhai    As Variant
    Dim NgNhap     As Variant
    Dim ThueTKy    As Variant
    Dim GTTinhThue As Variant
    Dim DaGhi      As Variant
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
    Dim NGANH      As Variant
    Dim DAGHI1     As Variant
    Dim AUTO_CAL   As Variant
    Dim DANHAN     As Variant
    Dim MADLT      As Variant
    Dim LanBS      As Variant
    Dim MANNKD     As Variant
    Dim THUECL     As Variant
    Dim DSoCL      As Variant
    Dim GTTTCL     As Variant
    Dim NGAYPS     As Variant
    Dim DAGHI2     As Variant
    
    
    sSQLCol = "MADTNT, MATHUE, MATM, MAPP, KYLBO, NGNOP, "
    sSQLCol = sSQLCol & "LOAITIEN, MABPQL, SHTEP, TTHTK, KYKKHAI, "
    sSQLCol = sSQLCol & "NGNHAP, THUETKY, GTTINHTHUE, DAGHI, "
    sSQLCol = sSQLCol & "KTTRUOC, DOANHSO, MAPP2, MAMUC, MAVACH,MATKHAI,HANNOP, THUETKY2,HANNOP2,TENDA,NGANH,DAGHI1,DAGHI2,AUTO_CAL,LAN_QUET, DANHAN,MADLT ,LANBS,MANNKD,THUECL,DSOCL,GTTTCL,NGAYPS"
    
        If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
    With fps
        .Sheet = 1
        
        'set DAGHI2 theo checkbox loi dinh danh
        .Col = .ColLetterToNumber("E")
        .Row = 36

        If .Value = 1 Then
            DAGHI2 = ".T."
        Else
            DAGHI2 = ".F."
        End If
        
        'set cac chi so
        
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), 30, kylbo

        If Trim(kylbo) = vbNullString Then
            kylbo = "''"
        Else

            If Len(Trim(kylbo)) = 6 Then
                kylbo = "'0" & kylbo & "'"
            Else
                kylbo = "'" & kylbo & "'"
            End If
        End If

        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
    
        If Trim(MABPQL) = vbNullString Then
            MABPQL = "''"
        Else
            MABPQL = "'" & MABPQL & "'"
        End If
        
        .GetText .ColLetterToNumber("M"), 32, NgNhap

        If Trim(NgNhap) = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If
        
'        .GetText .ColLetterToNumber("T"), 54, ThueTKy
'
'        If Trim(ThueTKy) = "0" Then
'            ThueTKy = "0"
'        End If
        
        .GetText .ColLetterToNumber("F"), 20, MADLT

        If Trim(MADLT) = vbNullString Then
            MADLT = "''"
        Else
            MADLT = "'" & MADLT & "'"
        End If

        '       lay so lan bs
        .GetText .ColLetterToNumber("T"), 6, LanBS

        If Trim(LanBS) = vbNullString Then
            LanBS = "0"
        End If

        If Trim(LAN_QUET) = vbNullString Then
            LAN_QUET = "1"
        End If
        
        ' Han nop: ngay thu 20 cua thang sau
        
        Dim thang As Variant
        Dim nam   As Variant
     
        thang = TAX_Utilities_Svr_New.Month ' Month(Now())
        nam = TAX_Utilities_Svr_New.Year 'Year(Now())

        If thang = 12 Then
            nam = nam + 1
            HANNOP = "20/01/" & nam
            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
        Else
            thang = thang + 1
            HANNOP = "20/" & thang & "/" & nam
            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
        End If

        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        MANNKD = "''"
        GTTinhThue = "0"

        DoanhSo = "0"
        ThueTKy2 = "0"
        HANNOP2 = HANNOP
        TENDA = "''"
        DAGHI1 = ".F."
        AUTO_CAL = ".F."
        THUECL = "0"
        MAPP2 = "'1'"
        DaGhi = ".T."
        KTTRUOC = "0"
        DSoCL = "0"
        GTTTCL = "0"
        NGAYPS = "CTOD('')"
        MATHUE = "'22'"
        MAPP = "'1'"
        DANHAN = "''"
        MaVach = ".T."
        matkhai = "'01/PHXD'"
        LoaiTien = "''"
        MABPQL = strMaBPQL
        SHTEP = "''"
        NGANH = TinhMaNganh
        
        ' sua MAMUC & MATM tu nam 2009
        If CLng(Mid(Right(kylbo, 5), 1, 4)) < 2009 Then
            maTM = "'01'"
            MaMuc = "'014'"
        Else
            MaMuc = "'1700'"
            maTM = "'1701'"
        End If
        
        Dim i As Integer
        
        Dim arrMaTMuc As Variant
        Dim arrThueTKy As Variant
        
        arrMaTMuc = Array("'2101'", "'2102'", "'2105'")
        MaMuc = "'2100'"
        
        Dim ThueTKy1, ThueTKy3, ThueTKy4 As Variant

                .GetText .ColLetterToNumber("T"), 55, ThueTKy1

        If Trim(ThueTKy1) = vbNullString Then
            ThueTKy1 = "0"
        End If
        
        .GetText .ColLetterToNumber("T"), 56, ThueTKy3

        If Trim(ThueTKy3) = vbNullString Then
            ThueTKy3 = "0"
        End If
        
        .GetText .ColLetterToNumber("T"), 57, ThueTKy4

        If Trim(ThueTKy4) = vbNullString Then
            ThueTKy4 = "0"
        End If
        
        arrThueTKy = Array(ThueTKy1, ThueTKy3, ThueTKy4)
        
        For i = 0 To 2
        
             maTM = arrMaTMuc(i)
               ThueTKy = arrThueTKy(i)
             
            sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & kylbo & "," & ngnop & ", "
            sSQLVal = sSQLVal & LoaiTien & ",'" & MABPQL & "'," & SHTEP & "," & tthtk & "," & kykkhai & ", "
            sSQLVal = sSQLVal & NgNhap & "," & ThueTKy & "," & GTTinhThue & "," & DaGhi & ", "
            sSQLVal = sSQLVal & KTTRUOC & "," & DoanhSo & "," & MAPP2 & "," & MaMuc & "," & MaVach & ", "
            sSQLVal = sSQLVal & matkhai & "," & HANNOP & "," & ThueTKy2 & "," & HANNOP2 & "," & TENDA & ",'" & NGANH & "'," & DAGHI1 & "," & DAGHI2 & "," & AUTO_CAL & "," & LAN_QUET & "," & DANHAN & "," & MADLT & "," & LanBS & "," & MANNKD & "," & THUECL & "," & DSoCL & "," & GTTTCL & "," & NGAYPS
       
            sSQL = "INSERT INTO TMP_TK" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
        Next
        
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
    Dim kylbo       As Variant
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
    Dim SONGAYNC    As Variant
    Dim SOTIENNC    As Variant
    Dim DANHAN      As Variant
    
    Dim i           As Integer, iCol As Long, iRow As Long
    Dim xmlNode     As MSXML.IXMLDOMNode

    Dim strFileName As String
    Dim fso         As New FileSystemObject

    'kiem tra ton tai tep *.dbf chua
    strFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"

    If fso.FileExists(strFileName) = False Then
        Dim strTepmau As String
        strTepmau = spathVat & "\tepmau\TMP_BS.DBF"
        fso.CopyFile strTepmau, strFileName, False
    End If

    sSQLCol = "MADTNT, KYKKHAI, MATKHAI, MATHUE, MAMUC, MATM, MAPP, NGNOP, TTHTK, KYLBO , MACT, MACT2, MACT3, STT, STT2, STTIN," & " SOKK, SODC, SOCL, SONGAYNC, SOTIENNC, LAN_QUET, DANHAN"

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If

    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"

        MaMuc = "'2100'"
        MATHUE = "' '"
        MAPP = "'1'"
        matkhai = "'01/PHXD'"
        maTM = "' '"
        DANHAN = "''"
        
        .GetText .ColLetterToNumber("E"), 32, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 30, kylbo

        If Trim(kylbo) = vbNullString Then
            kylbo = "CTOD('')"
        Else
            kylbo = "'" & kylbo & "'"
        End If

        ' Insert so dieu chinh tang

        mact = "''"
        MACT2 = "' '"
        MACT3 = "''"
        STT = 1
        STT2 = 0
        STTIN = "'I'"
        SOKK = "0"
        SODC = "0"
        SOCL = "0"
        SONGAYNC = "0"
        SOTIENNC = "0"

        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SONGAYNC & "," & SOTIENNC & "," & LAN_QUET & "," & DANHAN
                      
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

        .Sheet = .SheetCount - 1
        ' so ngay nop cham
        .GetText .ColLetterToNumber("BE"), 17, SONGAYNC

        If Trim(SONGAYNC) = vbNullString Then
            SONGAYNC = "0"
        Else
            SONGAYNC = SONGAYNC
        End If

        ' so tien nop cham
        .GetText .ColLetterToNumber("BE"), 18, SOTIENNC

        If Trim(SOTIENNC) = vbNullString Then
            SOTIENNC = "0"
        Else
            SOTIENNC = SOTIENNC
        End If
        
        i = 1
        .Row = 9

        Do
            mact = "''"
            .GetText .ColLetterToNumber("BE"), .Row, MACT2

            If Trim(MACT2) = vbNullString Then
                MACT2 = "' '"
                
            Else
                MACT2 = "'" & MACT2 & "'"
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

            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SONGAYNC & "," & SOTIENNC & "," & LAN_QUET & "," & DANHAN
                      
            sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If MACT2 <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "aa"

        ' Insert so dieu chinh giam

        MACT2 = "' '"
        MACT3 = "''"
        STT = 2
        STT2 = 0
        STTIN = "'II'"
        SOKK = "0"
        SODC = "0"
        SOCL = "0"

        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SONGAYNC & "," & SOTIENNC & "," & LAN_QUET & "," & DANHAN
                      
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)

        i = 1
        .Row = .Row + 3

        Do
            mact = "''"

            If Trim(MACT2) = vbNullString Then
                MACT2 = "' '"
                
            Else
                MACT2 = "'" & MACT2 & "'"
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

            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SONGAYNC & "," & SOTIENNC & "," & LAN_QUET & "," & DANHAN
                      
            sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If MACT2 <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
        Loop Until .Text = "bb"

        MACT3 = "''"
        STT = 0
        STT2 = 0
        STTIN = "' '"
        mact = "' '"
          
        ' Thue  GTGT con phai nop trong ky
        .Row = 1 + .Row
        
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
        
        MACT2 = "'40'"

        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SONGAYNC & "," & SOTIENNC & "," & LAN_QUET & "," & DANHAN
                      
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        
        ' Thue GTGT con duoc khau tru chuyen ky sau
        .Row = 1 + .Row
        MACT2 = "'45'"
        
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
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & tthtk & "," & kylbo & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SONGAYNC & "," & SOTIENNC & "," & LAN_QUET & "," & DANHAN
                      
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
    End With

    clsDAO.Disconnect
    InsertDTL_KHBS = vbNullString
   
End Function

Public Function InsertDTL() As String
    Dim sSQL    As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim bln     As Boolean
   
    Dim matkhai As Variant
    Dim madtnt  As Variant
    Dim kylbo   As Variant
    Dim kykkhai As Variant
    Dim ngnop   As Variant
    Dim mact    As Variant
    Dim donvi   As Variant
    Dim giatri  As Variant
    Dim tt      As Variant
    Dim giatri2 As Variant
    Dim ma      As Variant
    Dim DANHAN  As Variant
   
    sSQLCol = "matkhai, madtnt, kylbo, tthtk, mact, donvi, giatri, tt, giatri2, ma,ngnop,kykkhai,DANHAN,lan_quet"

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If

    'Du lieu muc I

    With fps
        .Sheet = 1
        
        matkhai = "'01/PHXD'"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        tt = "''"
        .GetText .ColLetterToNumber("E"), 30, kylbo

        If Trim(kylbo) = vbNullString Then
            kylbo = "CTOD('')"
        Else
            kylbo = "'" & kylbo & "'"
        End If
        
        .GetText .ColLetterToNumber("F"), 10, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        .GetText .ColLetterToNumber("E"), 32, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        DANHAN = "''"

        Dim arrMaCT() As Variant
        Dim arrMa()   As Variant
        Dim i         As Integer
        i = 0
        arrMaCT() = Array("'001'", "'002'", "'003'", "'004'", "'007'", "'008'", "'009'", "'010'", "'013'", "'014'", "'015'", "'016'")
        arrMa() = Array("''", "'21'", "'22'", "'23'", "' '", "'26'", "'27'", "'28'", "'31'", "'32'", "'33'", "'34'")
        
        .Col = .ColLetterToNumber("B")
        .Row = 46

        Do
            mact = arrMaCT(i)
            ma = arrMa(i)
            
            .GetText .ColLetterToNumber("N"), .Row, donvi
            
            If Trim(donvi) = vbNullString Then

                donvi = "''"
            Else

                donvi = "'" & TAX_Utilities_Svr_New.Convert(Trim(donvi), UNICODE, TCVN) & "'"
            End If
            
            .GetText .ColLetterToNumber("T"), .Row, giatri

            If giatri = vbNullString Then
                giatri = "0"
            End If
            
            giatri2 = giatri
            
            sSQLVal = matkhai & "," & madtnt & "," & kylbo & "," & tthtk & "," & mact & "," & donvi & "," & giatri & "," & tt & "," & giatri2 & "," & ma & "," & ngnop & "," & kykkhai & "," & DANHAN & "," & LAN_QUET

            sSQL = "INSERT INTO TMP_PHXD" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
            bln = clsDAO.ExecuteDLL(sSQL)
            
            .Row = .Row + 1
            i = i + 1
        Loop Until .Row = 58
        
    End With

    'Du lieu muc con cua I

    clsDAO.Disconnect
 
    InsertDTL = vbNullString
   
End Function

Public Function Prepared3() As Boolean

    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1

        'set check loi dinh danh
        If checkDLTTonTai = False Then
            .Col = .ColLetterToNumber("E")
            .Row = 36
            .Value = 1
        End If

        '        'Set NgayNhanToKhai
        '        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
        '        .Row = NGAY_NHAN_TO_KHAI_Y
        '        If strNgayNhanToKhai <> "" Then
        '            .Text = strNgayNhanToKhai
        '            UpdateCell fps, .Col, .Row, .Text
        '        End If

        .Col = .ColLetterToNumber(MA_BPQL_X)
        .Row = MA_BPQL_Y

        If strMaBPQL <> "" Then
            .Text = strMaBPQL
            UpdateCell fps, .Col, .Row, .Text
        End If

        '        .Col = .ColLetterToNumber("V")
        '        .Row = 30
        '        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        '
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

        .Col = .ColLetterToNumber("E")
        .Row = 30

        If Len(Month(Now())) = 1 Then
            .Text = "0" & Month(Now()) & "/" & Year(Now())
            sKyKeKhai = "0" & Month(Now()) & Year(Now())
        Else
            .Text = Month(Now()) & "/" & Year(Now())
            sKyKeKhai = Month(Now()) & Year(Now())
        End If

        UpdateCell fps, .Col, .Row, .Text

        .Col = .ColLetterToNumber("E")
        .Row = 32
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        UpdateCell fps, .Col, .Row, .Text
        .Lock = False

        .Col = .ColLetterToNumber("M")
        .Row = 32
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        UpdateCell fps, .Col, .Row, .Text

        ' ten nguoi nop the
        .Col = .ColLetterToNumber(TEN_GOI_X)
        .Row = TEN_GOI_Y

        If strTenGoi <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenGoi, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If

        ' dia chi
        .Col = .ColLetterToNumber(DIA_CHI_X)
        .Row = DIA_CHI_Y

        If strDchi <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strDchi, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If

        ' dien thoai
        .Col = .ColLetterToNumber(DIEN_THOAI_X)
        .Row = DIEN_THOAI_Y

        If strDThoai <> "" Then
            .Text = strDThoai
            UpdateCell fps, .Col, .Row, .Text
        End If

        ' Fax
        .Col = .ColLetterToNumber(FAX_X)
        .Row = FAX_Y

        If strFax <> "" Then
            .Text = strFax
            UpdateCell fps, .Col, .Row, .Text
        End If

        'Mail
        .Col = .ColLetterToNumber(MAIL_X)
        .Row = MAIL_Y

        If strFax <> "" Then
            .Text = strFax
            UpdateCell fps, .Col, .Row, .Text
        End If

        '------------- them thong tin dai ly thue ------------------------------
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
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strMST_DLT, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If

        '        'dia chi
        .Col = .ColLetterToNumber(DIA_CHI_DL_X)
        .Row = DIA_CHI_DL_Y

        If strDchi_DLT <> "" Then
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strDchi_DLT, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        ''        ' quan huyen
        .Col = .ColLetterToNumber(QUAN_HUYEN_DL_X)
        .Row = QUAN_HUYEN_DL_Y

        If strQHuyen_DLT <> "" Then
            .Text = strQHuyen_DLT
        End If
        
        ' thanh pho
        .Col = .ColLetterToNumber(TINH_TPHO_DL_X)
        .Row = TINH_TPHO_DL_Y

        If strTTPho_DLT <> "" Then
            .Text = strTTPho_DLT
        End If

        '        'dien thoai
        .Col = .ColLetterToNumber(DIEN_THOAI_DL_X)
        .Row = DIEN_THOAI_DL_Y

        If strDthoai_DLT <> vbNullString Then
            .Text = strDthoai_DLT
        End If

        '        'fax
        .Col = .ColLetterToNumber(FAX_DL_X)
        .Row = FAX_DL_Y

        If strFax_DLT <> "" Then
            .Text = strFax_DLT
        End If

        '        ' mail
        .Col = .ColLetterToNumber(MAIL_DL_X)
        .Row = MAIL_DL_Y

        If strMail_DLT <> "" Then
            .Text = strMail_DLT
        End If

        '        ' so hop dong
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

        ' ten bo phan quan ly
        '        .Col = .ColLetterToNumber(TEN_BPQL_X)
        '        .Row = TEN_BPQL_Y
        '        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        
        'so thu tu to khai
        .Col = .ColLetterToNumber(STT_X)
        .Row = STT_Y
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))

        ' lan quet
        .Col = .ColLetterToNumber("R")
        .Row = 37
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        .EventEnabled(EventAllEvents) = True
    End With

    Prepared3 = True
End Function

Public Sub Prepared4(ByVal dNgayDauKy As Date)
Dim strKyLapBo As String
Dim blnValid As Boolean

    blnValid = True
    
'    With fps
'        .EventEnabled(EventAllEvents) = False
'        .Sheet = 1
'        .Col = .ColLetterToNumber(KY_LAP_BO_X)
'        .Row = KY_LAP_BO_Y
'        strKyLapBo = .Text
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
'        .EventEnabled(EventAllEvents) = True
'    End With
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
     Dim NGAYNOP As Variant, ngayHienTai As Variant
    CheckValidData = True
    With fps
        .Sheet = 1
        
        ' check ngay nop khong dc lon hon ngay hien tai
        .GetText .ColLetterToNumber("E"), 32, NGAYNOP
        ngayHienTai = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        ngayHienTai = DateSerial(Int(Mid$(ngayHienTai, 7, 4)), Int(Mid$(ngayHienTai, 4, 2)), Int(Mid$(ngayHienTai, 1, 2)))
        NGAYNOP = DateSerial(Int(Mid$(NGAYNOP, 7, 4)), Int(Mid$(NGAYNOP, 4, 2)), Int(Mid$(NGAYNOP, 1, 2)))
        If NGAYNOP > ngayHienTai Then
            CheckValidData = False
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
'    Dim vSO_HIEU_TEP, vSO_TT_TK, vDANHAN, vGHI_CHU_LOI, vCO_GTRINH_02A
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
'    strSQL = strSQL & "'" & vDANHAN & "',"
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
                    .SetText Col, Row, Format_mmyyyy(CStr(varTemp))
'                    checkNgayNop varTemp, vNgayQyet
                Else
                    .SetActiveCell Col, Row
                End If
            Else
             .SetText Col, Row, ""
            End If
           UpdateCell fps, .Col, .Row, .Text
        End If
        If Col = .ColLetterToNumber("E") And Row = 23 Then
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



    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        End If
        sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_TK" & sKyKeKhai & _
            " WHERE MATKHAI = '01/PHXD'  AND MADTNT= " & "'" & strMST & "'" & _
            " AND KYKKHAI ='" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND (TTHTK = '1' OR TTHTK = '2')"
                 
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
          sSQL = "SELECT MADTNT, NGNOP FROM  TMP_TK" & vKeKhai_temp & "  WHERE MADTNT = " & "'" & strMST & "'" & " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'" & " AND MATKHAI = '01/PHXD'"
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

