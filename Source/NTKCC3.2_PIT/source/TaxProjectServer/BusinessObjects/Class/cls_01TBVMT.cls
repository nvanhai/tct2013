VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01TBVMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_Y = 22
Const KY_LAP_BO_X = "E"
Const NGAY_NHAN_TO_KHAI_Y = 24
Const NGAY_NHAN_TO_KHAI_X = "E"
Const NGUOI_SU_DUNG_Y = 24
Const NGUOI_SU_DUNG_X = "V"
Const HEADER_KY_LAP_BO_Y = 13
Const HEADER_KY_LAP_BO_X = "B"
Const PHONG_XU_LY_Y = 22
Const PHONG_XU_LY_X = "V"

'Longvh
Const NGAY_NOP_Y = 24
Const NGAY_NOP_X = "E"
Const NGAY_QUET_Y = 24
Const NGAY_QUET_X = "M"
Const PHONG_QL_X = "V"
Const PHONG_QL_Y = 26
Private rs As ADODB.Recordset
' thong tin din vi
Const MA_SO_THUE_X = "G"
Const MA_SO_THUE_Y = 8
Const TEN_GOI_X = "G"
Const TEN_GOI_Y = 7
Const DIA_CHI_X = "G"
Const DIA_CHI_Y = 9
Const QUAN_HUYEN_X = "F"
Const QUAN_HUYEN_Y = 11
Const TINH_TPHO_X = "T"
Const TINH_TPHO_Y = 11
Const DIEN_THOAI_X = "E"
Const DIEN_THOAI_Y = 12
Const FAX_X = "M"
Const FAX_Y = 12
Private tempKyLb As Variant
Private tempNgayNop As Variant
Const MAIL_X = "V"
Const MAIL_Y = 12
'thong tin Dai ly thue
Const MA_SO_THUE_DL_X = "G"
Const MA_SO_THUE_DL_Y = 15
Const TEN_GOI_DL_X = "H"
Const TEN_GOI_DL_Y = 14
Const DIA_CHI_DL_X = "G"
Const DIA_CHI_DL_Y = 16
Const QUAN_HUYEN_DL_X = "F"
Const QUAN_HUYEN_DL_Y = 17
Const TINH_TPHO_DL_X = "T"
Const TINH_TPHO_DL_Y = 17
Const DIEN_THOAI_DL_X = "E"
Const DIEN_THOAI_DL_Y = 18
Const FAX_DL_X = "M"
Const FAX_DL_Y = 18
Const MAIL_DL_X = "V"
Const MAIL_DL_Y = 18
Const SO_HDDL_DL_X = "I"
Const SO_HDDL_DL_Y = 19
Const NGAY_HDDL_DL_X = "S"
Const NGAY_HDDL_DL_Y = 19
'end

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
'Public strMaSoTep As String
Public isExistFile As Boolean
Public LAN_QUET As Variant
Const STT_X = "M"
Const STT_Y = 22
'Public strMaSoTep As String
Public strNgayNhanToKhai As String
Public strNguoiSuDung As String
Public strPhongXuLy As String
Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long
Private LOAI_TO_KHAI As String

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
'Private tempKyLb As Variant
'Private tempNgayNop As Variant
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
        .Row = 28
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100
        
        'Phong xu ly
        .Col = .ColLetterToNumber("V")
        .Row = 22
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
        
        ' Ky lap bo
        SetDateFormat fps, 1, 22, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 24, .ColLetterToNumber("E"), DDMMYYYY
        .Row = 24
        .Col = .ColLetterToNumber("E")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        '.Text = Day(Date) & "/" & Month(Date) & "/" & Year(Date)
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 24, .ColLetterToNumber("M"), DDMMYYYY
        .Row = 24
        .Col = .ColLetterToNumber("M")
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
    End With
    

    ' Lay ve gia tri cua HDR ID khi quet to khai
    hdrId = GetHdrId(spathVat)
    ' end

    
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean

    Dim maMh   As Variant
    Dim tenMh  As String
    Dim NgayPs As Variant

    Dim i      As Integer
    i = 0
    ReDim Preserve larrPhongXuLy(0)
    larrPhongXuLy(0) = "00"
    ReDim Preserve larrid(0)
    larrid(0) = "00"
    
    With fps
        .Sheet = 1
        
        .Row = 33
        .Col = .ColLetterToNumber("L")

        If .Text = "1" Then
            LOAI_TO_KHAI = "01PS/TBVMT"
            .GetText .ColLetterToNumber("M"), .Row, NgayPs
            TAX_Utilities_Svr_New.Month = Mid$(NgayPs, 4, 2)
            TAX_Utilities_Svr_New.Year = Right$(NgayPs, 4)
        Else
            LOAI_TO_KHAI = "01/TBVMT"
        End If
        
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
              
        .Sheet = 1
        .Row = 45

        Do
            maMh = ""
            tenMh = ""
            .GetText .ColLetterToNumber("AB"), .Row, maMh
            getTenMH maMh, tenMh
            .Col = .ColLetterToNumber("C")
            .Text = tenMh
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        
        Loop Until .Text = "aa"
        
        .Sheet = 2
        .Row = 35

        Do
            maMh = ""
            tenMh = ""
            .GetText .ColLetterToNumber("Y"), .Row, maMh
            getTenMH maMh, tenMh
            .Col = .ColLetterToNumber("D")
            .Text = tenMh
            UpdateCell fps, .Col, .Row, .Text
            
            .GetText .ColLetterToNumber("R"), .Row, maMh
            GetTenCQT maMh, tenMh
            .Col = .ColLetterToNumber("M")
            .Text = tenMh
            UpdateCell fps, .Col, .Row, .Text
            
            .Col = .ColLetterToNumber("C")
            .Row = .Row + 1
        
        Loop Until .Text = "aa"
        
    End With
    
    Prepared2 = True
End Function

Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
End Sub

Public Function Prepared3() As Boolean
    Dim So_tt As Variant

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

        ' thong tin dai ly thue
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
            .Text = Trim(TAX_Utilities_Svr_New.Convert(strMST_DLT, TCVN, UNICODE))
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
        .Col = .ColLetterToNumber(STT_X)
        .Row = STT_Y
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        
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
        
        

        If strMaBPQL <> "" Then
            .Col = .ColLetterToNumber(PHONG_QL_X)
            .Row = PHONG_QL_Y
            .Text = strMaBPQL
            UpdateCell fps, .Col, .Row, .Text
        End If

        
        'set check loi dinh danh
        If checkDLTTonTai = False Then
            .Col = .ColLetterToNumber("E")
            .Row = 28
            .Value = 1
        End If
        
        Dim j       As Integer
        Dim lastRow As Variant
        Dim strID, TenTN   As String
        Dim maTM1 As String
        .Sheet = 1
        .Row = 45

        Do

            .GetText .ColLetterToNumber("AB"), .Row, strID
            DataDM strID, , TenTN, , , , maTM1
            .SetText .ColLetterToNumber("C"), .Row, TenTN
            UpdateCell fps, .ColLetterToNumber("C"), .Row, TenTN
            
            .SetText .ColLetterToNumber("AC"), .Row, maTM1
            UpdateCell fps, .ColLetterToNumber("AC"), .Row, maTM1
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
           
            lastRow = .Row
        Loop Until .Text = "aa"

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

Private Sub DataDM(ByVal Id As String, _
                   Optional ByRef strIdQLT As String, _
                   Optional ByRef TenTN As String, _
                   Optional ByRef strDVT As String, _
                   Optional ByRef strThueSuat As String, _
                   Optional ByRef blnSuaThueSuat As Boolean, _
                   Optional ByRef maTM As String)
    Dim arrDanhsach()   As String
    Dim strDataFileName As String
    Dim xmlDOMdata      As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode         As MSXML.IXMLDOMNode

    'strDataFileName = GetAbsolutePath("..\InterfaceTemplates\Catalogue_DM_thue_BVMT.xml")
    strDataFileName = GetDMThueBVMT
    
    If xmlDOMdata.Load(GetAbsolutePath(strDataFileName)) Then
        Set xmlNodeListCell = xmlDOMdata.getElementsByTagName("Item")

        For Each xmlNode In xmlNodeListCell

            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    
                If Id = arrDanhsach(0) Then
                    TenTN = arrDanhsach(1)
                    strDVT = arrDanhsach(2)
                    strThueSuat = arrDanhsach(3)

                    maTM = arrDanhsach(4)


                    Exit Sub
                End If
            End If

        Next

    End If

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
    
        Dim NGAYNOP As Variant, ngayHienTai As Variant
    
    CheckValidData = True
    With fps
        .Sheet = 1
        
        ' check ngay nop khong dc lon hon ngay hien tai
        .GetText .ColLetterToNumber("E"), 24, NGAYNOP
        ngayHienTai = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        ngayHienTai = DateSerial(Int(Mid$(ngayHienTai, 7, 4)), Int(Mid$(ngayHienTai, 4, 2)), Int(Mid$(ngayHienTai, 1, 2)))
        NGAYNOP = DateSerial(Int(Mid$(NGAYNOP, 7, 4)), Int(Mid$(NGAYNOP, 4, 2)), Int(Mid$(NGAYNOP, 1, 2)))
        If NGAYNOP > ngayHienTai Then
            CheckValidData = False
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
    Dim sSQL      As String
    Dim sSQLCol   As String
    Dim sSQLVal   As String
    Dim bln       As Boolean
   
    Dim madtnt    As Variant
    Dim matkhai   As Variant
    Dim KYLBO     As Variant
    Dim ngnop     As Variant
    Dim kykkhai   As Variant
    Dim TTHTK     As Variant
    Dim NganhTN   As Variant
    Dim TENGOI    As Variant
    Dim donvi     As Variant
    Dim SOLUONG   As Variant
    Dim MUCPHI    As Variant
    Dim TIENPHI   As Variant
    Dim DANHAN    As Variant
    Dim SOLUONGCL As Variant
    Dim TIENPHICL As Variant
    Dim STT       As Variant
    Dim STT2      As Variant
    Dim STTIN     As Variant
    Dim mact      As Variant
    Dim MAHANG    As Variant
    Dim MaMuc As String
    Dim maTM As Variant
  
    sSQLCol = "MADTNT, MATKHAI, KYLBO, NGNOP, KYKKHAI, TTHTK, TENGOI, DVT, SOLUONG ,SOLUONGCL,TIENPHICL"
    sSQLCol = sSQLCol + ", MUCTHUE, TIENPHI, DANHAN,LAN_QUET,MAHANG,MAMUC,MATM"
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
    With fps
        .Sheet = 1
        'ma to khai
        matkhai = "'" & LOAI_TO_KHAI & "'"
        ' ma doi tuong nop thue
          
        .GetText .ColLetterToNumber("G"), 8, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
            
        'ky ke khai
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        ' ngay nop
        .GetText .ColLetterToNumber("E"), 24, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
        'ky lop bo
        .GetText .ColLetterToNumber("E"), 22, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If
        
        DANHAN = "' '"
        
        'trang thai to khai
        Dim TTHTK1 As Variant
        .GetText .ColLetterToNumber("C"), 33, TTHTK1

        If Trim(TTHTK1) = vbNullString Then
            TTHTK = "'" & 3 & "'"
        Else
            TTHTK = "'" & 1 & "'"
        End If
        
        SOLUONGCL = "0"
        TIENPHICL = "0"
        maTM = "'2001'"
        MaMuc = "'2000'"
        .Col = .ColLetterToNumber("B")
        .Row = 45
        Dim lastRow As Variant
        
        Do
            'Ma chi tieu
            mact = "'1000'"

            'ten loai khoang san
            .GetText .ColLetterToNumber("C"), .Row, TENGOI

            If Trim(TENGOI) = vbNullString Then
                TENGOI = "''"
            Else
                TENGOI = "'" & TAX_Utilities_Svr_New.Convert(Trim(TENGOI), UNICODE, TCVN) & "'"
            End If
            
            ' don vi tinh
            donvi = "''"
            .GetText .ColLetterToNumber("H"), .Row, donvi

            If Trim(donvi) = vbNullString Then

                donvi = "''"
            Else

                donvi = "'" & TAX_Utilities_Svr_New.Convert(Trim(donvi), UNICODE, TCVN) & "'"
            End If

            ' so luong
            .GetText .ColLetterToNumber("J"), .Row, SOLUONG

            If Trim(SOLUONG) = vbNullString Then

                SOLUONG = "0"
            End If

            ' nghanh tai nguyen
            .GetText .ColLetterToNumber("AB"), .Row, NganhTN

            If Trim(NganhTN) = vbNullString Then

                NganhTN = "''"
            Else
                NganhTN = "'" & NganhTN & "'"
            End If

            ' muc phi

            .GetText .ColLetterToNumber("O"), .Row, MUCPHI

            If Trim(MUCPHI) = vbNullString Then

                MUCPHI = "0"
            End If

            'tien phi
            .GetText .ColLetterToNumber("U"), .Row, TIENPHI

            If Trim(TIENPHI) = vbNullString Then

                TIENPHI = "0"
            End If
            
            'so thu tu:
            If .Row = 43 Then
                
                STT2 = "0"
                STTIN = "'I'"
            Else
                .GetText .ColLetterToNumber("B"), .Row, STT2

                If Trim(STT) = vbNullString Then
                    STT2 = "0"
                End If

                STTIN = "'" & STT2 & "'"
            End If

            'ma hang
            .GetText .ColLetterToNumber("AB"), .Row, MAHANG
            MAHANG = "'" & MAHANG & "'"
            
            STT = "1"
            
            'ma tieu muc
            .GetText .ColLetterToNumber("AC"), .Row, maTM
            maTM = "'" & maTM & "'"
            
            TIENPHICL = TIENPHI
            SOLUONGCL = SOLUONG
            
            sSQLVal = madtnt & "," & matkhai & "," & KYLBO & "," & ngnop & "," & kykkhai & "," & TTHTK & "," & TENGOI & "," & donvi & "," & SOLUONG & ","
            sSQLVal = sSQLVal & SOLUONGCL & "," & TIENPHICL & "," & MUCPHI & "," & TIENPHI & "," & DANHAN & "," & LAN_QUET & "," & MAHANG & "," & MaMuc & "," & maTM
            
            sSQL = "INSERT INTO TMP_BVMT" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If .Row <> 44 Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
            lastRow = .Row
        Loop Until .Text = "aa"
                
    End With

    clsDAO.Disconnect
    InsertDTL = vbNullString
End Function

'--------------------------------------------------------------------------------
' Project    :       Tax_Business_Svr_New
' Procedure  :       InsertDTL_By_CTTN
' Description:       [Insert bang TMP_T2, tmp_tncn_dtl theo chi tieu CTTN]
' Machine    :       KHAIHOAN
' Date-Time  :       7/22/2011-20:01:28
' Parameters :       pit : kiem tra truong PIT
'--------------------------------------------------------------------------------
Public Function InsertDTL_By_CTTN(hdr As TNCN_DTL, pit As Boolean) As Boolean
  
    Dim bln     As Boolean
    Dim sSQL1   As String
  

    If pit = True Then
        sSQL1 = InsertDTL_TGTC(hdr)
        bln = clsDAO.ExecuteDLL(sSQL1)
    
    End If

    InsertDTL_By_CTTN = bln
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

    sSQLCol = "MADTNT, KYKKHAI, MATKHAI, MATHUE, MAMUC, MATM, MAPP, NGNOP, TTHTK, KYLBO , MACT, MACT2, MACT3, STT, STT2, STTIN," & " SOKK, SODC, SOCL,SOTIENNC,  SONGAYNC, LAN_QUET, DANHAN"

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If

    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("G"), 8, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"

        MATHUE = "'21'"
        MAPP = "'2'"
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        maTM = "'2001'"
        MaMuc = "'2000'"
        
        .GetText .ColLetterToNumber("E"), 24, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 22, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
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
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                       
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
            
            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                           
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
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                       
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
            
            sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                           
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
        
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                       
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
        bln = clsDAO.ExecuteDLL(sSQL)
        .Row = .Row + 3
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
            
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & mact & "," & MACT2 & "," & MACT3 & "," & STT & "," & STT2 & "," & STTIN & "," & SOKK & "," & SODC & "," & SOCL & "," & SOTIENNC & "," & SONGAYNC & "," & LAN_QUET & "," & DANHAN
                           
        sSQL = "INSERT INTO TMP_BS" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        If MACT2 <> "''" Then
            bln = clsDAO.ExecuteDLL(sSQL)
        End If

    End With
   
    clsDAO.Disconnect
    InsertDTL_KHBS = vbNullString
End Function

Public Function InsertHDR() As String
    Dim sSQL         As String
    Dim sSQLCol      As String
    Dim sSQLVal      As String
    Dim bln          As Boolean
    
    Dim madtnt       As Variant
    Dim MATHUE       As Variant
    Dim maTM         As Variant
    Dim MAPP         As Variant
    Dim KYLBO        As Variant
    Dim ngnop        As Variant
    Dim LoaiTien     As Variant
    Dim MABPQL       As Variant
    Dim SHTEP        As Variant

    Dim kykkhai      As Variant
    Dim NgNhap       As Variant
    Dim ThueTKy      As Variant
    Dim GTTinhThue   As Variant
    Dim DaGhi        As Variant
    Dim KTTRUOC      As Variant
    Dim DoanhSo      As Variant
    Dim MAPP2        As Variant
    Dim MaMuc        As Variant
    Dim MaVach       As Variant
    Dim matkhai      As Variant
    Dim HANNOP       As Variant
    Dim ThueTKy2     As Variant
    Dim HANNOP2      As Variant
    Dim TENDA        As Variant
    Dim DAGHI1       As Variant
    Dim DAGHI2       As Variant
    Dim AUTO_CAL     As Variant
    Dim NGANH        As Variant
    Dim DANHAN       As Variant
    Dim MADLT        As Variant
    Dim LanBS        As Variant
    Dim THUECL       As Variant
    Dim DSoCL        As Variant
    Dim GTTTCL       As Variant
    Dim NgayPs       As Variant
    Dim MANNKD       As Variant
    Dim TENNV        As Variant
    Dim CHUNGCHI     As Variant
    'bien dung xu ly Ma Tieu Muc
    Dim arrMTM()     As String, arrMTMSave() As String
    Dim arrThueTKy() As String, arrThueTKySave() As Double
    Dim i, j As Integer
    Dim vTemp        As String
    Dim countThueTKy As Double
    Dim strMTM       As String
    Dim strThueTKy   As String
    Dim ROWNUM       As Variant
    Dim vTemp2       As String
    Dim NGAYNOP As Variant
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
    sSQLCol = "MADTNT, MADLT,MATHUE, MATM, MAPP, KYLBO, NGNOP, "
    sSQLCol = sSQLCol & "LOAITIEN, MABPQL, SHTEP, TTHTK, KYKKHAI, "
    sSQLCol = sSQLCol & "NGNHAP, THUETKY, GTTINHTHUE, DAGHI,  "
    sSQLCol = sSQLCol & "KTTRUOC, DOANHSO, MAPP2, MAMUC, MAVACH,MATKHAI,HANNOP, THUETKY2, HANNOP2, TENDA, DAGHI1 ,DAGHI2, AUTO_CAL, NGANH,LAN_QUET,LANBS, THUECL, NgayPS, DANHAN, DSOCL, GTTTCL, MANNKD, TENNV, CHUNGCHI"
   
    With fps
        .Sheet = 1
        
        .GetText .ColLetterToNumber("G"), .MaxRows - 4, TENNV
        .GetText .ColLetterToNumber("G"), .MaxRows - 2, CHUNGCHI
        'Ma tieu muc
        .Col = .ColLetterToNumber("B")
        .Row = 45
        strMTM = "~"
               
        Do
            ' Lay ma tieu muc tren cel
            .Col = .ColLetterToNumber("AC")
            'cuoi ma tieu muc
            vTemp = "~" & .Text & "~"
            vTemp2 = vTemp2 & .Text & "~"
            
            If InStr(1, strMTM, vTemp, vbTextCompare) = 0 Then
                strMTM = strMTM & .Text & "~"
            End If

            .Col = .ColLetterToNumber("U")
            strThueTKy = strThueTKy & .Text & "~"
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
            ROWNUM = .Row
        Loop Until .Text = "aa"
        
        arrMTMSave = Split(strMTM, "~")
        arrThueTKy = Split(strThueTKy, "~")
        arrMTM = Split(vTemp2, "~")
        
        ReDim arrThueTKySave(UBound(arrMTMSave))
        
        For i = 1 To UBound(arrMTMSave) - 1
            countThueTKy = 0

            For j = 0 To UBound(arrMTM) - 1

                If arrMTMSave(i) = arrMTM(j) Then
                    countThueTKy = countThueTKy + Val(Replace(arrThueTKy(j), ".", ""))
                End If

            Next

            arrThueTKySave(i) = countThueTKy
        Next

        ' so lan bo sung
        If TTHTK = "'" & 2 & "'" Then
            .GetText .ColLetterToNumber("I"), 33, LanBS
        Else
            LanBS = "0"
        End If

        ' ma doi tuong not thue
        .GetText .ColLetterToNumber("G"), 8, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        ' ma so thue dai ly
        .GetText .ColLetterToNumber("G"), 15, MADLT

        If Trim(MADLT) = vbNullString Then
            MADLT = "''"
        Else
            MADLT = "'" & MADLT & "'"
        End If
       
        ' ky lap bo
        .GetText .ColLetterToNumber("E"), 22, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        .GetText .ColLetterToNumber("E"), 24, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            NGAYNOP = ngnop
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
       
        ' MABPQL = "''"
        'ma bo phan quan ly
        .GetText .ColLetterToNumber(PHONG_QL_X), PHONG_QL_Y, MABPQL
        
        If Trim(MABPQL) = vbNullString Then
            MABPQL = "''"
        Else
            MABPQL = "'" & MABPQL & "'"
        End If

        SHTEP = "''"
         
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        NgNhap = Date
        
        If Trim(NgNhap) = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If

        LoaiTien = "''"
        MATHUE = "'24'"
        MAPP = "'1'"
        MAPP2 = "'1'"
        MaVach = ".T."
        matkhai = "'" & LOAI_TO_KHAI & "'"
        THUECL = "0"

        ' Han nop: ngay thu 20 cua thang sau
        
        Dim thang As Variant
        Dim nam   As Variant
     
        thang = Month(Now())
        nam = Year(Now())

        If thang = 12 Then
            nam = nam + 1
            HANNOP = "20/01/" & nam
        Else
            thang = thang + 1
            HANNOP = "20/" & thang & "/" & nam
        End If

        GTTinhThue = "0"
        KTTRUOC = "0"

        DoanhSo = "0"
        HANNOP2 = HANNOP
        TENDA = "''"
        DaGhi = ".T."
        
        DAGHI1 = ".F."

        If checkDLTTonTai = False Then
            DAGHI2 = ".T."
        Else
            DAGHI2 = ".F."
        End If

        AUTO_CAL = ".F."
        DANHAN = "''"
        DSoCL = "0"
        GTTTCL = "0"
        NGANH = TinhMaNganh
        
        MANNKD = "''"
        
        ' ngay phat sinh
        NgayPs = "CTOD('')"

        .Row = 33
        .Col = .ColLetterToNumber("L")

        If .Text = "1" Then
            .GetText .ColLetterToNumber("M"), .Row, NgayPs

            If Trim(NgayPs) = vbNullString Then
                NgayPs = "CTOD('')"
            Else
                NgayPs = ToDate(Trim(NgayPs), DDMMYYYY)
                NgayPs = "CTOD('" & Format(NgayPs, "mm/dd/yyyy") & "')"
            End If
            
            HANNOP = DateAdd("d", 10, NGAYNOP)
            HANNOP = ToString(HANNOP, "DD/MM/YYYY")
        End If

        If Trim(HANNOP) = vbNullString Then
            HANNOP = "CTOD('')"
        Else
            HANNOP = ToDate(Trim(HANNOP), DDMMYYYY)
            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
        End If
        HANNOP2 = HANNOP
        'thue gtgt phai nop
        '        .GetText .ColLetterToNumber("U"), 35, ThueTKy
        '
        '        If ThueTKy = vbNullString Then
        '            ThueTKy = "0"
        '        End If
        
        maTM = "'2001'"
        MaMuc = "'2000'"

        DoanhSo = "0"
        ThueTKy2 = "0"
        
        For i = 1 To UBound(arrMTMSave) - 1
            vTemp = arrMTMSave(i)
            
            maTM = "'" & vTemp & "'"
            ThueTKy = arrThueTKySave(i)

            If Trim(maTM) <> "" Or Trim(maTM) <> "''" Then
            
                sSQLVal = madtnt & "," & MADLT & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & ngnop & ", "
                sSQLVal = sSQLVal & LoaiTien & "," & MABPQL & "," & SHTEP & "," & TTHTK & "," & kykkhai & ", "
                sSQLVal = sSQLVal & NgNhap & "," & ThueTKy & "," & GTTinhThue & "," & DaGhi & ", "
                sSQLVal = sSQLVal & KTTRUOC & "," & DoanhSo & "," & MAPP2 & "," & MaMuc & "," & MaVach & ", "
                sSQLVal = sSQLVal & matkhai & "," & HANNOP & "," & ThueTKy2 & "," & HANNOP2 & "," & TENDA & "," & DAGHI1 & "," & DAGHI2 & "," & AUTO_CAL & ",'" & NGANH & "'," & LAN_QUET & "," & LanBS & "," & THUECL & "," & NgayPs & "," & DANHAN & "," & DSoCL & "," & GTTTCL & "," & MANNKD & ",'" & TENNV & "','" & CHUNGCHI & "'"
           
                sSQL = "INSERT INTO TMP_TK" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
     
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

        Next
        
    End With
     
    InsertHDR = vbNullString
    clsDAO.Disconnect
End Function
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
Public Function TKTT() As Boolean
    Dim sSQL            As String
    Dim rs              As ADODB.Recordset
    Dim fso             As New FileSystemObject
    Dim strFileNameHDR  As String
    Dim strFileNameDTL  As String
    Dim strFileNameHDR1 As String
    Dim strFileNameDTL1 As String
    Dim strFileNameBS As String
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BVMT" & sKyKeKhai & ".DBF"
    strFileNameHDR1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tmp_tncn_hdr.DBF"
    strFileNameDTL1 = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "tmp_tncn_dtl.DBF"
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
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_BVMT.DBF"
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


        sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyKeKhai & " WHERE MADTNT = " & "'" & strMST & "'" & " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='" & LOAI_TO_KHAI & "' "
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
    'hdrId = GetHdrId(spathVat)
    clsDAO.Disconnect

End Function

Public Function UpdateTHUETKY2() As Boolean
    Dim sSQL As String

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If
   
   sSQL = "UPDATE TMP_TZ" & sKyKeKhai & _
        " SET THUETKY2 = 1 " & _
        " WHERE MATHUE = '09' and MAPP = '02' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='02T/KK-TNCN'"

   clsDAO.ExecuteDLL sSQL

   UpdateTHUETKY2 = True

   clsDAO.Disconnect
End Function

Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL            As String
    Dim rs              As ADODB.Recordset
    Dim vKeKhai         As Variant
    Dim vKeKhai_temp    As Variant
    Dim vKyLB           As Variant
    Dim strTestFileName As String
    Dim fso             As New FileSystemObject
    Dim matkhai         As Variant
     
    Dim flag            As Boolean
    'kiem tra da co to khai chinh thuc chua
    flag = True
    vKyLB = DateSerial(Int(Year(Now())), Int(Month(Now())), 1)
    vKeKhai = DateSerial(Int(Val(TAX_Utilities_Svr_New.Year)), Int(Val(TAX_Utilities_Svr_New.Month)), 1)
    vKeKhai_temp = sKyKeKhai

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If

    Do While flag
   
        strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & vKeKhai_temp & ".DBF"

        'kiem tra neu co db thi moi thuc hien truy van
        If fso.FileExists(strTestFileName) Then
                   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & vKeKhai_temp & " WHERE MADTNT = " & "'" & strMST & "'" & " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='" & LOAI_TO_KHAI & "' "
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
   
   sSQL = "DELETE FROM TMP_TZ" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='02T/KK-TNCN'"
     
    
   clsDAO.ExecuteDLL sSQL
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 22, KYLBO
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
   sSQL = "DELETE FROM TMP_T2" & sKyKeKhai & _
        " WHERE MADTNT = " & "'" & strMST & "' AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' " & _
        " AND MATKHAI='02T/KK-TNCN'"
    
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
   Dim ngnop As Variant
   Dim NGNOPDB As Variant
   Dim sSQL As String
   Dim THUEGTGT_KT As Double
   Dim rs As ADODB.Recordset
   Dim i As Integer
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 24, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = ""
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
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
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TZ" & sKyKeKhai & _
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
                If Trim(NGNOPDB) <> vbNullString And Trim(NGNOPDB) = Trim(ngnop) Then
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
        If Col = .ColLetterToNumber("E") And Row = 24 Then
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
        If Col = .ColLetterToNumber("M") And Row = 24 Then
            .GetText Col, Row, varTemp
            If varTemp <> "" And varTemp <> "../...." Then
                If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
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
   
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"

    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
            clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
            clsDAO.Connect
        End If

        sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_TK" & sKyKeKhai & " WHERE MADTNT = " & "'" & strMST & "'" & " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='" & LOAI_TO_KHAI & "'" & " AND (TTHTK = '1' or TTHTK = '3' )"
    
        Set rs = clsDAO.Execute(sSQL)

        If Not rs Is Nothing Then
            TinhSoLanQuet = rs.Fields("LAN_QUET")
                      
        Else
            TinhSoLanQuet = 0
        End If

        clsDAO.Disconnect
    Else
        TinhSoLanQuet = 0
    End If
     
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
            TEMP = TEMP & "~'" & Trim(rs.Fields("tengoi")) & "'"
            TEMP = TEMP & "~'" & Trim(rs.Fields("dchi")) & "'"
        Else
            TEMP = "''~''~''"
        End If
        clsConn.Disconnect
        GetMaCqt = TEMP
End Function

'lay ten mat hang
Public Sub getTenMH(ByVal Id As String, Optional ByRef TenTN As String)
Dim arrDanhsach() As String
Dim strDataFileName As String
Dim xmlDOMdata As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

       'strDataFileName = "..\InterfaceTemplates\Catalogue_DM_thue_BVMT.xml"
        strDataFileName = GetDMThueBVMT
    
       If xmlDOMdata.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDOMdata.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If Id = arrDanhsach(0) Then
                            TenTN = arrDanhsach(1)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub


' Lay danh muc bieu thue BVMT theo HL ky ke khai
Public Function GetDMThueBVMT() As String
    Dim strReturn As String
    Dim strCatalogueName As String
    Dim strFileHLBieuThue As String
    Dim strMonth As String
    Dim strYear As String
    Dim vDateKK As Date
    Dim vDateHL As Date
    Dim MaxDate As Date
    Dim arrDateHL() As String
    Dim i As Integer
    Dim strHieuLucBT As String
    ' list danh sach cac file hieu luc bieu thue
    strFileHLBieuThue = "052015###022013###011900"
    arrDateHL = Split(strFileHLBieuThue, "###")
    ' lay ngay cuoi thang ke khai
    strMonth = TAX_Utilities_Svr_New.Month
    strYear = TAX_Utilities_Svr_New.Year
    If strMonth <> "" Then
        Select Case strMonth
            Case "01", "03", "05", "07", "08", "10", "12"
                vDateKK = Format("31/" & strMonth & "/" & strYear, "dd/mm/yyyy")
            Case "02"
                 If CInt(strYear) / 4 = CInt(strYear) \ 4 And CInt(strYear) \ 100 <> CInt(strYear) / 100 Then
                    vDateKK = Format("29/" & strMonth & "/" & strYear, "dd/mm/yyyy")
                Else
                    vDateKK = Format("28/" & strMonth & "/" & strYear, "dd/mm/yyyy")
                End If
            Case "04", "06", "09", "11"
                vDateKK = Format("30/" & strMonth & "/" & strYear, "dd/mm/yyyy")
        End Select
    End If
    
    ' kiem tra hieu luc voi ky KK
    For i = 0 To UBound(arrDateHL)
        vDateHL = Format("01/" & Left$(arrDateHL(i), 2) & "/" & Right$(arrDateHL(i), 4), "dd/mm/yyyy")
        If vDateKK >= vDateHL Then
            If vDateHL > MaxDate Then
                MaxDate = vDateHL
                strHieuLucBT = arrDateHL(i)
            End If
        End If
    Next
    ' ghep voi ten file DM
    strCatalogueName = "Catalogue_DM_thue_BVMT_" & strHieuLucBT
    
    strReturn = GetAbsolutePath("..\InterfaceTemplates\" & _
        strCatalogueName & ".xml")
        
    GetDMThueBVMT = strReturn
End Function
