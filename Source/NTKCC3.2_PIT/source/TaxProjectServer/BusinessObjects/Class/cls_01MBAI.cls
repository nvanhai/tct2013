VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01MBAI"
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
'Const HEADER_KY_LAP_BO_Y = 13
'Const HEADER_KY_LAP_BO_X = "B"
Const PHONG_XU_LY_Y = 22
Const PHONG_XU_LY_X = "V"
Const ROW_STATIC = 44
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
Const TEN_GOI_DL_X = "K"
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

Const HEADER_KY_LAP_BO_ROW = 13
Const HEADER_KY_LAP_BO_COL = "B"
'end

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
'Public strMaSoTep As String
Public isExistFile As Boolean
Public lan_quet As Variant
Const STT_X = "B"
Const STT_Y = 44
'Public strMaSoTep As String
Public strNgayNhanToKhai As String
Public strNguoiSuDung As String
Public strPhongXuLy As String
Private larrid() As String
Private larrPhongXuLy() As String
Private lSoPhongXL As Long
Public sKyKeKhai As String
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
Private MaMucKHBS As String

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

        'Ky lap bo
        SetDateFormat fps, 1, 22, .ColLetterToNumber("E"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 24, .ColLetterToNumber("E"), DDMMYYYY
        .Row = 24
        .Col = .ColLetterToNumber("E")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 24, .ColLetterToNumber("M"), DDMMYYYY
        .Row = 24
        .Col = .ColLetterToNumber("M")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft

    End With
    
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
Dim i As Integer, intIndexCombo As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim iCol As Long, iRow As Long
Dim vKHBS

    ReDim Preserve larrPhongXuLy(0)
    larrPhongXuLy(0) = "00"
    ReDim Preserve larrid(0)
    larrid(0) = "00"
    
    With fps
        .Sheet = 1
        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y
        If rsPXL.Fields.count > 0 Then
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

Private Function GetMaTieuMuc(Id As Variant) As String
    Dim xmlDOMdata As New MSXML.DOMDocument
    Dim xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim strDataFileName, strName As String
    Dim arrDanhsach() As String
    
    strDataFileName = GetAbsolutePath("..\InterfaceTemplates\Catalogue_PHLP_012009_01.xml")
    strName = ""
   If xmlDOMdata.Load(strDataFileName) Then
        Set xmlNodeListCell = xmlDOMdata.getElementsByTagName("Cell")
        For Each xmlNode In xmlNodeListCell
            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                If Id = arrDanhsach(1) Then
                    strName = arrDanhsach(2)
                    Exit For
                End If
            End If
        Next
       
    End If
    GetMaTieuMuc = strName
End Function

Private Function GetMaMuc(ByVal MaTieuMuc As Variant) As String
    Dim strDataFileName As String
    Dim pSQL            As String
    Dim strName         As String
    
    Dim rcord           As ADODB.Recordset
    Dim fso             As New FileSystemObject
    
    Dim clsDAOs As New TAX_Utilities_Svr_New.clsADO
    
    strDataFileName = spathVat & "\DB_HT\DMMT2.DBF"

    If fso.FileExists(strDataFileName) Then
        If clsDAOs.Connected = False Then
            ' Lay mamuc theo matm
            clsDAOs.CreateConnectionString spathVat & "\DB_HT\"
            clsDAOs.Connect
        End If
    
        pSQL = "SELECT MAMUC FROM DMMT2 WHERE MATM = '" & MaTieuMuc & "'"
       
        Set rcord = clsDAOs.Execute(pSQL)

        If Not rcord Is Nothing Then
            strName = rcord.Fields("MAMUC")
        End If
    End If
    
    GetMaMuc = strName
    
    clsDAOs.Disconnect
End Function
' Tra ve gia tri "Nganh--macap--machuong"
Private Function GetMaNganhTheoDTNT(ByVal MaDTNopThue As Variant) As String
    Dim strDataFileName As String
    Dim pSQL            As String
    Dim strName         As String
    
    Dim rcord           As ADODB.Recordset
    Dim fso             As New FileSystemObject
    Dim clsDAOs As New TAX_Utilities_Svr_New.clsADO
    
    strDataFileName = spathVat & "\DTNT\DTNT2.DBF"

    If fso.FileExists(strDataFileName) Then
        If clsDAOs.Connected = False Then
            ' Lay mamuc theo matm
            clsDAOs.CreateConnectionString spathVat & "\DTNT\"
            clsDAOs.Connect
        End If
    
        pSQL = "SELECT Nganh,Macap,Machuong FROM DTNT2 WHERE MADTNT = " & MaDTNopThue
       
        Set rcord = clsDAOs.Execute(pSQL)

        If Not rcord Is Nothing Then
            strName = rcord.Fields("Nganh") & "--" & rcord.Fields("Macap") & "--" & rcord.Fields("Machuong")
        End If
    End If
    
    GetMaNganhTheoDTNT = strName
    
    clsDAOs.Disconnect
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

        .Col = .ColLetterToNumber(NGAY_NOP_X)
        .Row = NGAY_NOP_Y
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        UpdateCell fps, .Col, .Row, .Value
        .Lock = False

        .Col = .ColLetterToNumber(NGAY_QUET_X)
        .Row = NGAY_QUET_Y
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        UpdateCell fps, .Col, .Row, .Value

        'Set NguoiSuDung
        .Col = .ColLetterToNumber(NGUOI_SU_DUNG_X)
        .Row = NGUOI_SU_DUNG_Y
        If strNguoiSuDung <> "" Then
            .Text = strNguoiSuDung
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
'        If Len(Month(Now())) = 1 Then
'            .Text = Year(Now())
'            sKyKeKhai = Year(Now())
'        Else
            .Text = Year(Now())
            sKyKeKhai = Year(Now())
        'End If
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
        
        .Col = .ColLetterToNumber(PHONG_QL_X)
        .Row = PHONG_QL_Y
        If strMaBPQL <> "" Then
            .Text = strMaBPQL
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        ' set thong tin dai ly thue
        ' ten dai ly thue
        .Col = .ColLetterToNumber(TEN_GOI_DL_X)
        .Row = TEN_GOI_DL_Y
        If strTen_DLT <> "" Then
            .Text = strTen_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' dia chi dai ly thue
        .Col = .ColLetterToNumber(DIA_CHI_DL_X)
        .Row = DIA_CHI_DL_Y
        If strDchi_DLT <> "" Then
            .Text = strDchi_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' dien thoai dai ly
        .Col = .ColLetterToNumber(DIEN_THOAI_DL_X)
        .Row = DIEN_THOAI_DL_Y
        If strDthoai_DLT <> "" Then
            .Text = strDthoai_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' fax dai ly thue
        .Col = .ColLetterToNumber(FAX_DL_X)
        .Row = FAX_DL_Y
        If strFax_DLT <> "" Then
            .Text = strFax_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' Email dai ly
        .Col = .ColLetterToNumber(MAIL_DL_X)
        .Row = MAIL_DL_Y
        If strMail_DLT <> "" Then
            .Text = strMail_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' so hop dong
        .Col = .ColLetterToNumber(SO_HDDL_DL_X)
        .Row = SO_HDDL_DL_Y
        If strSoHD_DLT <> "" Then
            .Text = strSoHD_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' ngay hop dong
        .Col = .ColLetterToNumber(NGAY_HDDL_DL_X)
        .Row = NGAY_HDDL_DL_Y
        If strNgayHD_DLT <> "" Then
            .Text = strNgayHD_DLT
            UpdateCell fps, .Col, .Row, .Text
        End If
             

                ' lan quet
        .Col = .ColLetterToNumber("M")
        .Row = 22
        strSoTTTKhai = Trim(TinhSoLanQuet)
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strSoTTTKhai, TCVN, UNICODE))
        .EventEnabled(EventAllEvents) = True
        
         ' set Phong quan ly
        .Col = .ColLetterToNumber(PHONG_QL_X)
        .Row = PHONG_QL_Y
        .Text = TAX_Utilities_Svr_New.Convert(strTenPQL, TCVN, UNICODE)
        UpdateCell fps, .Col, .Row, .Text
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

        ' Kiem tra neu ky lap bo trang se bao loi
        If Trim(strKyLapBo) = "" Or Trim(strKyLapBo) = "../...." Then
            blnValid = False
            .Col = .ColLetterToNumber("B")
            .Row = 19
            .formula = IIf(blnValid, "1", "0")
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 19
            .formula = IIf(blnValid, "1", "0")
            
            If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
                blnValid = False
            Else
                blnValid = True
            End If
            
            .Col = .ColLetterToNumber(HEADER_KY_LAP_BO_COL)
            .Row = HEADER_KY_LAP_BO_ROW
            
            If Not blnValid Then
                .formula = "0"
            Else
                .formula = "1"
            End If
            
            
'            ' Kiem tra xem ky ke khai trong cung ky hay khac ky
'            If (DateAdd("M", 1, dNgayDauKy) = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Or (dNgayDauKy = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Then
'                blnValid = True
'            Else
'                blnValid = False
'            End If
'
'            .Col = .ColLetterToNumber("B")
'            .Row = 17
'
'            If Not blnValid Then
'                .Formula = "0"
'            Else
'                .Formula = "1"
'            End If
            ' kiem tra ky lap bo khong duoc lon hon ky hien tai
            If DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) > Date Then
                blnValid = False
            Else
                blnValid = True
            End If
            .Col = .ColLetterToNumber("B")
            .Row = 18
            
            If Not blnValid Then
                .formula = "0"
            Else
                .formula = "1"
            End If
        End If
        
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

    strDataFileName = GetCatalogueFileName
    
    If xmlDOMdata.Load(GetAbsolutePath(strDataFileName)) Then
        Set xmlNodeListCell = xmlDOMdata.getElementsByTagName("Cell")

        For Each xmlNode In xmlNodeListCell

            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    
                If Id = arrDanhsach(1) Then
                    strIdQLT = arrDanhsach(0)
                    TenTN = arrDanhsach(2)
                    strDVT = arrDanhsach(3)
                    strThueSuat = arrDanhsach(4)
                    blnSuaThueSuat = IIf(arrDanhsach(5) = "1", True, False)

                    If Val(TAX_Utilities_Svr_New.Year) > 2014 Or (Val(TAX_Utilities_Svr_New.Year) = 2014 And Val(TAX_Utilities_Svr_New.Month) > 1) Then
                        maTM = arrDanhsach(8)
                    Else
                        maTM = "2625"
                    End If
                            
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

    Dim KYLBO     As Variant
    Dim madtnt    As Variant
    Dim ngnop     As Variant
    Dim MATHUE As Variant
    Dim MAPP    As Variant
    Dim TENCT   As Variant
    Dim MaMuc As Variant
    Dim maTM   As Variant
    Dim mact As Variant
    Dim STT As Variant
    Dim STT2    As Variant
    Dim STTIN   As Variant
    Dim TTHTK   As Variant
    
    Dim matkhai As Variant
    
    Dim kykkhai As Variant
    Dim BacMB   As Variant
    Dim SoThe   As Variant
    Dim VONDK   As Variant
    Dim ThuePN  As Variant
    Dim tt      As Variant
    Dim BACMB2  As Variant
    Dim SOTHE2  As Variant
    Dim VONDK2  As Variant
    Dim THUEPN2 As Variant
    
    'Dim TTHTK     As Variant
    Dim LOAILP   As Variant
    Dim LEPHI    As Variant
    Dim TYLE     As Variant
    Dim SOTIENLP   As Variant
    Dim SOTIENPN    As Variant
    Dim MACAP   As Variant
    Dim MACHUONG    As Variant
    Dim MAPHLP As Variant
    
    Dim DANHAN As Variant
    Dim lan_quet    As Variant
    
      
    sSQLCol = "kylbo,madtnt,ngnop,MATHUE,MAPP,TENCT,mamuc,matm,MACT,STT,STT2,STTIN,TTHTK,MATKHAI,kykkhai" & _
              ",BacMB,SoThe,VONDK,ThuePN,tt,BACMB2,SOTHE2,VONDK2,THUEPN2,DANHAN,lan_quet"
    
   
    
    With fps
        .Sheet = 1
        
        .GetText .ColLetterToNumber("G"), 8, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        matkhai = "'01/MBAI'"
        
        .GetText .ColLetterToNumber("E"), 22, KYLBO
        
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
       
        .GetText .ColLetterToNumber("E"), 24, ngnop
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
        MATHUE = "'10'"
       
        MAPP = "'1'"
        MaMuc = "'1800'"
                
        SoThe = "0"
        SOTHE2 = "0"
'        MACAP = GetMaNganhTheoDTNT(madtnt)
'        If Trim(MACAP) = vbNullString Then
'            MACAP = "''"
'        Else
'            MACAP = Split(MACAP, "--")(1)
'            MACAP = "'" & MACAP & "'"
'        End If
'
'        MACHUONG = GetMaNganhTheoDTNT(madtnt)
'        If Trim(MACHUONG) = vbNullString Then
'            MACHUONG = "''"
'        Else
'            MACHUONG = Split(MACHUONG, "--")(2)
'            MACHUONG = "'" & MACHUONG & "'"
'        End If
        
        'To khai da chuyen sang VAT = Y
        DANHAN = "''"
        kykkhai = "'" & TAX_Utilities_Svr_New.Year & "'"
        
        TTHTK = "''"
        
        lan_quet = TinhSoLanQuet + 1
        .Col = .ColLetterToNumber("B")
        .Row = 44
        
         If clsDAO.Connected = False Then
            clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
            clsDAO.Connect
        End If
        
        Do
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

            STT = "1"
            
            .GetText .ColLetterToNumber("C"), 49, TENCT
            
            If Trim(TENCT) = vbNullString Then
                TENCT = "''"
            Else
                TENCT = "'" & TENCT & "'"
            End If
            
           .GetText .ColLetterToNumber("M"), .Row, mact
           
           If Trim(mact) = vbNullString Then
                mact = "''"
           Else
                mact = "'" & mact & "'"
           End If
           
           .GetText .ColLetterToNumber("V"), 45, BacMB
            
            If Trim(BacMB) = vbNullString Then
                BacMB = "''"
            Else
                BacMB = "'" & BacMB & "'"
            End If
            
            If BacMB = "'1'" Then
                maTM = "'1801'"
            ElseIf BacMB = "'2'" Then
                maTM = "'1802'"
            ElseIf BacMB = "'3'" Then
                maTM = "'1803'"
            ElseIf BacMB = "'4'" Then
                maTM = "'1804'"
            End If
            
            .GetText .ColLetterToNumber("Q"), 45, VONDK
            
            If Trim(VONDK) = vbNullString Then
                VONDK = "0"
            Else
                VONDK = VONDK
            End If
            
            .GetText .ColLetterToNumber("AC"), 45, ThuePN
            
            If Trim(ThuePN) = vbNullString Then
                ThuePN = "0"
            Else
                ThuePN = ThuePN
            End If
            
            .GetText .ColLetterToNumber("V"), 49, BACMB2
            
            If Trim(BACMB2) = vbNullString Then
                BACMB2 = "''"
            Else
                BACMB2 = "'" & BACMB2 & "'"
            End If
            
            .GetText .ColLetterToNumber("Q"), 49, VONDK2
            
            If Trim(VONDK2) = vbNullString Then
                VONDK2 = "0"
            Else
                VONDK2 = VONDK2
            End If
            
            .GetText .ColLetterToNumber("AC"), 49, THUEPN2
            
            If Trim(THUEPN2) = vbNullString Then
                THUEPN2 = "0"
            Else
                THUEPN2 = THUEPN2
            End If
    

            
            tt = "''"
            
                    'Lan quet
            If Trim(lan_quet) = vbNullString Then
                lan_quet = "1"
            End If
            
            sSQLVal = KYLBO & "," & madtnt & "," & ngnop & "," & MATHUE & "," & MAPP & "," & TENCT & ","
            sSQLVal = sSQLVal & MaMuc & "," & maTM & "," & mact & "," & STT & "," & STT2 & "," & STTIN & "," & TTHTK & "," & matkhai & "," & kykkhai & ","
            sSQLVal = sSQLVal & BacMB & "," & SoThe & "," & VONDK & "," & ThuePN & "," & tt & ","
            sSQLVal = sSQLVal & BACMB2 & "," & SOTHE2 & "," & VONDK2 & "," & THUEPN2 & "," & DANHAN & "," & lan_quet
            
            sSQL = "INSERT INTO TMP_MBCT" & sKyKeKhai & _
                "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
         
            bln = clsDAO.ExecuteDLL(sSQL)
        
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
        
    End With

    clsDAO.Disconnect
    InsertDTL = vbNullString
End Function

'--------------------------------------------------------------------------------
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
    Dim madtnt As Variant
    Dim ngnop  As Variant
    Dim KYLBO  As Variant

    With fps
        .Sheet = 1
        
        .GetText .ColLetterToNumber("G"), 8, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
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

    End With

    InsertDTL_KHBS = InsertDTL_KHBS156(fps, madtnt, "01/MBAI", ngnop, KYLBO, spathVat, TTHTK, lan_quet, True, sKyKeKhai, MaMucKHBS)
    
    InsertHDR_KHBS
    
    InsertDTL_KHBS = vbNullString
End Function
Public Function InsertHDR_KHBS() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    Dim bln        As Boolean
    
    Dim madtnt As Variant
    Dim kykkhai As Variant
    Dim matkhai As Variant
    Dim MATHUE As Variant
    Dim MaMuc As Variant
    Dim maTM As Variant
    Dim MAPP As Variant
    Dim ngnop As Variant
    Dim TTHTK As Variant
    Dim KYLBO As Variant
    Dim songaycn As Variant
    Dim sotiencn As Variant
    Dim tienhoan As Variant
    Dim lenhhoan As String
    Dim ngayhoan As Variant
    Dim HANNOP As Variant
    Dim cqt As Variant
    Dim songay As Variant
    Dim SOTIEN As Variant
    Dim lydokhac As String
    Dim DANHAN As Variant
    
    With fps
        .Sheet = 1
        
        .GetText .ColLetterToNumber("G"), 8, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        MaMuc = "'2150'"
        maTM = "''"
        matkhai = "'01/MBAI'"
        
                MATHUE = "'20'"
                MAPP = "'2'"
           
        .GetText .ColLetterToNumber("E"), 24, ngnop

        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If

        .GetText .ColLetterToNumber("E"), 22, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
        
        'Lay sheet KHBS tren to khai
        .Sheet = .SheetCount - 1
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 28, songaycn
        If Trim(songaycn) = vbNullString Then
            songaycn = "0"
        End If
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 27, sotiencn
        If Trim(sotiencn) = vbNullString Then
            sotiencn = "0"
        End If
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 21, tienhoan
        If Trim(tienhoan) = vbNullString Then
            tienhoan = "0"
        End If
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 19, lenhhoan
        If Trim(lenhhoan) = vbNullString Then
            lenhhoan = "''"
        Else
            lenhhoan = "'" & TAX_Utilities_Svr_New.Convert(lenhhoan, UNICODE, TCVN) & "'"
        End If
        
        .GetText .ColLetterToNumber("E"), .MaxRows - 17, ngayhoan

        If Trim(ngayhoan) = vbNullString Then
            ngayhoan = "CTOD('')"
        Else
            ngayhoan = ToDate(Trim(ngayhoan), DDMMYYYY)
            ngayhoan = "CTOD('" & Format(ngayhoan, "mm/dd/yyyy") & "')"
        End If
        
        HANNOP = "CTOD('')"
        
        .GetText .ColLetterToNumber("BI"), .MaxRows - 13, cqt
        If Trim(cqt) = vbNullString Then
            cqt = "''"
        Else
            cqt = "'" & cqt & "'"
        End If
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 11, songay
        If Trim(songay) = vbNullString Then
            songay = "0"
        End If
        
        .GetText .ColLetterToNumber("BE"), .MaxRows - 9, SOTIEN
        If Trim(SOTIEN) = vbNullString Then
            SOTIEN = "0"
        End If
        
        .GetText .ColLetterToNumber("BD"), .MaxRows - 5, lydokhac
        If Trim(lydokhac) = vbNullString Then
            lydokhac = "''"
        Else
            lydokhac = "'" & TAX_Utilities_Svr_New.Convert(lydokhac, UNICODE, TCVN) & "'"
        End If
    
        DANHAN = "''"
        
        If clsDAO.Connected = False Then
            clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
            clsDAO.Connect
        End If
        
        
        sSQLCol = "madtnt,kykkhai,matkhai,MATHUE,MaMuc,maTM,MAPP,ngnop,TTHTK,KYLBO,songaycn,sotiencn,tienhoan," & _
            "lenhhoan,ngayhoan,HANNOP,cqt,songay,SOTIEN,lydokhac,DANHAN,LAN_QUET"
            
        sSQLVal = madtnt & "," & kykkhai & "," & matkhai & "," & MATHUE & "," & MaMuc & "," & maTM & "," & MAPP & "," & ngnop & "," & TTHTK & "," & KYLBO & "," & songaycn & "," & sotiencn & "," & tienhoan & "," & _
            lenhhoan & "," & ngayhoan & "," & HANNOP & "," & cqt & "," & songay & "," & SOTIEN & "," & lydokhac & "," & DANHAN & "," & lan_quet
        
        sSQL = "INSERT INTO TMP_ND" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

        bln = clsDAO.ExecuteDLL(sSQL)
        
    End With

    InsertHDR_KHBS = vbNullString
    
    clsDAO.Disconnect
    
End Function
Public Function InsertHDR() As String
    Dim sSQL         As String
    Dim sSQLCol      As String
    Dim sSQLVal      As String
    Dim bln          As Boolean
    
    Dim madtnt       As Variant
    Dim madlt        As Variant
    Dim MATHUE       As Variant
    Dim maTM         As Variant
    Dim MAPP         As Variant
    Dim KYLBO        As Variant
    Dim ngnop        As Variant
    Dim LoaiTien     As Variant
    Dim MABPQL       As Variant
    Dim SHTEP        As Variant
    
    Dim TTHTK        As Variant
    
    'TTHTK public
    Dim kykkhai      As Variant
    Dim ngnhap       As Variant
    Dim ThueTKy      As Variant
    Dim GTTinhThue   As Variant
    Dim DaGhi        As Variant
    Dim KTTRUOC      As Variant
    Dim DoanhSo      As Variant
    Dim MaMuc        As Variant
    Dim MaVach       As Variant
    Dim matkhai      As Variant
    Dim HANNOP       As Variant
    Dim ThueTKy2     As Variant
    Dim HANNOP2      As Variant
    Dim GhiChu        As Variant
    Dim daghi1       As Variant
    Dim DAGHI2       As Variant
    Dim AUTO_CAL     As Variant
    Dim nganh        As Variant
    Dim BacMB        As Variant
    Dim SoThe        As Variant
    Dim lanbs        As Variant
    Dim THUECL       As Variant
    Dim KHAIBS       As Variant
    Dim hthuc_nop As Variant ' not save
    Dim DANHAN       As Variant
    Dim TENNV        As Variant
    Dim CHUNGCHI     As Variant
    
    Dim lan_quet     As Variant
    Dim ITKHAI_ID   As Variant
    
    'LAN_QUET --public

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

    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
    
    sSQLCol = "madtnt,madlt,MATHUE,maTM,MAPP,KYLBO,ngnop,LoaiTien,MABPQL,SHTEP,TTHTK," & _
            "kykkhai,NgNhap,ThueTKy,GTTinhThue,DaGhi,KTTRUOC,DoanhSo,MaMuc,MaVach," & _
            "matkhai,HANNOP,ThueTKy2,HANNOP2,GhiChu,DAGHI1,DAGHI2,AUTO_CAL,NGANH," & _
            "BacMB,SoThe,LanBS,THUECL,KHAIBS,HTHUC_NOP,ITKHAI_ID, LAN_QUET"
    With fps
        .Sheet = 1
                
        .GetText .ColLetterToNumber("G"), 8, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
         .GetText .ColLetterToNumber("G"), 15, madlt
        If Trim(madlt) = vbNullString Then
            madlt = "''"
        Else
            madlt = "'" & madlt & "'"
        End If
        
        
        MATHUE = "'10'"
        'matm = "''"
        MAPP = "'1'"
        
        'TODO
        .GetText .ColLetterToNumber("E"), 22, KYLBO
        
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
        
        ' Doan nay tinh thuetky theo matm
        ' Matm duoc tinh theo maloaiphi
         strMTM = "~"
        
        .Col = .ColLetterToNumber("B")
        .Row = 44
        
        Do
            ' Lay ma tieu muc
            .GetText .ColLetterToNumber("H"), .Row, maTM
            vTemp = "~" & maTM & "~"
            vTemp2 = vTemp2 & maTM & "~"
            ' Kiem tra neu co maTM moi thi bo sung vao --> strMTM se luu cac maTM duy nhat
            If InStr(1, strMTM, vTemp, vbTextCompare) = 0 Then
                strMTM = strMTM & maTM & "~"
            End If
            
            'Ghep cac gia tri phi, le phi thu duoc
            .Col = .ColLetterToNumber("U")
            strThueTKy = strThueTKy & .Text & "~"
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
            
        Loop Until .Text = "aa"
        
        'Mang nay luu gia tri cac maTM duy nhat
        arrMTMSave = Split(strMTM, "~")
        'Mang nay luu gia tri tuong ung
        arrThueTKy = Split(strThueTKy, "~")
        arrMTM = Split(vTemp2, "~")
        
        'Mang nay luu cac tong ThueTKy theo tung maTM
        ReDim arrThueTKySave(UBound(arrMTMSave))
        
        'Tinh tong ThueTKy theo tung maTM
        For i = 1 To UBound(arrMTMSave) - 1
            countThueTKy = 0
            For j = 0 To UBound(arrMTM) - 1
                If arrMTMSave(i) = arrMTM(j) Then
                    countThueTKy = countThueTKy + Val(Replace(arrThueTKy(j), ".", ""))
                End If
            Next
            arrThueTKySave(i) = countThueTKy
        Next

        'MAPP = "'1'"
        
        .GetText .ColLetterToNumber("E"), 24, ngnop
        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
        
        LoaiTien = "''"
        
        
        .GetText .ColLetterToNumber(PHONG_QL_X), CInt(PHONG_QL_Y), MABPQL
        If Trim(MABPQL) = vbNullString Then
            MABPQL = "''"
        Else
            MABPQL = "'" & MABPQL & "'"
        End If
        
        
        SHTEP = "''"
        TTHTK = "''"
    
        kykkhai = "'" & TAX_Utilities_Svr_New.Year & "'"
        
        .GetText .ColLetterToNumber("M"), 24, ngnhap
        'NGNOP = Date
        If Trim(ngnhap) = vbNullString Then
            ngnhap = "CTOD('')"
        Else
            ngnhap = ToDate(Trim(ngnhap), DDMMYYYY)
            ngnhap = "CTOD('" & Format(ngnhap, "mm/dd/yyyy") & "')"
        End If
        
        
        .GetText .ColLetterToNumber("J"), 35, ThueTKy
        If Trim(ThueTKy) = vbNullString Then
            ThueTKy = "0"
        End If
    
    GTTinhThue = "0"
    DaGhi = ".F."
    KTTRUOC = "0"
    DoanhSo = "0"
    'MAPP2 = "'1'"
    
    MaMuc = "''"
    
    MaVach = ".T."
    matkhai = "'01/MBAI'"
    HANNOP = "CTOD('')"
    ThueTKy2 = "0"
    HANNOP2 = "CTOD('')"
    GhiChu = "''"
    
    daghi1 = ".F."
    
    .GetText .ColLetterToNumber("E"), 28, DAGHI2
    If DAGHI2 = vbNullString Or DAGHI2 = "" Or DAGHI2 = "0" Then
        DAGHI2 = ".F."
    Else
        DAGHI2 = ".T."
    End If
    
    
    AUTO_CAL = ".F."
  
  'Lay truong Nganh trong DTNT2 thuoc folder DTNT
    nganh = GetMaNganhTheoDTNT(madtnt)
    
    If Trim(nganh) = vbNullString Then
        nganh = "''"
    Else
        nganh = Split(nganh, "--")(0)
        nganh = "'" & nganh & "'"
    End If
    
            .GetText .ColLetterToNumber("V"), 45, BacMB
            
      If Trim(BacMB) = vbNullString Then
                BacMB = "''"
        Else
                BacMB = "'" & BacMB & "'"
        End If
            
            If BacMB = "'1'" Then
                maTM = "'1801'"
            ElseIf BacMB = "'2'" Then
                maTM = "'1802'"
            ElseIf BacMB = "'3'" Then
                maTM = "'1803'"
            ElseIf BacMB = "'4'" Then
                maTM = "'1804'"
            End If
    
    
    SoThe = "0"
    
    
    DANHAN = "''"
   
 
     '  LanBS
    .GetText .ColLetterToNumber("T"), 5, lanbs
    If lanbs = vbNullString Or lanbs = "0" Then
        lanbs = "0"
    Else
        lanbs = lanbs
    End If

    THUECL = ThueTKy
    'DSoCL = DoanhSo
    'GTTTCL = GTTinhThue
    KHAIBS = ".T."
   'NgayPS
   'NgayPs = "CTOD('')"
   
    'MANNKD = "''"
    
    'TENNV
    .GetText .ColLetterToNumber("G"), .MaxRows - 4, TENNV
    If Trim(TENNV) = vbNullString Then
        TENNV = "''"
    Else
        TENNV = "'" & TENNV & "'"
    End If
    'CHUNGCHI
    .GetText .ColLetterToNumber("G"), .MaxRows - 2, CHUNGCHI
    If Trim(CHUNGCHI) = vbNullString Then
        CHUNGCHI = "''"
    Else
        CHUNGCHI = "'" & CHUNGCHI & "'"
    End If
    hthuc_nop = "''"
    ITKHAI_ID = "0"
        
            'Lan quet
        If Trim(lan_quet) = vbNullString Then
            lan_quet = "1"
        End If
        
    End With


     'Tinh thuetky group by theo matm
'        For i = 1 To UBound(arrMTMSave) - 1
'            vTemp = arrMTMSave(i)
'            maTM = "'" & vTemp & "'"
'            ThueTKy = arrThueTKySave(i)
'            MaMuc = GetMaMuc(vTemp)
'            If Trim(MaMuc) = vbNullString Then
'                MaMuc = "''"
'            Else
'                MaMuc = "'" & MaMuc & "'"
'            End If
'
'            If Trim(ThueTKy) <> vbNullString Then
'                THUECL = ThueTKy
'            Else
'                THUECL = "0"
'            End If
'
'            If Trim(vTemp) <> "" Or Trim(vTemp) <> "''" Then
            
                       
            sSQLVal = madtnt & "," & madlt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & ngnop & "," & LoaiTien & "," & MABPQL & "," & SHTEP & "," & TTHTK & ", "
            sSQLVal = sSQLVal & kykkhai & "," & ngnhap & "," & ThueTKy & "," & GTTinhThue & "," & DaGhi & "," & KTTRUOC & "," & DoanhSo & "," & MaMuc & "," & MaVach & ","
            sSQLVal = sSQLVal & matkhai & "," & HANNOP & "," & ThueTKy2 & "," & HANNOP2 & "," & GhiChu & "," & daghi1 & "," & DAGHI2 & "," & AUTO_CAL & "," & nganh & ","
            sSQLVal = sSQLVal & BacMB & "," & SoThe & "," & lanbs & "," & THUECL & "," & KHAIBS & "," & hthuc_nop & "," & ITKHAI_ID & "," & lan_quet
    
                
            sSQL = "INSERT INTO TMP_TKMB" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                bln = clsDAO.ExecuteDLL(sSQL)
            'End If

       ' Next

    clsDAO.Disconnect
    InsertHDR = vbNullString
End Function

Public Function TinhMaNganh() As String
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim nganh As Variant
    Dim TEMP As New clsADO

        If TEMP.Connected = False Then
            TEMP.CreateConnectionString spathVat & "\dtnt\"
            TEMP.Connect
         End If
         sSQL = "SELECT nganh FROM dtnt2 WHERE MADTNT =  '" & strMST & "'"
        
        Set rs = TEMP.Execute(sSQL)
        If Not rs Is Nothing Then
            nganh = rs.Fields("nganh")
               
            TinhMaNganh = nganh
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
    Dim strFileNameBSHDR As String
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TKMB" & sKyKeKhai & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_MBCT" & sKyKeKhai & ".DBF"
    strFileNameBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"
    strFileNameBSHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_ND" & sKyKeKhai & ".DBF"
    
    'file HDR
    If fso.FileExists(strFileNameHDR) = False Then
        Dim strTepmauHDR As String
        strTepmauHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_TKMB.DBF"
        fso.CopyFile strTepmauHDR, strFileNameHDR, False
    End If

    'file dtl
    If fso.FileExists(strFileNameDTL) = False Then
        Dim strTepmauDTL As String
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_MBCT.DBF"
        fso.CopyFile strTepmauDTL, strFileNameDTL, False
    End If

    
    'file bo sung
        If fso.FileExists(strFileNameBS) = False Then
            Dim strTepmauBS As String
            strTepmauBS = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_BS.DBF"
            fso.CopyFile strTepmauBS, strFileNameBS, False
        End If
        
    If fso.FileExists(strFileNameBSHDR) = False Then
        Dim strTepmauBSHDR As String
        strTepmauBSHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_ND.DBF"
        fso.CopyFile strTepmauBSHDR, strFileNameBSHDR, False
    End If
    
    isExistFile = True
    
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If


        sSQL = "SELECT MADTNT, NGNOP FROM TMP_TKMB" & sKyKeKhai & " WHERE MADTNT = " & "'" & strMST & "'" & " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='01/MBAI' "
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


    lan_quet = TinhSoLanQuet + 1
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
   
        strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TKMB" & vKeKhai_temp & ".DBF"

        'kiem tra neu co db thi moi thuc hien truy van
        If fso.FileExists(strTestFileName) Then
                   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TKMB" & vKeKhai_temp & " WHERE MADTNT = " & "'" & strMST & "'" & " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='01/MBAI' "
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
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim LAN_QUET1 As Variant
    
   
        strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TKMB" & sKyKeKhai & ".DBF"
        If fso.FileExists(strFileNameHDR) = True Then
            If clsDAO.Connected = False Then
            clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
            clsDAO.Connect
            End If
             sSQL = "SELECT max(LAN_QUET) as LAN_QUET FROM TMP_TKMB" & sKyKeKhai & _
                    " WHERE MADTNT = " & "'" & strMST & "'" & _
                    " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='01/MBAI'" & " AND (TTHTK = '1' or TTHTK = '3' )"
    
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



