VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01TAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const KY_LAP_BO_Y = 23
Const KY_LAP_BO_X = "F"
Const NGAY_NHAN_TO_KHAI_Y = 25
Const NGAY_NHAN_TO_KHAI_X = "S"
Const NGUOI_SU_DUNG_Y = 25
Const NGUOI_SU_DUNG_X = "AG"
Const PHONG_XU_LY_Y = 23
Const PHONG_XU_LY_X = "AG"
Const NGAY_NOP_Y = 25
Const NGAY_NOP_X = "F"
Const NGAY_QUET_Y = 25
Const NGAY_QUET_X = "S"
Const TEN_BPQL_X = "AG"
Const TEN_BPQL_Y = 30
Const MA_BPQL_X = "AG"
Const MA_BPQL_Y = 27
Private rs As ADODB.Recordset
' thong tin din vi
Const MA_SO_THUE_X = "I"
Const MA_SO_THUE_Y = 7
Const TEN_GOI_X = "I"
Const TEN_GOI_Y = 5
Const DIA_CHI_X = "I"
Const DIA_CHI_Y = 9
'Const QUAN_HUYEN_X = "F"
'Const QUAN_HUYEN_Y = 11
'Const TINH_TPHO_X = "T"
'Const TINH_TPHO_Y = 11
Const DIEN_THOAI_X = "I"
Const DIEN_THOAI_Y = 11
Const FAX_X = "S"
Const FAX_Y = 11
Const MAIL_X = "AG"
Const MAIL_Y = 11
'thong tin Dai ly thue
Const MA_SO_THUE_DL_X = "I"
Const MA_SO_THUE_DL_Y = 15
Const TEN_GOI_DL_X = "I"
Const TEN_GOI_DL_Y = 13
Const DIA_CHI_DL_X = "I"
Const DIA_CHI_DL_Y = 17
'Const QUAN_HUYEN_DL_X = "F"
'Const QUAN_HUYEN_DL_Y = 17
'Const TINH_TPHO_DL_X = "T"
'Const TINH_TPHO_DL_Y = 17
Const DIEN_THOAI_DL_X = "I"
Const DIEN_THOAI_DL_Y = 19
Const FAX_DL_X = "S"
Const FAX_DL_Y = 19
Const MAIL_DL_X = "AG"
Const MAIL_DL_Y = 19
Const SO_HDDL_DL_X = "I"
Const SO_HDDL_DL_Y = 21
Const NGAY_HDDL_DL_X = "S"
Const NGAY_HDDL_DL_Y = 21
'end
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'dhdang xu lý lay ma phong xu ly
'ngay 05-08-2010
Public PHONG_XU_LY_Y1 As String
Public PHONG_XU_LY_X1 As String
Public isExistFile As Boolean
Public LAN_QUET As Variant
Const STT_X = "S"
Const STT_Y = 23
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
        .Col = .ColLetterToNumber("N")
        .Row = 29
        .CellType = CellTypeEdit
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 100
        
        ' Ky lap bo
        SetDateFormat fps, 1, KY_LAP_BO_Y, .ColLetterToNumber(KY_LAP_BO_X), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        'Phong xu ly
        .Col = .ColLetterToNumber(PHONG_XU_LY_X)
        .Row = PHONG_XU_LY_Y
        .CellType = CellTypeComboBox
        .TypeEditCharSet = TypeEditCharSetASCII
        .TypeMaxEditLen = 60
        
        ' Ngay nop
        SetDateFormat fps, 1, NGAY_NOP_Y, .ColLetterToNumber(NGAY_NOP_X), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = NGAY_NOP_Y
        .Col = .ColLetterToNumber(NGAY_NOP_X)
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, NGAY_QUET_Y, .ColLetterToNumber(NGAY_QUET_X), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = NGAY_QUET_Y
        .Col = .ColLetterToNumber(NGAY_QUET_X)
        .Text = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
        .TypeHAlign = TypeHAlignLeft
        
    End With
    Prepared1 = True
End Function

Public Function Prepared2(rsPXL As ADODB.Recordset) As Boolean
    Dim i                As Integer, intIndexCombo As Integer
    Dim strLTN           As Variant, Col7 As Variant, strLTNCu As Variant, strID As Variant
    Dim xmlCellNode      As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim varID            As Variant, varDVT As Variant, varThueSuat As Variant
    Dim iTenTN           As String, strDVT As String, strThueSuat As String
    Dim iCol             As Long, iRow As Long
    Dim strIdQLT         As String, blnSuaThueSuat As Boolean
    Dim strMucThueAnDinh As String
    Dim maTM             As String
    Dim ROWNUM           As Variant

    With fps
        .Sheet = .ActiveSheet
        i = 0
        .Row = 51

        Do
            strIdQLT = ""
            varID = Empty
            varDVT = Empty
            varThueSuat = Empty
            strDVT = vbNullString
            strThueSuat = vbNullString
            
            'Lay gia tri Ma
            .GetText .ColLetterToNumber("AN"), i + 51, varID
            DataDM varID, strIdQLT, iTenTN, strDVT, strThueSuat, blnSuaThueSuat, maTM

            If strIdQLT <> vbNullString Then

                'Kiem tra hieu luc danh muc tai nguyen
                If (Not blnSuaThueSuat) And (Val(strMucThueAnDinh) = 0) And CStr(varThueSuat) <> "" And strThueSuat <> "" Then

                    If Val(strThueSuat) <> Val(CStr(varThueSuat)) Then
                        DisplayMessage "0082", msOKOnly, miCriticalError
                        Exit Function
                    End If
                End If

                .Col = .ColLetterToNumber("D")
                .Text = iTenTN
                .SetText .ColLetterToNumber("D"), .Row, iTenTN
                '                    UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("AO")
                .Text = maTM
            ElseIf CStr(varID) <> vbNullString Then
                DisplayMessage "0082", msOKOnly, miCriticalError
                Exit Function
            End If
          
            i = i + 1
            .Col = .ColLetterToNumber("B")
            .Row = i + 51
        Loop Until .Text = "aa"

        ROWNUM = .Row
        i = 0
        .Row = ROWNUM + 3

        Do
            strIdQLT = ""
            varID = Empty
            varDVT = Empty
            varThueSuat = Empty
            strDVT = vbNullString
            strThueSuat = vbNullString

            'Lay gia tri Ma
            .GetText .ColLetterToNumber("AN"), i + ROWNUM + 3, varID
            DataDM varID, strIdQLT, iTenTN, strDVT, strThueSuat, blnSuaThueSuat, maTM

            If strIdQLT <> vbNullString Then

                'Kiem tra hieu luc danh muc tai nguyen
                If (Not blnSuaThueSuat) And (Val(strMucThueAnDinh) = 0) And CStr(varThueSuat) <> "" And strThueSuat <> "" Then

                    If Val(strThueSuat) <> Val(CStr(varThueSuat)) Then
                        DisplayMessage "0082", msOKOnly, miCriticalError
                        Exit Function
                    End If
                End If

                .Col = .ColLetterToNumber("D")
                .Text = iTenTN
                .SetText .ColLetterToNumber("D"), .Row, iTenTN
                '                    UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("AO")
                .Text = maTM
            ElseIf CStr(varID) <> vbNullString Then
                DisplayMessage "0082", msOKOnly, miCriticalError
                Exit Function
            End If

            i = i + 1
            .Col = .ColLetterToNumber("B")
            .Row = i + ROWNUM + 3
        Loop Until .Text = "bb"
         
        ROWNUM = .Row
        i = 0
        .Row = ROWNUM + 3

        Do
            strIdQLT = ""
            varID = Empty
            varDVT = Empty
            varThueSuat = Empty
            strDVT = vbNullString
            strThueSuat = vbNullString

            'Lay gia tri Ma
            .GetText .ColLetterToNumber("AN"), i + ROWNUM + 3, varID
            DataDM varID, strIdQLT, iTenTN, strDVT, strThueSuat, blnSuaThueSuat, maTM

            If strIdQLT <> vbNullString Then

                'Kiem tra hieu luc danh muc tai nguyen
                If (Not blnSuaThueSuat) And (Val(strMucThueAnDinh) = 0) And CStr(varThueSuat) <> "" And strThueSuat <> "" Then

                    If Val(strThueSuat) <> Val(CStr(varThueSuat)) Then
                        DisplayMessage "0082", msOKOnly, miCriticalError
                        Exit Function
                    End If
                End If

                .Col = .ColLetterToNumber("D")
                .Text = iTenTN
                .SetText .ColLetterToNumber("D"), .Row, iTenTN
                '                    UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("AO")
                .Text = maTM
            ElseIf CStr(varID) <> vbNullString Then
                DisplayMessage "0082", msOKOnly, miCriticalError
                Exit Function
            End If

            i = i + 1
            .Col = .ColLetterToNumber("B")
            .Row = i + ROWNUM + 3
        Loop Until .Text = "cc"

    End With
      
    i = 0
    ReDim Preserve larrPhongXuLy(0)
    larrPhongXuLy(0) = "00"
    ReDim Preserve larrid(0)
    larrid(0) = "00"
    
    With fps
        .Sheet = 1
        
        .Row = 35
        .Col = .ColLetterToNumber("L")

        If .Text = "1" Then
            LOAI_TO_KHAI = "01PS/TAIN"
            .Col = .ColLetterToNumber("M")
            TAX_Utilities_Svr_New.Month = Mid$(.Text, 4, 2)
            TAX_Utilities_Svr_New.Year = Right$(.Text, 4)
        Else
            LOAI_TO_KHAI = "01/TAIN"
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
    End With

    Prepared2 = True
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

        .Col = .ColLetterToNumber(TEN_BPQL_X)
        .Row = TEN_BPQL_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))

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
        
        .Col = .ColLetterToNumber(MA_BPQL_X)
        .Row = MA_BPQL_Y
        If strMaBPQL <> "" Then
            .Text = strMaBPQL
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(TEN_BPQL_X)
        .Row = TEN_BPQL_Y
        .Text = Trim(TAX_Utilities_Svr_New.Convert(strTenBpql, TCVN, UNICODE))
        ' lan quet
        'so thu tu to khai
        .Col = .ColLetterToNumber("S")
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

Public Function InsertDTL() As String
    Dim sSQL     As String
    Dim sSQLCol  As String
    Dim sSQLVal  As String
    Dim bln      As Boolean
    Dim KYLBO    As Variant
    Dim ngnop    As Variant
    Dim madtnt   As Variant
    Dim MATHUE   As Variant
    Dim MAPP     As Variant
    Dim maTM     As Variant
    Dim NganhTN  As Variant
    Dim SanLuong As Variant
    Dim DVTinh   As Variant
    Dim DONGIA   As Variant
    Dim GTTT     As Variant
    Dim THUESUAT As Variant
    Dim MucThue  As Variant
    Dim THUEPS   As Variant
    Dim MG       As Variant
    Dim THUETN   As Variant
    Dim matkhai  As Variant
    Dim kykkhai  As Variant
    Dim MaMuc    As Variant
    Dim DANHAN   As Variant
    Dim ROWNUM   As Integer
    Dim TTHTK1   As Variant
    Dim STT      As Variant
    Dim STT2     As Variant
    Dim STTIN    As Variant
    Dim mact     As Variant

    sSQLCol = "KYLBO, NGNOP, MADTNT, MATHUE, MAPP, MAMUC, MATM, NGANHTN, SANLUONG, DVTINH, DONGIA "
    sSQLCol = sSQLCol + ", GTTT, THUESUAT, MUCTHUE, THUEPS, MG, THUETN, TTHTK, MATKHAI, KYKKHAI, LAN_QUET, DANHAN, STT, STT2, STTIN, MaCT"
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
   
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber(MA_SO_THUE_X), MA_SO_THUE_Y, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber(NGAY_NOP_X), NGAY_NOP_Y, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber(KY_LAP_BO_X), KY_LAP_BO_Y, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
                
        MATHUE = "'04'"
        MAPP = "'1'"
        STT = "1"
        STT2 = "0"
        STTIN = "'I'"
        mact = "'1000'"
        
        ' sua MAMUC & MATM tu nam 2009
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'03'"
            MaMuc = "'012'"
        Else
            MaMuc = "'1550'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
             
        maTM = "''"
        NganhTN = "''"
        DVTinh = "''"
        SanLuong = "0"

        DONGIA = "0"
        GTTT = "0"
        THUESUAT = "0"
        MucThue = "0"
        THUEPS = "0"
        MG = "0"
        THUETN = "0"
    End With

    sSQLVal = KYLBO & "," & ngnop & "," & madtnt & "," & MATHUE & "," & MAPP & "," & MaMuc & "," & maTM & "," & NganhTN & "," & SanLuong & "," & DVTinh & "," & DONGIA & "," & GTTT & "," & THUESUAT & "," & MucThue & "," & THUEPS & "," & MG & "," & THUETN & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN & "," & STT & "," & STT2 & "," & STTIN & "," & mact
    sSQL = "INSERT INTO TMP_TA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    bln = clsDAO.ExecuteDLL(sSQL)
   
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber(MA_SO_THUE_X), MA_SO_THUE_Y, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber(NGAY_NOP_X), NGAY_NOP_Y, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber(KY_LAP_BO_X), KY_LAP_BO_Y, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
                
        MATHUE = "'04'"
        MAPP = "'1'"
        
        ' sua MAMUC & MATM tu nam 2009
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'03'"
            MaMuc = "'012'"
        Else
            MaMuc = "'1550'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        .Col = .ColLetterToNumber("B")
        .Row = 51
        
        Do
            STT = "1"
            STT2 = STT2 + 1
            STTIN = "'" & STT2 & "'"
            mact = "''"
            .Col = .ColLetterToNumber("AO")
            maTM = .Text

            If Trim(maTM) = vbNullString Then
                maTM = "''"
            Else
                maTM = "'" & maTM & "'"
            End If
               
            .Col = .ColLetterToNumber("AN")
            NganhTN = .Text

            If Trim(NganhTN) = vbNullString Then
                NganhTN = "''"
            Else
                NganhTN = "'" & NganhTN & "'"
            End If
                
            .Col = .ColLetterToNumber("M")
            DVTinh = .Text

            If Trim(DVTinh) = vbNullString Then
                DVTinh = "''"
            Else
                DVTinh = "'" & TAX_Utilities_Svr_New.Convert(Trim(DVTinh), UNICODE, TCVN) & "'"
            End If
                
            .Col = .ColLetterToNumber("P")
            SanLuong = .Value

            If Trim(SanLuong) = vbNullString Then
                SanLuong = "0"
            End If
               
            .Col = .ColLetterToNumber("S")

            DONGIA = .Value

            If Trim(DONGIA) = vbNullString Then

                DONGIA = "0"
            End If
               
            GTTT = CDbl(SanLuong) * CDbl(DONGIA)
               
            .Col = .ColLetterToNumber("X")
            THUESUAT = .Value

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            End If
                
            .Col = .ColLetterToNumber("AA")
            MucThue = .Value

            If Trim(MucThue) = vbNullString Then
                MucThue = "0"
            End If
                
            .Col = .ColLetterToNumber("AF")
            THUEPS = .Value

            If Trim(THUEPS) = vbNullString Then
                THUEPS = "0"
            End If
                
            .Col = .ColLetterToNumber("AI")
            MG = .Value

            If Trim(MG) = vbNullString Then
                MG = "0"
            End If
                
            .Col = .ColLetterToNumber("AM")
            THUETN = .Value

            If Trim(THUETN) = vbNullString Then
                THUETN = "0"
            End If
        
            sSQLVal = KYLBO & "," & ngnop & "," & madtnt & "," & MATHUE & "," & MAPP & "," & MaMuc & "," & maTM & "," & NganhTN & "," & SanLuong & "," & DVTinh & "," & DONGIA & "," & GTTT & "," & THUESUAT & "," & MucThue & "," & THUEPS & "," & MG & "," & THUETN & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN & "," & STT & "," & STT2 & "," & STTIN & "," & mact
            sSQL = "INSERT INTO TMP_TA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If maTM <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"

        ROWNUM = .Row
    End With

    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber(MA_SO_THUE_X), MA_SO_THUE_Y, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber(NGAY_NOP_X), NGAY_NOP_Y, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber(KY_LAP_BO_X), KY_LAP_BO_Y, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
                
        MATHUE = "'04'"
        MAPP = "'1'"
        STT = "2"
        STT2 = "0"
        STTIN = "'II'"
        mact = "'2000'"
        
        ' sua MAMUC & MATM tu nam 2009
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'03'"
            MaMuc = "'012'"
        Else
            MaMuc = "'1550'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
             
        maTM = "''"
        NganhTN = "''"
        DVTinh = "''"
        SanLuong = "0"

        DONGIA = "0"
        GTTT = "0"
        THUESUAT = "0"
        MucThue = "0"
        THUEPS = "0"
        MG = "0"
        THUETN = "0"
    End With

    sSQLVal = KYLBO & "," & ngnop & "," & madtnt & "," & MATHUE & "," & MAPP & "," & MaMuc & "," & maTM & "," & NganhTN & "," & SanLuong & "," & DVTinh & "," & DONGIA & "," & GTTT & "," & THUESUAT & "," & MucThue & "," & THUEPS & "," & MG & "," & THUETN & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN & "," & STT & "," & STT2 & "," & STTIN & "," & mact
    sSQL = "INSERT INTO TMP_TA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    bln = clsDAO.ExecuteDLL(sSQL)
   
    'du lieu secsion II
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber(MA_SO_THUE_X), MA_SO_THUE_Y, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber(NGAY_NOP_X), NGAY_NOP_Y, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber(KY_LAP_BO_X), KY_LAP_BO_Y, KYLBO
        
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
                
        MATHUE = "'04'"
        MAPP = "'1'"
        
        mact = "''"
         
        ' sua MAMUC & MATM tu nam 2009
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'03'"
            MaMuc = "'012'"
        Else
            MaMuc = "'1550'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        .Col = .ColLetterToNumber("B")
        .Row = ROWNUM + 3

        Do
            STT = "2"
            STT2 = STT2 + 1
            STTIN = "'" & STT2 & "'"
            .Col = .ColLetterToNumber("AO")
            maTM = .Text

            If Trim(maTM) = vbNullString Then
                maTM = "''"
            Else
                maTM = "'" & maTM & "'"
            End If
                    
            .Col = .ColLetterToNumber("AN")
            NganhTN = .Text

            If Trim(NganhTN) = vbNullString Then
                NganhTN = "''"
            Else
                NganhTN = "'" & NganhTN & "'"
            End If
                     
            .Col = .ColLetterToNumber("M")
            DVTinh = .Text

            If Trim(DVTinh) = vbNullString Then
                DVTinh = "''"
            Else
                DVTinh = "'" & DVTinh & "'"
            End If
                     
            .Col = .ColLetterToNumber("P")
            SanLuong = .Value

            If Trim(SanLuong) = vbNullString Then
                SanLuong = "0"
            End If
                    
            .Col = .ColLetterToNumber("S")

            DONGIA = .Value

            If Trim(DONGIA) = vbNullString Then

                DONGIA = "0"
            End If
                    
            GTTT = CDbl(SanLuong) * CDbl(DONGIA)
                    
            .Col = .ColLetterToNumber("X")
            THUESUAT = .Value

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            End If
                     
            .Col = .ColLetterToNumber("AA")
            MucThue = .Value

            If Trim(MucThue) = vbNullString Then
                MucThue = "0"
            End If
                     
            .Col = .ColLetterToNumber("AF")
            THUEPS = .Value

            If Trim(THUEPS) = vbNullString Then
                THUEPS = "0"
            End If
                     
            .Col = .ColLetterToNumber("AI")
            MG = .Value

            If Trim(MG) = vbNullString Then
                MG = "0"
            End If
                     
            .Col = .ColLetterToNumber("AM")
            THUETN = .Value

            If Trim(THUETN) = vbNullString Then
                THUETN = "0"
            End If
         
            sSQLVal = KYLBO & "," & ngnop & "," & madtnt & "," & MATHUE & "," & MAPP & "," & MaMuc & "," & maTM & "," & NganhTN & "," & SanLuong & "," & DVTinh & "," & DONGIA & "," & GTTT & "," & THUESUAT & "," & MucThue & "," & THUEPS & "," & MG & "," & THUETN & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN & "," & STT & "," & STT2 & "," & STTIN & "," & mact
            sSQL = "INSERT INTO TMP_TA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If maTM <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "bb"
        ROWNUM = .Row
    End With
      
      
        With fps
        .Sheet = 1
        .GetText .ColLetterToNumber(MA_SO_THUE_X), MA_SO_THUE_Y, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber(NGAY_NOP_X), NGAY_NOP_Y, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber(KY_LAP_BO_X), KY_LAP_BO_Y, KYLBO
        
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
                
        MATHUE = "'04'"
        MAPP = "'1'"
        STT = "3"
        STT2 = "0"
        STTIN = "'III'"
        mact = "'3000'"
        
        ' sua MAMUC & MATM tu nam 2009
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'03'"
            MaMuc = "'012'"
        Else
            MaMuc = "'1550'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
             
        maTM = "''"
        NganhTN = "''"
        DVTinh = "''"
        SanLuong = "0"

        DONGIA = "0"
        GTTT = "0"
        THUESUAT = "0"
        MucThue = "0"
        THUEPS = "0"
        MG = "0"
        THUETN = "0"
    End With

    sSQLVal = KYLBO & "," & ngnop & "," & madtnt & "," & MATHUE & "," & MAPP & "," & MaMuc & "," & maTM & "," & NganhTN & "," & SanLuong & "," & DVTinh & "," & DONGIA & "," & GTTT & "," & THUESUAT & "," & MucThue & "," & THUEPS & "," & MG & "," & THUETN & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN & "," & STT & "," & STT2 & "," & STTIN & "," & mact
    sSQL = "INSERT INTO TMP_TA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
    bln = clsDAO.ExecuteDLL(sSQL)
    
    'du lieu secsion III
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber(MA_SO_THUE_X), MA_SO_THUE_Y, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
        
        .GetText .ColLetterToNumber(NGAY_NOP_X), NGAY_NOP_Y, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
                    
        .GetText .ColLetterToNumber(KY_LAP_BO_X), KY_LAP_BO_Y, KYLBO
        
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "CTOD('')"
        Else
            KYLBO = "'" & KYLBO & "'"
        End If
                
        MATHUE = "'04'"
        MAPP = "'1'"
        
        mact = "''"
         
        ' sua MAMUC & MATM tu nam 2009
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'03'"
            MaMuc = "'012'"
        Else
            MaMuc = "'1550'"
        End If
        
        matkhai = "'" & LOAI_TO_KHAI & "'"
        DANHAN = "''"
        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        .Col = .ColLetterToNumber("B")
        .Row = ROWNUM + 3

        Do
            STT = "3"
            STT2 = STT2 + 1
            STTIN = "'" & STT2 & "'"
            .Col = .ColLetterToNumber("AO")
            maTM = .Text

            If Trim(maTM) = vbNullString Then
                maTM = "''"
            Else
                maTM = "'" & maTM & "'"
            End If
                    
            .Col = .ColLetterToNumber("AN")
            NganhTN = .Text

            If Trim(NganhTN) = vbNullString Then
                NganhTN = "''"
            Else
                NganhTN = "'" & NganhTN & "'"
            End If
                     
            .Col = .ColLetterToNumber("M")
            DVTinh = .Text

            If Trim(DVTinh) = vbNullString Then
                DVTinh = "''"
            Else
                DVTinh = "'" & DVTinh & "'"
            End If
                     
            .Col = .ColLetterToNumber("P")
            SanLuong = .Value

            If Trim(SanLuong) = vbNullString Then
                SanLuong = "0"
            End If
                    
            .Col = .ColLetterToNumber("S")

            DONGIA = .Value

            If Trim(DONGIA) = vbNullString Then

                DONGIA = "0"
            End If
                    
            GTTT = CDbl(SanLuong) * CDbl(DONGIA)
                    
            .Col = .ColLetterToNumber("X")
            THUESUAT = .Value

            If Trim(THUESUAT) = vbNullString Then
                THUESUAT = "0"
            End If
                     
            .Col = .ColLetterToNumber("AA")
            MucThue = .Value

            If Trim(MucThue) = vbNullString Then
                MucThue = "0"
            End If
                     
            .Col = .ColLetterToNumber("AF")
            THUEPS = .Value

            If Trim(THUEPS) = vbNullString Then
                THUEPS = "0"
            End If
                     
            .Col = .ColLetterToNumber("AI")
            MG = .Value

            If Trim(MG) = vbNullString Then
                MG = "0"
            End If
                     
            .Col = .ColLetterToNumber("AM")
            THUETN = .Value

            If Trim(THUETN) = vbNullString Then
                THUETN = "0"
            End If
         
            sSQLVal = KYLBO & "," & ngnop & "," & madtnt & "," & MATHUE & "," & MAPP & "," & MaMuc & "," & maTM & "," & NganhTN & "," & SanLuong & "," & DVTinh & "," & DONGIA & "," & GTTT & "," & THUESUAT & "," & MucThue & "," & THUEPS & "," & MG & "," & THUETN & "," & TTHTK & "," & matkhai & "," & kykkhai & "," & LAN_QUET & "," & DANHAN & "," & STT & "," & STT2 & "," & STTIN & "," & mact
            sSQL = "INSERT INTO TMP_TA" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"

            If maTM <> "''" Then
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "cc"

    End With

    clsDAO.Disconnect
   
    InsertDTL = vbNullString
   
End Function
Public Function InsertDTL_KHBS() As String
    Dim sSQL As String
    Dim sSQLCol As String
    Dim sSQLVal As String
    Dim bln  As Boolean
    Dim madtnt As Variant
    Dim ngnop As Variant
    Dim KYLBO As Variant
    Dim kykkhai As Variant
    Dim matkhai As Variant
    Dim MATHUE As Variant
    Dim MaMuc As Variant
    Dim maTM As Variant
    Dim MAPP As Variant
    Dim mact As Variant
    Dim MACT2 As Variant
    Dim MACT3 As Variant
    Dim STT As Variant
    Dim STT2 As Variant
    Dim STTIN As Variant
    Dim SOKK As Variant
    Dim SODC As Variant
    Dim SOCL As Variant
    Dim SONGAYNC As Variant
    Dim SOTIENNC As Variant
    Dim i As Integer, iCol As Long, iRow As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim strFileName As String
    Dim fso As New FileSystemObject
    Dim DANHAN As Variant
    Dim arrMTM() As String, arrMTMSave() As String
    Dim arrVitri() As Integer
    Dim idx, idx1, countMTM As Integer
    Dim vTemp As String, vVitriTrung As String
    Dim ROWNUM As Variant
   'kiem tra ton tai tep *.dbf chua
    strFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_BS" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileName) = False Then
    Dim strTepmau As String
        strTepmau = spathVat & "\tepmau\TMP_BS.DBF"
        fso.CopyFile strTepmau, strFileName, False
    End If

   sSQLCol = "MADTNT, KYKKHAI, MATKHAI, MATHUE, MAMUC, MATM, MAPP, NGNOP, TTHTK, KYLBO , MACT, MACT2, MACT3, STT, STT2, STTIN," & _
             " SOKK, SODC, SOCL,SOTIENNC, SONGAYNC, LAN_QUET, DANHAN"

   If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
   End If


   With fps
        
        .Sheet = 1
        'Lay ds ma tieu muc
       ' layt ma tieu muc theo yeu cau chi Em
        'ngay 14/09/2010
        .Col = .ColLetterToNumber("B")
        .Row = 51
        Do
            ' Lay ma tieu muc tren cel
            .Col = .ColLetterToNumber("AO")
            vTemp = vTemp & .Text & "~"
                     
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
         Loop Until .Text = "cc"
         arrMTM = Split(vTemp, "~")
         vTemp = ""
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
         'the end
         
        .GetText .ColLetterToNumber(MA_SO_THUE_X), MA_SO_THUE_Y, madtnt
        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If

        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
        matkhai = "'" & LOAI_TO_KHAI & "'"
        MATHUE = "'01'"
        .GetText .ColLetterToNumber(KY_LAP_BO_X), KY_LAP_BO_Y, KYLBO
        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else
             If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'03'"
            MaMuc = "'012'"
        Else
            MaMuc = "'1550'"
        '    matm = "''"
        End If
        vTemp = arrMTMSave(i - 1)
        maTM = "'" & vTemp & "'"
        MAPP = "'2'"
        
        .GetText .ColLetterToNumber(NGAY_NOP_X), NGAY_NOP_Y, ngnop
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
        DANHAN = "''"
        
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

'Xoa phan tu cua 1 mang tai vi tri idx

Private Function DelPosition(idx As Integer, arrDel() As String)
    Dim i As Integer
    For i = idx To UBound(arrDel) - 1
        arrDel(i) = arrDel(i + 1)
    Next
    ReDim Preserve arrDel(UBound(arrDel) - 1)
End Function

Public Function InsertHDR() As String
    Dim sSQL       As String
    Dim sSQLCol    As String
    Dim sSQLVal    As String
    
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
    Dim DAGHI2     As Variant
    Dim AUTO_CAL   As Variant
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
    Dim TEMP       As Variant
    Dim TEMP2      As Variant
    Dim THUECL     As Variant
    Dim NGANH      As Variant
    Dim bln        As Boolean
    Dim SanLuong   As Variant
    Dim DVTinh     As Variant
    Dim DONGIA     As Variant
    Dim GTTT       As Variant
    Dim THUESUAT   As Variant
    Dim MucThue    As Variant
    Dim THUEPS     As Variant
    Dim MG         As Variant
    Dim THUETN     As Variant
    Dim ROWNUM     As Variant
    Dim madlt      As Variant
    Dim LanBS      As Variant
    Dim solanBS    As Variant
    Dim TTHTK1     As Variant
    Dim DANHAN     As Variant
    Dim DSoCL      As Variant
    Dim GTTTCL     As Variant
    Dim MANNKD     As Variant
    Dim NgayPs     As Variant
    Dim NGAYNOP As Variant
    
    Dim TENNV      As Variant
    Dim CHUNGCHI   As Variant

    'test
    Dim arrMTM()   As String, arrMTMSave() As String
    Dim arrVitri() As Integer
    Dim idx, idx1, i, countMTM As Integer
    Dim vTemp As String, vVitriTrung As String
    'end test
    
    sSQLCol = "MADTNT, MATHUE, MATM, MAPP, KYLBO, NGNOP, "
    sSQLCol = sSQLCol & "LOAITIEN, MABPQL, SHTEP, TTHTK, KYKKHAI, "
    sSQLCol = sSQLCol & "NGNHAP, THUETKY, GTTINHTHUE, DAGHI,  "
    sSQLCol = sSQLCol & "KTTRUOC, DOANHSO, MAPP2, MAMUC, MAVACH,MATKHAI,HANNOP, THUETKY2, HANNOP2, TENDA, DAGHI1 ,DAGHI2, AUTO_CAL, NGANH,LAN_QUET, MADLT, LANBS, THUECL,DSOCL, GTTTCL, MANNKD, DANHAN, NgayPS, TENNV, CHUNGCHI"
   
    If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
    End If
       
    With fps
        .Sheet = 1
        
        .GetText .ColLetterToNumber("T"), .MaxRows - 5, TENNV
        .GetText .ColLetterToNumber("T"), .MaxRows - 3, CHUNGCHI

        'Lay ds ma tieu muc
        'dhdang phan lay vao mang
        'ngay 14/09/2010
        .Col = .ColLetterToNumber("B")
        .Row = 51

        Do
            ' Lay ma tieu muc tren cel
            .Col = .ColLetterToNumber("AO")
            vTemp = vTemp & .Text & "~"
                     
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "cc"

        If vTemp <> "" Then
            arrMTM = Split(vTemp, "~")
            'vTemp = ""
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
        ' ma dai ly thue
        .GetText .ColLetterToNumber(MA_SO_THUE_DL_X), MA_SO_THUE_DL_Y, madlt

        If Trim(madlt) = vbNullString Then
            madlt = "''"
        Else
            madlt = "'" & madlt & "'"
        End If

        'trang thai to khai
        .GetText .ColLetterToNumber("V"), 4, TTHTK1

        If Trim(TTHTK1) = vbNullString Then
            TTHTK = "'" & 2 & "'"
        End If

        ' so lan bo sung
        If Trim(TTHTK1) = vbNullString Then
            .GetText .ColLetterToNumber("AD"), 4, solanBS
            LanBS = solanBS
        Else
            LanBS = "0"
        End If

        ' ma doi tuong not thue
        .GetText .ColLetterToNumber(MA_SO_THUE_X), MA_SO_THUE_Y, madtnt

        If Trim(madtnt) = vbNullString Then
            madtnt = "''"
        Else
            madtnt = "'" & madtnt & "'"
        End If
       
        .GetText .ColLetterToNumber(KY_LAP_BO_X), KY_LAP_BO_Y, KYLBO

        If Trim(KYLBO) = vbNullString Then
            KYLBO = "''"
        Else

            If Len(Trim(KYLBO)) = 6 Then
                KYLBO = "'0" & KYLBO & "'"
            Else
                KYLBO = "'" & KYLBO & "'"
            End If
        End If

        .GetText .ColLetterToNumber(NGAY_NOP_X), NGAY_NOP_Y, ngnop

        'NGNOP = Date
        If Trim(ngnop) = vbNullString Then
            ngnop = "CTOD('')"
        Else
            ngnop = ToDate(Trim(ngnop), DDMMYYYY)
            NGAYNOP = ngnop
            ngnop = "CTOD('" & Format(ngnop, "mm/dd/yyyy") & "')"
        End If
       
        .GetText .ColLetterToNumber(MA_BPQL_X), MA_BPQL_Y, MABPQL

        If Trim(MABPQL) = vbNullString Then
            MABPQL = "''"
        Else
            MABPQL = "'" & MABPQL & "'"
        End If
        
        '        .GetText .ColLetterToNumber("R"), 23, SHTEP
        '        If Trim(SHTEP) = vbNullString Then
        '            SHTEP = "''"
        '        Else
        '            SHTEP = "'" & SHTEP & "'"
        '        End If
        SHTEP = "''"
        
        .Row = 35
        .Col = .ColLetterToNumber("L")

        If .Text = "1" Then
            .Col = .ColLetterToNumber("M")
            NgayPs = .Text
            HANNOP = DateAdd("d", 10, NGAYNOP)
            HANNOP = ToString(HANNOP, "DD/MM/YYYY")
        End If

        kykkhai = "'" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"

        NgNhap = Date

        If Trim(NgNhap) = vbNullString Then
            NgNhap = "CTOD('')"
        Else
            NgNhap = "CTOD('" & Format(NgNhap, "mm/dd/yyyy") & "')"
        End If

        LoaiTien = "''"
        MATHUE = "'04'"
        MAPP = "'1'"
        MAPP2 = "'1'"
        MaVach = ".T."
        matkhai = "'" & LOAI_TO_KHAI & "'"
        THUECL = "0"
        ' ngay phat sinh
        '        .GetText .ColLetterToNumber("AG"), 42, NGAYPS

        If (NgayPs) = vbNullString Then
            NgayPs = "CTOD('')"
        Else
            NgayPs = ToDate(Trim(NgayPs), DDMMYYYY)
            NgayPs = "CTOD('" & Format(NgayPs, "mm/dd/yyyy") & "')"
        End If

        ' sua MAMUC & MATM tu nam 2009
        If CLng(Mid(Right(KYLBO, 5), 1, 4)) < 2009 Then
            maTM = "'03'"
            MaMuc = "'012'"
        Else
            MaMuc = "'1550'"
            '    matm = "''"
        End If
       
        If Trim(HANNOP) = vbNullString Then
            HANNOP = "CTOD('')"
        Else
            HANNOP = ToDate(Trim(HANNOP), DDMMYYYY)
            HANNOP = "CTOD('" & Format(HANNOP, "mm/dd/yyyy") & "')"
        End If

        DANHAN = "''"
        DSoCL = "0"
        GTTTCL = "0"
        .GetText .ColLetterToNumber("AC"), 34, MANNKD

        If (MANNKD) = vbNullString Then
            MANNKD = "''"
        Else
            MANNKD = "'" & MANNKD & "'"
        End If
        
        For i = 0 To UBound(arrMTMSave) - 1
            ThueTKy = 0
            GTTinhThue = 0
            vTemp = arrMTMSave(i)
            maTM = "'" & vTemp & "'"
            .Col = .ColLetterToNumber("B")
            .Row = 51

            Do
                .Col = .ColLetterToNumber("AO")

                If vTemp = Trim(.Text) Then
                    .Col = .ColLetterToNumber("AM")
                    TEMP = .Value

                    If Trim(TEMP) = vbNullString Then
                        TEMP = "0"
                    End If

                    ThueTKy = ThueTKy + CDbl(TEMP)
                            
                    .Col = .ColLetterToNumber("P")
                    TEMP = .Value

                    If Trim(TEMP) = vbNullString Then
                        TEMP = "0"
                    End If

                    .Col = .ColLetterToNumber("S")
                    TEMP2 = .Value

                    If Trim(TEMP2) = vbNullString Then
                        TEMP2 = "0"
                    End If

                    GTTinhThue = GTTinhThue + CDbl(TEMP) * CDbl(TEMP2)
                End If

                .Col = .ColLetterToNumber("B")
                .Row = .Row + 1
            Loop Until .Text = "cc"

            ROWNUM = .Row
            
            ThueTKy2 = "0"
            KTTRUOC = "0"

            DoanhSo = GTTinhThue
            HANNOP2 = HANNOP
            TENDA = "''"
            DaGhi = ".T."
            MaVach = ".T."
            DAGHI1 = ".F."

            If checkDLTTonTai = False Then
                DAGHI2 = ".T."
            Else
                DAGHI2 = ".F."
            End If

            AUTO_CAL = ".F."
            NGANH = TinhMaNganh

            If Trim(maTM) <> "" Or Trim(maTM) <> "''" Then
                sSQLVal = madtnt & "," & MATHUE & "," & maTM & "," & MAPP & "," & KYLBO & "," & ngnop & ", "
                sSQLVal = sSQLVal & LoaiTien & "," & MABPQL & "," & SHTEP & "," & TTHTK & "," & kykkhai & ", "
                sSQLVal = sSQLVal & NgNhap & "," & ThueTKy & "," & GTTinhThue & "," & DaGhi & ", "
                sSQLVal = sSQLVal & KTTRUOC & "," & DoanhSo & "," & MAPP2 & "," & MaMuc & "," & MaVach & ", "
                sSQLVal = sSQLVal & matkhai & "," & HANNOP & "," & ThueTKy2 & "," & HANNOP2 & "," & TENDA & "," & DAGHI1 & "," & DAGHI2 & "," & AUTO_CAL & ",'" & NGANH & "'," & LAN_QUET & "," & madlt & "," & LanBS & "," & THUECL & "," & DSoCL & "," & GTTTCL & "," & MANNKD & "," & DANHAN & "," & NgayPs & ",'" & TENNV & "','" & CHUNGCHI & "'"
                   
                sSQL = "INSERT INTO TMP_TK" & sKyKeKhai & "( " & sSQLCol & " ) VALUES( " & sSQLVal & " )"
                bln = clsDAO.ExecuteDLL(sSQL)
            End If

        Next

    End With
    
    clsDAO.Disconnect
    InsertHDR = vbNullString
    
End Function
Public Function TKTT() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim fso As New FileSystemObject
    Dim strFileNameHDR As String
    Dim strFileNameDTL As String
    Dim strFileNameBS As String
    
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
    strFileNameDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TA" & sKyKeKhai & ".DBF"
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
        strTepmauDTL = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "\tepmau\TMP_TA.DBF"
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
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyKeKhai & _
        " WHERE MATHUE = '04' AND  MAPP = '1' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='" & LOAI_TO_KHAI & "'"
    
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
   
   sSQL = "UPDATE TMP_TK" & sKyKeKhai & _
        " SET THUETKY2 = 1 " & _
        " WHERE MATHUE = '04' and MAPP = '1' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='" & LOAI_TO_KHAI & "'"

   clsDAO.ExecuteDLL sSQL

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
   
   sSQL = "DELETE FROM TMP_TK" & sKyKeKhai & _
        " WHERE MATHUE = '04'  AND MAPP = '1' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='" & LOAI_TO_KHAI & "'"
    
   clsDAO.ExecuteDLL sSQL
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber(KY_LAP_BO_X), KY_LAP_BO_Y, KYLBO
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
   sSQL = "DELETE FROM TMP_TA" & sKyKeKhai & _
        " WHERE MATHUE = '04' AND  MAPP = '1' AND MADTNT = " & "'" & strMST & "' AND KYLBO = " & KYLBO & _
        " AND MATKHAI='" & LOAI_TO_KHAI & "' AND KYKKHAI ='" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'"
    
   clsDAO.ExecuteDLL sSQL
   
   XoaTKTT = True
   
   clsDAO.Disconnect

End Function

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

                    If Val(TAX_Utilities_Svr_New.Year) > 2010 Then
                        maTM = arrDanhsach(7)
                    ElseIf Val(TAX_Utilities_Svr_New.Month) >= 4 And Val(TAX_Utilities_Svr_New.Year) = 2010 Then
                        '                                MATM = arrDanhsach(8)
                        '                            Else
                        maTM = arrDanhsach(7)
                    End If

                    Exit Sub
                End If
            End If

        Next

    End If

End Sub

Public Sub GetParams(ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String) 'ByRef strPMaSoTep, ByRef strPNgayNhanToKhai As String, ByRef strPPhongXuLy As String)
'    strPMaSoTep = strMaSoTep
    strPNgayNhanToKhai = strNgayNhanToKhai
    strPPhongXuLy = strPhongXuLy
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

'Private Sub fps_KeyPress(KeyAscii As Integer)
'    If fps.ActiveCol = fps.ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) And fps.ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then _
        CellEditFormatNumber fps, fps.ActiveSheet, fps.ActiveCol, fps.ActiveRow, KeyAscii
'    With fps
'        If .ActiveSheet = 1 And .ActiveCol = .ColLetterToNumber(HEADER_SO_TT_TRONG_TEP_X) _
'           And .ActiveRow = HEADER_SO_TT_TRONG_TEP_Y Then
'            If KeyAscii = 46 Or KeyAscii = 45 Then ' KeyAscii of '.' character
'                KeyAscii = 0
'            End If
'        End If
'    End With
'End Sub


Public Function isToKhaiChinhThuc() As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim vKeKhai As Variant
    Dim vKeKhai_temp As Variant
    Dim vKyLB As Variant
    'them bien de kiem tra su ton tai cua bang dl
    Dim strTestFileName As String
    Dim fso As New FileSystemObject
    
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
   Do While flag
        strTestFileName = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & vKeKhai_temp & ".DBF"
        'kiem tra neu ton tai bang thi moi thuc hien truy van
        If fso.FileExists(strTestFileName) Then
             sSQL = "SELECT * FROM TMP_TK" & vKeKhai_temp & _
             " WHERE MATHUE = '04' AND  MADTNT = " & "'" & strMST & "'" & _
             " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "'" & _
             " AND (TTHTK = '1' or TTHTK = '3' or TTHTK = '4') AND MATKHAI = '" & LOAI_TO_KHAI & "'"
        
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

Public Function CheckValidData() As Boolean
    Dim varCheckValue As Variant, varNoteValue As Variant
    
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("I"), 13, varCheckValue
        .GetText .ColLetterToNumber("T"), 13, varNoteValue
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
Public Function TKTTNGNOP() As Boolean
   Dim ngnop As Variant
   Dim NGNOPDB As Variant
   Dim sSQL As String
   Dim THUEGTGT_KT As Double
   Dim rs As ADODB.Recordset
   Dim i As Integer
   
   With fps
        .Sheet = 1
        .GetText .ColLetterToNumber(NGAY_NOP_X), NGAY_NOP_Y, ngnop
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
   
   sSQL = "SELECT MADTNT, NGNOP FROM TMP_TK" & sKyKeKhai & _
        " WHERE MATHUE = '04' AND  MAPP = '1' AND MADTNT = " & "'" & strMST & "'" & _
        " AND KYKKHAI <> '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='" & LOAI_TO_KHAI & "'"
    
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

Public Function TKRB() As Boolean

   TKRB = True

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
        If Col = .ColLetterToNumber(NGAY_NOP_X) And Row = NGAY_NOP_Y Then
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
        If Col = .ColLetterToNumber(NGAY_QUET_X) And Row = NGAY_QUET_Y Then
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
        If Col = .ColLetterToNumber(KY_LAP_BO_X) And Row = KY_LAP_BO_Y Then
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
    strFileNameHDR = spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile") & "TMP_TK" & sKyKeKhai & ".DBF"
    If fso.FileExists(strFileNameHDR) = True Then
        If clsDAO.Connected = False Then
        clsDAO.CreateConnectionString spathVat & GetAttribute(TAX_Utilities_Svr_New.NodeMenu, "DirFile")
        clsDAO.Connect
        End If
                  
            sSQL = "SELECT max(LAN_QUET) as LAN_QUET  FROM TMP_TK" & sKyKeKhai & _
                " WHERE MATHUE = '04' AND  MAPP = '1' AND MADTNT = " & "'" & strMST & "'" & _
                " AND KYKKHAI = '" & TAX_Utilities_Svr_New.Month & "/" & TAX_Utilities_Svr_New.Year & "' AND MATKHAI='" & LOAI_TO_KHAI & "' AND (TTHTK = '1' or TTHTK = '2' or TTHTK = '3' or TTHTK = '4')"
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

