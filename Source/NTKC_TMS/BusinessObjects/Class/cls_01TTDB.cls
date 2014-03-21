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
Const MA_SO_THUE_X = "I"
Const MA_SO_THUE_Y = 7
Const NGAY_NOP_Y = 25
Const NGAY_NOP_X = "F"
Const NGAY_QUET_Y = 25
Const NGAY_QUET_X = "S"
Const TEN_GOI_X = "I"
Const TEN_GOI_Y = 5
Const DIA_CHI_X = "I"
Const DIA_CHI_Y = 9
Const MA_BPQL_X = "AG"
Const MA_BPQL_Y = 27
Const DIEN_THOAI_X = "I"
Const DIEN_THOAI_Y = 11
Const FAX_X = "S"
Const FAX_Y = 11

Const KY_LAP_BO_Y = 23
Const KY_LAP_BO_X = "F"

Const HEADER_KY_LAP_BO_ROW = 13
Const HEADER_KY_LAP_BO_COL = "B"

Const NGAY_NHAN_TO_KHAI_Y = 25
Const NGAY_NHAN_TO_KHAI_X = "F"
Const NGUOI_SU_DUNG_Y = 25
Const NGUOI_SU_DUNG_X = "AG"
Const HEADER_SO_TT_TRONG_TEP_X = "M"
Const HEADER_SO_TT_TRONG_TEP_Y = 39
Const PHONG_XU_LY_Y = 23
Const PHONG_XU_LY_X = "AG"
Const CHECK_COL = "G"
Const CHECK_ROW = 29

Const CHECK_COL1 = "N"
Const CHECK_ROW1 = 49

' thong tin dai ly thue
Const TEN_DL_ROW = 13
Const TEN_DL_COL = "I"

Const DIA_CHI_DL_ROW = 17
Const DIA_CHI_DL_COL = "I"

Const DIEN_THOAI_DL_ROW = 19
Const DIEN_THOAI_DL_COL = "I"

Const FAX_DL_ROW = 19
Const FAX_DL_COL = "S"

Const EMAIL_DL_ROW = 19
Const EMAIL_DL_COL = "AG"

Const SO_HOP_DONG_DL_ROW = 21
Const SO_HOP_DONG_DL_COL = "I"

Const NGAY_HOP_DONG_DL_ROW = 21
Const NGAY_HOP_DONG_DL_COL = "S"

' Thong tin nganh nghe KD va to khai PS
Const NGANH_NGHE_KD_ROW = 33
Const NGANH_NGHE_KD_COL = "M"
' Ngay phat sinh
Const NGAY_PS_ROW = 33
Const NGAY_PS_COL = "L"


Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
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
Public sKyKeKhai As String

Public strTenDL As String
Public strDiaChiDL As String
Public strDienThoaiDL As String
Public strFaxDL As String
Public strEmailDL As String
Public strSoHopDongDL As String
Public strNgayHopDongDL As String

Public Function Prepared1() As Boolean
    With fps
        .Sheet = 1
        
        'Ghi chu
        .Col = .ColLetterToNumber("S")
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
    
        'Ky lap bo
        SetDateFormat fps, 1, 23, .ColLetterToNumber("F"), MMYYYY
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay nop
        SetDateFormat fps, 1, 25, .ColLetterToNumber("F"), DDMMYYYY
        .Sheet = 1
        .Row = 25
        .Col = .ColLetterToNumber("F")
        .SetText .Col, .Row, Format(Date, "dd/mm/yyyy")
        '.Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        ' Ngay quet
        SetDateFormat fps, 1, 25, .ColLetterToNumber("S"), DDMMYYYY
        .Row = 25
        .Col = .ColLetterToNumber("S")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
    
    End With
    
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
Dim vKHBS

    With fps
        lCurrSheet = .Sheet
        .Sheet = 1
            i = 0
            blnLoadThueSuat = True
            strCheck = GetAttribute(TAX_Utilities_Srv_New.Data(.Sheet - 1).nodeFromID( _
                     GetCellID(fps, .ColLetterToNumber(CHECK_COL1), CHECK_ROW1)), "Value")
            Do
                
                varID = Empty
                ' nvhai sua lay ma TTDB thay cho lay ID -> phuc vu QLT hach toan thue
               .GetText .ColLetterToNumber("C"), i + 53, varID
               If varID <> "" Then
                    .Row = i + 53
                    varThueSuat = .Value
                    
                    DataDM varID, strIdQLT, iTenHHDV, strDVT, strThueSuat, blnSuaDVT
                    
                    If strIdQLT <> vbNullString Then
                        ' Lay ten update vao dung vi tri cua ma Bieu thue
                        .Col = .ColLetterToNumber("D")
                        .Row = i + 53
                        .Text = iTenHHDV
                        UpdateCell fps, .ColLetterToNumber("F"), i + 31, .Text
                        
                        ' Lay ID update lai vao data file dung vao vi tri cua ma Bieu thue
                        .Col = .ColLetterToNumber("C")
                        .Row = i + 53
                        .Text = strIdQLT
                        UpdateCell fps, .ColLetterToNumber("C"), .Row, .Text
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
            blnLoadThueSuat = True
'            strCheck = GetAttribute(TAX_Utilities_Srv_New.Data(.Sheet - 1).nodeFromID( _
'                     GetCellID(fps, .ColLetterToNumber(CHECK_COL1), CHECK_ROW1)), "Value")
            Do
                varID = Empty
               .GetText .ColLetterToNumber("I"), i + 25, varID
               If varID <> "" Then
                    .Row = i + 25
                    varThueSuat = .Value
                    
                    DataDM varID, strIdQLT, iTenHHDV, strDVT, strThueSuat, blnSuaDVT
                    
                    If strIdQLT <> vbNullString Then
                        ' Lay ten update vao dung ten cua bieu thue
                        .Col = .ColLetterToNumber("L")
                        .Row = i + 25
                        .Text = iTenHHDV
                        UpdateCell fps, .ColLetterToNumber("L"), i + 25, .Text
                        
'                        ' Lay ID update lai vao data file dung vao vi tri cua ma Bieu thue
'                        .Col = .ColLetterToNumber("I")
'                        .Row = i + 25
'                        .Text = strIdQLT
'                        UpdateCell fps, .ColLetterToNumber("I"), .Row, .Text
                        
                    ElseIf CStr(varID) <> vbNullString Then
                        DisplayMessage "0081", msOKOnly, miCriticalError
                        Exit Function
                    End If
               End If
                    .RowHeight(i + 25) = .MaxTextRowHeight(i + 25)
                    'UpdateCell fps, .Col, .Row, .Text
                    i = i + 1
                    .Col = .ColLetterToNumber("B")
                    .Row = i + 25
                    If Trim(.Text) = "III" Then blnLoadThueSuat = False
                    
            Loop Until .Row >= .MaxRows
            
            ' Kiem tra xem co KHBS ko. Neu co thi Update lai MCT cua KHBS chinh la ID cua tung loai Hang hoa, dich vu cua bieu thue TTDB
            ' Trong KHBS, Cot C se luu ca mot dong bao gom Ten, Ma, Thue suat cua cac loai hang hoa
            ' Lay Ma hang hoa sau do lay ID tuong ung voi hang hoa do de cap nhat vao Cot D trong KHBS
            ' Nhu vay khi hien thi trong NTK thi MCT cua gia tri dieu chinh la ID cua Hang hoa chu ko phai chi tieu 10 (Chi tieu tong thue phat sinh)
'            vKHBS = TAX_Utilities_Srv_New.NodeValidity.childNodes(3).Attributes.getNamedItem("Active").nodeValue
'            If (Val(vKHBS) = 1) Then
'                .Sheet = 4
'                i = 0
'                Do
'                   varID = Empty
'                   .GetText .ColLetterToNumber("C"), i + 24, varID
'                   varID = LTrim(RTrim(Mid(varID, 201, 10)))
'                   If varID <> "" Then
'                        .Row = i + 24
'                        ' Lay Ma update lai vao data file dung vao vi tri cua Ma so chi tieu
'                        .Col = .ColLetterToNumber("E")
'                        .Row = i + 24
'                        .Text = varID
'                        UpdateCell fps, .ColLetterToNumber("E"), .Row, .Text
'                   End If
'                   i = i + 1
'                   .Row = .Row + 1
'                Loop Until .Row >= .MaxRows
'            End If
                
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
        If rsPXL.Fields.count > 0 Then
            Do While Not rsPXL.EOF
                .TypeComboBoxIndex = -1
                .TypeComboBoxString = TAX_Utilities_Srv_New.Convert(rsPXL.Fields(1).Value, TCVN, UNICODE)
                i = i + 1
                ReDim Preserve larrPhongXuLy(i)
                larrPhongXuLy(i) = TAX_Utilities_Srv_New.Convert(rsPXL.Fields(1).Value, TCVN, UNICODE)
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

Private Sub DataDM(ByVal Id As String, Optional ByRef strIdQLT As String, Optional ByRef TenTN As String, Optional ByRef strDVT As String, Optional ByRef strThueSuat As String, Optional ByRef blnSuaThueSuat As Boolean)
Dim arrDanhsach() As String
Dim strDataFileName As String
Dim xmlDOMdata As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = GetCatalogueFileName
       ' tu thang 4 /2010 danh muc bieu thue thay doi (mat hang bia co ma chuyen tu 103 -> 10300)
       ' cac phien ban tu 2.5.1 ve truoc van su dung danh muc cu co ma 103
       ' vi vay khi nhan du lieu tu thang 4/2010 bang NTK 2.5.2 phai chuyen thanh ma 10300
       
       If (Val(TAX_Utilities_Srv_New.Month) >= 1 And Val(TAX_Utilities_Srv_New.Year) >= 2010) And Id = "103" Then
            Id = "10300"
       End If
       If xmlDOMdata.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDOMdata.getElementsByTagName("Cell")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If Id = arrDanhsach(1) Then
                            strIdQLT = arrDanhsach(1)
                            TenTN = arrDanhsach(2)
                            strDVT = arrDanhsach(3)
                            strThueSuat = arrDanhsach(4)
                            blnSuaThueSuat = IIf(arrDanhsach(5) = "1", True, False)
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

Public Function Prepared3() As Boolean
    Dim strNgayPS As Variant
    With fps
        '.EventEnabled(EventAllEvents) = False
        'Set MaSoTep
        .Sheet = 1
        'Set NgayNhanToKhai
        .Col = .ColLetterToNumber(NGAY_NHAN_TO_KHAI_X)
        .Row = NGAY_NHAN_TO_KHAI_Y
        If strNgayNhanToKhai <> "" Then
            .Text = strNgayNhanToKhai
            UpdateCell fps, .Col, .Row, .Text
        End If

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
        
        .Col = .ColLetterToNumber(NGAY_PS_COL)
        .Row = NGAY_PS_ROW
        .GetText .Col, .Row, strNgayPS
        If Trim(strNgayPS) = "" Or Trim(strNgayPS) = vbNullString Then
            .Col = .ColLetterToNumber("AH")
            .Row = 3
            .Text = "[X]"
            .Col = .ColLetterToNumber("AP")
            .Row = 3
            .Text = ""
        Else
            .Col = .ColLetterToNumber("AH")
            .Row = 3
            .Text = ""
            .Col = .ColLetterToNumber("AP")
            .Row = 3
            .Text = "[X]"
        End If
        
'        .Col = .ColLetterToNumber(KY_LAP_BO_X)
'        .Row = KY_LAP_BO_Y
'        If Len(Month(Now())) = 1 Then
'            .Text = "0" & Month(Now()) & "/" & Year(Now())
'            sKyKeKhai = "0" & Month(Now()) & Year(Now())
'        Else
'            .Text = Month(Now()) & "/" & Year(Now())
'            sKyKeKhai = Month(Now()) & Year(Now())
'        End If
'        UpdateCell fps, .Col, .Row, .Text
'
'        .Col = .ColLetterToNumber(NGAY_NOP_X)
'        .Row = NGAY_NOP_Y
'        .Value = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
'         UpdateCell fps, .Col, .Row, .Value
'        .Lock = False
        
'        .Col = .ColLetterToNumber(NGAY_QUET_X)
'        .Row = NGAY_QUET_Y
'        .Value = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
'        UpdateCell fps, .Col, .Row, .Value
       
        .Col = .ColLetterToNumber(TEN_GOI_X)
        .Row = TEN_GOI_Y
        If strTenGoi <> "" Then
            .Text = Trim(TAX_Utilities_Srv_New.Convert(strTenGoi, TCVN, UNICODE))
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Col = .ColLetterToNumber(DIA_CHI_X)
        .Row = DIA_CHI_Y
        If strDchi <> "" Then
            .Text = Trim(TAX_Utilities_Srv_New.Convert(strDchi, TCVN, UNICODE))
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
        
         ' set thong tin dai ly thue
        ' ten dai ly thue
        .Col = .ColLetterToNumber(TEN_DL_COL)
        .Row = TEN_DL_ROW
        If strTenDL <> "" Then
            .Text = strTenDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' dia chi dai ly thue
        .Col = .ColLetterToNumber(DIA_CHI_DL_COL)
        .Row = DIA_CHI_DL_ROW
        If strDiaChiDL <> "" Then
            .Text = strDiaChiDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' dien thoai dai ly
        .Col = .ColLetterToNumber(DIEN_THOAI_DL_COL)
        .Row = DIEN_THOAI_DL_ROW
        If strDienThoaiDL <> "" Then
            .Text = strDienThoaiDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' fax dai ly thue
        .Col = .ColLetterToNumber(FAX_DL_COL)
        .Row = FAX_DL_ROW
        If strFaxDL <> "" Then
            .Text = strFaxDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' Email dai ly
        .Col = .ColLetterToNumber(EMAIL_DL_COL)
        .Row = EMAIL_DL_ROW
        If strEmailDL <> "" Then
            .Text = strEmailDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' so hop dong
        .Col = .ColLetterToNumber(SO_HOP_DONG_DL_COL)
        .Row = SO_HOP_DONG_DL_ROW
        If strSoHopDongDL <> "" Then
            .Text = strSoHopDongDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        ' ngay hop dong
        .Col = .ColLetterToNumber(NGAY_HOP_DONG_DL_COL)
        .Row = NGAY_HOP_DONG_DL_ROW
        If strNgayHopDongDL <> "" Then
            .Text = strNgayHopDongDL
            UpdateCell fps, .Col, .Row, .Text
        End If
        
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
        
'        If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
'            blnValid = False
'        End If
'        .Col = .ColLetterToNumber(HEADER_KY_LAP_BO_COL)
'        .Row = HEADER_KY_LAP_BO_ROW
'
'        If Not blnValid Then
'            .Formula = "0"
'        Else
'            .Formula = "1"
'        End If
'
'        ' Kiem tra xem ky ke khai trong cung ky hay khac ky
'        If (DateAdd("M", 1, dNgayDauKy) = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Or (dNgayDauKy = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Then
'            blnValid = True
'        Else
'            blnValid = False
'        End If
'
'        .Col = .ColLetterToNumber("B")
'        .Row = 17
'
'        If Not blnValid Then
'            .Formula = "0"
'        Else
'            .Formula = "1"
'        End If
        ' Kiem tra neu ky lap bo trang se bao loi
        If Trim(strKyLapBo) = "" Or Trim(strKyLapBo) = "../...." Then
            blnValid = False
            .Col = .ColLetterToNumber("B")
            .Row = 19
            .Formula = IIf(blnValid, "1", "0")
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 19
            .Formula = IIf(blnValid, "1", "0")
            
            If dNgayDauKy > DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) Then
                blnValid = False
            Else
                blnValid = True
            End If
            .Col = .ColLetterToNumber(HEADER_KY_LAP_BO_COL)
            .Row = HEADER_KY_LAP_BO_ROW
            
            If Not blnValid Then
                .Formula = "0"
            Else
                .Formula = "1"
            End If
            
            
            ' Kiem tra xem ky ke khai trong cung ky hay khac ky
            If (DateAdd("M", 1, dNgayDauKy) = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Or (dNgayDauKy = DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1)) Then
                blnValid = True
            Else
                blnValid = False
            End If
            
            .Col = .ColLetterToNumber("B")
            .Row = 17

            If Not blnValid Then
                .Formula = "0"
            Else
                .Formula = "1"
            End If
            ' kiem tra ky lap bo khong duoc lon hon ky hien tai
            If DateSerial(CInt(Right$(strKyLapBo, 4)), CInt(Left$(strKyLapBo, 2)), 1) > Date Then
                blnValid = False
            Else
                blnValid = True
            End If
            .Col = .ColLetterToNumber("B")
            .Row = 18
            
            If Not blnValid Then
                .Formula = "0"
            Else
                .Formula = "1"
            End If
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

Public Function CheckValidData() As Boolean
    Dim varCheckValue As Variant, varNoteValue As Variant
    Dim varNgayNop As Variant
    CheckValidData = True
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("G"), 29, varCheckValue
        .GetText .ColLetterToNumber("S"), 29, varNoteValue
        '
        .GetText .ColLetterToNumber("F"), 25, varNgayNop
        If CStr(varCheckValue) = "1" And Trim(CStr(varNoteValue)) = "" Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 14, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 14, "1"
        End If
        '
        If CStr(Trim(varNgayNop)) = "" Or CStr(Trim(varNgayNop)) = "../../...." Then
            .Sheet = .SheetCount
            .SetText 2, 12, "0"
            CheckValidData = False
        Else
            .Sheet = .SheetCount
            .SetText 2, 12, "1"
        End If
        
    End With
End Function

Public Function GenerateSQL_Header(xmlDOMdata As MSXML.DOMDocument, strSQL_HDR As String, vHdrID As Variant, vKKBS As Byte, ByVal dNgayDauKy As Date) As String
    Dim xmlList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlAttribute As MSXML.IXMLDOMAttribute
    Dim iRowID As Long, strSQL As String, strTempSQL As String
    Dim dDate As Date, strDate() As String
    Dim vTIN, vTEN_DTNT, vDIA_CHI, vLOAI_TKHAI, vNGAY_NOP, vKyLB
    Dim vKYKK, vNGAY_CAP_NHAT, vNGUOI_CAP_NHAT, vCO_LOI_DDANH
    Dim vSO_HIEU_TEP, vSO_TT_TK, vDA_NHAN, vGHI_CHU_LOI, vCO_GTRINH_02A
    Dim vCO_GTRINH_02B, vCO_GTRINH_02C
    Dim vPHONG_XU_LY
    Dim i As Long, j As Long
    Dim strMaPhongXuLy As String
    Dim sQCT As String
    Dim vTIN_DL, vSO_HOP_DONG, vNGAY_HOP_DONG, vLAN_BS, vCHINH_THUC, vNGANH_NGHE_KD, vTO_KHAI_LAN_PS, vNGAY_PS As Variant
    
On Error GoTo ErrHandle
    strSQL = strSQL_HDR
    Set xmlList = xmlDOMdata.getElementsByTagName("Cell")
    For Each xmlNode In xmlList
        With xmlNode.Attributes
        
        If Trim(GetAttribute(xmlNode, "MCT")) = vbNullString Then
            Select Case Trim(GetAttribute(xmlNode, "CellID"))
                Case "I_7"
                    vTIN = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "I_5"
                    vTEN_DTNT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "I_9"
                    vDIA_CHI = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "F_25"
                    vNGAY_NOP = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "F_23"
                    vKyLB = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "S_25"
                    vNGAY_CAP_NHAT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "AG_25"
                    vNGUOI_CAP_NHAT = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "G_29"
                    vCO_LOI_DDANH = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "T_39"
                    ' vSO_HIEU_TEP = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                    ' Tu phien ban 1.3.0 song tep se thuc hien bang tay=> So hieu tep luon bang ""
                    vSO_HIEU_TEP = ""
                Case HEADER_SO_TT_TRONG_TEP_X & "_" & HEADER_SO_TT_TRONG_TEP_Y
                    vSO_TT_TK = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case "S_29"
                    vGHI_CHU_LOI = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
                Case PHONG_XU_LY_X & "_" & PHONG_XU_LY_Y
                    vPHONG_XU_LY = Trim(TAX_Utilities_Srv_New.Convert(GetAttribute(xmlNode, "Value"), UNICODE, TCVN))
            End Select
        End If
        End With
    Next
    If TAX_Utilities_Srv_New.isCheckQCT = True Then
        sQCT = "'Y'"
    Else
        sQCT = "null"
    End If
    ' ngay phat sinh
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber(NGAY_PS_COL), NGAY_PS_ROW, vNGAY_PS
    End With
    
    strSQL = strSQL & "'" & vHdrID & "',"
    strSQL = strSQL & vKKBS & ","
    strSQL = strSQL & "'" & vTIN & "',"
    strSQL = strSQL & "'" & vTEN_DTNT & "',"
    strSQL = strSQL & "'" & vDIA_CHI & "',"
    'vLOAI_TKHAI = TAX_Utilities_Srv_New.NodeMenu.Attributes.getNamedItem("ID").nodeValue
    vLOAI_TKHAI = "01_TTDB11"
    strSQL = strSQL & "'" & vLOAI_TKHAI & "',"
    strSQL = strSQL & "To_date('" & vNGAY_NOP & "','dd/mm/yyyy'),"
    
    'Ky/Quy LB
    If Trim(TAX_Utilities_Srv_New.Month) <> "" Then
        'Ngay dau ky lap bo va ngay cuoi ky lap bo
        strDate = Split(vKyLB, "/")
        dDate = DateSerial(Val(strDate(1)), Val(strDate(0)), 1)
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
        dDate = DateAdd("m", 1, dDate)
        dDate = DateAdd("d", -1, dDate)
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    End If
    
    'Ky/ Quy KK
    If Trim(TAX_Utilities_Srv_New.Month) <> "" And Trim(vNGAY_PS) = "" Then
        'Ngay dau ky ke khai va ngay cuoi ky ke khai
        'strDate = Split(TAX_Utilities_Srv_New.Month & "/" & TAX_Utilities_Srv_New.Year, "/")
        'dDate = DateSerial(Val(strDate(1)), Val(strDate(0)), 1)
        dDate = dNgayDauKy
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
        dDate = DateAdd("m", 1, dDate)
        dDate = DateAdd("d", -1, dDate)
        strSQL = strSQL & "To_date('" & Format(dDate, "dd/mm/yyyy") & "','dd/mm/yyyy'),"
    ElseIf Trim(vNGAY_PS) <> "" Then
            strSQL = strSQL & "To_date('" & vNGAY_PS & "','dd/mm/yyyy'),"
            strSQL = strSQL & "To_date('" & vNGAY_PS & "','dd/mm/yyyy'),"
    End If
    
    strSQL = strSQL & "To_date('" & vNGAY_CAP_NHAT & "','dd/mm/yyyy'),"
    strSQL = strSQL & "'" & vNGUOI_CAP_NHAT & "',"
    strSQL = strSQL & "'" & vCO_LOI_DDANH & "',"
    strSQL = strSQL & "'" & vSO_HIEU_TEP & "',"
    strSQL = strSQL & "'" & vSO_TT_TK & "',"
    
    strSQL = strSQL & "'" & vDA_NHAN & "',"
    strSQL = strSQL & "'" & vGHI_CHU_LOI & "',"
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'',"
    ' Them danh dau 3 phu luc cho anh Huy
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'',"
    ' Ket thuc
    strSQL = strSQL & "null,null,"
    
    With fps
        For i = 1 To lSoPhongXL
            If vPHONG_XU_LY = TAX_Utilities_Srv_New.Convert(larrPhongXuLy(i), UNICODE, TCVN) Then
                strMaPhongXuLy = larrid(i)
                Exit For
            End If
        Next
        
       .Sheet = 1
       ' Ma dai ly thue
        .GetText .ColLetterToNumber("I"), 15, vTIN_DL
        .GetText .ColLetterToNumber(SO_HOP_DONG_DL_COL), SO_HOP_DONG_DL_ROW, vSO_HOP_DONG
        .GetText .ColLetterToNumber(NGAY_HOP_DONG_DL_COL), NGAY_HOP_DONG_DL_ROW, vNGAY_HOP_DONG
        .GetText .ColLetterToNumber("J"), 33, vLAN_BS
        .GetText .ColLetterToNumber("D"), 33, vCHINH_THUC
        If Trim(vCHINH_THUC) = "1" Then
            vLAN_BS = "0"
        End If
        
        If Len(Trim(vTIN_DL)) = 13 Then
            vTIN_DL = Left(vTIN_DL, 10) & "-" & Right(vTIN_DL, 3)
        End If
        
        
        .GetText .ColLetterToNumber(NGANH_NGHE_KD_COL), NGANH_NGHE_KD_ROW, vNGANH_NGHE_KD
        vNGANH_NGHE_KD = "'" & vNGANH_NGHE_KD & "'"
        ' ngay phat sinh
        '.GetText .ColLetterToNumber(NGAY_PS_COL), NGAY_PS_ROW, vNGAY_PS
        If Trim(vNGAY_PS) = "" Or vNGAY_PS = vbNullString Then
            vNGAY_PS = "null"
            vTO_KHAI_LAN_PS = "null"
        Else
            vNGAY_PS = "'" & vNGAY_PS & "'"
            vNGAY_PS = "to_date(" & vNGAY_PS & ",'dd/mm/yyyy')"
            vTO_KHAI_LAN_PS = "1"
        End If
        
'        vTO_KHAI_LAN_PS = "null"
'        vNGAY_PS = "null"
        
        If Trim(vNGAY_HOP_DONG) = "" Or vNGAY_HOP_DONG = vbNullString Then
            vNGAY_HOP_DONG = "null"
        Else
            vNGAY_HOP_DONG = "'" & vNGAY_HOP_DONG & "'"
            vNGAY_HOP_DONG = "to_date(" & vNGAY_HOP_DONG & ",'dd/mm/yyyy')"
        End If
    End With
    strSQL = strSQL & "'" & strMaPhongXuLy & "', null,"
    'strSQL = strSQL & "null)"
    strSQL = strSQL & "null,"
    strSQL = strSQL & "'" & Trim(vTIN_DL) & "','" & Trim(vSO_HOP_DONG) & "'," & vNGAY_HOP_DONG & "," & Trim(vLAN_BS) & "," & vNGANH_NGHE_KD & "," & vTO_KHAI_LAN_PS & "," & vNGAY_PS & "," & sQCT & ")"
    
    GenerateSQL_Header = strSQL
    
    Exit Function
ErrHandle:
    SaveErrorLog "cls_01TTDB", "GenerateSQL_Header", Err.Number, Err.Description
End Function

Public Function KiemTraKhoaSo(ByVal strDate As String) As Boolean
    Dim dNgayKhoaSo As Date, dKyLapBo As Date
    
    KiemTraKhoaSo = True
    dNgayKhoaSo = Format(strDate, "dd/mm/yyyy")
    dKyLapBo = Format("01/" & GetAttribute(TAX_Utilities_Srv_New.Data(0).nodeFromID("F_23"), "Value"), "dd/mm/yyyy")
    
    If dNgayKhoaSo < dKyLapBo Then
        KiemTraKhoaSo = False
    End If
End Function

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    With fps
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("F") And Row = 25 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    Else
                        .SetActiveCell Col, Row
                    End If
'                Else
'                 .SetActiveCell Col, Row
                End If
               .Col = Col
               .Row = Row
               UpdateCell fps, .Col, .Row, .Text
            End If
            'ky lap bo
            
            If Col = .ColLetterToNumber("F") And Row = 23 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../...." Then
                    If Format_mmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_mmyyyy(CStr(varTemp))
                    Else
                        .SetActiveCell Col, Row
                    End If
'                Else
'                 .SetActiveCell Col, Row
                End If
               .Col = Col
               .Row = Row
               UpdateCell fps, .Col, .Row, .Text
            End If
        End If
    End With
End Sub
