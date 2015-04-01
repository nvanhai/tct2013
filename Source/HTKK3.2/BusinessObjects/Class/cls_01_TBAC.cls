VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01TBAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrSolanBosung As String
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
Dim strDataFileName As String
Dim arrDanhsach() As String
Dim strComboHien As String
Dim strCombo As String
Dim MSTDN As String
Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_loai_HD.xml")
    ' testttt svn
    ' Lay danh muc loai hoa don
    ' 15/11/2010
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListCell
            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                If (TAX_Utilities_v1.Year = 2014 And Val(TAX_Utilities_v1.Month) < 6) Or Val(TAX_Utilities_v1.Year) < 2014 Then
                    strCombo = strCombo + CPab(arrDanhsach(0), 10) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(2), 200) + Chr$(9)
                    strComboHien = strComboHien + arrDanhsach(1) + Chr$(9)
                Else
                    If arrDanhsach(1) <> "06HDXK" Then
                        strCombo = strCombo + CPab(arrDanhsach(0), 10) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(2), 200) + Chr$(9)
                        strComboHien = strComboHien + arrDanhsach(1) + Chr$(9)
                    End If
                End If
            End If
        Next
    End If
    
     With fps
        SetDateFormat fps, 1, 46, .ColLetterToNumber("AA"), DDMMYYYY
        SetDateFormat fps, 1, 35, .ColLetterToNumber("T"), DDMMYYYY
        SetDateFormat fps, 1, 35, .ColLetterToNumber("AE"), DDMMYYYY
        'SetKyHieuHDFormat fps, 1, 35, .ColLetterToNumber("L")
        .Sheet = 1

        ' Set gia tri ten loai hoa don
        .Col = .ColLetterToNumber("C")
        .Row = 35
        .TypeComboBoxList = strComboHien
        ' Set gia tri cho combo an
        .Col = .ColLetterToNumber("AG")
        .Row = 35
        .TypeComboBoxList = strCombo
       
    End With
    
    
    
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    'bien ten dvcq
    Dim strDvcq As String
    Dim strTenDvcq As String
    Dim countRowDynamic As Variant
    Dim nguoiky As Variant
    Dim strId As Variant
    Dim indexCombo As Integer
    
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim iCol As Long, iRow As Long
    
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
         mCurrentSheet = .Sheet
        .Row = 35
        Do
            countRowDynamic = countRowDynamic + 1
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
        'lay ten va MST DVCQ tren Header
            strDvcq = GetDVCQ()
            strTenDvcq = GetTenDVCQ()
            .SetText .ColLetterToNumber("F"), countRowDynamic + 41, strDvcq
            .SetText .ColLetterToNumber("F"), countRowDynamic + 39, strTenDvcq
        'lay ten nguoi ky tren header
        .Row = countRowDynamic + 47
        .Col = .ColLetterToNumber("AA")
        If Trim(.Text) = "" Then
            nguoiky = GetNguoiKy()
            .SetText .ColLetterToNumber("AA"), countRowDynamic + 47, nguoiky
        End If
        
         'Xoa cong thuc cua o Chu ky
        Set xmlCellNode = TAX_Utilities_v1.Data(0).nodeFromID(GetCellID(fps, .ColLetterToNumber("D"), 35))
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode.lastChild
        Set xmlCellNode = xmlCellNode.firstChild
        Set xmlCellNode = xmlCellNode.lastChild
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        .Row = iRow
        .Col = iCol
        .Formula = ""
        
        ' Set format ngay ky
        .Row = countRowDynamic + 45
        .Col = .ColLetterToNumber("AA")
        '.Text = Format(Date, "dd/mm/yyyy")
        .Text = TAX_Utilities_v1.Day & "/" & TAX_Utilities_v1.Month & "/" & TAX_Utilities_v1.Year
        '.TypeHAlign = TypeHAlignLeft
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub



Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
 Dim strId As Variant
 Dim intIndexCombo As Integer
 
    fps.EventEnabled(EventAllEvents) = False
    With fps
        
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("C") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("AG")
            .typeComboboxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber("AG")
            UpdateCell fps, .Col, .Row, .Text
            
            .GetText .ColLetterToNumber("AG"), .Row, strId
         
            .Col = .ColLetterToNumber("N")
            If Trim(Left(strId, 1)) = "0" Then
                .SetText .ColLetterToNumber("J"), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell fps, .ColLetterToNumber("J"), .Row, .Text
                .Col = .ColLetterToNumber("J")
                .TypeMaxEditLen = 13
                .Col = .ColLetterToNumber("D")
                .SetText .ColLetterToNumber("D"), .Row, Trim(Mid(strId, 21))
                UpdateCell fps, .ColLetterToNumber("D"), .Row, .Text
            ElseIf Trim(Left(strId, 1)) = "3" Then
                .SetText .ColLetterToNumber("J"), .Row, ""
                UpdateCell_sheet fps, 1, .ColLetterToNumber("J"), .Row, .Text
                .Col = .ColLetterToNumber("J")
                .TypeMaxEditLen = 20
                .Col = .ColLetterToNumber("D")
                .SetText .ColLetterToNumber("D"), .Row, ""
                UpdateCell_sheet fps, 1, .ColLetterToNumber("D"), .Row, .Text
            ElseIf Trim(Left(strId, 1)) = "4" Then
                .SetText .ColLetterToNumber("J"), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell_sheet fps, 1, .ColLetterToNumber("J"), .Row, .Text
                .Col = .ColLetterToNumber("J")
                .TypeMaxEditLen = 13
                .Col = .ColLetterToNumber("D")
                .SetText .ColLetterToNumber("D"), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 1, .ColLetterToNumber("D"), .Row, .Text
            Else
                .Col = .ColLetterToNumber("D")
                .SetText .ColLetterToNumber("D"), .Row, Trim(Mid(strId, 21))
                .Col = .ColLetterToNumber("J")
                .TypeMaxEditLen = 20
                UpdateCell fps, .ColLetterToNumber("D"), .Row, .Text
                .SetText .ColLetterToNumber("J"), .Row, Trim(Mid(strId, 11, 3))
                UpdateCell fps, .ColLetterToNumber("J"), .Row, .Text
            End If
            .SetText .ColLetterToNumber("AJ"), .Row, Trim(Left(strId, 1))
            UpdateCell fps, .ColLetterToNumber("AJ"), .Row, .Text
                
            .EventEnabled(EventChange) = True
         End If
    End With
   fps.EventEnabled(EventAllEvents) = True
End Sub

Private Sub fps_KeyPress(KeyAscii As Integer)
     With fps
        If .Col = .ColLetterToNumber("P") Or .Col = .ColLetterToNumber("R") Then
            If (KeyAscii < 48 And KeyAscii <> 3 And KeyAscii <> 22 And KeyAscii <> 13 And KeyAscii <> 8) Or KeyAscii > 57 Then
                KeyAscii = 0
            End If
        End If
    End With
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Integer, iCol As Long, iRow As Long
    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 1 Then
                   fps.EventEnabled(EventAllEvents) = True
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 35
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 35
                         i = i + 1
                    Loop Until .Text = "aa"
                    .SetActiveCell iCol, iRow

                    .Row = iRow
                    .Col = iCol
                     fps.EventEnabled(EventAllEvents) = False
                     .SetActiveCell iCol, iRow
            End If
         End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim tuSo As Variant
    Dim denso As Variant
    Dim strId As Variant
    Dim strLoai As Variant
    Dim strTemp As Variant
    Dim countRowDynamic As Integer
    Dim strSoHieu As Variant

    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
        ' check ngay bat dau su dung
            If Col = .ColLetterToNumber("T") Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    Else
                        .SetFocus
                        .SetActiveCell Col, Row
                    End If
                Else
                 .SetText Col, Row, ""
                End If
                .Col = Col
                .Row = Row
                UpdateCell fps, Col, Row, .Text
            End If
        'check ngay in hop dong
            If Col = .ColLetterToNumber("AE") Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    Else
                        .SetFocus
                        .SetActiveCell Col, Row
                    End If
                Else
                 .SetText Col, Row, ""
                End If
                .Col = Col
                .Row = Row
                UpdateCell fps, Col, Row, .Text
            End If
        ' check ngay ky TBAC
        .Row = 35
            Do
                countRowDynamic = countRowDynamic + 1
                .Col = .ColLetterToNumber("B")
                .Row = .Row + 1
            Loop Until .Text = "aa"
                If Col = .ColLetterToNumber("AA") And Row = countRowDynamic + 45 Then
                     .GetText Col, Row, varTemp
                     If varTemp <> "" And varTemp <> "../../...." Then
                         If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                             .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                         Else
                             .SetFocus
                             .SetActiveCell Col, Row
                         End If
                     Else
                      .SetText Col, Row, ""
                     End If
                     .Col = Col
                     .Row = Row
                     UpdateCell fps, Col, Row, .Text
                 End If
'         'khoa o V , Z , AC, AD
'         .GetText .ColLetterToNumber("K"), countRowDynamic + 34, strSoHieu
'            If Mid(Trim(strSoHieu), 6, 1) <> "P" Then
'                .Col = .ColLetterToNumber("V")
'                .Row = countRowDynamic + 34
'                .Lock = True
'                .Col = .ColLetterToNumber("Z")
'                .Row = countRowDynamic + 34
'                .Lock = True
'                .Col = .ColLetterToNumber("AC")
'                .Row = countRowDynamic + 34
'                .Lock = True
'                .Col = .ColLetterToNumber("AD")
'                .Row = countRowDynamic + 34
'                .Lock = True
'            ElseIf Mid(Trim(strSoHieu), 6, 1) = "P" Then
'                 .Col = .ColLetterToNumber("V")
'                .Row = countRowDynamic + 34
'                .Lock = False
'                .Col = .ColLetterToNumber("Z")
'                .Row = countRowDynamic + 34
'                .Lock = False
'                .Col = .ColLetterToNumber("AC")
'                .Row = countRowDynamic + 34
'                .Lock = False
'                .Col = .ColLetterToNumber("AD")
'                .Row = countRowDynamic + 34
'                .Lock = False
'            End If
          ' Format dinh dang tu so den so co 7 ky tu
            If Col = .ColLetterToNumber("P") Or Col = .ColLetterToNumber("R") Then
                 .GetText Col, Row, varTemp
                 .SetText Col, Row, FormatSoHD(CStr(varTemp))
                 .Col = Col
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
             End If
          ' Ho tro tinh so luong
                
                .GetText .ColLetterToNumber("P"), Row, tuSo
                .GetText .ColLetterToNumber("R"), Row, denso
                If Trim(tuSo) <> "" And Trim(denso) <> "" Then
                    ' so HD 0000000 se khong tinh
                    If Val(tuSo) = 0 And Val(denso) = 0 Then
                        .SetText .ColLetterToNumber("N"), Row, 0
                    ElseIf Val(tuSo) = 0 And Val(denso) <> 0 Then
                        .SetText .ColLetterToNumber("N"), Row, Val(denso) - Val(tuSo)
                    Else
                        .SetText .ColLetterToNumber("N"), Row, Val(denso) - Val(tuSo) + 1
                    End If
                    .Col = Col
                    .Row = Row
                    UpdateCell fps, .Col, .Row, .Text
                ElseIf Trim(tuSo) = "" And Trim(denso) = "" Then
                    .SetText .ColLetterToNumber("N"), Row, ""
                    .Col = Col
                    .Row = Row
                    UpdateCell fps, .Col, .Row, .Text
                End If
          ' Chuyen ky hieu HD thanh chu cai in hoa
            If Col = .ColLetterToNumber("L") Then
                 .GetText Col, Row, varTemp
                 .SetText Col, Row, FormatKyHieu(CStr(varTemp))
                 .Col = Col
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            
            .GetText .ColLetterToNumber("AG"), .Row, strId
            strLoai = Trim(Left(strId, 1))
            
            ' set gia tri cho check KH NN/XC
            strMauSoHD_01GTKT = Trim(Mid(strId, 11, 6))
           'Kiem tra cau truc cua ky hieu neu sai set cot an =1 dung =0
            If Col = .ColLetterToNumber("L") Then
                 .GetText Col, Row, varTemp
                 .SetText .ColLetterToNumber("AJ"), Row, CheckSoHD(CStr(varTemp), strLoai)
                 .Col = .ColLetterToNumber("AJ")
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
            End If
           ' Chuyen Mau so thanh chu cai in hoa
                If Col = .ColLetterToNumber("J") Then
                    .GetText Col, Row, varTemp
                    '.SetText Col, Row, FormatKyHieu(CStr(varTemp))
                    .SetText Col, Row, UCase(varTemp)
                    .Col = Col
                    .Row = Row
                    UpdateCell fps, .Col, .Row, .Text
                End If
           'Kiem tra cau truc cua Mau so
           If Col = .ColLetterToNumber("J") Then
                .GetText .ColLetterToNumber("AG"), .Row, strId
                strLoai = Trim(Left(strId, 1))
                If strLoai = "0" Then
                    strTemp = Trim(Mid(strId, 11, 6))
                ElseIf strLoai = "4" Then
                    strTemp = Trim(Mid(strId, 11, 6))
                Else
                    strTemp = Trim(Mid(strId, 11, 3))
                End If
                .GetText Col, Row, varTemp
                 .SetText .ColLetterToNumber("AL"), Row, CheckMauSoHD(CStr(varTemp), CStr(strLoai), CStr(strTemp))
                 .Col = .ColLetterToNumber("AL")
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub


'Calling this function whenever we don't need this class anymore.
'No parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        
        CheckDynamicError 'Set Exception Error on cells of interface
            
        UpdateSheets
        
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub


Private Sub CheckDynamicError()
    Dim strCheck As String
    Dim errMauSoHD As Variant, errKyHieuHD As Variant, errSoLuong As Variant, errSoLienHD As Variant, errDVCQ As Variant, errNgayKy As Variant, errMsgTusoDenSo As Variant, errMsgMSTTrungNhau As Variant, errSoLienHDDT As Variant
    Dim errKhoangTN As Variant, errLengthHD As Variant, errEmty As Variant, errThongTinDN As Variant, errCheckNgaySD As Variant, errCheckNgayIn As Variant, errMSTDN As Variant, errMST As Variant, errMstDVCQ As Variant
    
    Dim errNgayInTB As Variant
    
    Dim strTenHD As Variant, strMauSoHD As Variant, strKyHieuHD As Variant, strTuSo As Variant, strDenso As Variant, strDvcq As Variant, strTenDvcq As Variant, strMST As Variant
    Dim strSoLuong As Variant, strTenDN As Variant, strMstDN As Variant, strDateSD As Variant, DateNgayKy As Variant, strNguoiKy As Variant
    Dim iFlagMauSoHD As Integer, iFlagSoLienHD As Integer, iFlagEmty As Integer, iFlagDVCQ As Integer, iFlagNgayKy As Integer, iFlagNgayBatDau As Integer, iFlagNgayIn As Integer, iFlagTusoDenSo As Integer, iFlagNgayInTB As Integer
    Dim iWhiteRow As Integer
    Dim strId As Variant, strLoai As String, strTemp As String, result As String
    Dim str() As String
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Dim NgaySD As Date
    Dim strDateNHT As Date
    Dim NgayIn As Date
    Dim ngayHienTai As Date
    Dim CheckNgaySD As Variant
    Dim CheckNgaySD1 As Variant
    Dim errCheckNgaySD1 As Variant
    Dim iFlagNgayBatDau1 As Integer
    Dim CheckNgayKy As Variant
    Dim CheckNgayIn As Variant
    Dim SohieuHD As Variant
    Dim mstDVCQ As Variant
    Dim tenDVCQ As Variant
    Dim MST As Variant
    Dim Col As Long, Row As Long
    Dim countRowDynamic As Integer
    Dim Tuso2 As Variant
    Dim denso2 As Variant
    Dim strTenHD2 As Variant
    Dim strMauSoHD2 As Variant
    Dim strKyHieuHD2 As Variant
    Dim chkTTDVCQ As Variant
    
    ' tra ve mang cac row bi trung khoang
    str = checkTrung(fps, 35, "P", "R", "B", "D", "J", "L", "", False, 1)
    With fps
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 15, errMSTDN
        .GetText .ColLetterToNumber("E"), 16, errMST
        .GetText .ColLetterToNumber("E"), 17, errMsgMSTTrungNhau
        .GetText .ColLetterToNumber("E"), 18, errMauSoHD
        .GetText .ColLetterToNumber("E"), 19, errKyHieuHD
        .GetText .ColLetterToNumber("E"), 20, errSoLienHD
        .GetText .ColLetterToNumber("E"), 21, errKhoangTN
        .GetText .ColLetterToNumber("E"), 22, errLengthHD
        .GetText .ColLetterToNumber("E"), 23, errMsgTusoDenSo
        .GetText .ColLetterToNumber("E"), 24, errEmty
        .GetText .ColLetterToNumber("E"), 26, errCheckNgaySD1
        .GetText .ColLetterToNumber("E"), 28, errCheckNgayIn
        .GetText .ColLetterToNumber("E"), 27, errDVCQ
        .GetText .ColLetterToNumber("E"), 29, errThongTinDN
        .GetText .ColLetterToNumber("E"), 30, errNgayKy
        .GetText .ColLetterToNumber("E"), 31, errMstDVCQ
        
        .GetText .ColLetterToNumber("E"), 32, errNgayInTB
        
        .GetText .ColLetterToNumber("E"), 33, errCheckNgaySD
        
        .GetText .ColLetterToNumber("E"), 34, errSoLienHDDT
        
        
        .GetText .ColLetterToNumber("AI"), 3, chkTTDVCQ
        ' Kiem tra tren sheet 1
        .Sheet = 1
        .Row = 4
        .Col = .ColLetterToNumber("F")
        .GetText .ColLetterToNumber("F"), 4, strMST
        .Row = 35
        .Col = .ColLetterToNumber("B")
        i = 1
        ' Dem so row dynamic
        Do
            .Col = .ColLetterToNumber("P")
            .CellNote = ""
            .BackColor = mNonErrorColor
            .Col = .ColLetterToNumber("R")
            .CellNote = ""
            .BackColor = mNonErrorColor
            countRowDynamic = countRowDynamic + 1
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until UCase(.Text) = "AA"
        
        .GetText .ColLetterToNumber("AA"), countRowDynamic + 47, strNguoiKy
        .GetText .ColLetterToNumber("AA"), countRowDynamic + 45, DateNgayKy
        .GetText .ColLetterToNumber("F"), countRowDynamic + 39, strTenDvcq
        .GetText .ColLetterToNumber("F"), countRowDynamic + 41, strDvcq
         ' check nguoi ky bat buoc nhap
            If Trim(strNguoiKy) = "" Then
                .Col = .ColLetterToNumber("AA")
                .Row = countRowDynamic + 47
                .CellNote = errEmty
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                .Col = .ColLetterToNumber("AA")
                .Row = countRowDynamic + 47
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
        'chek ngay ky " cho phep sua nhung ko duoc lon hon ngay hien tai"
            ngayHienTai = DateSerial(Mid(Trim(DateNgayKy), 7, 4), Mid(Trim(DateNgayKy), 4, 2), Mid(Trim(DateNgayKy), 1, 2))
            strDateNHT = DateTime.Now
            CheckNgayKy = DateDiff("d", strDateNHT, ngayHienTai)
            If CheckNgayKy > 0 Then
                .Col = .ColLetterToNumber("AA")
                .Row = countRowDynamic + 45
                .CellNote = errNgayKy
                .BackColor = mErrorColor
                iFlagNgayKy = 1
            Else
                .Col = .ColLetterToNumber("AA")
                .Row = countRowDynamic + 45
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
       
        'check thong tin don vi chu quan
        'neu  MST don vi chu quan trung voi mst trong thong tin Doanh nghiep thi khong hien thi thong tin don vi chu quan
        'dntai neu check vao don vi chu quan tren header thi moi bat buoc nhap dl
        If Trim(chkTTDVCQ) = "1" Or UCase(Trim(chkTTDVCQ)) = "X" Then
        MST = Trim(strMST)
        mstDVCQ = Trim(strDvcq)
        If MST <> mstDVCQ Then
                ' ten don vi chu quan
                If Trim(strTenDvcq) = "" Then
                    .Col = .ColLetterToNumber("F")
                    .Row = countRowDynamic + 39
                    .CellNote = errDVCQ
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                Else
                    .Col = .ColLetterToNumber("F")
                    .Row = countRowDynamic + 39
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                If Trim(strDvcq) = "" Then
                    .Col = .ColLetterToNumber("F")
                    .Row = countRowDynamic + 41
                    .CellNote = errDVCQ
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                ElseIf Trim(strDvcq) <> vbNullString Then
                     strCheck = ""
                        If Len(strDvcq) = 10 Or Len(strDvcq) = 13 Then
                            strCheck = CheckTaxCode(Mid(strDvcq, 1, 1), Mid(strDvcq, 2, 1), Mid(strDvcq, 3, 1), Mid(strDvcq, 4, 1), Mid(strDvcq, 5, 1), Mid(strDvcq, 6, 1), Mid(strDvcq, 7, 1), Mid(strDvcq, 8, 1), Mid(strDvcq, 9, 1), Mid(strDvcq, 10, 1), Mid(strDvcq, 11, 1), Mid(strDvcq, 12, 1), Mid(strDvcq, 13, 1))
                            If Len(strDvcq) = 10 And strCheck <> "000" Then
                                .Col = .ColLetterToNumber("F")
                                .Row = countRowDynamic + 41
                                .CellNote = errMstDVCQ
                                .BackColor = mErrorColor
                                iFlagDVCQ = 1
                            ElseIf Len(strDvcq) = 13 And strCheck <> "000" Then
                                .Col = .ColLetterToNumber("F")
                                .Row = countRowDynamic + 41
                                .CellNote = errMstDVCQ
                                .BackColor = mErrorColor
                                iFlagDVCQ = 1
                            Else
                                .Col = .ColLetterToNumber("F")
                                .Row = countRowDynamic + 41
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            End If
                        ElseIf Len(strDvcq) <> 10 Or Len(strDvcq) <> 13 Then
                            .Col = .ColLetterToNumber("F")
                            .Row = countRowDynamic + 41
                            .CellNote = errMstDVCQ
                            .BackColor = mErrorColor
                            iFlagDVCQ = 1
                        End If
                Else
                        .Col = .ColLetterToNumber("F")
                        .Row = countRowDynamic + 41
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                End If
        Else
            .Col = .ColLetterToNumber("F")
            .Row = countRowDynamic + 39
            .CellNote = ""
            .BackColor = mNonErrorColor
            .Col = .ColLetterToNumber("F")
            .Row = countRowDynamic + 41
            .CellNote = ""
            .BackColor = mNonErrorColor
            End If
            
        End If
        .Row = 35
        Do
        j = .Row
             ' Lay cac gia tri cua 1 row
             iWhiteRow = 1
            .GetText .ColLetterToNumber("D"), .Row, strTenHD
            .GetText .ColLetterToNumber("J"), .Row, strMauSoHD
            .GetText .ColLetterToNumber("L"), .Row, strKyHieuHD
            .GetText .ColLetterToNumber("P"), .Row, strTuSo
            .GetText .ColLetterToNumber("R"), .Row, strDenso
            .GetText .ColLetterToNumber("N"), .Row, strSoLuong
            .GetText .ColLetterToNumber("AG"), .Row, strId
            .GetText .ColLetterToNumber("W"), .Row, strTenDN
            .GetText .ColLetterToNumber("AA"), .Row, strMstDN
            '.GetText .ColLetterToNumber("AD"), .Row, strSoHD
            .GetText .ColLetterToNumber("T"), .Row, strDateSD
            '.GetText .ColLetterToNumber("AE"), .Row, strDateIn
            strLoai = Trim(Left(strId, 1))
            
            ' set gia tri cho check KH NN/XC
            strMauSoHD_01GTKT = Trim(Mid(strId, 11, 6))
            ' end
            
        If Trim(strTenHD) <> "" Or Trim(strMauSoHD) <> "" Or Trim(strKyHieuHD) <> "" Or Trim(strId) <> "" Or Trim(strMstDN) <> "" Or Trim(strTuSo) <> "" Or Trim(strDenso) <> "" Or _
        Trim(strTenDN) <> "" Then
            ' Khong phai dong trang
            iWhiteRow = 0
        End If
     ' check ngay " ngay su dung - ngay in hop dong >= 5"
        If Trim(strTenHD) <> "" Then
              If Trim(strDateSD) = "" And iWhiteRow = 0 Then
                     .Col = .ColLetterToNumber("T")
                     .CellNote = errEmty
                     .BackColor = mErrorColor
                     iFlagEmty = 1
              Else
                      NgaySD = DateSerial(Mid(Trim(strDateSD), 7, 4), Mid(Trim(strDateSD), 4, 2), Mid(Trim(strDateSD), 1, 2))
                      ngayHienTai = DateSerial(Mid(Trim(DateNgayKy), 7, 4), Mid(Trim(DateNgayKy), 4, 2), Mid(Trim(DateNgayKy), 1, 2))
                      CheckNgaySD = DateDiff("d", ngayHienTai, NgaySD)
                     If CheckNgaySD < 0 Then
                         .Col = .ColLetterToNumber("T")
'                         .CellNote = errCheckNgaySD
'                         .BackColor = mAlertColor
'                         iFlagNgayIn = 0
                         .CellNote = errCheckNgaySD
                         .BackColor = mErrorColor
                         iFlagNgayIn = 1
                      Else
                          .Col = .ColLetterToNumber("T")
                          .CellNote = ""
                          .BackColor = mNonErrorColor
                     End If
               End If
               
               ' kiem tra ngay thong bao phat hanh va ngay su dung
               If Trim(strDateSD) = "" And iWhiteRow = 0 Then
                     .Col = .ColLetterToNumber("T")
                     .CellNote = errEmty
                     .BackColor = mErrorColor
                     iFlagEmty = 1
              Else
                      NgaySD = DateSerial(Mid(Trim(strDateSD), 7, 4), Mid(Trim(strDateSD), 4, 2), Mid(Trim(strDateSD), 1, 2))
                      ngayHienTai = DateSerial(Mid(Trim(DateNgayKy), 7, 4), Mid(Trim(DateNgayKy), 4, 2), Mid(Trim(DateNgayKy), 1, 2))
                      CheckNgaySD = DateDiff("d", ngayHienTai, NgaySD)
                     If CheckNgaySD < 5 Then
                         .Col = .ColLetterToNumber("T")
'                         .CellNote = errCheckNgaySD
'                         .BackColor = mAlertColor
'                         iFlagNgayIn = 0
                         .CellNote = errCheckNgaySD1
                         .BackColor = mErrorColor
                         iFlagNgayBatDau = 1
                      Else
                          .Col = .ColLetterToNumber("T")
                          .CellNote = ""
                          .BackColor = mNonErrorColor
                     End If
               End If
'        Else
'            .Col = .ColLetterToNumber("D")
'            .CellNote = errEmty
'            .BackColor = mErrorColor
'            iFlagEmty = 1
        End If
        'Kiem tra cau truc cua ky hieu
                Dim varTemp As Variant
                .GetText .ColLetterToNumber("L"), .Row, varTemp
                SohieuHD = CheckSoHD(CStr(varTemp), strLoai)
            If (Trim(SohieuHD) = "" Or Trim(SohieuHD) = "../..." Or Trim(varTemp) = vbNullString) And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("L")
                .CellNote = errEmty
                .BackColor = mErrorColor
                iFlagEmty = 1
            ElseIf (Trim(SohieuHD) <> "" Or Trim(SohieuHD) <> "../...") And iWhiteRow = 0 Then
                If Trim(SohieuHD) = "1" Or Trim(SohieuHD) = "2" Then
                      .Col = .ColLetterToNumber("L")
                      .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                      .CellNote = errKyHieuHD    'static
                      .BackColor = mErrorColor
                      iFlagEmty = 1
                      'capture the first error cell
                Else
                      .Col = .ColLetterToNumber("L")
                      .CellNote = ""
                      .BackColor = mNonErrorColor
                End If
            Else
                .Col = .ColLetterToNumber("L")
                .CellNote = ""
                .BackColor = mNonErrorColor
'                iFlagEmty = 1
            End If
       ' neu ky tu thu 6 = "P" thi hien thi phan thong tin doanh nghiep( ten, mst, so hop dong va ngay ky)
        If Mid(Trim(strKyHieuHD), 6, 1) = "P" And iWhiteRow = 0 Then
            '   ngay in phai nho hon ngay thong bao phat hanh ( ngay Ky)
'                If Trim(strDateIn) = "" Then
'                      .Col = .ColLetterToNumber("AE")
'                      .CellNote = errThongTinDN
'                      .BackColor = mErrorColor
'                      iFlagEmty = 1
'                Else
'                NgayIn = DateSerial(Mid(Trim(strDateIn), 7, 4), Mid(Trim(strDateIn), 4, 2), Mid(Trim(strDateIn), 1, 2))
'                ngayHienTai = DateSerial(Mid(Trim(DateNgayKy), 7, 4), Mid(Trim(DateNgayKy), 4, 2), Mid(Trim(DateNgayKy), 1, 2))
'                CheckNgayIn = DateDiff("d", NgayIn, ngayHienTai)
'                    If Trim(strDateIn) <> "" And CheckNgayIn < 0 Then
'                        .Col = .ColLetterToNumber("AE")
'                        .CellNote = errCheckNgayIn
'                        .BackColor = mErrorColor
'                        iFlagNgayBatDau = 1
'                    Else
'                        .Col = .ColLetterToNumber("AE")
'                        .CellNote = ""
'                        .BackColor = mNonErrorColor
'                    End If
'                End If
          '  kiem tra ten doanh nghiep
               If Trim(strTenDN) = "" Then
                   .Col = .ColLetterToNumber("W")
                   .CellNote = errThongTinDN
                   .BackColor = mErrorColor
                   iFlagEmty = 1
               Else
                   .Col = .ColLetterToNumber("W")
                   .CellNote = ""
                   .BackColor = mNonErrorColor
               End If
        '  kiem tra ma so thue doanh nghiep
                If Trim(strMstDN) = "" Then
                   .Col = .ColLetterToNumber("AA")
                   .CellNote = errThongTinDN
                   .BackColor = mErrorColor
                   iFlagEmty = 1
        ' kiem tra mst
                ElseIf Trim(strMstDN) <> vbNullString Then
                    strMstDN = Replace(strMstDN, " ", "")
                    If Len(strMstDN) = 14 And InStr(strMstDN, "-") > 0 Then
                        strMstDN = Left$(strMstDN, 10) & Right$(strMstDN, 3)
                    End If
                    strMstDN = Replace(strMstDN, "-", "")
                    strCheck = ""
                    
                    
                    If Len(strMstDN) = 10 Or Len(strMstDN) = 13 Then
                        If strMstDN <> Replace(strMstDN, " ", "") Then
                            strCheck = ""
                        Else
                            strCheck = CheckTaxCode(Mid(strMstDN, 1, 1), Mid(strMstDN, 2, 1), Mid(strMstDN, 3, 1), Mid(strMstDN, 4, 1), Mid(strMstDN, 5, 1), Mid(strMstDN, 6, 1), Mid(strMstDN, 7, 1), Mid(strMstDN, 8, 1), Mid(strMstDN, 9, 1), Mid(strMstDN, 10, 1), Mid(strMstDN, 11, 1), Mid(strMstDN, 12, 1), Mid(strMstDN, 13, 1))
                        End If
                        If Len(strMstDN) = 10 And strCheck <> "000" Then
                            .Col = .ColLetterToNumber("AA")
                            .CellNote = errMSTDN
                            .BackColor = mErrorColor
                            iFlagEmty = 1
                        ElseIf Len(strMstDN) = 13 And strCheck <> "000" Then
                            .Col = .ColLetterToNumber("AA")
                            .CellNote = errMSTDN
                            .BackColor = mErrorColor
                            iFlagEmty = 1
                        Else
                            .Col = .ColLetterToNumber("AA")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    ElseIf Len(strMstDN) <> 10 Or Len(strMstDN) <> 13 Then
                        .Col = .ColLetterToNumber("AA")
                        .CellNote = errMSTDN
                        .BackColor = mErrorColor
                        iFlagEmty = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("AA")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
        ' kiem tra so hop dong Doanh nghiep
'                If Trim(strSoHD) = "" Then
'                    .Col = .ColLetterToNumber("AD")
'                    .CellNote = errThongTinDN
'                    .BackColor = mErrorColor
'                    iFlagEmty = 1
'                Else
'                    .Col = .ColLetterToNumber("AD")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                End If
         Else
            .Col = .ColLetterToNumber("W")
            .CellNote = ""
            .BackColor = mNonErrorColor
            .Col = .ColLetterToNumber("AA")
            .CellNote = ""
            .BackColor = mNonErrorColor
            .Col = .ColLetterToNumber("AD")
            .CellNote = ""
            .BackColor = mNonErrorColor
            .Col = .ColLetterToNumber("AE")
            .CellNote = ""
            .BackColor = mNonErrorColor
        End If
        If Mid(Trim(strKyHieuHD), 6, 1) <> "P" And iWhiteRow = 0 Then
                ' ngay ky
'                If Trim(strDateIn) = "" Then
'                    .Col = .ColLetterToNumber("AE")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                Else
'                    NgayIn = DateSerial(Mid(Trim(strDateIn), 7, 4), Mid(Trim(strDateIn), 4, 2), Mid(Trim(strDateIn), 1, 2))
'                    ngayHienTai = DateSerial(Mid(Trim(DateNgayKy), 7, 4), Mid(Trim(DateNgayKy), 4, 2), Mid(Trim(DateNgayKy), 1, 2))
'                    CheckNgayIn = DateDiff("d", NgayIn, ngayHienTai)
'                    If CheckNgayIn < 0 And Trim(strDateIn) <> "" Then
'                            .Col = .ColLetterToNumber("AE")
'                            .CellNote = errCheckNgayIn
'                            .BackColor = mErrorColor
'                            'iFlagNgayBatDau = 0
'                            iFlagNgayBatDau = 1
'                    Else
'                        .Col = .ColLetterToNumber("AE")
'                        .CellNote = ""
'                        .BackColor = mNonErrorColor
'                    End If
'                End If
            ' bat cau truc MST doanh nghiep
                If Trim(strMstDN) <> vbNullString Then
                    strMstDN = Replace(strMstDN, " ", "")
                    If Len(strMstDN) = 14 And InStr(strMstDN, "-") > 0 Then
                        strMstDN = Left$(strMstDN, 10) & Right$(strMstDN, 3)
                    End If
                    strMstDN = Replace(strMstDN, "-", "")
                    
                    strCheck = ""
                    If Len(strMstDN) = 10 Or Len(strMstDN) = 13 Then
                        If strMstDN <> Replace(strMstDN, " ", "") Then
                            strCheck = ""
                        Else
                            strCheck = CheckTaxCode(Mid(strMstDN, 1, 1), Mid(strMstDN, 2, 1), Mid(strMstDN, 3, 1), Mid(strMstDN, 4, 1), Mid(strMstDN, 5, 1), Mid(strMstDN, 6, 1), Mid(strMstDN, 7, 1), Mid(strMstDN, 8, 1), Mid(strMstDN, 9, 1), Mid(strMstDN, 10, 1), Mid(strMstDN, 11, 1), Mid(strMstDN, 12, 1), Mid(strMstDN, 13, 1))
                        End If
                        
                        If Len(strMstDN) = 10 And strCheck <> "000" Then
                            .Col = .ColLetterToNumber("AA")
                            .CellNote = errMSTDN
                            .BackColor = mErrorColor
                            iFlagEmty = 1
                        ElseIf (Len(strMstDN) = 13 And strCheck <> "000") Or Not IsNumeric(strMstDN) Then
                            .Col = .ColLetterToNumber("AA")
                            .CellNote = errMSTDN
                            .BackColor = mErrorColor
                            iFlagEmty = 1
                        Else
                            .Col = .ColLetterToNumber("AA")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    ElseIf Len(strMstDN) <> 10 Or Len(strMstDN) <> 13 Then
                        .Col = .ColLetterToNumber("AA")
                        .CellNote = errMSTDN
                        .BackColor = mErrorColor
                        iFlagEmty = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("AA")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
'        Else
'            .Col = .ColLetterToNumber("V")
'            .CellNote = ""
'            .BackColor = mNonErrorColor
'            .Col = .ColLetterToNumber("Z")
'            .CellNote = ""
'            .BackColor = mNonErrorColor
'            .Col = .ColLetterToNumber("AC")
'            .CellNote = ""
'            .BackColor = mNonErrorColor
'            .Col = .ColLetterToNumber("AD")
'            .CellNote = ""
'            .BackColor = mNonErrorColor
        End If
         'Ten HD
            If Trim(strTenHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("D")
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                .Col = .ColLetterToNumber("D")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
          
        ' Mau so HD
            If Trim(strMauSoHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("J")
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            ElseIf Trim(strMauSoHD) <> "" And iWhiteRow = 0 Then
                    'strLoai = Trim(Left(strId, 1))
                        If strLoai = "0" Then
                            strTemp = Trim(Mid(strId, 11, 6))
                        ElseIf strLoai = "4" Then
                            strTemp = Trim(Mid(strId, 11, 6))
                        Else
                            strTemp = Trim(Mid(strId, 11, 3))
                        End If
                    result = CheckMauSoHD(CStr(strMauSoHD), strLoai, strTemp)
                        If result = "1" And iWhiteRow = 0 Then
                            .Col = .ColLetterToNumber("J")
                            .CellNote = errMauSoHD  'static
                            .BackColor = mErrorColor
                            iFlagMauSoHD = 1
                        ElseIf result = "2" And iWhiteRow = 0 Then
                            If Right$(Trim$(strKyHieuHD), 1) = "E" Then
                                If CheckSoLienHDDT(CStr(strMauSoHD), strLoai, strTemp, CStr(strKyHieuHD)) = "3" Then
                                     .Col = .ColLetterToNumber("J")
                                    .CellNote = errSoLienHDDT  'static
                                    .BackColor = mAlertColor
                                    '.BackColor = mErrorColor
                                    'iFlagSoLienHD = 2
                                Else
                                    .Col = .ColLetterToNumber("J")
                                    .CellNote = ""
                                    .BackColor = mNonErrorColor
                                End If
                            Else
                                .Col = .ColLetterToNumber("J")
                                .CellNote = errSoLienHD  'static
                                .BackColor = mErrorColor
                                iFlagSoLienHD = 1
                            End If
                        Else
                            If Right$(Trim$(strKyHieuHD), 1) = "E" Then
                                If CheckSoLienHDDT(CStr(strMauSoHD), strLoai, strTemp, CStr(strKyHieuHD)) = "3" Then
                                     .Col = .ColLetterToNumber("J")
                                    .CellNote = errSoLienHDDT  'static
                                    .BackColor = mAlertColor
'                                    .BackColor = mErrorColor
'                                    iFlagSoLienHD = 2
                                Else
                                    .Col = .ColLetterToNumber("J")
                                    .CellNote = ""
                                    .BackColor = mNonErrorColor
                                End If
                            Else
                                .Col = .ColLetterToNumber("J")
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            End If
                        End If
            Else
                    .Col = .ColLetterToNumber("J")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
            End If
         ' check du lieu "trang"
                If (Trim(strTenHD) <> "" Or Trim(SohieuHD) <> "" Or Trim(strMauSoHD) <> "") And iWhiteRow = 0 Then
                        If Trim(strTuSo) = "" And Trim(strDenso) = "" Then
                           .Col = .ColLetterToNumber("P")
                           .CellNote = errEmty
                           .BackColor = mErrorColor
                           iFlagEmty = 1
                           .Col = .ColLetterToNumber("R")
                           .CellNote = errEmty
                           .BackColor = mErrorColor
                           iFlagEmty = 1
                        ElseIf Trim(strTuSo) = "" And Trim(strDenso) <> "" Then
                            .Col = .ColLetterToNumber("P")
                            .CellNote = errEmty
                            .BackColor = mErrorColor
                            iFlagEmty = 1
                        
                        ElseIf Trim(strTuSo) <> "" And Trim(strDenso) = "" Then
                            .Col = .ColLetterToNumber("R")
                            .CellNote = errEmty
                            .BackColor = mErrorColor
                            iFlagEmty = 1
                     'check tu so < den so
                       ElseIf Trim(strTuSo) <> "" And Trim(strDenso) <> "" Then
                             strTuSo = Val(strTuSo)
                             strDenso = Val(strDenso)
                            If strTuSo > strDenso Then
                                .Col = .ColLetterToNumber("P")
                                .CellNote = errMsgTusoDenSo
                                .BackColor = mErrorColor
                                'iFlagEmty = 1
                                .Col = .ColLetterToNumber("R")
                                .CellNote = errMsgTusoDenSo
                                .BackColor = mErrorColor
                                iFlagTusoDenSo = 1
                              Else
                              End If
                    ' kiem tra trung khoang HD
                             If UBound(str) > 0 Then
                                For k = 0 To UBound(str)
                                    If j = str(k) Then
                                        .Col = .ColLetterToNumber("P")
                                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                        .CellNote = errKhoangTN     'static
                                        .BackColor = mErrorColor
                                         .Col = .ColLetterToNumber("R")
                                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                        .CellNote = errKhoangTN     'static
                                        .BackColor = mErrorColor
                                        iFlagEmty = 1
                                     Exit For
                                    End If
                                Next k
                             Else
'                                .Col = .ColLetterToNumber("O")
'                                .CellNote = ""
'                                .BackColor = mNonErrorColor
'                                .Col = .ColLetterToNumber("Q")
'                                .CellNote = ""
'                                .BackColor = mNonErrorColor
                            End If
                       Else
                                .Col = .ColLetterToNumber("P")
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                                .Col = .ColLetterToNumber("R")
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                       End If
                 End If
             
             .Col = .ColLetterToNumber("B")
             .Row = i + 35
                    i = i + 1
            Loop Until .Text = "aa"
            ' set lai trang thai cua HDR
              .Sheet = .SheetCount
             .SetText .ColLetterToNumber("B"), 24, IIf(iFlagEmty = 1, "0", "1")
             .SetText .ColLetterToNumber("B"), 18, IIf(iFlagMauSoHD = 1, "0", "1")
             
             .SetText .ColLetterToNumber("B"), 20, IIf(iFlagSoLienHD = 1, "0", "1")
             .SetText .ColLetterToNumber("B"), 34, IIf(iFlagSoLienHD = 2, "0", "1")
             
             .SetText .ColLetterToNumber("B"), 27, IIf(iFlagDVCQ = 1, "0", "1")
             .SetText .ColLetterToNumber("B"), 30, IIf(iFlagNgayKy = 1, "0", "1")
             .SetText .ColLetterToNumber("B"), 26, IIf(iFlagNgayBatDau = 1, "0", "1")
             '.SetText .ColLetterToNumber("B"), 28, IIf(iFlagNgayIn = 1, "0", "1")
             .SetText .ColLetterToNumber("B"), 28, IIf(iFlagNgayBatDau = 1, "0", "1")
             .SetText .ColLetterToNumber("B"), 23, IIf(iFlagTusoDenSo = 1, "0", "1")
             
             .SetText .ColLetterToNumber("B"), 33, IIf(iFlagNgayIn = 1, "0", "1")
            'CheckErrorMST
    End With
End Sub
Sub CheckErrorMST()
    Dim vError1 As Variant, vError2 As Variant, vError3 As Variant
    Dim MST1 As Variant, MST2 As Variant, MST3 As Variant, MST4 As Variant
    Dim MST5 As Variant, MST6 As Variant, MST7 As Variant, MST8 As Variant
    Dim MST9 As Variant, MST10 As Variant, MST11 As Variant, MST12 As Variant, MST13 As Variant
    Dim iCurrentSheet As Integer, strCheck As String
    Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    With fps
        
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 15, vError1
        .GetText .ColLetterToNumber("E"), 16, vError2
        .GetText .ColLetterToNumber("E"), 17, vError3
        
        .GetText .ColLetterToNumber(SxMST1Col), SxMST1Row, MST1
        .GetText .ColLetterToNumber(SxMST2Col), SxMST2Row, MST2
        .GetText .ColLetterToNumber(SxMST3Col), SxMST3Row, MST3
        .GetText .ColLetterToNumber(SxMST4Col), SxMST4Row, MST4
        .GetText .ColLetterToNumber(SxMST5Col), SxMST5Row, MST5
        .GetText .ColLetterToNumber(SxMST6Col), SxMST6Row, MST6
        .GetText .ColLetterToNumber(SxMST7Col), SxMST7Row, MST7
        .GetText .ColLetterToNumber(SxMST8Col), SxMST8Row, MST8
        .GetText .ColLetterToNumber(SxMST9Col), SxMST9Row, MST9
        .GetText .ColLetterToNumber(SxMST10Col), SxMST10Row, MST10
        .GetText .ColLetterToNumber(SxMST11Col), SxMST11Row, MST11
        .GetText .ColLetterToNumber(SxMST12Col), SxMST12Row, MST12
        .GetText .ColLetterToNumber(SxMST13Col), SxMST13Row, MST13
        
        strCheck = CheckTaxCode(MST1, MST2, MST3, MST4, MST5, MST6, MST7, MST8, MST9, MST10, MST11, MST12, MST13)
        iFlagTaxCode1 = CInt(Left(strCheck, 1))
        strCheck = Right(strCheck, Len(strCheck) - 1)
        iFlagTaxCode2 = CInt(Left(strCheck, 1))
        strCheck = Right(strCheck, Len(strCheck) - 1)
        iFlagTaxCode3 = CInt(strCheck)
        
        If iFlagTaxCode1 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 15
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 15
            .Text = "1"
        End If
        
        If iFlagTaxCode2 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 16
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 16
            .Text = "1"
        End If
        
        
        If iFlagTaxCode3 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 17
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 17
            .Text = "1"
        End If
        

        .Sheet = 1
        .Col = .ColLetterToNumber("E")
        .Row = 4
        .CellNote = ""
        .BackColor = mFormColor
        If iFlagTaxCode1 = 1 Then
            .CellNote = .CellNote & "> " & vError1
            .BackColor = mErrorColor
        End If
        If iFlagTaxCode2 = 1 Then
            .CellNote = .CellNote & IIf(Trim(.CellNote) = "", "", vbCrLf) & "> " & vError2
            .BackColor = mErrorColor
        End If
        If iFlagTaxCode3 = 1 Then
            .CellNote = .CellNote & IIf(Trim(.CellNote) = "", "", vbCrLf) & "> " & vError3
            .BackColor = mErrorColor
        End If

        
        .Sheet = iCurrentSheet
    End With
End Sub

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

Public Sub SetActiveSheet()

End Sub

Private Sub kiemTraDuLieuImport()

End Sub

Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean
    Dim totalCell  As Long
    Dim countRowReset As Long
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    totalCell = TAX_Utilities_v1.Data(0).getElementsByTagName("Cell").length
    countRowReset = (totalCell - 5) / 14
    With fps
        For Each xmlNodeReset In TAX_Utilities_v1.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            .Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            If (lRow < 35 Or lRow > 35 + countRowReset) And mCurrentSheet = 1 Then
            Else
                .Col = lCol
                .Row = lRow
                
                Select Case .CellType
                    Case CellTypeCheckBox
                        .Text = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                    Case CellTypeComboBox
                        .Text = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                    Case CellTypeNumber
                        .value = 0
                        UpdateCell fps, lCol, lRow, "0"
                    Case Else
                        .value = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                End Select
                'mAdjustData = IIf(IsUpdate = True, IsUpdate, mAdjustData)
                TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
            End If
        Next
    End With
End Function

Public Sub SetData()

End Sub
Private Sub UpdateSheets()
    Dim varTemp As Variant
    Dim ssSheet As Integer
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet
        If (GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "Active") <> "0") Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing

End Sub


Public Function GetDVCQ() As String
    Dim xmlDomHeader As New MSXML.DOMDocument
    Dim strMST As String
    
    xmlDomHeader.Load GetAbsolutePath(TAX_Utilities_v1.DataFolder & "Header_01.xml")
    strMST = GetAttribute(xmlDomHeader.nodeFromID("C_34"), "Value")
    If Len(strMST) = 13 Then
        strMST = Left$(strMST, 10) & "-" & Right$(strMST, 3)
    End If
    'GetDVCQ = GetAttribute(xmlDomHeader.nodeFromID("C_34"), "Value")
    GetDVCQ = strMST
    
    Set xmlDomHeader = Nothing
End Function
Public Function GetTenDVCQ() As String
    Dim xmlDomHeader As New MSXML.DOMDocument
    
    xmlDomHeader.Load GetAbsolutePath(TAX_Utilities_v1.DataFolder & "Header_01.xml")
    GetTenDVCQ = GetAttribute(xmlDomHeader.nodeFromID("C_32"), "Value")
    
    Set xmlDomHeader = Nothing
End Function
Public Function GetNguoiKy() As String
    Dim xmlDomHeader As New MSXML.DOMDocument
    
    xmlDomHeader.Load GetAbsolutePath(TAX_Utilities_v1.DataFolder & "Header_01.xml")
    GetNguoiKy = GetAttribute(xmlDomHeader.nodeFromID("C_26"), "Value")
    
    Set xmlDomHeader = Nothing
End Function


' ham set thong tin sau khi tai XML
Public Sub finishImportXML()
    Dim maHD    As Variant
    Dim tenHD  As String
    Dim arrTemp() As String
    With fps
        .EventEnabled(EventAllEvents) = False
          .Sheet = 1
          mCurrentSheet = .Sheet
            .Row = 35
            Do
                .GetText .ColLetterToNumber("J"), .Row, maHD
                 GetMaHoaDon maHD, tenHD
                 If tenHD <> "" Then
                    arrTemp = Split(tenHD, "###")
                    If UBound(arrTemp) > 0 Then
                        .Col = .ColLetterToNumber("C")
                        .Text = arrTemp(0)
                        UpdateCell fps, .Col, .Row, .Text
                        .Col = .ColLetterToNumber("AG")
                        .Text = CStr(arrTemp(1))
                        UpdateCell fps, .Col, .Row, .Text
                    End If
                 End If
                 
                 
                 .Col = .ColLetterToNumber("B")
                 .Row = .Row + 1
            Loop Until .Text = "aa"
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
