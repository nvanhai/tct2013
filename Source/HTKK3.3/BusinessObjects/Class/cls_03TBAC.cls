VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_03TBAC"
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
Dim strQuanHeNNT As String
Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

 strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_loai_HD.xml")
        ' Lay danh muc loai hoa don
        ' 15/11/2010
        If xmlDomData.Load(strDataFileName) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    strCombo = strCombo + CPab(arrDanhsach(0), 10) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(2), 200) + Chr$(9)
'                    strCombo = strCombo + CPab(arrDanhsach(1), 10) + Chr$(9)
                    strComboHien = strComboHien + arrDanhsach(1) + Chr$(9)
                End If
            Next
        End If
    
     With fps
        SetDateFormat fps, 1, 8, .ColLetterToNumber("O"), DDMMYYYY
        SetDateFormat fps, 1, 49, .ColLetterToNumber("R"), DDMMYYYY
        'SetKyHieuHDFormat fps, 1, 43, .ColLetterToNumber("S")
        'SetKyHieuHDFormat fps, 1, 43, .ColLetterToNumber("S")
          .Sheet = 1
          .Row = 43
          ' set ma hoa don
        .Col = .ColLetterToNumber("C")
        .TypeComboBoxList = strComboHien
        .Text = strQuanHeNNT
        .Col = .ColLetterToNumber("AI")
        .TypeComboBoxList = strCombo
        .Sheet = 1
        .Row = 49
       
        ' Set format ngay ky
        .Col = .ColLetterToNumber("R")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        ' Set format ngay thong bao huy
        .Row = 8
        .Col = .ColLetterToNumber("O")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        ' set gio huy
        .Col = .ColLetterToNumber("L")
        .Text = Hour(Time)
        .TypeHAlign = TypeHAlignLeft
        ' set phut huy
        .Col = .ColLetterToNumber("N")
        .Text = Minute(Time)
        .TypeHAlign = TypeHAlignLeft
        
        ' Set gia tri cho combo an
        .Col = .ColLetterToNumber("AI")
        .Row = 43
        .TypeComboBoxList = strCombo
        
    End With
    
    
    
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
Dim i As Integer

Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim xmlNode As MSXML.IXMLDOMNode
Dim iCol As Long, iRow As Long

Dim strGio, strPhut, strNgay As Variant

Dim totalRow As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        
        ' set gia tri de ket xuat XML
        .GetText .ColLetterToNumber("O"), 8, strNgay
        .GetText .ColLetterToNumber("L"), 8, strGio
        .GetText .ColLetterToNumber("N"), 8, strPhut
        
        If Len(strGio) = 1 Then
            strGio = "0" & strGio
        End If
        
        If Len(strPhut) = 1 Then
            strPhut = "0" & strPhut
        End If
        .Row = 9
        .Col = .ColLetterToNumber("O")
        .Text = strNgay & "T" & strGio & ":" & strPhut & ":00Z"
        ' end set gia tri de ket xuat XML
        
        i = 1
        .Col = .ColLetterToNumber("B")
        .Row = 43
        Do
'             .Col = .ColLetterToNumber("C")
'             .Text = "0"
             
             .Col = .ColLetterToNumber("B")
             .Row = i + 43
             totalRow = i
             i = i + 1
        Loop Until .Text = "aa"
        'dhdang sua theo yc chi en ngay 25/01
        .Row = totalRow + 46
        .Col = .ColLetterToNumber("R")
        If Trim(.Text) = "" Then
            .SetText .ColLetterToNumber("R"), totalRow + 46, GetNguoiKy()
        End If
        
         'Xoa cong thuc cua o Chu ky
        Set xmlCellNode = TAX_Utilities_v2.Data(0).nodeFromID(GetCellID(fps, .ColLetterToNumber("E"), 43))
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode.lastChild
        Set xmlCellNode = xmlCellNode.firstChild
        Set xmlCellNode = xmlCellNode.firstChild
        Set xmlCellNode = xmlCellNode.nextSibling
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        .Row = iRow
        .Col = iCol
        .Formula = ""
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
 Dim strId As Variant
 Dim intIndexCombo As Integer
'Dim Col7 As Variant
'Dim iRow As Long, iCol As Long, GTSua As String
'Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
'Dim tCol As Long, tRow As Long
'Dim cCol As Long, cRow As Long
'Dim cValue As Variant, temp As Variant, varTemp As Variant
'Dim cValueG As Variant
'Dim cTagF As String
'Dim cValueF As Variant
'Dim i As Integer
'Dim xmlNodeH15 As MSXML.IXMLDOMNode

Dim strGio, strPhut, strNgay As Variant

    Dim strTuSo As Variant
    Dim strDenso As Variant
    fps.EventEnabled(EventAllEvents) = False
    With fps
        
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("C") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("AI")
            .typeComboboxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber("AI")
            UpdateCell_sheet fps, 1, .Col, .Row, .Text
             
            
            .GetText .ColLetterToNumber("AI"), .Row, strId
            
           If Trim(Left(strId, 1)) = "0" Then
                .SetText .ColLetterToNumber("O"), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell_sheet fps, 1, .ColLetterToNumber("O"), .Row, .Text
                .Col = .ColLetterToNumber("O")
                .TypeMaxEditLen = 13
                .Col = .ColLetterToNumber("E")
                .SetText .ColLetterToNumber("E"), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 1, .ColLetterToNumber("E"), .Row, .Text
            ElseIf Trim(Left(strId, 1)) = "3" Then
                .SetText .ColLetterToNumber("O"), .Row, ""
                UpdateCell_sheet fps, 1, .ColLetterToNumber("O"), .Row, .Text
                .Col = .ColLetterToNumber("O")
                .TypeMaxEditLen = 20
                .Col = .ColLetterToNumber("E")
                .SetText .ColLetterToNumber("E"), .Row, ""
                UpdateCell_sheet fps, 1, .ColLetterToNumber("E"), .Row, .Text
            ElseIf Trim(Left(strId, 1)) = "4" Then
                .SetText .ColLetterToNumber("O"), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell_sheet fps, 1, .ColLetterToNumber("O"), .Row, .Text
                .Col = .ColLetterToNumber("O")
                .TypeMaxEditLen = 13
                .Col = .ColLetterToNumber("E")
                .SetText .ColLetterToNumber("E"), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 1, .ColLetterToNumber("E"), .Row, .Text
            Else
                .SetText .ColLetterToNumber("O"), .Row, Trim(Mid(strId, 11, 3))
                UpdateCell_sheet fps, 1, .ColLetterToNumber("O"), .Row, .Text
                .Col = .ColLetterToNumber("O")
                .TypeMaxEditLen = 20
                .Col = .ColLetterToNumber("E")
                .SetText .ColLetterToNumber("E"), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 1, .ColLetterToNumber("E"), .Row, .Text
            End If
'            .SetText .ColLetterToNumber("AJ), .Row, Trim(Left(strId, 1))
            'UpdateCell_sheet fps, 1, .ColLetterToNumber("AA"), .Row, "0"
            .GetText .ColLetterToNumber("V"), .Row, strTuSo
            .GetText .ColLetterToNumber("Y"), .Row, strDenso
            If Trim(CStr(strTuSo)) = "" And Trim(CStr(strDenso)) = "" Then
                .SetText .ColLetterToNumber("AA"), .Row, "0"
                UpdateCell_sheet fps, 1, .ColLetterToNumber("AA"), .Row, "0"
            End If
            
            .EventEnabled(EventChange) = True
         End If
         
         If (.Col = .ColLetterToNumber("O") Or .Col = .ColLetterToNumber("L") Or .Col = .ColLetterToNumber("N")) And .Row = 8 Then
            ' set gia tri de ket xuat XML
            .GetText .ColLetterToNumber("O"), 8, strNgay
            .GetText .ColLetterToNumber("L"), 8, strGio
            .GetText .ColLetterToNumber("N"), 8, strPhut
            
            If Len(strGio) = 1 Then
                strGio = "0" & strGio
            End If
            
            If Len(strPhut) = 1 Then
                strPhut = "0" & strPhut
            End If
            .Row = 9
            .Col = .ColLetterToNumber("O")
            .Text = strNgay & "T" & strGio & ":" & strPhut & ":00Z"
            ' end set gia tri de ket xuat XML
         End If
    End With
   fps.EventEnabled(EventAllEvents) = True
End Sub

Private Sub fps_KeyPress(KeyAscii As Integer)
    With fps
        If .Col = .ColLetterToNumber("V") Or .Col = .ColLetterToNumber("Y") Then
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
                    .Row = 43
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 43
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
    
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
        'check ngay thong bao huy TBAC
            If Col = .ColLetterToNumber("O") And Row = 8 Then
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
        ' Dem so dong dong
        .Row = 43
        Do
            countRowDynamic = countRowDynamic + 1
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
            
        ' check ngay ky TBAC
            If Col = .ColLetterToNumber("R") And Row = countRowDynamic + 48 Then
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
          ' Format dinh dang tu so den so co 7 ky tu
            If Col = .ColLetterToNumber("V") Or Col = .ColLetterToNumber("Y") Then
                 .GetText Col, Row, varTemp
                 .SetText Col, Row, FormatSoHD(CStr(varTemp))
                 .Col = Col
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
           ' Ho tro tinh so luong
                .GetText .ColLetterToNumber("V"), Row, tuSo
                .GetText .ColLetterToNumber("Y"), Row, denso
                If tuSo <> "" And denso <> "" Then
                    If Val(tuSo) = 0 And Val(denso) = 0 Then
                        .SetText .ColLetterToNumber("AA"), Row, 0
                    ElseIf Val(tuSo) = 0 And Val(denso) <> 0 Then
                        .SetText .ColLetterToNumber("AA"), Row, Val(denso) - Val(tuSo)
                    Else
                        .SetText .ColLetterToNumber("AA"), Row, Val(denso) - Val(tuSo) + 1
                    End If
                    
                    .Col = Col
                    .Row = Row
                    UpdateCell fps, .Col, .Row, .Text
                Else
                    .SetText .ColLetterToNumber("AA"), Row, "0"
                    .Col = Col
                    .Row = Row
                    UpdateCell fps, .Col, .Row, .Text
                End If
                
            End If
            
          ' Chuyen ky hieu HD thanh chu cai in hoa
            If Col = .ColLetterToNumber("S") Then
                 .GetText Col, Row, varTemp
                 .SetText Col, Row, FormatKyHieu(CStr(varTemp))
                 .Col = Col
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
            End If
            .GetText .ColLetterToNumber("AI"), Row, strId
            strLoai = Trim(Left(strId, 1))
           'Kiem tra cau truc cua ky hieu neu sai set cot an =1 dung =0
            If Col = .ColLetterToNumber("S") Then
                 .GetText Col, Row, varTemp
                 .SetText .ColLetterToNumber("AH"), Row, CheckSoHD(CStr(varTemp), strLoai)
                 .Col = .ColLetterToNumber("AH")
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
            End If
           'Kiem tra cau truc cua Mau so
           If Col = .ColLetterToNumber("O") Then
'                .GetText .ColLetterToNumber("AI"), Row, strId
'                strLoai = Trim(Left(strId, 1))
                If strLoai = "0" Then
                    strTemp = Trim(Mid(strId, 11, 6))
                ElseIf strLoai = "4" Then
                    strTemp = Trim(Mid(strId, 11, 6))
                Else
                    strTemp = Trim(Mid(strId, 11, 3))
                End If
                .GetText Col, Row, varTemp
                .SetText Col, Row, UCase(varTemp)
                 .SetText .ColLetterToNumber("AJ"), Row, CheckMauSoHD(CStr(varTemp), CStr(strLoai), CStr(strTemp))
                 .Col = .ColLetterToNumber("AJ")
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

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    Dim strCheck As String
    Dim errMauSoHD As Variant, errKyHieuHD As Variant, errSoLuong As Variant, errSoLienHD As Variant
    Dim errKhoangGN As Variant, errLengthHD As Variant, errEmty As Variant, errNgayKy As Variant, errNgayHuy As Variant, errTrungKhoang As Variant
    Dim errNgayHuyHL As Variant, errNgayKyHL As Variant
    Dim strTenHD As Variant, strMauSoHD As Variant, strKyHieuHD As Variant, strTuSo As Variant, strDenso As Variant
    Dim strTenHD1 As Variant, strMauSoHD1 As Variant, strKyHieuHD1 As Variant, strTuso1 As Variant, strDenso1 As Variant
    Dim strSoLuong As Variant, strNguoiKy As Variant, strNguoiLapBieu As Variant, strNgayKy As Variant, strNgayHuy As Variant
    Dim iFlagTenHD As Integer, iFlagMauSoHD As Integer, iFlagKyHieu As Integer, iFlagSoLienHD As Integer, iFlagEmty As Integer
    Dim iWhiteRow As Integer, iFlagLengHD As Integer, iFlagSoLuong As Integer, iFlagNgayKy As Integer, iFlagNgayHuy As Integer, iFlagNgayHuyHL As Integer, iFlagNgayKyHL As Integer
    Dim iFlagKhoangTSDSTrung As Integer, iFocusFlag As Integer
    
    Dim errSoLienHDDT As Variant
    
    Dim strId As Variant, strLoai As String, strTemp As String, result As String
    Dim i As Integer, countDynamicRow As Integer, j As Integer, k As Integer
    
    Dim str() As String
    ' tra ve mang cac row bi trung khoang
    str = checkTrung(fps, 43, "V", "Y", "B", "E", "O", "S", "", False, 1)
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 14, errMauSoHD
        .GetText .ColLetterToNumber("E"), 15, errKyHieuHD
        .GetText .ColLetterToNumber("E"), 16, errSoLienHD
        .GetText .ColLetterToNumber("E"), 17, errKhoangGN
        .GetText .ColLetterToNumber("E"), 18, errLengthHD
        .GetText .ColLetterToNumber("E"), 19, errSoLuong
        .GetText .ColLetterToNumber("E"), 20, errEmty
        .GetText .ColLetterToNumber("E"), 21, errNgayHuy
        .GetText .ColLetterToNumber("E"), 22, errNgayKy
        .GetText .ColLetterToNumber("E"), 23, errTrungKhoang

        .GetText .ColLetterToNumber("E"), 27, errNgayHuyHL
        .GetText .ColLetterToNumber("E"), 28, errNgayKyHL
        
        .GetText .ColLetterToNumber("E"), 29, errSoLienHDDT
        
        
        'set flag bang 0
        iFlagTenHD = 0
        iFlagMauSoHD = 0
        iFlagKyHieu = 0
        iFlagSoLienHD = 0
        iFlagEmty = 0
        iFlagSoLuong = 0
        iFlagNgayKy = 0
        iFlagNgayHuy = 0
        iFlagKhoangTSDSTrung = 0
        
        ' dem so row dynamic
        countDynamicRow = 0
        ' Kiem tra tren sheet 1
        .Sheet = 1
        .Col = .ColLetterToNumber("B")
        .Row = 43
        'i = 1
        Do
            j = .Row
            countDynamicRow = countDynamicRow + 1
             ' Lay cac gia tri cua 1 row
             iWhiteRow = 1
            .GetText .ColLetterToNumber("E"), .Row, strTenHD
            .GetText .ColLetterToNumber("O"), .Row, strMauSoHD
            .GetText .ColLetterToNumber("S"), .Row, strKyHieuHD
            .GetText .ColLetterToNumber("V"), .Row, strTuSo
            .GetText .ColLetterToNumber("Y"), .Row, strDenso
            .GetText .ColLetterToNumber("AA"), .Row, strSoLuong
            .GetText .ColLetterToNumber("AI"), .Row, strId
            strLoai = Trim(Left(strId, 1))
            
            
            ' set gia tri cho check KH NN/XC
            strMauSoHD_01GTKT = Trim(Mid(strId, 11, 6))
            ' end
            
            If Trim(strTenHD) <> "" Or Trim(strMauSoHD) <> "" Or Trim(strKyHieuHD) <> "" Or Trim(strId) <> "" Or Trim(strTuSo) <> "" Or Trim(strDenso) <> "" Then
                ' Khong phai dong trang
                iWhiteRow = 0
            End If
            
            'Ten HD
            If Trim(strTenHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("E")
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                .Col = .ColLetterToNumber("E")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            ' Mau so HD
            If Trim(strMauSoHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("O")
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
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
                    .Col = .ColLetterToNumber("O")
                    .CellNote = errMauSoHD  'static
                    .BackColor = mErrorColor
                    iFlagMauSoHD = 1
                ElseIf result = "2" And iWhiteRow = 0 Then
                    If Right$(Trim$(strKyHieuHD), 1) = "E" Then
                        If CheckSoLienHDDT(CStr(strMauSoHD), strLoai, strTemp, CStr(strKyHieuHD)) = "3" Then
                             .Col = .ColLetterToNumber("O")
                            .CellNote = errSoLienHDDT  'static
'                            .BackColor = mErrorColor
'                            iFlagSoLienHD = 2
                             .BackColor = mAlertColor
                        Else
                            .Col = .ColLetterToNumber("O")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    Else
                        .Col = .ColLetterToNumber("O")
                        .CellNote = errSoLienHD  'static
                        .BackColor = mErrorColor
                        iFlagSoLienHD = 1
                    End If
                Else
                     If Right$(Trim$(strKyHieuHD), 1) = "E" Then
                        If CheckSoLienHDDT(CStr(strMauSoHD), strLoai, strTemp, CStr(strKyHieuHD)) = "3" Then
                             .Col = .ColLetterToNumber("O")
                            .CellNote = errSoLienHDDT  'static
'                            .BackColor = mErrorColor
'                            iFlagSoLienHD = 2
                             .BackColor = mAlertColor
                        Else
                            .Col = .ColLetterToNumber("O")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    Else
                        .Col = .ColLetterToNumber("O")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                End If
            End If
            ' check ky hieu
            If Trim(strKyHieuHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("S")
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                result = CheckSoHD(CStr(strKyHieuHD), strLoai)
                If result = "1" And iWhiteRow = 0 Then
                    .Col = .ColLetterToNumber("S")
                    .CellNote = errKyHieuHD  'static
                    .BackColor = mErrorColor
                    iFlagKyHieu = 1
                ElseIf result = "2" And iWhiteRow = 0 Then
                    .Col = .ColLetterToNumber("S")
                    .CellNote = errLengthHD  'static
                    .BackColor = mErrorColor
                    iFlagLengHD = 1
                Else
                    .Col = .ColLetterToNumber("S")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            End If
            ' kiem tra tu so den so
            If iWhiteRow = 0 Then
                If Trim(strTuSo) = "" And Trim(strDenso) = "" Then
                    ' tu so
                    .Col = .ColLetterToNumber("V")
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    ' den so
                    .Col = .ColLetterToNumber("Y")
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                ElseIf Trim(strTuSo) = "" And Trim(strDenso) <> "" Then
                    ' tu so
                    .Col = .ColLetterToNumber("V")
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                    ' den so
                    .Col = .ColLetterToNumber("Y")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                ElseIf Trim(strTuSo) <> "" And Trim(strDenso) = "" Then
                    .Col = .ColLetterToNumber("Y")
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                    ' tu so
                    .Col = .ColLetterToNumber("V")
                    .CellNote = ""
                    .BackColor = mNonErrorColor

                Else
                    If (Val(strDenso) - Val(strTuSo) >= 0 And Val(strDenso) > 0) Or (Val(strTuSo) = 0 And Val(strDenso) = 0) Then
                        ' tu so
                        .Col = .ColLetterToNumber("V")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        ' den so
                        .Col = .ColLetterToNumber("Y")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    Else
                        .Col = .ColLetterToNumber("V")
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        .Col = .ColLetterToNumber("Y")
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        iFlagSoLuong = 1
                    End If
                End If
            Else
                ' tu so
                .Col = .ColLetterToNumber("V")
                .CellNote = ""
                .BackColor = mNonErrorColor
                ' den so
                .Col = .ColLetterToNumber("Y")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            ' kiem tra trung khoang HD
            If UBound(str) > 0 And iWhiteRow = 0 Then
                For i = 0 To UBound(str)
                    If j = str(i) Then
                        .Col = .ColLetterToNumber("V")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errTrungKhoang     'static
                        .BackColor = mErrorColor
                        
                         .Col = .ColLetterToNumber("Y")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errTrungKhoang     'static
                        .BackColor = mErrorColor
                        iFlagKhoangTSDSTrung = 1
                        Exit For
                    End If
                Next i
            End If
            ' end check trung
            .Row = j
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
                
        ' Kiem tra du lieu nguoi dai dien
         .GetText .ColLetterToNumber("R"), countDynamicRow + 46, strNguoiKy
         .GetText .ColLetterToNumber("D"), countDynamicRow + 46, strNguoiLapBieu
         .GetText .ColLetterToNumber("R"), countDynamicRow + 48, strNgayKy
         .GetText .ColLetterToNumber("O"), 8, strNgayHuy
         ' kiem tra nguoi ky bat buoc nhap
         If Trim(strNguoiKy) = "" Then
            .Row = countDynamicRow + 46
            .Col = .ColLetterToNumber("R")
            .CellNote = errEmty  'static
            .BackColor = mErrorColor
            iFlagEmty = 1
         Else
            .Row = countDynamicRow + 46
            .Col = .ColLetterToNumber("R")
            .CellNote = ""
            .BackColor = mNonErrorColor
         End If
         'kiem tra ngay ky
         If Trim(strNgayKy) = "" Or Trim(strNgayKy) = "../../...." Then
            .Row = countDynamicRow + 48
            .Col = .ColLetterToNumber("R")
            .CellNote = errEmty  'static
            .BackColor = mErrorColor
            iFlagEmty = 1
         Else
                strNgayKy = DateSerial(CInt(Mid$(strNgayKy, 7, 4)), CInt(Mid$(strNgayKy, 4, 2)), CInt(Mid$(strNgayKy, 1, 2)))
                If strNgayKy > Now Then
                    .Col = .ColLetterToNumber("R")
                    .Row = countDynamicRow + 48
                    .CellNote = errNgayKy  'static
                    .BackColor = mErrorColor
                    iFlagNgayKy = 1
                Else
                    .Row = countDynamicRow + 48
                    .Col = .ColLetterToNumber("R")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
         End If
         ' kiem tra ngay huy
        If Trim(strNgayHuy) = "" Or Trim(strNgayHuy) = "../../...." Then
            .Row = 8
            .Col = .ColLetterToNumber("O")
            .CellNote = errEmty  'static
            .BackColor = mErrorColor
            iFlagEmty = 1
         Else
                strNgayHuy = DateSerial(CInt(Mid$(strNgayHuy, 7, 4)), CInt(Mid$(strNgayHuy, 4, 2)), CInt(Mid$(strNgayHuy, 1, 2)))
                If strNgayHuy > Now Then
                    .Col = .ColLetterToNumber("O")
                    .Row = 8
                    .CellNote = errNgayHuy  'static
                    .BackColor = mErrorColor
                    iFlagNgayHuy = 1
                ElseIf strNgayHuy < DateSerial(2011, 1, 1) Then
                    .Col = .ColLetterToNumber("O")
                    .Row = 8
                    .CellNote = errNgayHuyHL  'static
                    .BackColor = mErrorColor
                    iFlagNgayHuyHL = 1
                ElseIf strNgayHuy > strNgayKy Then
                    .Col = .ColLetterToNumber("O")
                    .Row = 8
                    .CellNote = errNgayKyHL  'static
                    .BackColor = mErrorColor
                    iFlagNgayKyHL = 1
                Else
                    .Row = 8
                    .Col = .ColLetterToNumber("O")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
         End If
                
        ' set lai trang thai cua HDR
          .Sheet = .SheetCount
         .SetText .ColLetterToNumber("B"), 20, IIf(iFlagEmty = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 14, IIf(iFlagMauSoHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 16, IIf(iFlagSoLienHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 15, IIf(iFlagKyHieu = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 18, IIf(iFlagLengHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 19, IIf(iFlagSoLuong = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 21, IIf(iFlagNgayHuy = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 22, IIf(iFlagNgayKy = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 23, IIf(iFlagKhoangTSDSTrung = 1, "0", "1")
         
         .SetText .ColLetterToNumber("B"), 27, IIf(iFlagNgayHuyHL = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 28, IIf(iFlagNgayKyHL = 1, "0", "1")
         
         .SetText .ColLetterToNumber("B"), 29, IIf(iFlagSoLienHD = 2, "0", "1")
         
        ' CheckErrorMST
        '.ReDraw = True
        '.Visible = True
        .EventEnabled(EventAllEvents) = True
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
        .GetText .ColLetterToNumber("E"), 12, vError1
        .GetText .ColLetterToNumber("E"), 13, vError2
        .GetText .ColLetterToNumber("E"), 14, vError3
        
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
            .Row = 12
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 12
            .Text = "1"
        End If
        
        If iFlagTaxCode2 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 13
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 13
            .Text = "1"
        End If
        
        
        If iFlagTaxCode3 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 14
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 14
            .Text = "1"
        End If
        
        .Sheet = 1
        .Col = .ColLetterToNumber("F")
        .Row = 5
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

Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean
    Dim totalCell  As Long
    Dim countRowReset As Long
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    totalCell = TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length
    countRowReset = (totalCell - 8) / 9
    With fps
        For Each xmlNodeReset In TAX_Utilities_v2.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            .Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            If (lRow < 43 Or lRow > 43 + countRowReset) And mCurrentSheet = 1 Then
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
                TAX_Utilities_v2.AdjustData(mCurrentSheet - 1) = True
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
    Dim arrDate() As String
    
    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet
        If (GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "Active") <> "0") Then
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
                If lCol = .ColLetterToNumber("O") And lRow = 8 Then
                     If varTemp <> "" Then
                         TAX_Utilities_v2.FirstDay = varTemp
                     End If
                End If
             Next
        End If
      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing

End Sub

Public Function GetNguoiKy() As String
    Dim xmlDomHeader As New MSXML.DOMDocument
    
    xmlDomHeader.Load GetAbsolutePath(TAX_Utilities_v2.DataFolder & "Header_01.xml")
    GetNguoiKy = GetAttribute(xmlDomHeader.nodeFromID("C_26"), "Value")
    
    Set xmlDomHeader = Nothing
End Function





' ham set thong tin sau khi tai XML
Public Sub finishImportXML()
    Dim maHD    As Variant
    Dim tenHD  As String
    Dim arrTemp() As String
    Dim arrTemp1() As String
    Dim arrNgay() As String
    
    Dim strGio, strPhut, strNgay As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
          .Sheet = 1
          mCurrentSheet = .Sheet
          ' lay ve gio phut ngay
          .Row = 9
          .GetText .ColLetterToNumber("O"), .Row, maHD
          If maHD <> "" Then
              arrTemp1 = Split(maHD, "T")
              If UBound(arrTemp1) > 0 Then
                    ' set ngay
                    .Row = 8
                    .Col = .ColLetterToNumber("O")
                    arrNgay = Split(arrTemp1(0), "-")
                    If UBound(arrNgay) > 1 Then
                        .Text = arrNgay(2) & "/" & arrNgay(1) & "/" & arrNgay(0)
                        UpdateCell fps, .Col, .Row, .Text
                    End If
                    
                    .Col = .ColLetterToNumber("L")
                    .Text = Val(Left$(arrTemp1(1), 2))
                    UpdateCell fps, .Col, .Row, .Text
                    
                    .Col = .ColLetterToNumber("N")
                    .Text = Val(Mid$(arrTemp1(1), 4, 2))
                    UpdateCell fps, .Col, .Row, .Text
              End If
          End If
          
          '
            .Row = 43
            Do
                .GetText .ColLetterToNumber("O"), .Row, maHD
                 GetMaHoaDon maHD, tenHD
                 If tenHD <> "" Then
                    arrTemp = Split(tenHD, "###")
                    If UBound(arrTemp) > 0 Then
                        .Col = .ColLetterToNumber("C")
                        .Text = arrTemp(0)
                        UpdateCell fps, .Col, .Row, .Text
                        .Col = .ColLetterToNumber("AI")
                        .Text = CStr(arrTemp(1))
                        UpdateCell fps, .Col, .Row, .Text
                        
                        .Col = .ColLetterToNumber("AJ")
                        .Text = Left$(CStr(arrTemp(1)), 1)
                         UpdateCell fps, .Col, .Row, .Text
                    End If
                 End If
                 
                 
                 .Col = .ColLetterToNumber("B")
                 .Row = .Row + 1
            Loop Until .Text = "aa"
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
