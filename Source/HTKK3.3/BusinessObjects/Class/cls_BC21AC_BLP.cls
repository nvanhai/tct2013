VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_BC21AC_BLP"
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
Dim strQuanHeNNT As String
Dim strCombo As String

Dim strCboGhiChu As String
Dim strCboValueGhiChu As String

Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

    strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_loai_BLP.xml")
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
        .Sheet = 1
        SetDateFormat fps, 1, 8, .ColLetterToNumber("O"), DDMMYYYY
        .SetText .ColLetterToNumber("O"), 8, Now
'        .SetText .ColLetterToNumber("K"), 8, Hour(Time)
'        .SetText .ColLetterToNumber("M"), 8, Minute(Time)
        ' set ngay mat, chay, hong hoa don
        .Col = .ColLetterToNumber("O")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        'set Mau hoa don
        .Row = 43
        .Col = .ColLetterToNumber("C")
        .TypeComboBoxList = strComboHien
        .Text = strQuanHeNNT
        .Col = .ColLetterToNumber("AI")
        .TypeComboBoxList = strCombo
        
        .Row = 51
        ' Set format ngay ky
        .Col = .ColLetterToNumber("AB")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        
        ' Set gia tri cho combo an
        .Col = .ColLetterToNumber("AI")
        .Row = 43
        .TypeComboBoxList = strCombo
        
        ' Set gia tri cho truong ghi chu
        strCboGhiChu = GetAttribute(GetMessageCellById("0297"), "Msg") + Chr$(9) + GetAttribute(GetMessageCellById("0298"), "Msg")
        .Row = 43
        .Col = .ColLetterToNumber("AC")
        .TypeComboBoxList = strCboGhiChu
        
        strCboValueGhiChu = "02" + Chr$(9) + "04"
        .Row = 43
        .Col = .ColLetterToNumber("AL")
        .TypeComboBoxList = strCboValueGhiChu
        
        'dntai  10/05/2011   them dinh dang vao cot mau so
        'SetKyHieuHDFormat fps, 1, 43, .ColLetterToNumber("s")
        
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


Dim totalRow As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        i = 1
        .Col = .ColLetterToNumber("B")
        .Row = 43
        Do
'             .Col = .ColLetterToNumber("E")
'             .Text = "0"
             
             .Col = .ColLetterToNumber("B")
             .Row = i + 43
             totalRow = i
             i = i + 1
        Loop Until .Text = "aa"
        'dhdang sua theo yc chi en ngay 25/01
        .Row = totalRow + 48
        .Col = .ColLetterToNumber("AB")
        If Trim(.Text) = "" Then
            .SetText .ColLetterToNumber("AB"), totalRow + 48, GetNguoiKy()
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
                .SetText .ColLetterToNumber("O"), .Row, Trim(Mid(strId, 11, 5))
                UpdateCell_sheet fps, 1, .ColLetterToNumber("O"), .Row, .Text
                .Col = .ColLetterToNumber("O")
                '.TypeMaxEditLen = 13
                .Col = .ColLetterToNumber("E")
                .SetText .ColLetterToNumber("E"), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 1, .ColLetterToNumber("E"), .Row, .Text
'            ElseIf Trim(Left(strId, 1)) = "3" Then
'                .SetText .ColLetterToNumber("O"), .Row, ""
'                UpdateCell_sheet fps, 1, .ColLetterToNumber("O"), .Row, .Text
'                .Col = .ColLetterToNumber("O")
'                .TypeMaxEditLen = 20
'                .Col = .ColLetterToNumber("E")
'                .SetText .ColLetterToNumber("E"), .Row, ""
'                UpdateCell_sheet fps, 1, .ColLetterToNumber("E"), .Row, .Text
'            ElseIf Trim(Left(strId, 1)) = "4" Then
'                .SetText .ColLetterToNumber("O"), .Row, Trim(Mid(strId, 11, 6))
'                UpdateCell_sheet fps, 1, .ColLetterToNumber("O"), .Row, .Text
'                .Col = .ColLetterToNumber("O")
'                .TypeMaxEditLen = 13
'                .Col = .ColLetterToNumber("E")
'                .SetText .ColLetterToNumber("E"), .Row, Trim(Mid(strId, 21))
'                UpdateCell_sheet fps, 1, .ColLetterToNumber("E"), .Row, .Text
'
'
'            Else
'                .SetText .ColLetterToNumber("O"), .Row, Trim(Mid(strId, 11, 3))
'                UpdateCell_sheet fps, 1, .ColLetterToNumber("O"), .Row, .Text
'                .Col = .ColLetterToNumber("O")
'                .TypeMaxEditLen = 20
'                .Col = .ColLetterToNumber("E")
'                .SetText .ColLetterToNumber("E"), .Row, Trim(Mid(strId, 21))
'                UpdateCell_sheet fps, 1, .ColLetterToNumber("E"), .Row, .Text
            End If
            
            ' Set loai HD
            .SetText .ColLetterToNumber("AK"), .Row, Trim(Mid(strId, 1, 1))
            UpdateCell fps, .ColLetterToNumber("AK"), .Row, .Text
            
            .EventEnabled(EventChange) = True
         End If
         
         ' set gia tri ma cua truong ghi chu
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("AC") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("AL")
            .typeComboboxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber("AL")
            UpdateCell_sheet fps, 1, .Col, .Row, .Text
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
    
    Dim maLoaiBLP As Variant
    
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
        ' check ngay ky TBAC
            Dim totalRowD As Variant
            Dim m As Variant
            m = 1
            .Col = .ColLetterToNumber("B")
            .Row = 43
            Do
                 .Col = .ColLetterToNumber("B")
                 .Row = m + 43
                 totalRowD = m
                 m = m + 1
            Loop Until .Text = "aa"
            If Col = .ColLetterToNumber("AB") And Row = totalRowD + 50 Then
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
                ElseIf tuSo = "" And denso = "" Then
                .SetText .ColLetterToNumber("AA"), Row, ""
                .Col = Col
                .Row = Row
                UpdateCell fps, .Col, .Row, .Text
                End If
                
                
            End If
            
          ' Chuyen ky hieu HD thanh chu cai in hoa
            If Col = .ColLetterToNumber("S") Then
                 .GetText Col, Row, varTemp
                 .SetText Col, Row, FormatKyHieuBLP(CStr(varTemp))
                 .Col = Col
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
            End If
           ' chuyen Mau so thanh chu hoa
            If Col = .ColLetterToNumber("O") Then
                 .GetText Col, Row, varTemp
                 '.SetText Col, Row, FormatKyHieuBLP(CStr(varTemp))
                 .SetText Col, Row, UCase(varTemp)
                 .Col = Col
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
            End If
           
            
           'Kiem tra cau truc cua ky hieu neu sai set cot an =1 dung =0
            If Col = .ColLetterToNumber("S") Then
                 .GetText Col, Row, varTemp
                 .GetText .ColLetterToNumber("C"), Row, maLoaiBLP
                 .SetText .ColLetterToNumber("AH"), Row, CheckSoBLP(CStr(varTemp), maLoaiBLP)
                 .Col = .ColLetterToNumber("AH")
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
            End If
           'Kiem tra cau truc cua Mau so
           If Col = .ColLetterToNumber("O") Then
                .GetText .ColLetterToNumber("AI"), .Row, strId
                strLoai = Trim(Left(strId, 1))
                If strLoai = "0" Then
                    strTemp = Trim(Mid(strId, 11, 5))
'                ElseIf strLoai = "4" Then
'                    strTemp = Trim(Mid(strId, 11, 6))
'                Else
'                    strTemp = Trim(Mid(strId, 11, 3))
                End If
                .GetText Col, Row, varTemp
                 .SetText .ColLetterToNumber("AJ"), Row, CheckMauSoBLP(CStr(varTemp), CStr(strLoai), CStr(strTemp))
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
    Dim MST1 As Variant, MST2 As Variant, MST3 As Variant, MST4 As Variant
    Dim MST5 As Variant, MST6 As Variant, MST7 As Variant, MST8 As Variant
    Dim MST9 As Variant, MST10 As Variant, MST11 As Variant, MST12 As Variant, MST13 As Variant
    Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer, iFlagSoLuong As Integer
    Dim strCheck As String
    Dim errMauSoHD As Variant, errKyHieuHD As Variant, errSoLuong As Variant, errSoLienHD As Variant, vErrNgayKy As Variant, vErrNgayPhatHien As Variant, vErrNgayPhatHienHL As Variant, vErrNgayKyHL As Variant
    Dim errKhoangGN As Variant, errLengthHD As Variant, errEmty As Variant, errGhiChu As Variant
    Dim strTenHD As Variant, strMauSoHD As Variant, strKyHieuHD As Variant, strTuSo As Variant, strDenso As Variant
    Dim strSoLuong As Variant
    Dim iFlagTenHD As Integer, iFlagMauSoHD As Integer, iFlagKyHieu As Integer, iFlagSoLienHD As Integer, iFlagEmty As Integer, iFlagKhoangTSDSTrung As Integer, iFlagNgayKy As Integer, iFlagNgayPhatHien As Integer, iFlagGhiChu As Integer
    Dim iFlagNgayPhatHienHL As Integer, iFlagNgayKyHL As Integer
    
    
    
    Dim iWhiteRow As Integer
    Dim strId As Variant, strLoai As String, strTemp As String, result As String
    Dim i, j As Integer
    
    Dim Tuso1 As Variant
    Dim Tuso2 As Variant
    Dim denso1 As Variant
    Dim denso2 As Variant
    Dim Tenloaihd1 As Variant
    Dim Tenloaihd2 As Variant
    Dim mauso1 As Variant
    Dim mauso2 As Variant
    Dim kyhieu1 As Variant
    Dim kyhieu2 As Variant
    Dim SohieuHD As Variant
    Dim ghiChu As Variant
    
    Dim errLyDoEmty, errDonViBCEmty As Variant
    
    With fps
        '.Visible = False
        '.ReDraw = False
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 14, errMauSoHD
        .GetText .ColLetterToNumber("E"), 15, errKyHieuHD
        .GetText .ColLetterToNumber("E"), 16, errSoLienHD
        .GetText .ColLetterToNumber("E"), 17, errKhoangGN
        .GetText .ColLetterToNumber("E"), 18, errLengthHD
        .GetText .ColLetterToNumber("E"), 19, errSoLuong
        .GetText .ColLetterToNumber("E"), 20, errEmty
        .GetText .ColLetterToNumber("E"), 22, vErrNgayKy
        .GetText .ColLetterToNumber("E"), 23, vErrNgayPhatHien
        .GetText .ColLetterToNumber("E"), 24, errGhiChu
        .GetText .ColLetterToNumber("E"), 25, vErrNgayPhatHienHL
        .GetText .ColLetterToNumber("E"), 26, vErrNgayKyHL
        
        .GetText .ColLetterToNumber("E"), 27, errLyDoEmty
        .GetText .ColLetterToNumber("E"), 28, errDonViBCEmty
        
        'set flag bang 0
        iFlagTenHD = 0
        iFlagMauSoHD = 0
        iFlagKyHieu = 0
        iFlagSoLienHD = 0
        iFlagEmty = 0
        iFlagSoLuong = 0
        iFlagKhoangTSDSTrung = 0
        iFlagNgayKy = 0
        iFlagNgayPhatHien = 0
        iFlagGhiChu = 0
        
        .Sheet = 1
        .Col = .ColLetterToNumber("B")
        .Row = 43
        Do
            .Col = .ColLetterToNumber("O")
            .CellNote = ""
            .BackColor = mNonErrorColor
            
            .Col = .ColLetterToNumber("S")
            .CellNote = ""
            .BackColor = mNonErrorColor
            
            .Col = .ColLetterToNumber("V")
            .CellNote = ""
            .BackColor = mNonErrorColor

            .Col = .ColLetterToNumber("Y")
            .CellNote = ""
            .BackColor = mNonErrorColor
            
            .Col = .ColLetterToNumber("AC")
            .CellNote = ""
            .BackColor = mNonErrorColor
            
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
            ' Check cho den het cac dong co du lieu thi thoi
        Loop Until UCase(.Text) = "AA"
        
        ' Kiem tra tren sheet 1
        .Sheet = 1
        .Col = .ColLetterToNumber("B")
        .Row = 43
        i = 1
        Do
             ' Lay cac gia tri cua 1 row
             iWhiteRow = 1
            .GetText .ColLetterToNumber("E"), .Row, strTenHD
            .GetText .ColLetterToNumber("E"), .Row, Tenloaihd1
            .GetText .ColLetterToNumber("O"), .Row, strMauSoHD
            .GetText .ColLetterToNumber("S"), .Row, strKyHieuHD
            .GetText .ColLetterToNumber("O"), .Row, mauso1
            .GetText .ColLetterToNumber("S"), .Row, kyhieu1
            .GetText .ColLetterToNumber("V"), .Row, strTuSo
            .GetText .ColLetterToNumber("Y"), .Row, strDenso
            .GetText .ColLetterToNumber("V"), .Row, Tuso1
            .GetText .ColLetterToNumber("Y"), .Row, denso1
            .GetText .ColLetterToNumber("AA"), .Row, strSoLuong
            .GetText .ColLetterToNumber("AI"), .Row, strId
            .GetText .ColLetterToNumber("AC"), .Row, ghiChu
            
            If Trim(strTenHD) <> "" Or Trim(strMauSoHD) <> "" Or Trim(strKyHieuHD) <> "" Or Trim(strId) <> "" Or Trim(strId) <> "" Or Trim(strTuSo) <> "" Or Trim(strDenso) <> "" Then
                ' Khong phai dong trang
                iWhiteRow = 0
            End If
            
            ' check ghi chu khong duoc de trang
            If Trim(CStr(ghiChu)) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("AC")
                .CellNote = errGhiChu  'static
                .BackColor = mErrorColor
                iFlagGhiChu = 1
            End If
            'tu so den so bat buoc nhap
            If Trim(Tuso1) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("V")
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                'blCheck_S3 = False
                iFlagEmty = 1
'            Else
'                .Col = .ColLetterToNumber("V")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
            End If
            
            If Trim(denso1) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("Y")
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                'blCheck_S3 = False
                iFlagEmty = 1
'            Else
'                .Col = .ColLetterToNumber("Y")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
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
                .Col = .ColLetterToNumber("O")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            strLoai = Trim(Left(strId, 1))
            If strLoai = "0" Then
                strTemp = Trim(Mid(strId, 11, 5))
'            ElseIf strLoai = "4" Then
'                strTemp = Trim(Mid(strId, 11, 6))
'            Else
'                strTemp = Trim(Mid(strId, 11, 3))
            End If
            result = CheckMauSoBLP(CStr(strMauSoHD), strLoai, strTemp)
            If result = "1" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("O")
                .CellNote = errMauSoHD  'static
                .BackColor = mErrorColor
                iFlagMauSoHD = 1
            ElseIf result = "2" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("O")
                .CellNote = errSoLienHD  'static
                .BackColor = mErrorColor
                iFlagSoLienHD = 1
            ElseIf Trim(strMauSoHD) <> "" Then
                .Col = .ColLetterToNumber("O")
                .CellNote = ""
                .BackColor = mNonErrorColor
            
            End If
            Tuso1 = Val(Tuso1)
            denso1 = Val(denso1)
            'kiem tra tu so <= den so
            If (Tuso1 > denso1) And Trim(strTuSo) <> "" And Trim(strDenso) <> "" Then
                    .Col = .ColLetterToNumber("V")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = errSoLuong    'static
                    .BackColor = mErrorColor
                    .Col = .ColLetterToNumber("Y")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = errSoLuong    'static
                    .BackColor = mErrorColor
                    iFlagSoLuong = 1
'            Else
'                    .Col = .ColLetterToNumber("V")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'
'                    .Col = .ColLetterToNumber("Y")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
            End If
            
            'Kiem tra cau truc cua ky hieu
            Dim varTemp As Variant
            Dim maLoaiBLP As Variant
            .GetText .ColLetterToNumber("S"), .Row, varTemp
            .GetText .ColLetterToNumber("C"), .Row, maLoaiBLP
            SohieuHD = CheckSoBLP(CStr(varTemp), maLoaiBLP)
            

            If (Trim(SohieuHD) = "1" Or Trim(SohieuHD) = "2" Or Trim(varTemp) = vbNullString) And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("S")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = errKyHieuHD    'static
                .BackColor = mErrorColor
                iFlagKyHieu = 1
                'capture the first error cell
            Else
                .Col = .ColLetterToNumber("S")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            
            ' Kiem tra trung nhau tu so den so
            'bat dau
            Do
                j = .Row + 1
                .GetText .ColLetterToNumber("V"), j, Tuso2
                .GetText .ColLetterToNumber("Y"), j, denso2
                Tuso2 = Val(Tuso2)
                denso2 = Val(denso2)
                .GetText .ColLetterToNumber("E"), j, Tenloaihd2
                .GetText .ColLetterToNumber("O"), j, mauso2
                .GetText .ColLetterToNumber("S"), j, kyhieu2
                'Dim vMST2 As Variant
                '.GetText .ColLetterToNumber("E"), .Row, vMST2

                    If (((Tuso1 - Tuso2) * (denso1 - Tuso2)) <= 0 Or ((Tuso1 - denso2) * (denso1 - denso2)) <= 0 Or ((Tuso1 > Tuso2) And (denso1 < denso2))) And ((Tenloaihd1 = Tenloaihd2) And (mauso1 = mauso2) And (kyhieu1 = kyhieu2)) Then
                        .Row = i + 42
                        .Col = .ColLetterToNumber("V")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errKhoangGN     'static
                        .BackColor = mErrorColor


                        .Row = i + 42
                        .Col = .ColLetterToNumber("Y")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errKhoangGN     'static
                        .BackColor = mErrorColor


                        .Row = j
                        .Col = .ColLetterToNumber("V")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errKhoangGN     'static
                        .BackColor = mErrorColor


                        .Row = j
                        .Col = .ColLetterToNumber("Y")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errKhoangGN     'static
                        .BackColor = mErrorColor
                        
                        iFlagKhoangTSDSTrung = 1
                     End If
                .Row = j
                .Col = .ColLetterToNumber("B")
                ' Check cho den het cac dong co du lieu thi thoi
            Loop Until UCase(.Text) = "AA"

            '--------------------------------------------------
        'check ngay ky , nguoi ky
        'dhdang sua
        Dim ngayky As Variant
        Dim ngayphathien As Variant
        Dim nguoiky As Variant
        Dim totalRowD As Variant
        Dim m As Variant
        
        Dim lyDoMatChayHong As Variant
        Dim donViBC As Variant
        Dim iFlagLyDoEmpty As Integer
        Dim iFlagDonViBCEmpty As Integer
        
        m = 1
        .Col = .ColLetterToNumber("B")
        .Row = 43
        Do
             .Col = .ColLetterToNumber("B")
             .Row = m + 43
             totalRowD = m
             m = m + 1
        Loop Until .Text = "aa"
        
        
        .GetText .ColLetterToNumber("AB"), totalRowD + 48, nguoiky
        .GetText .ColLetterToNumber("AB"), totalRowD + 50, ngayky
        .GetText .ColLetterToNumber("O"), 8, ngayphathien
        
        .GetText .ColLetterToNumber("I"), totalRowD + 45, lyDoMatChayHong
        .GetText .ColLetterToNumber("I"), totalRowD + 47, donViBC
        
        If Trim(nguoiky) = "" Then
                .Col = .ColLetterToNumber("AB")
                .Row = totalRowD + 48
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                'blCheck_S3 = False
                iFlagEmty = 1
        Else
                .Row = totalRowD + 48
                .Col = .ColLetterToNumber("AB")
                .CellNote = ""
                .BackColor = mNonErrorColor
        End If
        
        If Trim(ngayky) = "" Then
                .Col = .ColLetterToNumber("AB")
                .Row = totalRowD + 50
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                'blCheck_S3 = False
                iFlagEmty = 1
        Else
                ngayky = DateSerial(CInt(Mid$(ngayky, 7, 4)), CInt(Mid$(ngayky, 4, 2)), CInt(Mid$(ngayky, 1, 2)))
                If ngayky > Now Then
                    .Col = .ColLetterToNumber("AB")
                    .Row = totalRowD + 50
                    .CellNote = vErrNgayKy  'static
                    .BackColor = mErrorColor
                    iFlagNgayKy = 1
                    'iFlagEmty = 1
                ElseIf ngayky < DateSerial(2011, 1, 1) Then
                    .Col = .ColLetterToNumber("AB")
                    .Row = totalRowD + 50
                    .CellNote = vErrNgayKyHL   'static
                    .BackColor = mErrorColor
                    iFlagNgayKyHL = 1
                Else
                    .Row = totalRowD + 50
                    .Col = .ColLetterToNumber("AB")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
        End If
        
        'check ngay phat hien
        If Trim(ngayphathien) = "" Then
                .Col = .ColLetterToNumber("O")
                .Row = 8
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                'blCheck_S3 = False
                iFlagEmty = 1
        Else
                ngayphathien = DateSerial(CInt(Mid$(ngayphathien, 7, 4)), CInt(Mid$(ngayphathien, 4, 2)), CInt(Mid$(ngayphathien, 1, 2)))
                If ngayphathien > ngayky Then
                    .Col = .ColLetterToNumber("O")
                    .Row = 8
                    .CellNote = vErrNgayPhatHien  'static
                    .BackColor = mErrorColor
                    iFlagNgayPhatHien = 1
                    'iFlagEmty = 1
                ElseIf ngayphathien < DateSerial(2011, 1, 1) Then
                    .Col = .ColLetterToNumber("O")
                    .Row = 8
                    .CellNote = vErrNgayPhatHienHL  'static
                    .BackColor = mErrorColor
                    iFlagNgayPhatHienHL = 1
                Else
                    .Row = 8
                    .Col = .ColLetterToNumber("O")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
        End If
        
        ' check ly do mat chay hong
        If Trim(lyDoMatChayHong) = "" Then
                .Col = .ColLetterToNumber("I")
                .Row = totalRowD + 45
                .CellNote = errLyDoEmty  'static
                .BackColor = mErrorColor
                'blCheck_S3 = False
                iFlagLyDoEmpty = 1
        Else
                .Row = totalRowD + 45
                .Col = .ColLetterToNumber("I")
                .CellNote = ""
                .BackColor = mNonErrorColor
        End If
        
        ' check don vi BC
        If Trim(donViBC) = "" Then
                .Col = .ColLetterToNumber("I")
                .Row = totalRowD + 47
                .CellNote = errDonViBCEmty  'static
                .BackColor = mErrorColor
                'blCheck_S3 = False
                iFlagDonViBCEmpty = 1
        Else
                .Row = totalRowD + 47
                .Col = .ColLetterToNumber("I")
                .CellNote = ""
                .BackColor = mNonErrorColor
        End If
'end
'------------------------------------------------------------------------
         
             .Col = .ColLetterToNumber("B")
             .Row = i + 43
             i = i + 1
        Loop Until .Text = "aa"
        ' set lai trang thai cua HDR
          .Sheet = .SheetCount
         .SetText .ColLetterToNumber("B"), 20, IIf(iFlagEmty = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 14, IIf(iFlagMauSoHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 16, IIf(iFlagSoLienHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 19, IIf(iFlagSoLuong = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 15, IIf(iFlagKyHieu = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 17, IIf(iFlagKhoangTSDSTrung = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 22, IIf(iFlagNgayKy = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 23, IIf(iFlagNgayPhatHien = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 24, IIf(iFlagGhiChu = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 25, IIf(iFlagNgayPhatHienHL = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 26, IIf(iFlagNgayKyHL = 1, "0", "1")
         
         .SetText .ColLetterToNumber("B"), 27, IIf(iFlagLyDoEmpty = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 28, IIf(iFlagDonViBCEmpty = 1, "0", "1")
        'CheckErrorMSTs
        '.ReDraw = True
        '.Visible = True
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
    countRowReset = (totalCell - 6) / 12
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
    Dim maBLP    As Variant
    Dim tenBLP  As String
    Dim arrTemp() As String
    
    With fps
        .EventEnabled(EventAllEvents) = False
          .Sheet = 1
          mCurrentSheet = .Sheet
          '
            .Row = 43
            Do
                .GetText .ColLetterToNumber("O"), .Row, maBLP
                 GetMaBLP maBLP, tenBLP
                 If tenBLP <> "" Then
                    arrTemp = Split(tenBLP, "###")
                    If UBound(arrTemp) > 0 Then
                        .Col = .ColLetterToNumber("C")
                        .Text = arrTemp(0)
                        UpdateCell fps, .Col, .Row, .Text
                        .Col = .ColLetterToNumber("AI")
                        .Text = CStr(arrTemp(1))
                        UpdateCell fps, .Col, .Row, .Text
                        
                        .Col = .ColLetterToNumber("AK")
                        .Text = Left$(CStr(arrTemp(1)), 1)
                         UpdateCell fps, .Col, .Row, .Text
                    End If
                 End If
                 ' set lai ma cho truong ghi chu
                 .GetText .ColLetterToNumber("AC"), .Row, maBLP
                 If maBLP = GetAttribute(GetMessageCellById("0297"), "Msg") Then
                    .Col = .ColLetterToNumber("AL")
                    .Text = "02"
                    UpdateCell fps, .Col, .Row, .Text
                 ElseIf maBLP = GetAttribute(GetMessageCellById("0298"), "Msg") Then
                    .Col = .ColLetterToNumber("AL")
                    .Text = "04"
                    UpdateCell fps, .Col, .Row, .Text
                 End If
                 
                 .Col = .ColLetterToNumber("B")
                 .Row = .Row + 1
            Loop Until .Text = "aa"
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
