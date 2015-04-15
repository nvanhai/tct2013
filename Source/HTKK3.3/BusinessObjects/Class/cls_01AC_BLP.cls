VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01AC_BLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "To khai quyet toan thu nhap ca nhan mau 04-TNCN" interface sheets
'this Class is belong to TAX_Business_v2 project which will be compline to DLL

Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrSolanBosung As String
Public flgLoadToKhai As Boolean
Private totalRow As Long
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    With fps

        Dim xmlDocument As New MSXML.DOMDocument
        Dim xmlNode As MSXML.IXMLDOMNode
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_loai_BLP.xml"))
        Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
        Dim strQuanHeNNTID As String
        Dim strQuanHeNNT As String
        Dim strDataFileName As String
        Dim arrDanhsach() As String
        Dim strComboHien As String
        Dim strCombo As String
        Dim strCombo1 As String
        Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
        Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
       .Sheet = 1
       strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_loai_BLP.xml")
        ' Lay danh muc loai hoa don
        ' 18/02/2014
        If xmlDomData.Load(strDataFileName) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    If arrDanhsach(0) <> "3" Then
                        strCombo = strCombo + CPab(arrDanhsach(0), 10) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(2), 200) + Chr$(9)
    '                    strCombo = strCombo + CPab(arrDanhsach(1), 10) + Chr$(9)
                        strComboHien = strComboHien + arrDanhsach(1) + Chr$(9)
                    End If
                End If
            Next
        End If
       .Sheet = 1
       
'        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
'        For Each xmlNode In xmlNodeListMap
'            strQuanHeNNT = strQuanHeNNT + GetAttribute(xmlNode, "Value") + Chr$(9)
'        Next
'        strQuanHeNNTID = "01GTKT" + Chr$(9) + "02GTTT" + Chr$(9) + "06HDXK" + Chr$(9) + "07KPTQ" + Chr$(9) + "03XKNB" + Chr$(9) + "04HGDL" + Chr$(9) + "01/" + Chr$(9) + "02/" + Chr$(9) + "01/" + Chr$(9) + "02/" + Chr$(9) + "01/" + Chr$(9) + "02/" + Chr$(9)
        .Row = 22
        .Col = .ColLetterToNumber("H")
        
        .TypeComboBoxList = strComboHien
        .Text = strQuanHeNNT
        .Col = .ColLetterToNumber("N")
        .TypeComboBoxList = strCombo
      
        SetDateFormat fps, 1, 22, .ColLetterToNumber("G"), DDMMYYYY
         'SetKyHieuHDFormat fps, 1, 22, .ColLetterToNumber("K")
        .Row = 32
        ' Set format ngay ky
        SetDateFormat fps, 1, 32, .ColLetterToNumber("J"), DDMMYYYY
        .Col = .ColLetterToNumber("J")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        ' Set dinh dang ky hieu hoa don
  '      SetKyHieuHDFormat fps, 1, 22, .ColLetterToNumber("K")
        
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
    
    
    With fps
        .Sheet = 1
        .EventEnabled(EventAllEvents) = False
        i = 1
        .Col = .ColLetterToNumber("B")
        .Row = 22
        Do
'             .Col = .ColLetterToNumber("C")
'             .Text = "0"
             
             .Col = .ColLetterToNumber("B")
             .Row = i + 22
             totalRow = i
             i = i + 1
        Loop Until .Text = "aa"
        'dhdang sua theo yc chi en ngay 25/01
        .Row = totalRow + 29
        .Col = .ColLetterToNumber("J")
        If Trim(.Text) = "" Then
            .SetText .ColLetterToNumber("J"), totalRow + 29, GetNguoiKy()
        End If
        
        
        
        
         'Xoa cong thuc cua o Chu ky
        Set xmlCellNode = TAX_Utilities_v2.Data(0).nodeFromID(GetCellID(fps, .ColLetterToNumber("D"), 22))
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode.lastChild
        Set xmlCellNode = xmlCellNode.firstChild
        Set xmlCellNode = xmlCellNode.firstChild
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        .Row = iRow
        .Col = iCol
        .Formula = ""
        
        
        ' set lai gia tri tu ngay den ngay
        .Sheet = 1
        .Row = 17
        .Col = .ColLetterToNumber("D")
        .Formula = "TRIM(Header!T6)"
        
        .Row = 17
        .Col = .ColLetterToNumber("E")
        .Formula = "TRIM(Header!T7)"
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Function GetValidatedCellIndex(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long) As Integer
    
    
End Function

Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    fps.EventEnabled(EventAllEvents) = False
    If Col = fps.ColLetterToNumber("B") And Row = 15 Then
            If fps.CellType = CellTypeCheckBox Then
                fps.Row = Row
                fps.Col = Col
                fps.Sheet = 1
                UpdateCell_sheet fps, fps.Sheet, fps.Col, fps.Row, IIf(fps.value = 1, "x", vbNullString) 'fps.Value
            End If
    End If
    fps.EventEnabled(EventAllEvents) = True
    CellChange Col, Row
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    Dim strId As Variant
    Dim intIndexCombo As Integer
    fps.EventEnabled(EventAllEvents) = False
    With fps
        .Sheet = 1
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("H") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("N")
            .typeComboboxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber("N")
            UpdateCell_sheet fps, 1, .Col, .Row, .Text
             
            
            .GetText .ColLetterToNumber("N"), .Row, strId
            
           If Trim(Left(strId, 1)) = "0" Then
                .SetText .ColLetterToNumber("I"), .Row, Trim(Mid(strId, 11, 5))
                UpdateCell_sheet fps, 1, .ColLetterToNumber("I"), .Row, .Text
                .Col = .ColLetterToNumber("I")
                '.TypeMaxEditLen = 13
                .Col = .ColLetterToNumber("J")
                .SetText .ColLetterToNumber("J"), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 1, .ColLetterToNumber("J"), .Row, .Text
'            ElseIf Trim(Left(strId, 1)) = "3" Then
'                .SetText .ColLetterToNumber("J"), .Row, ""
'                UpdateCell_sheet fps, 1, .ColLetterToNumber("J"), .Row, .Text
'                .Col = .ColLetterToNumber("J")
'                .TypeMaxEditLen = 20
'                .Col = .ColLetterToNumber("I")
'                .SetText .ColLetterToNumber("I"), .Row, ""
'                UpdateCell_sheet fps, 1, .ColLetterToNumber("I"), .Row, .Text
'            ElseIf Trim(Left(strId, 1)) = "4" Then
'                .SetText .ColLetterToNumber("J"), .Row, Trim(Mid(strId, 11, 6))
'                UpdateCell_sheet fps, 1, .ColLetterToNumber("J"), .Row, .Text
'                .Col = .ColLetterToNumber("J")
'                .TypeMaxEditLen = 13
'                .Col = .ColLetterToNumber("I")
'                .SetText .ColLetterToNumber("I"), .Row, Trim(Mid(strId, 21))
'                UpdateCell_sheet fps, 1, .ColLetterToNumber("I"), .Row, .Text
'            Else
'                .SetText .ColLetterToNumber("J"), .Row, Trim(Mid(strId, 11, 3))
'                UpdateCell_sheet fps, 1, .ColLetterToNumber("J"), .Row, .Text
'                .Col = .ColLetterToNumber("J")
'                .TypeMaxEditLen = 20
'                .Col = .ColLetterToNumber("I")
'                .SetText .ColLetterToNumber("I"), .Row, Trim(Mid(strId, 21))
'                UpdateCell_sheet fps, 1, .ColLetterToNumber("I"), .Row, .Text
            End If
            
'            .SetText .ColLetterToNumber("P"), .Row, Trim(Left(strId, 1))
'            UpdateCell fps, .ColLetterToNumber("P"), .Row, .Text
            
            .EventEnabled(EventChange) = True
         End If
    End With
   fps.EventEnabled(EventAllEvents) = True
End Sub

Public Sub kiemTraDuLieuImport()
   
End Sub


Private Sub fps_KeyPress(KeyAscii As Integer)
     With fps
        If .Col = .ColLetterToNumber("L") Or .Col = .ColLetterToNumber("M") Then
            If (KeyAscii < 48 And KeyAscii <> 3 And KeyAscii <> 22 And KeyAscii <> 13 And KeyAscii <> 8) Or KeyAscii > 57 Then
                KeyAscii = 0
            End If
        End If
    End With
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, iCol As Long, iRow As Long
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tCol As Long, tRow As Long
    With fps
        .EventEnabled(EventAllEvents) = False
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 1 Then
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 22
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 22
                         totalRow = i
                         i = i + 1
                    Loop Until .Text = "aa"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    
                    .SetActiveCell iCol, iRow
            End If
        End If
  
       .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim varMST As Variant, varMST5A As Variant, varMST5B As Variant
    Dim sumThuNhap5A As Variant, sumThuNhap5B As Variant
    Dim tuSo As Variant
    Dim denso As Variant
    Dim strId As Variant
    Dim strLoai As Variant
    Dim strTemp As Variant
    
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
           
            If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("G") And Row >= 22 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    'If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                              .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    'End If
                Else
                 .SetText Col, Row, ""
                End If
            End If
        ' check ngay ky TBAC
            Dim totalRowD As Variant
            Dim m As Variant
            m = 1
            .Col = .ColLetterToNumber("B")
            .Row = 22
            Do
                 .Col = .ColLetterToNumber("B")
                 .Row = m + 22
                 totalRowD = m
                 m = m + 1
            Loop Until .Text = "aa"
            If Col = .ColLetterToNumber("J") And (Row = totalRowD + 31) Then
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
         ' Chuyen ky hieu HD thanh chu cai in hoa
            If Col = .ColLetterToNumber("K") And (Row <> totalRowD + 31) And (Row <> totalRowD + 29) Then
                 .GetText Col, Row, varTemp
                 .SetText Col, Row, FormatKyHieuBLP(CStr(varTemp))
                 .Col = Col
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
            End If
           ' chuyen mau so thanh chu hoa
           If Col = .ColLetterToNumber("I") And (Row <> totalRowD + 31) And (Row <> totalRowD + 29) Then
                 .GetText Col, Row, varTemp
                 '.SetText Col, Row, FormatKyHieuBLP(CStr(varTemp))
                 .SetText Col, Row, UCase(varTemp)
                 .Col = Col
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
           End If
           
           
           ' Format dinh dang tu so den so co 7 ky tu
            If Col = .ColLetterToNumber("L") Or Col = .ColLetterToNumber("M") Then
                 .GetText Col, Row, varTemp
                 .SetText Col, Row, FormatSoHD(CStr(varTemp))
                 .Col = Col
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
           ' Ho tro tinh so luong
                .GetText .ColLetterToNumber("L"), Row, tuSo
                .GetText .ColLetterToNumber("M"), Row, denso
                If tuSo <> "" And denso <> "" Then
                    If Val(tuSo) = 0 And Val(denso) = 0 Then
                        .SetText .ColLetterToNumber("Q"), Row, 0
                    ElseIf Val(tuSo) = 0 And Val(denso) <> 0 Then
                        .SetText .ColLetterToNumber("Q"), Row, Val(denso) - Val(tuSo)
                    Else
                        .SetText .ColLetterToNumber("Q"), Row, Val(denso) - Val(tuSo) + 1
                    End If
                    
                    .Col = Col
                    .Row = Row
                    UpdateCell fps, .Col, .Row, .Text
                ElseIf tuSo = "" And denso = "" Then
                .SetText .ColLetterToNumber("Q"), Row, ""
                .Col = Col
                .Row = Row
                UpdateCell fps, .Col, .Row, .Text
                End If
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
    
End Sub

'Calling this function whenever we don't need this class anymore.
'No parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
    'If flgLoadToKhai = True Then Exit Sub
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        
        SetData
        ' 4/1/2011 dhdang sua
        UpdateSheets
        CheckDynamicError 'Set Exception Error on cells of interface
            
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
        
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    With fps
        'If flgLoadToKhai = True Then Exit Sub
        
        'CheckErrorMST
        ' Neu ton tai file data bang ke 05A thi kiem tra tren bang ke 5A
 '       If IIf(TAX_Utilities_v2.NodeValidity.childNodes(.SheetCount - 3).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            checkSheet1
'        Else
'            clearCheckSheet1
'        End If
    End With
End Sub

Private Sub checkSheet1()

   Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    Dim strCheck As String, iCurrentSheet As Integer
    
    Dim vMST As Variant, vErrMsgMST As Variant, vErrMsgMSTReq As Variant
    Dim vErrMsgMSTTrungCQCT As Variant, vErrMsgMSTTrungNhau As Variant, vErrMsgSoHD As Variant, vErrNgayHD As Variant, vErrNgayKy As Variant
    Dim vHoVaTen As Variant, vErrMsgHoVaTen As Variant
    Dim vTongThuNhap As Variant
    Dim vThuNhapGiamThue As Variant, vErrMsgThuNhapGiamThue As Variant
    Dim vSoThueDuocGiam As Variant, vErrMsgSoThueDuocGiam As Variant
    Dim vSoThueDaKhauTru As Variant, vErrMsgSoThueDaKhauTru As Variant, vErrMsgTusoDenSo As Variant
    Dim errMauSoHD As Variant
    Dim errSoLienHD As Variant
    Dim vLastRow As Variant
    Dim vMSTCQCT As Variant
    Dim vMST1 As Variant
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
     Dim iWhiteRow As Integer
    Dim strSoLuong As Variant
    Dim ngayHD As Variant
    Dim strId As Variant, strLoai As String, strTemp As String, result As String
    
    Dim iFlagvHoVaTen As Integer, iFlagvMst As Integer, iFocusFlag As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String
    Dim i, j As Integer
    Dim blCheck_S1, blCheck_S2, blCheck_S3, blCheck_S4, blCheck_S5, blCheck_S6, blCheck_S7, blCheck_S8, blCheck_S9 As Boolean
    
    Dim vTuNgayBC As Variant
    Dim errNgayBCHL As Variant
    Dim blCheck_NgayBCHL As Boolean
    
    Dim vSoHopDong As Variant
    
    Dim iWhiteRow1 As Variant
    
    With fps
        .EventEnabled(EventAllEvents) = False
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        ' Lay thong bao ve sai ma so thue trong sheet header
        .GetText .ColLetterToNumber("E"), 12, vErrMsgMST
        .GetText .ColLetterToNumber("E"), 13, vErrMsgMSTReq
        .GetText .ColLetterToNumber("E"), 14, vErrMsgMSTTrungNhau
        .GetText .ColLetterToNumber("E"), 15, vErrMsgSoHD
        .GetText .ColLetterToNumber("E"), 16, vErrMsgTusoDenSo
        .GetText .ColLetterToNumber("E"), 17, errMauSoHD
        .GetText .ColLetterToNumber("E"), 18, errSoLienHD
        .GetText .ColLetterToNumber("E"), 19, vErrNgayHD
        .GetText .ColLetterToNumber("E"), 20, vErrNgayKy
        .GetText .ColLetterToNumber("E"), 21, errNgayBCHL
        
        .GetText .ColLetterToNumber("T"), 6, vTuNgayBC

        
        
        ' Lay thong bao ve ma so thue bi trung voi co quan chi tra trong sheet header
        Dim ngaydauky As Variant
        Dim ngaycuoiky As Variant
        .GetText .ColLetterToNumber("T"), 6, ngaydauky
        .GetText .ColLetterToNumber("T"), 7, ngaycuoiky
        .Sheet = 1
        
        '.GetText .ColLetterToNumber("E"), 9, vMSTCQCT
        blCheck_S1 = True
        blCheck_S2 = True
        blCheck_S3 = True
        blCheck_S4 = True
        blCheck_S5 = True
        blCheck_S6 = True
        blCheck_S7 = True
        blCheck_S8 = True
        blCheck_S9 = True
        
        blCheck_NgayBCHL = True
        
        .Row = 22
        Do
            .Col = .ColLetterToNumber("L")
            .CellNote = ""
            .BackColor = mNonErrorColor
            
            .Col = .ColLetterToNumber("M")
            .CellNote = ""
            .BackColor = mNonErrorColor
            
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
            ' Check cho den het cac dong co du lieu thi thoi
        Loop Until UCase(.Text) = "AA"
        
        .Row = 22
        Do
            j = .Row
            
            .GetText .ColLetterToNumber("D"), .Row, vHoVaTen
            .GetText .ColLetterToNumber("C"), .Row, vMST
            .GetText .ColLetterToNumber("L"), .Row, Tuso1
            .GetText .ColLetterToNumber("M"), .Row, denso1
            .GetText .ColLetterToNumber("J"), .Row, Tenloaihd1
            .GetText .ColLetterToNumber("I"), .Row, mauso1
            .GetText .ColLetterToNumber("K"), .Row, kyhieu1
            
            .GetText .ColLetterToNumber("F"), .Row, vSoHopDong
            .GetText .ColLetterToNumber("G"), .Row, ngayHD
            
            iWhiteRow = 1
            iWhiteRow1 = 1
            
            .GetText .ColLetterToNumber("Q"), .Row, strSoLuong
            .GetText .ColLetterToNumber("N"), .Row, strId
            'check mau so hoa don
            If Trim(Tenloaihd1) <> "" Or Trim(mauso1) <> "" Or Trim(kyhieu1) <> "" Or Trim(Tuso1) <> "" Or Trim(denso1) <> "" Or Trim(strId) <> "" Or Trim(vMST) <> "" Then
                ' Khong phai dong trang
                iWhiteRow = 0
            End If
            
            ' check dong duoc phep in
            
             If Trim(Tenloaihd1) <> "" Or Trim(mauso1) <> "" Or Trim(kyhieu1) <> "" Or Trim(Tuso1) <> "" Or Trim(denso1) <> "" Then
                ' Khong phai dong trang
                iWhiteRow1 = 0
            End If
            
            'Ten HD
            If Trim(Tenloaihd1) = "" And iWhiteRow = 0 And iWhiteRow1 = 0 Then
                .Col = .ColLetterToNumber("J")
                .CellNote = vErrMsgMSTReq  'static
                .BackColor = mErrorColor
                blCheck_S3 = False
            Else
                .Col = .ColLetterToNumber("J")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            
            'tu so den so bat buoc nhap
            If Trim(Tuso1) = "" And iWhiteRow = 0 And iWhiteRow1 = 0 Then
                .Col = .ColLetterToNumber("L")
                .CellNote = vErrMsgMSTReq  'static
                .BackColor = mErrorColor
                blCheck_S3 = False
'                'iFlagEmty = 1
'            Else
'                .Col = .ColLetterToNumber("L")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
            End If
            
            If Trim(denso1) = "" And iWhiteRow = 0 And iWhiteRow1 = 0 Then
                .Col = .ColLetterToNumber("M")
                .CellNote = vErrMsgMSTReq  'static
                .BackColor = mErrorColor
                blCheck_S3 = False
                'iFlagEmty = 1
'            Else
'                .Col = .ColLetterToNumber("L")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
            End If
            '---------------------------
            ' Check so Hop dong
            If Trim(vSoHopDong) = "" And iWhiteRow = 0 Then
               .Col = .ColLetterToNumber("F")
                .CellNote = vErrMsgMSTReq  'static
                .BackColor = mErrorColor
                blCheck_S3 = False
            Else
               .Col = .ColLetterToNumber("F")
                .CellNote = ""  'static
                .BackColor = mNonErrorColor
            End If
            
            'check ngay HD
            If Trim(ngayHD) <> "" And iWhiteRow = 0 Then
            
                If DateSerial(CInt(Mid$(ngayHD, 7, 4)), CInt(Mid$(ngayHD, 4, 2)), CInt(Mid$(ngayHD, 1, 2))) > DateSerial(CInt(Mid$(ngaycuoiky, 7, 4)), CInt(Mid$(ngaycuoiky, 4, 2)), CInt(Mid$(ngaycuoiky, 1, 2))) Or DateSerial(CInt(Mid$(ngayHD, 7, 4)), CInt(Mid$(ngayHD, 4, 2)), CInt(Mid$(ngayHD, 1, 2))) < DateSerial(CInt(Mid$(ngaydauky, 7, 4)), CInt(Mid$(ngaydauky, 4, 2)), CInt(Mid$(ngaydauky, 1, 2))) Then
                    .Col = .ColLetterToNumber("G")
                    .CellNote = vErrNgayHD  'static
                    .BackColor = mAlertColor
                    'iFlagEmty = 1
                Else
                    .Col = .ColLetterToNumber("G")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            ElseIf (Trim(ngayHD) = "" Or Trim(ngayHD) = "../../....") And iWhiteRow = 0 Then
                    .Col = .ColLetterToNumber("G")
                    .CellNote = vErrMsgMSTReq
                    .BackColor = mErrorColor
                    blCheck_S3 = False
                    'iFlagEmty = 1
            End If
            'Ten HD
            If Trim(Tenloaihd1) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("C")
                .CellNote = vErrMsgMSTReq  'static
                .BackColor = mErrorColor
                'iFlagEmty = 1
            Else
                .Col = .ColLetterToNumber("C")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            ' Mau so HD
            If Trim(mauso1) = "" And iWhiteRow = 0 And iWhiteRow1 = 0 Then
                .Col = .ColLetterToNumber("I")
                .CellNote = vErrMsgMSTReq  'static
                .BackColor = mErrorColor
               ' iFlagEmty = 1
            Else
                .Col = .ColLetterToNumber("I")
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
            result = CheckMauSoBLP(CStr(mauso1), strLoai, strTemp)
            If result = "1" And iWhiteRow = 0 And iWhiteRow1 = 0 Then
                .Col = .ColLetterToNumber("I")
                .CellNote = errMauSoHD  'static
                .BackColor = mErrorColor
                'iFlagMauSoHD = 1
                blCheck_S7 = False
            ElseIf result = "2" And iWhiteRow = 0 And iWhiteRow1 = 0 Then
                .Col = .ColLetterToNumber("I")
                .CellNote = errSoLienHD  'static
                .BackColor = mErrorColor
                'iFlagSoLienHD = 1
                blCheck_S8 = False
            Else
                .Col = .ColLetterToNumber("I")
                .CellNote = ""
                .BackColor = mNonErrorColor
            
            End If
            '/-----------------------------------
            Tuso1 = Val(Tuso1)
            denso1 = Val(denso1)
            'kiem tra tu so <= den so
            If Tuso1 > denso1 Then
                    .Col = .ColLetterToNumber("L")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgTusoDenSo    'static
                    .BackColor = mErrorColor
                     .Col = .ColLetterToNumber("M")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgTusoDenSo    'static
                    .BackColor = mErrorColor
                    blCheck_S6 = False
            End If
            'Kiem tra cau truc cua ky hieu
            Dim varTemp As Variant
            Dim maLoaiBLP As Variant
            .GetText .ColLetterToNumber("K"), .Row, varTemp
'            .SetText .ColLetterToNumber("O"), .Row, CheckSoBLP(CStr(varTemp), strLoai)
'            .Col = .ColLetterToNumber("O")
'            UpdateCell fps, .Col, .Row, .Text
            .GetText .ColLetterToNumber("H"), .Row, maLoaiBLP
            SohieuHD = CheckSoBLP(CStr(varTemp), maLoaiBLP)
            

            If (Trim(SohieuHD) = "1" Or Trim(SohieuHD) = "2") And iWhiteRow = 0 And iWhiteRow1 = 0 Then
                .Col = .ColLetterToNumber("K")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsgSoHD    'static
                .BackColor = mErrorColor
                iFlagvHoVaTen = 1
                blCheck_S5 = False
                'capture the first error cell
                If iFocusFlag = 0 Then
                    strFocusSheetName = "'" & .SheetName & "'"
                    strFocusCol = str(.Col)
                    strFocusRow = str(.Row)
                    iFocusFlag = 1
                End If
            Else
                .Col = .ColLetterToNumber("K")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            
'            If (Trim(vHoVaTen) <> vbNullString Or Trim(vMST) <> vbNullString) And iWhiteRow = 0 Then
             If iWhiteRow = 0 Then
'                If Trim(vHoVaTen) = vbNullString Then
'                    .Col = .ColLetterToNumber("D")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgMSTReq    'static
'                    .BackColor = mErrorColor
'                    iFlagvHoVaTen = 1
'                    blCheck_S1 = False
'                    'capture the first error cell
'                    If iFocusFlag = 0 Then
'                        strFocusSheetName = "'" & .SheetName & "'"
'                        strFocusCol = str(.Col)
'                        strFocusRow = str(.Row)
'                        iFocusFlag = 1
'                    End If
'                Else
'                    .Col = .ColLetterToNumber("D")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                End If
                
            'Check ma so thue tren to khai 02A/TNCN
            If Trim(vMST) <> vbNullString Then
                   'vMST = Replace(vMST, " ", "")
    '               vMST = Replace(vMST, "-", "")
    '                vMST = Left(vMST, 10)
                    
                    If Len(vMST) = 14 And InStr(vMST, "-") > 0 Then
                        vMST = Left$(vMST, 10) & Right$(vMST, 3)
                    End If
                    vMST = Replace(vMST, "-", "")
    
    
                   strCheck = ""
                   If vMST <> Replace(vMST, " ", "") Then
                        strCheck = ""
                   Else
                        If Len(vMST) = 10 Or Len(vMST) = 13 Then
                             strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                        End If
                   End If
                   ' Kiem tra dinh dang theo cau truc cua co quan thue
                   If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                        .Col = .ColLetterToNumber("C")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgMST   'static
                        .BackColor = mErrorColor
                        iFlagvMst = 1
                        blCheck_S2 = False
                        'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
                    Else
                         .Col = .ColLetterToNumber("C")
                         .CellNote = ""
                         .BackColor = mNonErrorColor
                    End If
            Else
                    .Col = .ColLetterToNumber("C")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMSTReq    'static
                    .BackColor = mErrorColor
                    iFlagvMst = 1
                    blCheck_S3 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
            End If
            Else
            .Col = .ColLetterToNumber("C")
            .CellNote = ""
            .BackColor = mNonErrorColor
            .Col = .ColLetterToNumber("D")
            .CellNote = ""
            .BackColor = mNonErrorColor
            End If
            
             ' Kiem tra trung nhau tu so den so
            'bat dau
                    Do
                        i = .Row + 1
                        .GetText .ColLetterToNumber("L"), i, Tuso2
                        .GetText .ColLetterToNumber("M"), i, denso2
                        Tuso2 = Val(Tuso2)
                        denso2 = Val(denso2)
                        .GetText .ColLetterToNumber("J"), i, Tenloaihd2
                        .GetText .ColLetterToNumber("I"), i, mauso2
                        .GetText .ColLetterToNumber("K"), i, kyhieu2
                        Dim vMST2 As Variant
                        .GetText .ColLetterToNumber("C"), i, vMST2
                        
                            If (((Tuso1 - Tuso2) * (denso1 - Tuso2)) <= 0 Or ((Tuso1 - denso2) * (denso1 - denso2)) <= 0 Or ((Tuso1 > Tuso2) And (denso1 < denso2))) And ((Tenloaihd1 = Tenloaihd2) And (mauso1 = mauso2) And (kyhieu1 = kyhieu2) And (vMST = vMST2)) Then
                                .Row = j
                                .Col = .ColLetterToNumber("L")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrMsgMSTTrungNhau     'static
                                .BackColor = mErrorColor
                                
                                
                                .Row = j
                                .Col = .ColLetterToNumber("M")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrMsgMSTTrungNhau     'static
                                .BackColor = mErrorColor
                                
                                
                                .Row = i
                                .Col = .ColLetterToNumber("L")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrMsgMSTTrungNhau     'static
                                .BackColor = mErrorColor
                               
                                
                                .Row = i
                                .Col = .ColLetterToNumber("M")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrMsgMSTTrungNhau     'static
                                .BackColor = mErrorColor
                                
    
                                'iFlagvMst = 1
                                blCheck_S4 = False
                                'capture the first error cell
                                If iFocusFlag = 0 Then
                                    strFocusSheetName = "'" & .SheetName & "'"
                                    strFocusCol = str(.Col)
                                    strFocusRow = str(.Row)
                                    iFocusFlag = 1
                                End If
                                Else
    '                            .Row = i
    '                            .Col = .ColLetterToNumber("O")
    '                            .CellNote = ""
    '                            .BackColor = mNonErrorColor
    '                            .Text = .Text - 1
                             End If
                        .Row = i
                        .Col = .ColLetterToNumber("B")
                        ' Check cho den het cac dong co du lieu thi thoi
                    Loop Until UCase(.Text) = "AA"
                    
                    '--------------------------------------------------
            .Row = j
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
            ' Check cho den het cac dong co du lieu thi thoi
        Loop Until UCase(.Text) = "AA"
      
        'check ngay ky , nguoi ky
        'dhdang sua
        Dim ngayky As Variant
        Dim nguoiky As Variant
        Dim totalRowD As Variant
        
        i = 1
        .Col = .ColLetterToNumber("B")
        .Row = 22
        Do
             .Col = .ColLetterToNumber("B")
             .Row = i + 22
             totalRowD = i
             i = i + 1
        Loop Until .Text = "aa"
        
        
        .GetText .ColLetterToNumber("J"), totalRowD + 29, nguoiky
        .GetText .ColLetterToNumber("J"), totalRowD + 31, ngayky
        
        
        If Trim(nguoiky) = "" Then
                .Col = .ColLetterToNumber("J")
                .Row = totalRowD + 29
                .CellNote = vErrMsgMSTReq  'static
                .BackColor = mErrorColor
                blCheck_S3 = False
                'iFlagEmty = 1
        Else
                .Row = totalRowD + 29
                .Col = .ColLetterToNumber("J")
                .CellNote = ""
                .BackColor = mNonErrorColor
        End If
        
        If Trim(ngayky) = "" Then
                .Col = .ColLetterToNumber("J")
                .Row = totalRowD + 31
                .CellNote = vErrMsgMSTReq  'static
                .BackColor = mErrorColor
                blCheck_S3 = False
                'iFlagEmty = 1
        Else
                ngayky = DateSerial(CInt(Mid$(ngayky, 7, 4)), CInt(Mid$(ngayky, 4, 2)), CInt(Mid$(ngayky, 1, 2)))
                If ngayky > Now Then
                    .Col = .ColLetterToNumber("J")
                    .Row = totalRowD + 31
                    .CellNote = vErrNgayKy  'static
                    .BackColor = mErrorColor
                    blCheck_S9 = False
                    'iFlagEmty = 1
                ElseIf ngayky < DateSerial(CInt(Mid$(vTuNgayBC, 7, 4)), CInt(Mid$(vTuNgayBC, 4, 2)), CInt(Mid$(vTuNgayBC, 1, 2))) Then
                    .Col = .ColLetterToNumber("J")
                    .Row = totalRowD + 31
                    .CellNote = errNgayBCHL  'static
                    .BackColor = mErrorColor
                    blCheck_NgayBCHL = False
                Else
                    .Row = totalRowD + 31
                    .Col = .ColLetterToNumber("J")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
        End If
        
        

        ' Trong truong hop blCheck_S2 khong co MST nao bi loi thi set lai gia tri cua Header la 1
        If blCheck_S2 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 12, "1"
            '.SetText .ColLetterToNumber("B"), 13, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 12, "0"
            '.SetText .ColLetterToNumber("B"), 13, "0"
            
        End If
                If blCheck_S3 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 13, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 13, "0"
        End If
        
        
        If blCheck_S4 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 14, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 14, "0"
        End If
        
        If blCheck_S5 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "0"
        End If
        
        If blCheck_S6 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 16, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 16, "0"
        End If
        
        If blCheck_S7 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 17, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 17, "0"
        End If
        
        If blCheck_S8 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 18, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 18, "0"
        End If
        
        If blCheck_S9 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 18, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 18, "0"
        End If
        
        If blCheck_NgayBCHL = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 21, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 21, "0"
        End If
        
        .Sheet = iCurrentSheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub clearCheckSheet1()
    
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
        
        .Sheet = iCurrentSheet
    End With
End Sub

Public Sub SetActiveSheet()
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v2.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "01"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
                    If lCol = fps.ColLetterToNumber("C") Or _
                       lCol = fps.ColLetterToNumber("D") Then
                        strValue = GetAttribute(xmlCellNode, "Value")
                        If Not IsNullValue(strValue) Then
                            blnNullValue = False
                            Exit For
                        End If
                    End If
                Next
            
                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If
        End Select
    Next
    
    Set xmlCellNode = Nothing
    Set xmlNodeValid = Nothing
End Sub
Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean
    Dim totalCell  As Long
    Dim countRowReset As Long
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    totalCell = TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length
    countRowReset = (totalCell - 5) / 14
    With fps
        For Each xmlNodeReset In TAX_Utilities_v2.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            .Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            If (lRow < 22 Or lRow > 22 + countRowReset) And mCurrentSheet = 1 Then
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

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()

End Sub

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

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
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               'dhdang sua de ko set lai gia tri cho nut check
               If lCol <> .ColLetterToNumber("B") And lRow <> 15 Then
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
                End If
             Next

      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

Public Sub FinishImport()

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
            .Row = 22
            Do
                .GetText .ColLetterToNumber("I"), .Row, maBLP
                 GetMaBLP maBLP, tenBLP
                 If tenBLP <> "" Then
                    arrTemp = Split(tenBLP, "###")
                    If UBound(arrTemp) > 0 Then
                        .Col = .ColLetterToNumber("H")
                        .Text = arrTemp(0)
                        UpdateCell fps, .Col, .Row, .Text
                        .Col = .ColLetterToNumber("N")
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
