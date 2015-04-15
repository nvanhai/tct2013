VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01BVMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Thu Nhap Doanh Nghiep" interface sheets
'this Class is belong to TAX_Business_v2 project which will be compline to DLL
'Coder: PhuocHK
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Private Const S1Txt1Row = 9
Private Const S1Txt1Col = "I"
Private Const S1Txt1aRow = 10
Private Const S1Txt1aCol = "I"
Private Const S1Txt1bRow = 11
Private Const S1Txt1bCol = "I"
Private Const S1Txt1cRow = 12
Private Const S1Txt1cCol = "I"
Dim vt As String, va As String, vb As String, vc As String
Dim strComboHien As String, strCombo As String
Dim arrTaiNguyen As New Scripting.Dictionary, mblKhoiDau As Boolean
Dim arrTaiNguyenGoc As New Scripting.Dictionary
Dim blKiemTraDieuChinhCu As Boolean
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
Dim mblKeyF5 As Boolean
Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim arrData() As String
Dim arrDanhSachKHBS() As String
Dim TenNVThue As Variant, SoChungChi As Variant
Public loaiKHBS As String
Public FlagThaydoi As String
Public Strloaitk As String
Public StrSolanBosung As String
Private Const Row_MSTDN = 14
Private Const Col_MSTDN = "I"

' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim strDataFileName As String
    Dim strDDKTTN, strDonvitinh As String
    Dim arrDanhsach() As String, blChon As Boolean
    blChon = False
    strDataFileName = GetCatalogueFileName(blChon)
    
    
    With fps
        SetDateFormat fps, 1, 62, .ColLetterToNumber("R"), DDMMYYYY
        .Sheet = 1
        .Row = 62
        .Col = .ColLetterToNumber("R")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        SetDateFormat fps, 2, 31, .ColLetterToNumber("BE"), DDMMYYYY
        SetDateFormat fps, 2, 46, .ColLetterToNumber("BG"), DDMMYYYY
        .Sheet = 2
        .Row = 46
        .Col = .ColLetterToNumber("BG")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
    End With
    arrTaiNguyen.RemoveAll
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
        For Each xmlNode In xmlNodeListCell
            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                If (blChon = True And arrDanhsach(6) = "1") Or blChon = False Then
                    strCombo = strCombo + CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 10) + CPab(arrDanhsach(5), 1) + Chr$(9)
                    arrTaiNguyen.Add arrDanhsach(1), CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 10) + CPab(arrDanhsach(5), 1) + Chr$(9)
                    strComboHien = strComboHien + arrDanhsach(2) + Chr$(9)
                End If
            End If
        Next
        Dim xmlDocument As New MSXML.DOMDocument
        Dim xmlNode1 As MSXML.IXMLDOMNode
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_29122007_2.xml"))
        Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
        For Each xmlNode1 In xmlNodeListMap
            strDonvitinh = strDonvitinh + GetAttribute(xmlNode1, "value") + Chr$(9)
        Next
    End If

    strDataFileName = GetCatalogueTemplateFile
    
    arrTaiNguyenGoc.RemoveAll
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
        For Each xmlNode In xmlNodeListCell
            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                If (blChon = True And arrDanhsach(6) = "1") Or blChon = False Then
                    arrTaiNguyenGoc.Add arrDanhsach(1), CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 10) + CPab(arrDanhsach(5), 1) + Chr$(9)
                End If
            End If
        Next
    End If
    
    With fps
        'them danh muc hang vao combobox danh muc 1
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        .Col = .ColLetterToNumber("AB")
        .Row = 45
        .TypeComboBoxList = strCombo 'CPab("Maple", 100) + Chr$(9) + CPab("Oak", 100) + Chr$(9) + CPab("Ash", 100)
        'Combo Hien
        .Col = .ColLetterToNumber("C")
        .Row = 45
        .TypeComboBoxList = strComboHien
        .Col = .ColLetterToNumber("J")
        .TypeComboBoxList = strDonvitinh
        'them danh muc hang vao combobox danh muc 2
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        .Col = .ColLetterToNumber("AB")
        .Row = 49
        .TypeComboBoxList = strCombo 'CPab("Maple", 100) + Chr$(9) + CPab("Oak", 100) + Chr$(9) + CPab("Ash", 100)
        'Combo Hien
        .Col = .ColLetterToNumber("C")
        .Row = 49
        .TypeComboBoxList = strComboHien
        .Col = .ColLetterToNumber("J")
        .TypeComboBoxList = strDonvitinh
        End With

        Set xmlNodeListCell = Nothing
        ' Lay dia diem khai thac tai nguyen
        If TAX_Utilities_v2.Month = "01" Then 'Thang 1 ve thang 12 nam truoc
            strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & "12" & Trim(str(CInt(TAX_Utilities_v2.Year) - 1)) & ".xml"
        Else
            strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") & "_" & Right("00" & Trim(str(CInt(TAX_Utilities_v2.Month) - 1)), 2) & TAX_Utilities_v2.Year & ".xml"
        End If
        
   mblKeyF5 = False
   
               ' set CQT ra quyet dinh hoan
    setCQTQuanLyHoanThue fps
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
Dim i As Integer, intIndexCombo As Integer
Dim strLTN As Variant, Col7 As Variant, strLTNCu As Variant, strId As Variant
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode
Dim iCol As Long, iRow As Long, GTSua As String, tRow As Long
Dim strComboName, strComboMCT, strCombo As String
Dim count1 As Long
Dim tongNTNN As Variant, tienDoiThanhChu As String
    With fps
        mblKhoiDauToKhai = True
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
         'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("C")
            .Row = 35
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 35
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 35
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("C")
            .Row = 35
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 35
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 35
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
         'tong tien bang chu
         .GetText .ColLetterToNumber("T"), 37, tongNTNN
        If str(tongNTNN) <> "" Then
            tongNTNN = Replace(tongNTNN, ".", "")
            tongNTNN = Replace(tongNTNN, ",", "")
            tienDoiThanhChu = doiTien(CStr(tongNTNN))
            .SetText .ColLetterToNumber("AB"), 37, tienDoiThanhChu
            UpdateCell fps, .ColLetterToNumber("AB"), 37, tienDoiThanhChu
        End If
        'set gia tri cho cot so tien bang chu
        .SetText .ColLetterToNumber("B"), 55, tienDoiThanhChu
        UpdateCell fps, .ColLetterToNumber("B"), 55, tienDoiThanhChu
        
        .Sheet = 1
            mCurrentSheet = 1

            Set xmlCellNode = TAX_Utilities_v2.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("C"), 45))

            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode.lastChild
            Set xmlCellNode = xmlCellNode.firstChild

            'set ten nhan vien dai ly
            Set xmlCellNode = xmlCellNode.firstChild
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            .Row = iRow
            .Col = iCol
            .Formula = ""
            'set nguoi ky
            Set xmlCellNode = xmlCellNode.nextSibling
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            .Row = iRow
            .Col = iCol
            .Formula = ""
            'set chung chi hanh nghe
            Set xmlCellNode = xmlCellNode.nextSibling
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            .Row = iRow
            .Col = iCol
            .Formula = ""
            'set ngay ky
            Set xmlCellNode = xmlCellNode.nextSibling
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            .Row = iRow
            .Col = iCol
            .Formula = ""
        
        'Lay gia tri khoi dau cho bien strGiaTriCu
        .GetText .ColLetterToNumber("C"), 45, strGiaTriCu
       
    End With
    
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
Dim strId As Variant, strLTN As Variant
Dim intIndexCombo As Integer
Dim Col7 As Variant
Dim iRow As Long, iCol As Long, GTSua As String
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tCol As Long, tRow As Long
Dim cCol As Long, cRow As Long
Dim cValue As Variant, temp As Variant, varTemp As Variant
Dim cValueG As Variant
Dim cTagF As String
Dim cValueF As Variant
Dim i As Integer
Dim mblLoi As Boolean



Dim xmlNodeH15 As MSXML.IXMLDOMNode
     fps.EventEnabled(EventAllEvents) = False
    With fps
       .Sheet = 1
       'vttoan
       'khi tron danh muc khoang san mac dinh don vi tinh la kg va cho phep sua sua
       If Col = .ColLetterToNumber("C") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("AB")
            .typeComboboxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber("AB")
            UpdateCell fps, .Col, .Row, .Text
            'Chon lai combo cua colF
            .Col = .ColLetterToNumber("AC")
            .GetText .ColLetterToNumber("AB"), Row, strId
            .Text = Trim(Mid(strId, 200, 10))
             UpdateCell fps, .Col, .Row, .Text
             .Col = .ColLetterToNumber("AD")
             .Text = Trim(Left(strId, 200))
             UpdateCell fps, .Col, .Row, .Text
             .Col = .ColLetterToNumber("J")
             .Text = Trim(Mid(strId, 210, 10))
             UpdateCell fps, .Col, .Row, .Text
             If strId <> "" Then
                .Col = .ColLetterToNumber("AE")
                .Text = strId
                UpdateCell fps, .Col, .Row, .Text
             End If
       End If
       mblLoi = False
            'Huy chon Combo to khai khong hop le va ra thong bao
        blGiaTriCu = False
            If blGiaTriCu And Col = .ColLetterToNumber("C") Then 'And Not mblKeyF5 Then
                If Not .Visible Then
                    .Visible = True
                    DisplayMessage "0227", msOKOnly, miWarning
                    .Visible = False
                Else
                    DisplayMessage "0227", msOKOnly, miWarning
                End If
                
                .Sheet = 1
                mCurrentSheet = .Sheet
                .Row = RowGiaTriCu
                .Col = ColGiaTriCu
                .Text = strGiaTriCu
                .Sheet = 1
                mCurrentSheet = .Sheet
                .Row = RowGiaTriCu
                .Col = ColGiaTriCu
                UpdateCell fps, .Col, .Row, .Text
                blGiaTriCu = False
                mblLoi = True
            End If
'KHBS
     If .ActiveSheet = .SheetCount - 1 Then
         Dim cValueOld As Variant
         Dim cellid As String
         Dim xmlCellNodeData As MSXML.IXMLDOMNode
        ' Dim i As Integer
             fps.EventEnabled(EventAllEvents) = False
             .Sheet = .SheetCount - 1
              cCol = .Col
              cRow = .Row
             .Col = .ColLetterToNumber("BD")
             .Row = .Row
             .GetText .ColLetterToNumber("BC"), .Row, varTemp

       If varTemp <> "" Then
          Dim arr() As String
            For i = 0 To UBound(arrDanhSachKHBS)
               If varTemp = Left(arrDanhSachKHBS(i), InStr(arrDanhSachKHBS(i), "~") - 1) Then
                   cellid = "AN_" & Right(arrDanhSachKHBS(i), Len(arrDanhSachKHBS(i)) - InStr(arrDanhSachKHBS(i), "~") - 2)
                  For Each xmlCellNodeData In TAX_Utilities_v2.DataKHBS.getElementsByTagName("Cell")
                       If GetAttribute(xmlCellNodeData, "CellID") = cellid Then
                            cValue = CDbl(cValue) + CDbl(GetAttribute(xmlCellNodeData, "Value"))
                            Exit For
                        End If
                   Next
               End If
            Next

              .Col = .ColLetterToNumber("BF")
              .Row = Row
               .value = cValue
               UpdateCell fps, .Col, .Row, .value
               .Col = .ColLetterToNumber("BG")
               .Row = Row
               UpdateCell fps, .Col, .Row, .value
               .Col = .ColLetterToNumber("BH")
               .Row = Row
               UpdateCell fps, .Col, .Row, .value


            .Sheet = .SheetCount - 1
            Set xmlNodeH15 = TAX_Utilities_v2.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
                 ParserCellID fps, GetAttribute(xmlNodeH15, "CellID"), cCol, cRow
                .Col = cCol
                .Row = cRow
                UpdateCell fps, .Col, .Row, .value
                If .value < 0 Then
                    .Col = .ColLetterToNumber("BE")
                    .Row = cRow + 3
                    .value = 0
                    UpdateCell fps, .Col, .Row, 0
                 Else
                    .Col = .ColLetterToNumber("BE")
                    .Row = cRow + 3
                    UpdateCell fps, .Col, .Row, .value
                End If
            fps.EventEnabled(EventAllEvents) = True
         End If
    End If
    End With
   fps.EventEnabled(EventAllEvents) = True
End Sub

Private Sub doiTien2()
    Dim tongNTNN As Variant, tienDoiThanhChu As String

    With fps
         .EventEnabled(EventAllEvents) = False
        'If .ActiveSheet = 1 Then
            .Sheet = 1
            mCurrentSheet = .Sheet
            .GetText .ColLetterToNumber("T"), 37, tongNTNN
            tongNTNN = Replace(tongNTNN, ".", "")
            tongNTNN = Replace(tongNTNN, ",", "")
            tienDoiThanhChu = doiTien(CStr(tongNTNN))

            Dim i As Variant
            i = 1
            .Col = .ColLetterToNumber("B")
            .Row = 45
            Do
                 .Text = str(i)
                 .Col = .ColLetterToNumber("B")
                 .Row = i + 45
                 i = i + 1
            Loop Until .Text = "aa"
            .Row = .Row
            i = 1
            .Row = .Row + 3
            Do
                .Text = str(i)
                .Col = .ColLetterToNumber("B")
                .Row = .Row + 1
                i = i + 1
            Loop Until .Text = "aa"
            .SetText .ColLetterToNumber("B"), .Row + 5, tienDoiThanhChu
            UpdateCell fps, .ColLetterToNumber("B"), .Row + 5, tienDoiThanhChu
            
            .SetText .ColLetterToNumber("AB"), 37, tienDoiThanhChu
            UpdateCell fps, .ColLetterToNumber("AB"), 37, tienDoiThanhChu
            
            
            
        .EventEnabled(EventAllEvents) = True
    End With

End Sub

Private Sub fps_BeforeEditMode(ByVal Col As Long, ByVal Row As Long, ByVal UserAction As FPUSpreadADO.BeforeEditModeActionConstants, CursorPos As Variant, Cancel As Variant)
Dim iRow As Long, iCol As Long
With fps
    If .Sheet = 1 Then
        iRow = .ActiveRow
        iCol = .ActiveCol
        GetCellSpan fps, iCol, iRow
        If iCol = fps.ColLetterToNumber("C") Then
            fps.Col = fps.ColLetterToNumber("C")
            fps.Row = iRow
            strGiaTriCu = fps.Text
        End If
    End If
End With
End Sub

'Chuc nang: thay doi gia tri cua Combo
'Doi voi to khai
'- Xac dinh nhom cua mat hang
'- Kiem tra mat hang da duoc chon chua tren to khai va ban dieu chinh
    '- kiem tra trung Ten TN trong to khai
    '- Neu co thi dat co blGiaTriCu = True
    '- kiem tra da khai trong ban dieu chinh
    '- neu trung thi dat co blKiemTraDieuChinhCu = True
'Doi voi ban diem chinh
'- kiem tra ten hang va ky da duoc nhap chua
    'neu da duoc nhap thi dat co blDieuChinhCu = True
Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim intIndexCombo As Integer
Dim temp As Variant, itemp As Variant, iRow As Long, iCol As Long, iColRow As String
Dim tempTen As Variant, tempKy As Variant, itempTen As Variant, itempKy As Variant, itempId As Variant
Dim i As Integer, aSection As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim aCol As Long, aRow As Long, iSection As Integer

Dim intIndexCuc As Integer, intIndexChiCuc As Integer
Dim varIdCucThue As Variant
Dim xmlDomData As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode
Dim arrDanhsach() As String
Dim tenCucThue As String
Dim maCucThue As String
Dim tenChiCucThue As String
Dim maChiCucThue As String

With fps

    .Sheet = .ActiveSheet
    mCurrentSheet = .Sheet
    Select Case .Sheet
        Case 1
        'Lay vi tri O ban dau
        aCol = .ActiveCol
        aRow = .ActiveRow
        GetCellSpan fps, aCol, aRow
            
        'Xac dinh nhom cua mat hang
        iSection = 1
        aSection = 0
        Set xmlCellNode = TAX_Utilities_v2.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("C"), 45))
        Set xmlCellsNode = xmlCellNode.parentNode
        Do
            Do
                For i = 1 To xmlCellsNode.childNodes.length
                    ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
                    If iCol = aCol And iRow = aRow And aSection = 0 Then
                        aSection = iSection
                        Exit For
                    End If
                Next
                If xmlCellsNode.nextSibling Is Nothing Then Exit Do
                If aSection > 0 Then Exit Do
                Set xmlCellsNode = xmlCellsNode.nextSibling
            Loop Until xmlCellsNode Is Nothing
            If Not xmlCellsNode.parentNode.nextSibling Is Nothing Then
                Set xmlCellsNode = xmlCellsNode.parentNode.nextSibling.firstChild
            End If
            If GetAttribute(xmlCellsNode.parentNode, "Dynamic") = "1" Then
                iSection = iSection + 1
            End If
        Loop Until xmlCellsNode.parentNode.nextSibling Is Nothing
        
        'Kiem tra mat hang da duoc chon chua tren to khai va ban dieu chinh
        Set xmlCellNode = TAX_Utilities_v2.Data(1 - 1).nodeFromID(GetCellID(fps, Col, Row))
        Set xmlCellNode = xmlCellNode.parentNode.parentNode.firstChild.firstChild
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        blGiaTriCu = False
        blKiemTraDieuChinhCu = False
        If Col = fps.ColLetterToNumber("C") Then
            'kiem tra trung Ten TN trong to khai
            fps.Col = Col
            fps.Row = Row
            temp = fps.Text
'            Set xmlCellNode = TAX_Utilities_v2.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 13))
            Set xmlCellsNode = xmlCellNode.parentNode
'            Do
'            If GetAttribute(xmlCellsNode.parentNode, "Dynamic") = "1" Then
            
            i = 1
            .Col = .ColLetterToNumber("B")
            ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
            .Row = iRow
          
'            Do
'                 .Col = .ColLetterToNumber("C")
'                 itemp = .Text
'                 If .Row <> Row Then
'                    If itemp = temp Then 'Neu co thi dat co blGiaTriCu = True
'                        RowGiaTriCu = Row
'                        ColGiaTriCu = .ColLetterToNumber("C")
'                        .Row = Row
'                        .Col = .ColLetterToNumber("C")
'                        blGiaTriCu = True
'                        Exit Sub
'                    End If
'                 End If
'                 .Col = .ColLetterToNumber("B")
'                 .Row = i + iRow
'                 i = i + 1
'            Loop Until .Text = "aa"
            
            .Sheet = 1
            mCurrentSheet = 1
        End If
'KHBS
'        Case .SheetCount - 1
'                   'lay ve dong combobox
'            Set xmlCellNode = TAX_Utilities_v2.Data(.SheetCount - 2).nodeFromID(GetCellID(fps, .ColLetterToNumber("BI"), 5))
'            Set xmlCellNode = xmlCellNode.parentNode
'            Set xmlCellNode = xmlCellNode.parentNode
'            Set xmlCellNode = xmlCellNode.parentNode.lastChild
'            Set xmlCellNode = xmlCellNode.firstChild
'            Set xmlCellNode = xmlCellNode.firstChild
'            For i = 1 To 6
'                Set xmlCellNode = xmlCellNode.nextSibling
'            Next i
'            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'
'
'            .Sheet = .SheetCount - 1
'            mCurrentSheet = .SheetCount - 1
'
'            .Col = Col
'            .Row = Row
'
'            If Col = .ColLetterToNumber("BE") And Row = iRow Then 'Khi chon Combo
'                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
'                intIndexCuc = .TypeComboBoxCurSel
'                .Col = .ColLetterToNumber("BI")
'                .Row = Row
'                .TypeComboBoxCurSel = intIndexCuc
'                'Nhap lai Combo cot Q
'                .Col = .ColLetterToNumber("BI")
'                UpdateCell fps, .Col, .Row, .Text
'                ' Lay Id cua cuc thue vua chon o tren
'                .GetText .Col, .Row, varIdCucThue
'
'                If varIdCucThue <> "" Or varIdCucThue <> vbNullString Then
'                    ' Lay thong tin cho chi cuc thue
'                    varIdCucThue = Left$(varIdCucThue, 3)
'
'                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
'                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
'                        For Each xmlNode In xmlNodeListCell
'                            If GetAttribute(xmlNode, "Value") <> "" Then
'                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
'
'                                If arrDanhsach(0) = "0" And arrDanhsach(2) = varIdCucThue Then
'                                    tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
'                                    maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
'                                End If
'                            End If
'                        Next
'                        Set xmlDomData = Nothing
'                        Set xmlNodeListCell = Nothing
'                        Set xmlNode = Nothing
'                    End If
'
'                    .Row = iRow + 2
'                    .Col = .ColLetterToNumber("BE")
'                    .TypeComboBoxList = tenChiCucThue
'
'                    .Row = iRow + 2
'                    .Col = .ColLetterToNumber("BI")
'                    .TypeComboBoxList = maChiCucThue
'                End If
'            End If
'
'            If Col = .ColLetterToNumber("BE") And Row = iRow + 2 Then 'Khi chon Combo
'                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
'                intIndexChiCuc = .TypeComboBoxCurSel
'                .Col = .ColLetterToNumber("BI")
'                .Row = Row
'                .TypeComboBoxCurSel = intIndexChiCuc
'                'Nhap lai Combo cot Q
'                .Col = .ColLetterToNumber("BI")
'                UpdateCell fps, .Col, .Row, .Text
'            End If
        End Select
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
            fps.EventEnabled(EventAllEvents) = False
                 i = 1
                 .Col = .ColLetterToNumber("B")
                 .Row = 45
                 Do
                      .Text = str(i)
                      .Col = .ColLetterToNumber("B")
                      .Row = i + 45
                      i = i + 1
                 Loop Until .Text = "aa"
                 .Row = .Row
                 i = 1
                 .Row = .Row + 3
                 Do
                     .Text = str(i)
                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                     i = i + 1
                 Loop Until .Text = "aa"
            End If
            fps.EventEnabled(EventAllEvents) = True
         End If
         If .ActiveSheet = .SheetCount - 1 Then
                If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
                 fps.EventEnabled(EventAllEvents) = False
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 9
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("BE")
                         .Text = "10"
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 9
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 3
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("BE")
                         .Text = "10"
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
                End If
                 fps.EventEnabled(EventAllEvents) = True
            End If
    End With
End Sub
Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim CellNguoiKy As String
    Dim tongNTNN As Variant, tienDoiThanhChu As String
    
    CellNguoiKy = GetAttribute(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length - 1), "CellID")
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            .GetText .ColLetterToNumber("T"), 37, tongNTNN
            tongNTNN = Replace(tongNTNN, ".", "")
            tongNTNN = Replace(tongNTNN, ",", "")
            tienDoiThanhChu = doiTien(CStr(tongNTNN))
            .SetText .ColLetterToNumber("AB"), 37, tienDoiThanhChu
            UpdateCell fps, .ColLetterToNumber("AB"), 37, tienDoiThanhChu
            Dim i As Variant
            i = 1
            .Col = .ColLetterToNumber("B")
            .Row = 45
            Do
                 .Text = str(i)
                 .Col = .ColLetterToNumber("B")
                 .Row = i + 45
                 i = i + 1
            Loop Until .Text = "aa"
            .Row = .Row
            i = 1
            .Row = .Row + 3
            Do
                .Text = str(i)
                .Col = .ColLetterToNumber("B")
                .Row = .Row + 1
                i = i + 1
            Loop Until .Text = "aa"
            .SetText .ColLetterToNumber("B"), .Row + 5, tienDoiThanhChu
            UpdateCell fps, .ColLetterToNumber("B"), .Row + 5, tienDoiThanhChu
        
            If Col = .ColLetterToNumber("R") And Row = .MaxRows - 1 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                        .TypeHAlign = TypeHAlignLeft
                    Else
                               .SetActiveCell Col, Row
                               .SetFocus
                    End If
                Else
                    .SetText Col, Row, ""
                End If
                .Col = Col
                .Row = Row
                UpdateCell fps, .Col, .Row, .Text
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
    'CellChange Col, Row
End Sub
Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
  
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
                    
        .EventEnabled(EventAllEvents) = False
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .Sheet = 1
        Dim i As Integer
        i = 1
         .Col = .ColLetterToNumber("B")
         .Row = 45
         Do
              .Text = str(i)
              .Col = .ColLetterToNumber("B")
              .Row = i + 45
              i = i + 1
         Loop Until .Text = "aa"
         .Row = .Row
         i = 1
         .Row = .Row + 3
         Do
             .Text = str(i)
             .Col = .ColLetterToNumber("B")
             .Row = .Row + 1
             i = i + 1
         Loop Until .Text = "aa"
         
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
Dim Error As Variant
Dim Dvtinh As Variant
Dim loaiks As Variant
Dim sluong As Variant
Dim mphi As Variant
Dim iFlagEmty As Variant
Dim countRow As Integer
    With fps
      .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 15, Error
      .Sheet = 1
      Dim i As Integer
     i = 1
     countRow = 0
     .Row = 45
     .Col = .ColLetterToNumber("B")
      Do
        .GetText .ColLetterToNumber("L"), .Row, sluong
        .GetText .ColLetterToNumber("Q"), .Row, mphi
        .GetText .ColLetterToNumber("C"), .Row, loaiks
        .GetText .ColLetterToNumber("J"), .Row, Dvtinh
            If Trim(sluong) <> 0 Or Trim(mphi) <> 0 Or Trim(Dvtinh) <> "" Then
                If Trim(loaiks) = "" Then
                    .Col = .ColLetterToNumber("C")
                    .CellNote = Error
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                Else
                    .Col = .ColLetterToNumber("C")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            Else
                .ColLetterToNumber ("C")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
      .Col = .ColLetterToNumber("B")
      .Row = 45 + i
      i = i + 1
      countRow = countRow + 1
      Loop Until .Text = "aa"
      
      i = 1
      .Row = 48 + countRow
      .Col = .ColLetterToNumber("B")
      Do
        .GetText .ColLetterToNumber("L"), .Row, sluong
        .GetText .ColLetterToNumber("Q"), .Row, mphi
        .GetText .ColLetterToNumber("C"), .Row, loaiks
        .GetText .ColLetterToNumber("J"), .Row, Dvtinh
            If Trim(sluong) <> 0 Or Trim(mphi) <> 0 Or Trim(Dvtinh) <> "" Then
                If Trim(loaiks) = "" Then
                    .Col = .ColLetterToNumber("C")
                    .CellNote = Error
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                Else
                    .Col = .ColLetterToNumber("C")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            Else
                .ColLetterToNumber ("C")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
      .Col = .ColLetterToNumber("B")
      .Row = 48 + countRow + i
      i = i + 1
      Loop Until .Text = "aa"
     .Sheet = .SheetCount
     .SetText .ColLetterToNumber("B"), 15, IIf(iFlagEmty = 1, "0", "1")
  
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

Public Sub SetActiveSheet()
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    Dim strValue3 As String
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v2.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
             Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v2.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v2.Data(1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Then
                            blnNullValue = False
                    End If

                    If blnNullValue Then
                        'SetAttribute xmlNodeValid, "Active", "0"
                        'FlagThaydoi = "1"
                    End If
        End Select
    Next
    
    Set xmlCellNode = Nothing
    Set xmlNodeValid = Nothing
End Sub
Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol         As Long, lRow As Long
    Dim IsUpdate     As Boolean
    Dim idtkhai      As Variant
    Dim totalCell1   As Long, countRowReset1 As Long
    Dim a            As Long
    Dim sheetTemp    As Integer
    
    With fps
        
        sheetTemp = mCurrentSheet
        mCurrentSheet = .ActiveSheet
        .Col = .ColLetterToNumber("Q")
        .Row = 13
        .Text = ""
        UpdateCell fps, .Col, .Row, .Text
         
        .Col = .ColLetterToNumber("H")
        .Row = 15
        .Text = ""
        UpdateCell fps, .Col, .Row, .Text
        

'          a = (TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length - 12) / 6

        For Each xmlNodeReset In TAX_Utilities_v2.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            fps.Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            fps.Col = lCol
            fps.Row = lRow

                If .ActiveSheet = 1 Then
                    .Sheet = mCurrentSheet
                    If (lRow < 40 Or lRow > .MaxRows - 10) Then
                        GoTo next1
                    End If
                ElseIf .ActiveSheet = 2 Then
                    .Sheet = mCurrentSheet
                    If (lRow < 5 Or lRow > 21) Then
                        GoTo next1
                    End If
                End If
                
                Select Case fps.CellType

                        Case CellTypeCheckBox
                            fps.Text = vbNullString
                            UpdateCell fps, lCol, lRow, vbNullString

                        Case CellTypeComboBox
                            fps.Text = vbNullString
                            UpdateCell fps, lCol, lRow, vbNullString

                        Case CellTypeNumber
                            fps.value = 0
                            UpdateCell fps, lCol, lRow, vbNullString

                        Case Else
                            fps.value = vbNullString
                            UpdateCell fps, lCol, lRow, vbNullString
                    End Select


                
            TAX_Utilities_v2.AdjustData(mCurrentSheet - 1) = IIf(IsUpdate = True, IsUpdate, TAX_Utilities_v2.AdjustData(mCurrentSheet - 1))
next1:
        Next
       mCurrentSheet = sheetTemp
    End With
End Function
Public Sub SetData()
Dim tongNTNN As Variant, tienDoiThanhChu As String
'set lai du lieu cot gia tri tien bang chu o cuoi to khai.
    With fps
    .Sheet = 1
    mCurrentSheet = 1
    'tong tien bang chu
         .GetText .ColLetterToNumber("T"), 37, tongNTNN
        If str(tongNTNN) <> "" Then
            tongNTNN = Replace(tongNTNN, ".", "")
            tongNTNN = Replace(tongNTNN, ",", "")
            tienDoiThanhChu = doiTien(CStr(tongNTNN))
            .SetText .ColLetterToNumber("AB"), 37, tienDoiThanhChu
            UpdateCell fps, .ColLetterToNumber("AB"), 37, tienDoiThanhChu
        End If
        Dim i As Integer
        i = 1
         .Col = .ColLetterToNumber("B")
         .Row = 45
         Do
              .Text = str(i)
              .Col = .ColLetterToNumber("B")
              .Row = i + 45
              i = i + 1
         Loop Until .Text = "aa"
         .Row = .Row
         i = 1
         .Row = .Row + 3
         Do
             .Text = str(i)
             .Col = .ColLetterToNumber("B")
             .Row = .Row + 1
             i = i + 1
         Loop Until .Text = "aa"
         
        .SetText .ColLetterToNumber("B"), .Row + 5, tienDoiThanhChu
        UpdateCell fps, .ColLetterToNumber("B"), .Row + 5, tienDoiThanhChu
    End With
End Sub

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

Public Sub Prepare_KHBS()
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    Dim strPath As String
    Dim strComboName As String
    Dim strComboMCT As String
    Dim strCombo As String
    Dim blChon As Boolean
    Dim arrDanhsach() As String
    fps.Sheet = fps.SheetCount - 1
    fps.ActiveSheet = fps.Sheet
    
    blChon = False
    strPath = GetCatalogueFileName(blChon)
    
    'combo box khbs
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Cell")
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                If (blChon = True And arrDanhsach(6) = "1") Or blChon = False Then
                    strCombo = strCombo + CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 10) + CPab(arrDanhsach(5), 1) + Chr$(9)
                  '  strCombo = strCombo + CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 10) + Chr$(9)
                    strComboName = strComboName + arrDanhsach(2) + Chr$(9)
                    strComboMCT = strComboMCT + "10" + Chr$(9)
                End If
            End If
        Next
    End If
    
    With fps
    End With
    
End Sub


Public Sub UpdateChangeKHBS()
    Dim varTemp As Variant
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    Dim xmlCell As MSXML.IXMLDOMNode
    Dim xmlCellUpdate As MSXML.IXMLDOMNode
    Dim xmlCellGoc As MSXML.IXMLDOMNode
    Dim strId As String
    With fps
       If loaiKHBS = "frmKHBS_TT" Then
            .Sheet = 1
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateLastKHBSCell fps, lCol, lRow, varTemp
            Next
       Else
            For Each xmlCell In TAX_Utilities_v2.Data(1).getElementsByTagName("Cell")
               If Left(GetAttribute(xmlCell, "CellID"), 2) = "BC" Then
                    strId = Trim(Mid(GetAttribute(xmlCell, "Value"), 200, 10))
                        
                    For Each xmlCellGoc In TAX_Utilities_v2.DataKHBS.getElementsByTagName("Cell")
                        If Left(GetAttribute(xmlCellGoc, "CellID"), 1) = "AC" Then
                            If GetAttribute(xmlCellGoc, "Value") = strId Then
                            'update secsion
                               Set xmlCellUpdate = seachCellNode(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell"), GetAttribute(xmlCellGoc, "CellID"))
                               updateSecsion xmlCellUpdate.parentNode
                                'updateSecsion (xmlCellGoc.parentNode)
                            Else
                            'insert secsion
                            
                            End If
                            
                        End If
                    Next
               
               End If
            Next
       End If
        
        
        .Sheet = mCurrentSheet
    End With
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

Private Sub updateSecsion(xmlCellNode As MSXML.IXMLDOMNode)
Dim xmlCellNodeData As MSXML.IXMLDOMNode
Dim varTemp As Variant
Dim lCol As Long, lRow As Long
With fps
    For Each xmlCellNodeData In xmlCellNode.childNodes
     If Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "L" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "Q" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 2) = "U" Then
        ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
        .Sheet = 1
        .GetText lCol, lRow, varTemp
        UpdateLastKHBSCell fps, lCol, lRow, varTemp
     End If
    Next
End With
End Sub



Private Sub fillData()
    Dim varTemp As Variant
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    Dim arrDanhsach() As String
    Dim xmlCell As MSXML.IXMLDOMNode
    With fps
        Dim i As Integer
        Dim strFormular As String
        ReDim arrDanhSachKHBS(TAX_Utilities_v2.Data(0).getElementsByTagName("Section")(1).childNodes.length - 1)
        i = 0
        For Each xmlCellNodeData In TAX_Utilities_v2.DataKHBS.getElementsByTagName("Cell")
           If Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "AB" Then
                    ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .GetText lCol, lRow, varTemp
                     arrDanhSachKHBS(i) = varTemp & "~" & GetAttribute(xmlCellNodeData, "CellID")
                     i = i + 1
            End If
        Next
      .Sheet = .SheetCount - 1
    End With
End Sub
Private Function seachCellNode(xmlCellNodelist As MSXML.IXMLDOMNodeList, cellid As String) As MSXML.IXMLDOMNode
Dim nodeReturn As MSXML.IXMLDOMNode
    For Each nodeReturn In xmlCellNodelist
        If GetAttribute(nodeReturn, "CellID") = cellid Then
            Set seachCellNode = nodeReturn
            Exit For
            Exit Function
        End If
    Next
End Function

Sub fps_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Long, TipWidth As Long, TipText As String, ShowTip As Boolean)
    ' Set tip to display and set tip's content
      If Col = fps.ColLetterToNumber("G") Then
        ShowTip = True
        TipText = TipText
      End If
End Sub
Public Function getValueTK(strFileLoadData As String) As String

    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim result As String
    Dim iDom As Integer
    Dim MY As Date

    iDom = 0
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        If .Load(strFileLoadData) = True Then
            Set xmlNode = .nodeFromID("T_37")
            result = "CT6_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01BVMT", "getValueTK", Err.number, Err.Description
End Function

Public Function getDieuChinhGiam(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue, strDieuChinhGiam, strDieuChinhTang As String
    Dim tempArrOld() As String
    Dim tempArrCurr() As String
    Dim tempArrOldValue() As String
    Dim tempArrCurrValue() As String
    Dim i, j As Integer
    Dim strMaCT As String
    Dim arrMapCT() As String ' Luu gia tri cua dong chi tieu
    Dim arrMapCTValue() As String ' Luu gia tri cua cac cell chi tieu
    strMaCT = getMapCT
    With fps
                .Sheet = 1
                .GetText .ColLetterToNumber("T"), 37, varTemp
                strCurrentValue = "CT6_" & Trim(varTemp)
                ' tinh chenh lech giua 2 lan dieu chinh
                tempArrOld = Split(strOldValue, "~")
                tempArrCurr = Split(strCurrentValue, "~")
                arrMapCT = Split(strMaCT, "~")
                strDieuChinhGiam = ""
                strDieuChinhTang = ""
                For i = 0 To UBound(tempArrOld)
                    tempArrOldValue = Split(tempArrOld(i), "_")
                    tempArrCurrValue = Split(tempArrCurr(i), "_")
                    ' lay cac gia tri cua tung cot trong sheet KHBS
                    For j = 0 To UBound(arrMapCT)
                        arrMapCTValue = Split(arrMapCT(j), "_")
                        ' phan tu so 0 la ma chi tieu
                        If arrMapCTValue(0) = Mid(tempArrOldValue(0), 3) Then
                            strMaCT = arrMapCT(j)
                            Exit For
                        End If
                    Next j

                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
                        ' dieu chinh giam
                        strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
                        ' dieu chinh tang
                        strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                Next i
                If strDieuChinhGiam <> "" Then
                    strDieuChinhGiam = Mid(strDieuChinhGiam, 2, Len(strDieuChinhGiam) - 1)
                End If
                If strDieuChinhTang <> "" Then
                    strDieuChinhTang = Mid(strDieuChinhTang, 2, Len(strDieuChinhTang) - 1)
                End If
                getDieuChinhGiam = strDieuChinhGiam & "###" & strDieuChinhTang
      End With
End Function
Function getMapCT() As String
    On Error GoTo ErrHandle
    
    Dim tempValue As Variant
    Dim strPath As String
    Dim i As Integer
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT_TT28.xml"
    
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "ID") = GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") Then
                Exit For
            End If
        Next
    End If
    
    Set xmlNodeListItem = xmlNode.childNodes
    
    For Each xmlNodeItem In xmlNodeListItem
         tempValue = tempValue & "~" & GetAttribute(xmlNodeItem, "MCT") & "_" & GetAttribute(xmlNodeItem, "Caption")
    Next
    tempValue = Mid(tempValue, 2)
    getMapCT = tempValue
    Exit Function
ErrHandle:
    SaveErrorLog "01_BVMT", "getMapCT", Err.number, Err.Description
End Function

' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("T"), 37, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function
' Lay ve chi tieu 32 o to khai lan truoc
Public Function getValueCTDC(strFileLoadData As String) As String

    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode

    'Dim strOriginDataFileName As String
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim result As String
    Dim iDom As Integer
    Dim MY As Date

    iDom = 0
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        If .Load(strFileLoadData) = True Then
            Set xmlNode = .nodeFromID("T_37")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01BVMT", "getValueCTDC", Err.number, Err.Description
End Function
'******************************
'Set trang thai Flag KHBS active
Public Sub SetActiveKHBS() 'ByRef xmlNodeValid As MSXML.IXMLDOMNode)
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    Dim strValue3 As String
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v2.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
                Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v2.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v2.Data(1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v2.Data(1).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v2.Data(1).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v2.Data(1).getElementsByTagName("Cell").length > 31 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
                            blnNullValue = False
                    End If

                    If blnNullValue Then
                        SetAttribute xmlNodeValid, "Active", "0"
                        FlagThaydoi = "1"
                    End If
        End Select
    Next
    
    Set xmlCellNode = Nothing
    Set xmlNodeValid = Nothing
End Sub


' 27-01-2011
Private Sub UpdateSheets()
    Dim varTemp As Variant
    Dim ssSheet As Integer
    Dim blCheckS2 As Boolean
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet
         For Each xmlCellNodeData In TAX_Utilities_v2.Data(.Sheet - 1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
         Next
        
       ' Update lai sheet KHBS
        .Sheet = .SheetCount - 1
        mCurrentSheet = .Sheet
        blCheckS2 = IIf(TAX_Utilities_v2.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheckS2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(.Sheet - 1).getElementsByTagName("Cell")
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
'doc vao mot so va tra ve chu cua so do
Private Function doiSoThanhChu(ByVal so As String) As String
    Dim result As String
    result = ""
    Select Case so
            Case "0": result = ""
            Case "1": result = GetAttribute(GetMessageCellById("0204"), "Msg")
            Case "2": result = GetAttribute(GetMessageCellById("0205"), "Msg")
            Case "3": result = GetAttribute(GetMessageCellById("0206"), "Msg")
            Case "4": result = GetAttribute(GetMessageCellById("0207"), "Msg")
            Case "5": result = GetAttribute(GetMessageCellById("0208"), "Msg")
            Case "6": result = GetAttribute(GetMessageCellById("0209"), "Msg")
            Case "7": result = GetAttribute(GetMessageCellById("0210"), "Msg")
            Case "8": result = GetAttribute(GetMessageCellById("0211"), "Msg")
            Case "9": result = GetAttribute(GetMessageCellById("0212"), "Msg")
    End Select
    doiSoThanhChu = result
End Function
'ham doi tien tu so thanh chu
Private Function doiTien(ByVal so As String) As String
    Dim GTtien As String
    Dim value As String
    Dim i As Integer
    'luu lai gia tri so ban dau de xu ly
    Dim soTam2 As String
    Dim soTam As String
    soTam = so
    soTam2 = "1"
    For i = 0 To Len(so) - 1 Step 1

            If Len(so) Mod 3 = 2 And Mid(so, 1, 1) = "1" Then
                value = GetAttribute(GetMessageCellById("0213"), "Msg")
            Else
                value = doiSoThanhChu(Mid(so, 1, 1))
            End If
            doiTien = doiTien + value
            If (Len(so) Mod 3 = 0) Then
                soTam2 = so
            End If
            If (Len(so) Mod 3 = 0) And Mid(so, 1, 1) <> "0" Then
                doiTien = doiTien + GetAttribute(GetMessageCellById("0214"), "Msg")
            ElseIf Len(so) Mod 3 = 2 And Mid(so, 1, 1) <> "0" And Mid(so, 1, 1) <> "1" Then
                doiTien = doiTien + GetAttribute(GetMessageCellById("0219"), "Msg")
            ElseIf Len(so) Mod 3 = 2 And Mid(so, 1, 1) = "0" And Mid(so, 2, 1) <> "0" Then
                doiTien = doiTien + GetAttribute(GetMessageCellById("0220"), "Msg")
            ElseIf Len(so) Mod 3 = 1 And Len(so) <> 1 And Val(Mid(soTam2, 1, 3)) <> 0 Then
                doiTien = doiTien + giaTriTien(so)
            End If
            so = Mid(so, 2, Len(so) - 1)
    Next
    doiTien = Trim(doiTien) + GetAttribute(GetMessageCellById("0221"), "Msg")
    doiTien = UCase(Left(doiTien, 1)) + Mid(doiTien, 2)
End Function
'xac dinh menh gia tien
Private Function giaTriTien(ByVal tien As String) As String
    If Len(tien) < 4 Then
        giaTriTien = GetAttribute(GetMessageCellById("0214"), "Msg")
    ElseIf Len(tien) >= 4 And Len(tien) < 7 Then
        giaTriTien = GetAttribute(GetMessageCellById("0215"), "Msg")
    ElseIf Len(tien) >= 7 And Len(tien) < 10 Then
        giaTriTien = GetAttribute(GetMessageCellById("0216"), "Msg")
    ElseIf Len(tien) >= 10 And Len(tien) < 13 Then
        giaTriTien = GetAttribute(GetMessageCellById("0217"), "Msg")
    ElseIf Len(tien) = "13" Then
        giaTriTien = GetAttribute(GetMessageCellById("0218"), "Msg")
    End If
End Function


' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 16
    denRowDL = 21
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "20_30_2"
    ' end
End Sub

' ham set thong tin sau khi tai XML
Public Sub finishImportXML()
    doiTien2
End Sub
