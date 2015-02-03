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
'This Class is used to individual features of "Thu Nhap Doanh Nghiep" interface sheets
'this Class is belong to TAX_Business_v1 project which will be compline to DLL
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1


Dim vt As String, va As String, vb As String, vc As String
Dim strComboHien As String, strCombo As String
Dim arrTaiNguyen As New Scripting.Dictionary, mblKhoiDau As Boolean
Dim arrTaiNguyenGoc As New Scripting.Dictionary
Dim blKiemTraDieuChinhCu As Boolean
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
Dim strDVTCu As Variant, strMucThueCu As Variant
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
Private Const Row_MSTDN = 17
Private Const Col_MSTDN = "I"

Private Const SheetPL01_1 = 2
Private Const rowStart = 35
Private Const Col_CQT_Cuc = "M"
Private Const Col_Ma_CQT_Cuc = "Q"
Private Const Col_CQT_CCuc = "O"
Private Const Col_Ma_CQT_CCuc = "R"
Private Const Col_HH = "D"
Private Const Col_Ma_HH = "Y"



Public StrTKThang_PS As String

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
    Dim strDataFileName As String
    Dim xmlNodeListItems As MSXML.IXMLDOMNodeList
    Dim strCboTenHH As String
    Dim strCboMaHHAn As String
    
    Dim strCboTenHHThan As String
    Dim strCboMaHHAnThan As String
    
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim i As Integer
    Dim fldList() As String
    strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_DM_thue_BVMT.xml")
    
    With fps
        SetDateFormat fps, 1, 57, .ColLetterToNumber("R"), DDMMYYYY
        .Sheet = 1
        .Row = 57
        .Col = .ColLetterToNumber("R")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        SetDateFormat fps, 3, 31, .ColLetterToNumber("BE"), DDMMYYYY
        SetDateFormat fps, 3, 46, .ColLetterToNumber("BG"), DDMMYYYY
        .Sheet = 3
        .Row = 46
        .Col = .ColLetterToNumber("BG")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
    End With

    If xmlDomData.Load(strDataFileName) Then
      Set xmlNodeListItems = xmlDomData.getElementsByTagName("Item")
      For Each xmlNode In xmlNodeListItems
          fldList = Split(GetAttribute(xmlNode, "Value"), "###")
          strCboTenHH = strCboTenHH + fldList(1) + Chr$(9)
          strCboMaHHAn = strCboMaHHAn + fldList(0) + "~" + fldList(1) + "~" + fldList(2) + "~" + fldList(3) + Chr$(9)
          If Left$(Trim(fldList(0)), 4) = "0102" Then
                strCboTenHHThan = strCboTenHHThan + fldList(1) + Chr$(9)
                strCboMaHHAnThan = strCboMaHHAnThan + fldList(0) + "~" + fldList(1) + "~" + fldList(2) + "~" + fldList(3) + Chr$(9)
          End If
      Next
    End If

    
    With fps
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        .Col = .ColLetterToNumber("AB")
        .Row = 44
        .TypeComboBoxList = strCboMaHHAn
        'Combo Hien
        .Col = .ColLetterToNumber("C")
        .Row = 44
        .TypeComboBoxList = strCboTenHH
        
        .Sheet = SheetPL01_1
        .Col = .ColLetterToNumber("Y")
        .Row = rowStart
        .TypeComboBoxList = strCboMaHHAnThan
        'Combo Hien
        .Col = .ColLetterToNumber("D")
        .Row = rowStart
        .TypeComboBoxList = strCboTenHHThan
     End With
        
     setCQTQuanLy
     
                 ' set CQT ra quyet dinh hoan
    setCQTQuanLyHoanThue fps
        
End Sub
Private Sub setCQTQuanLy()
        Dim xmlDomData As New MSXML.DOMDocument
        Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
        Dim xmlNode As MSXML.IXMLDOMNode
        Dim arrDanhsach() As String
        Dim tenCucThue As String
        Dim maCucThue As String
        Dim tenChiCucThue As String
        Dim maChiCucThue As String
        
        Dim loaiCq As String
        Dim maLoaiCq As String
        
        If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    
                    If arrDanhsach(0) = "1" Then
                        'tenCucThue = tenCucThue + arrDanhsach(3) + Chr$(9)
                        tenCucThue = tenCucThue + Right$(arrDanhsach(3), Len(arrDanhsach(3)) - 6) + Chr$(9)
                        maCucThue = maCucThue + arrDanhsach(1) + Chr$(9)
                    End If
                    'If arrDanhsach(0) = "0" Then
                        tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
                        maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
                    'End If
                End If
            Next
            Set xmlDomData = Nothing
            Set xmlNodeListCell = Nothing
            Set xmlNode = Nothing
        End If
        
        ' setloai
        If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_DM_LoaiCq.xml")) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    loaiCq = loaiCq + arrDanhsach(1) + Chr$(9)
                    maLoaiCq = maLoaiCq + arrDanhsach(0) + Chr$(9)
                End If
            Next
            Set xmlDomData = Nothing
            Set xmlNodeListCell = Nothing
            Set xmlNode = Nothing
        End If
        
        With fps
            .Sheet = SheetPL01_1 'set thong tin CQL
            '.EventEnabled(EventAllEvents) = False
            
            .Row = rowStart
            .Col = .ColLetterToNumber(Col_CQT_Cuc)
            .TypeComboBoxList = tenCucThue
            
            .Row = rowStart
            .Col = .ColLetterToNumber(Col_Ma_CQT_Cuc)
            .TypeComboBoxList = maCucThue
            
            .Row = rowStart
            .Col = .ColLetterToNumber(Col_CQT_CCuc)
            .TypeComboBoxList = tenChiCucThue
    
            .Row = rowStart
            .Col = .ColLetterToNumber(Col_Ma_CQT_CCuc)
            .TypeComboBoxList = maChiCucThue
            '.EventEnabled(EventAllEvents) = True
        End With
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
         ' to khai thang, lan phat sinh
         If StrTKThang_PS = "TK_THANG" Then
            .Col = .ColLetterToNumber("L")
            .Row = 35
            .Text = 0
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("M")
            .Row = 35
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf StrTKThang_PS = "TK_LANPS" Then
            .Col = .ColLetterToNumber("M")
            .Row = 35
            .Text = TAX_Utilities_v1.Day & "/" & TAX_Utilities_v1.Month & "/" & TAX_Utilities_v1.Year
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("L")
            .Row = 35
            .Text = "1"
            UpdateCell fps, .Col, .Row, .Text
        End If
            
        ' to khai chinh thuc, bo sung
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
        
        .Sheet = 1
            mCurrentSheet = 1

            Set xmlCellNode = TAX_Utilities_v1.Data(0).nodeFromID(GetCellID(fps, .ColLetterToNumber("C"), 44))

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
        
    
        
    End With
    
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    Dim strId           As Variant
    Dim intIndexCombo   As Integer
    Dim txmlCellNode    As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
    Dim temp As Variant, varTemp As Variant
    Dim i               As Integer
    Dim mblLoi          As Boolean
    
    Dim arrHH() As String
    
    fps.EventEnabled(EventAllEvents) = False

    With fps
        If .ActiveSheet = 1 Then
            .Sheet = 1
            If Col = .ColLetterToNumber("C") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                .Col = .ColLetterToNumber("AB")
                .typeComboboxCurSel = intIndexCombo
    
                .Col = .ColLetterToNumber("AC")
                .GetText .ColLetterToNumber("AB"), Row, strId
                arrHH = Split(strId, "~")
                ' Set ma hang hoa
                .Text = arrHH(0)
                UpdateCell fps, .Col, Row, .Text
                ' set gia tri don vi tinh
                .Col = .ColLetterToNumber("J")
                .Text = arrHH(2)
                UpdateCell fps, .Col, Row, .Text
                ' set muc thue
                .Col = .ColLetterToNumber("Q")
                .Text = arrHH(3)
                UpdateCell fps, .Col, Row, .Text
                
                
    
            End If
            'Huy chon Combo to khai khong hop le va ra thong bao
            If blGiaTriCu And Col = .ColLetterToNumber("C") Then 'And Not mblKeyF5 Then
                If Not .Visible Then
                    .Visible = True
                    DisplayMessage "0261", msOKOnly, miWarning
                    .Visible = False
                Else
                    DisplayMessage "0261", msOKOnly, miWarning
                End If
                    
                .Sheet = 1
                mCurrentSheet = .Sheet
                ' Ten hang hoa
                .Row = RowGiaTriCu
                .Col = ColGiaTriCu
                .Text = strGiaTriCu
                UpdateCell fps, .Col, .Row, .Text
                ' Don vi tinh
                .Row = RowGiaTriCu
                .Col = .ColLetterToNumber("J")
                .Text = strDVTCu
                UpdateCell fps, .Col, .Row, .Text
                ' Muc thue cu
                .Row = RowGiaTriCu
                .Col = .ColLetterToNumber("Q")
                .Text = strMucThueCu
                UpdateCell fps, .Col, .Row, .Text
                
                
                blGiaTriCu = False
                mblLoi = True
            End If
         ElseIf .ActiveSheet = 2 Then
            .Sheet = 2
            If Col = .ColLetterToNumber("D") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                .Col = .ColLetterToNumber("Y")
                .typeComboboxCurSel = intIndexCombo
    
                .Col = .ColLetterToNumber("Z")
                .GetText .ColLetterToNumber("Y"), Row, strId
                arrHH = Split(strId, "~")
                ' Set ma hang hoa
                .Text = arrHH(0)
                UpdateCell fps, .Col, Row, .Text
               
                ' set muc thue
                .Col = .ColLetterToNumber("W")
                .Text = arrHH(3)
                UpdateCell fps, .Col, Row, .Text
                
                
    
            End If
            'Huy chon Combo to khai khong hop le va ra thong bao
            If blGiaTriCu And Col = .ColLetterToNumber("D") Then 'And Not mblKeyF5 Then
                If Not .Visible Then
                    .Visible = True
                    DisplayMessage "0261", msOKOnly, miWarning
                    .Visible = False
                Else
                    DisplayMessage "0261", msOKOnly, miWarning
                End If
                    
                .Sheet = 2
                mCurrentSheet = .Sheet
                ' Ten hang hoa
                .Row = RowGiaTriCu
                .Col = ColGiaTriCu
                .Text = strGiaTriCu
                UpdateCell fps, .Col, .Row, .Text
            
                
                ' Muc thue cu
                .Row = RowGiaTriCu
                .Col = .ColLetterToNumber("W")
                .Text = strMucThueCu
                UpdateCell fps, .Col, .Row, .Text
                
                
                blGiaTriCu = False
                mblLoi = True
            End If
         End If
    End With

    fps.EventEnabled(EventAllEvents) = True
    setChangeComb Col, Row
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
            
            fps.Col = fps.ColLetterToNumber("J")
            fps.Row = iRow
            strDVTCu = fps.Text
            
            fps.Col = fps.ColLetterToNumber("Q")
            fps.Row = iRow
            strMucThueCu = fps.Text
            
            
        End If
    End If
End With
End Sub


Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim temp As Variant, itemp As Variant, iRow As Long, iCol As Long
Dim intIndexCombo As Integer
Dim i As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode


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
    If .ActiveSheet = 1 Then
              ' Kiem tra trung ten
'           Set xmlCellNode = TAX_Utilities_v1.Data(0).nodeFromID(GetCellID(fps, Col, Row))
'           Set xmlCellNode = xmlCellNode.parentNode.parentNode.firstChild.firstChild
'           ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'           blGiaTriCu = False
'
'            If Col = fps.ColLetterToNumber("C") Then
'                'kiem tra trung Ten NNKD trong to khai
'                fps.Col = Col
'                fps.Row = Row
'                temp = fps.Text
'                Set xmlCellsNode = xmlCellNode.parentNode
'                i = 1
'                .Col = .ColLetterToNumber("B")
'                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
'                .Row = iRow
'
'                Do
'                     .Col = .ColLetterToNumber("C")
'                     itemp = .Text
'                     If .Row <> Row Then
'                        If itemp = temp Then 'Neu co thi dat co blGiaTriCu = True
'                            RowGiaTriCu = Row
'                            ColGiaTriCu = .ColLetterToNumber("C")
'                            .Row = Row
'                            .Col = .ColLetterToNumber("C")
'                            blGiaTriCu = True
'                            Exit Sub
'                        End If
'                     End If
'                     .Col = .ColLetterToNumber("B")
'                     .Row = i + iRow
'                     i = i + 1
'                Loop Until .Text = "aa"
'                .Sheet = 1
'                mCurrentSheet = 1
'            End If
     ElseIf .ActiveSheet = SheetPL01_1 Then
        .Col = Col
        .Row = Row
        
         If Col = .ColLetterToNumber(Col_CQT_Cuc) And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber(Col_Ma_CQT_Cuc)
            .typeComboboxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
        End If
        
         If Col = .ColLetterToNumber(Col_CQT_CCuc) And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber(Col_Ma_CQT_CCuc)
            .typeComboboxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        If Col = .ColLetterToNumber(Col_HH) And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber(Col_Ma_HH)
            .typeComboboxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
        End If
    ElseIf .ActiveSheet = .SheetCount - 1 Then
                    Set xmlCellNode = TAX_Utilities_v1.Data(.SheetCount - 2).nodeFromID(GetCellID(fps, .ColLetterToNumber("BI"), 5))
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode.lastChild
            Set xmlCellNode = xmlCellNode.firstChild
            Set xmlCellNode = xmlCellNode.firstChild
            For i = 1 To 6
                Set xmlCellNode = xmlCellNode.nextSibling
            Next i
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
   
   
            .Sheet = .SheetCount - 1
            mCurrentSheet = .SheetCount - 1
            
            .Col = Col
            .Row = Row
            
            If Col = .ColLetterToNumber("BE") And Row = iRow Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
                intIndexCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("BI")
                .Row = Row
                .typeComboboxCurSel = intIndexCuc
                'Nhap lai Combo cot Q
                .Col = .ColLetterToNumber("BI")
                UpdateCell fps, .Col, .Row, .Text
                ' Lay Id cua cuc thue vua chon o tren
                .GetText .Col, .Row, varIdCucThue
                
                If varIdCucThue <> "" Or varIdCucThue <> vbNullString Then
                    ' Lay thong tin cho chi cuc thue
                    varIdCucThue = Left$(varIdCucThue, 3)
                    
                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
                        For Each xmlNode In xmlNodeListCell
                            If GetAttribute(xmlNode, "Value") <> "" Then
                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                                
                                If arrDanhsach(0) = "0" And arrDanhsach(2) = varIdCucThue Then
                                    tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
                                    maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
                                End If
                            End If
                        Next
                        Set xmlDomData = Nothing
                        Set xmlNodeListCell = Nothing
                        Set xmlNode = Nothing
                    End If
                    
                    .Row = iRow + 2
                    .Col = .ColLetterToNumber("BE")
                    .TypeComboBoxList = tenChiCucThue
                    
                    .Row = iRow + 2
                    .Col = .ColLetterToNumber("BI")
                    .TypeComboBoxList = maChiCucThue
                End If
            End If
                        
            If Col = .ColLetterToNumber("BE") And Row = iRow + 2 Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
                intIndexChiCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("BI")
                .Row = Row
                .typeComboboxCurSel = intIndexChiCuc
                'Nhap lai Combo cot Q
                .Col = .ColLetterToNumber("BI")
                UpdateCell fps, .Col, .Row, .Text
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
                fps.EventEnabled(EventAllEvents) = False
                i = 1
                .Col = .ColLetterToNumber("B")
                .Row = 44

                Do
                    .Text = str(i)
                    .Col = .ColLetterToNumber("B")
                    .Row = i + 44
                    i = i + 1
                Loop Until .Text = "aa"

                fps.EventEnabled(EventAllEvents) = True
            End If

            If .ActiveSheet = 2 Then
                fps.EventEnabled(EventAllEvents) = False
                    i = 1
                    .Col = .ColLetterToNumber("C")
                    .Row = rowStart
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("C")
                         .Row = i + rowStart
                         i = i + 1
                    Loop Until .Text = "aa"
                    fps.EventEnabled(EventAllEvents) = True
                    .SetActiveCell iCol, iRow
            End If
            
        End If

    End With

End Sub
Private Sub setChangeComb(ByVal Col As Long, Row As Long)
    Dim varTemp As Variant
    Dim intIndexCuc As Integer
    Dim intIndexChiCuc As Integer
    Dim varIdCucThue As Variant
    
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim arrDanhsach() As String
    Dim tenChiCucThue As String
    Dim maChiCucThue As String
    
  '  Dim checkDL As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        If .ActiveSheet = SheetPL01_1 Then
            .Sheet = SheetPL01_1
            mCurrentSheet = SheetPL01_1
            
            .Col = Col
            .Row = Row
            
            If Col = .ColLetterToNumber(Col_CQT_Cuc) Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot G va gia tri Index cot S
                intIndexCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber(Col_Ma_CQT_Cuc)
                .Row = Row
                .typeComboboxCurSel = intIndexCuc
                'Nhap lai Combo ma cuc
                .Col = .ColLetterToNumber(Col_Ma_CQT_Cuc)
                UpdateCell fps, .Col, .Row, .Text
                ' Lay Id cua cuc thue vua chon o tren
                fps.GetText .Col, .Row, varIdCucThue
                
                'reset chi cuc
                .Col = .ColLetterToNumber(Col_CQT_CCuc)
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                
                .Col = .ColLetterToNumber(Col_Ma_CQT_CCuc)
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                
                If varIdCucThue <> "" Or varIdCucThue <> vbNullString Then
                    ' Lay thong tin cho chi cuc thue
                    varIdCucThue = Left$(varIdCucThue, 3)
                    
                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
                        For Each xmlNode In xmlNodeListCell
                            If GetAttribute(xmlNode, "Value") <> "" Then
                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                                
                                If arrDanhsach(0) = "0" And arrDanhsach(2) = varIdCucThue Then
                                    tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
                                    maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
                                End If
                            End If
                        Next
                        Set xmlDomData = Nothing
                        Set xmlNodeListCell = Nothing
                        Set xmlNode = Nothing
                    End If
                    
                    .Row = Row
                    .Col = .ColLetterToNumber(Col_CQT_CCuc)
                    .TypeComboBoxList = tenChiCucThue
                    
                    .Row = Row
                    .Col = .ColLetterToNumber(Col_Ma_CQT_CCuc)
                    .TypeComboBoxList = maChiCucThue
                   
                End If
            End If
                        
            If Col = .ColLetterToNumber(Col_CQT_CCuc) Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot L va gia tri Index cot R
                intIndexChiCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber(Col_Ma_CQT_CCuc)
                .Row = Row
                .typeComboboxCurSel = intIndexChiCuc
                'Nhap lai Combo cot R
                .Col = .ColLetterToNumber(Col_Ma_CQT_CCuc)
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            If Col = .ColLetterToNumber(Col_HH) Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot L va gia tri Index cot R
                intIndexChiCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber(Col_Ma_HH)
                .Row = Row
                .typeComboboxCurSel = intIndexChiCuc
                'Nhap lai Combo cot R
                .Col = .ColLetterToNumber(Col_Ma_HH)
                UpdateCell fps, .Col, .Row, .Text
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim CellNguoiKy As String
    Dim tongNTNN As Variant, tienDoiThanhChu As String
    
    CellNguoiKy = GetAttribute(TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")(TAX_Utilities_v1.Data(0).getElementsByTagName("Cell").length - 1), "CellID")
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
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
         .Row = 44
         Do
              .Text = str(i)
              .Col = .ColLetterToNumber("B")
              .Row = i + 44
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
Dim errMST, errEmty, errMSTDN, errMst10, mErrMSTTrung As Variant
Dim vMST, vMSTCheckTrung As Variant
Dim iFlagCheckMaSTDN As Integer, iFlagCheckMaSTDN10 As Integer, iFlagCheckMaST As Integer, iFlagMSTTrung As Integer
Dim rowDes As Long
Dim blCheck_S2 As Boolean

Dim vMsgErrorTongTyLe As Variant

Dim countRow As Integer
    With fps
      .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 15, Error
        .GetText .ColLetterToNumber("E"), 18, errMST
        .GetText .ColLetterToNumber("E"), 17, errEmty
        .GetText .ColLetterToNumber("E"), 19, errMSTDN
        .GetText .ColLetterToNumber("E"), 20, errMst10
        .GetText .ColLetterToNumber("E"), 21, mErrMSTTrung
        
        .GetText .ColLetterToNumber("E"), 22, vMsgErrorTongTyLe
      .Sheet = 1
      Dim i As Integer
     i = 1
     .Row = 44
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
      .Row = 44 + i
      i = i + 1
      countRow = countRow + 1
      Loop Until .Text = "aa"
      
      
      
         .Sheet = 2
        If .SheetVisible = True Then
            mCurrentSheet = .Sheet
            Dim strTenDN As Variant, strMaST As Variant, strValue As Variant
            Dim strHangHoa, strCQTCapCuc, strCQTCapCC, strSoThanTTNoiDia, strTongSoThanTT, strSanLuongThanDP, strMucThue As Variant
            Dim iFlagEmpty As Integer
            Dim strKyHieu As Variant, strCheckMST As String
            Dim strTongTLPhanBo As Variant
            
            Dim flagTongTyLe As Integer
            Const colorCell = 16709097
            
            Dim iFlagCheckData As Boolean
            
            Dim iWhiteRow As Integer
            iWhiteRow = 0
            
            '.GetText .ColLetterToNumber("U"), 30, strTongTLPhanBo
            
            
'            ' kiem tra tong ty le phan bo phai bang 100
'            If Val(strTongTLPhanBo) <> 100 And Val(strTongTLPhanBo) <> 0 Then
'                .Col = .ColLetterToNumber("U")
'                .Row = 35
'                .BackColor = mErrorColor
'                .CellNote = vMsgErrorTongTyLe
'                fLagTongTyLe = 1
'            Else
'                .Col = .ColLetterToNumber("U")
'                .Row = 35
'                .BackColor = mNonErrorColor
'                .CellNote = ""
'            End If
            
            iFlagEmpty = 0
            'check not empty ten doanh nghiep
            .Row = rowStart
            Do
                .GetText .ColLetterToNumber("D"), .Row, strHangHoa
                .GetText .ColLetterToNumber("F"), .Row, strTenDN
                .GetText .ColLetterToNumber("L"), .Row, strMaST
                .GetText .ColLetterToNumber("M"), .Row, strCQTCapCuc
                .GetText .ColLetterToNumber("O"), .Row, strCQTCapCC
                .GetText .ColLetterToNumber("S"), .Row, strSoThanTTNoiDia
                .GetText .ColLetterToNumber("T"), .Row, strTongSoThanTT
                .GetText .ColLetterToNumber("V"), .Row, strSanLuongThanDP
                .GetText .ColLetterToNumber("W"), .Row, strMucThue
                If Trim$(strHangHoa) <> "" Or Trim$(strTenDN) <> "" Or Trim$(strMaST) <> "" Or Trim$(strCQTCapCuc) <> "" Or Trim$(strCQTCapCC) <> "" _
                Or Val(strSoThanTTNoiDia) <> 0 Or Val(strTongSoThanTT) <> 0 Or Val(strSanLuongThanDP) <> 0 Or Val(strMucThue) <> 0 Then iWhiteRow = 1
            
                If iWhiteRow <> 0 Then
                
                    .Col = .ColLetterToNumber("F")
                    strValue = .Text
                    'check du lieu bat buoc phai nhap
                    If (strValue = "") Then
                        .Col = .ColLetterToNumber("F")
                        .BackColor = mErrorColor
                        .CellNote = errEmty
                        iFlagEmpty = 1
                    Else
                        .Col = .ColLetterToNumber("F")
                        .BackColor = mNonErrorColor
                        .CellNote = ""
                    End If
                    'lay du lieu nhap vao
                
                    
                    ' clear cell not
                    .Col = .ColLetterToNumber("L")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
        
                    
                    .Col = .ColLetterToNumber("L")
                    vMST = .Text
                    
                    vMSTCheckTrung = vMST
                    
                    If vMST <> vbNullString Then
                        vMST = Replace(vMST, " ", "")
                        'check MST 14 ky tu
                        If Len(vMST) = 14 And Mid(vMST, 11, 1) = "-" Then
                            vMST = Replace(vMST, "-", "")
                        End If
                        strCheckMST = ""
                        If Len(vMST) = 10 Or Len(vMST) = 13 Then
                            strCheckMST = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
'                            If Len(vMST) = 10 And strCheckMST = "000" Then
'                                .Col = .ColLetterToNumber("F")
'                                .BackColor = mErrorColor
'                                .CellNote = errMst10
'                                iFlagCheckMaSTDN10 = 1
'                            Else
                            If Len(vMST) = 10 And strCheckMST <> "000" Then
                                .Col = .ColLetterToNumber("L")
                                .BackColor = mErrorColor
                                .CellNote = errMST
                                iFlagCheckMaST = 1
                            ElseIf Len(vMST) = 13 And (strCheckMST <> "000" Or IsNumeric(vMST) = False) Then
                                .Col = .ColLetterToNumber("L")
                                .BackColor = mErrorColor
                                .CellNote = errMST
                                iFlagCheckMaST = 1
                            Else
                                .Col = .ColLetterToNumber("L")
                                .BackColor = mNonErrorColor
                                .CellNote = ""
                                'iFlagCheckMaST = 0
                            End If
                            
                            
                            
                        ElseIf Len(vMST) <> 10 Or Len(vMST) <> 13 Then
                               .Col = .ColLetterToNumber("L")
                                .BackColor = mErrorColor
                                .CellNote = errMST
                                iFlagCheckMaST = 1
                        
                        End If
                        
                        
                        
                        'check MST trung nhau
                         countRow = .SearchCol(.ColLetterToNumber("L"), -1, -1, vMSTCheckTrung, SearchFlagsNone)
                         rowDes = .SearchCol(.ColLetterToNumber("L"), countRow, -1, vMSTCheckTrung, SearchFlagsNone)
                        If rowDes > -1 Then
                            .Col = .ColLetterToNumber("L")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = mErrMSTTrung      'static
                            .BackColor = mErrorColor
                            iFlagMSTTrung = 1
                        End If
                    Else
                             .Col = .ColLetterToNumber("L")
        '                    .BackColor = mNonErrorColor
        '                    .CellNote = ""
                            'iFlagCheckMaST = 0
                            .BackColor = mErrorColor
                            .CellNote = errEmty
                            iFlagEmpty = 1
                    End If
                    
                    
                    'danh dau neu dong co chua du lieu
                    iFlagCheckData = True
                End If
            
            .Col = .ColLetterToNumber("C")
            .Row = .Row + 1
            Loop Until .Text = "aa"
        End If
 
        
         'neu dong co chua du lieu thi kiem tra cac rang buoc trong cac chi tieu
        If iFlagCheckData Then
           'lay tong ty le chia TNCT
            .GetText .ColLetterToNumber("U"), .MaxRows - 12, strTongTLPhanBo
            'check tong TNCT phai = 100
            If Val(strTongTLPhanBo) <> 100 Then
                .Col = .ColLetterToNumber("U")
                .Row = .MaxRows - 12
                .BackColor = mErrorColor
                .CellNote = vMsgErrorTongTyLe
                flagTongTyLe = 1
            Else
                .Col = .ColLetterToNumber("U")
                .Row = .MaxRows - 12
                .BackColor = colorCell
                .CellNote = ""
            End If
        End If
 
        
     blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        'If blCheck_S2 Then
        If .SheetVisible = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 20, IIf(iFlagCheckMaSTDN10 = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 17, IIf(iFlagEmpty = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 18, IIf(iFlagCheckMaST = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 21, IIf(iFlagMSTTrung = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 22, IIf(flagTongTyLe = 1, "0", "1")
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 20, "1"
            .SetText .ColLetterToNumber("B"), 17, "1"
            .SetText .ColLetterToNumber("B"), 18, "1"
            .SetText .ColLetterToNumber("B"), 21, "1"
            .SetText .ColLetterToNumber("B"), 22, "1"
        End If
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
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "01-1"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
                    If lCol = fps.ColLetterToNumber("F") Or _
                       lCol = fps.ColLetterToNumber("L") Then
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
        
             Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(2).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(2).getElementsByTagName("Cell").length - 1)
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
'        .Col = .ColLetterToNumber("Q")
'        .Row = 13
'        .Text = ""
'        UpdateCell fps, .Col, .Row, .Text
'
'        .Col = .ColLetterToNumber("H")
'        .Row = 15
'        .Text = ""
'        UpdateCell fps, .Col, .Row, .Text
        

'          a = (TAX_Utilities_v1.Data(0).getElementsByTagName("Cell").length - 12) / 6

        For Each xmlNodeReset In TAX_Utilities_v1.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            fps.Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            fps.Col = lCol
            fps.Row = lRow

                If .ActiveSheet = 1 Then
                    .Sheet = mCurrentSheet
                    If (lRow < 40 Or lRow > .MaxRows - 10) Then
                        GoTo next1
                    End If
                ElseIf .ActiveSheet = 3 Then
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


                
            TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = IIf(IsUpdate = True, IsUpdate, TAX_Utilities_v1.AdjustData(mCurrentSheet - 1))
next1:
        Next
       mCurrentSheet = sheetTemp
    End With

End Function
Public Sub SetData()
   Dim varTemp As Variant
    Dim ssSheet As Integer
    Dim lCol As Long, lRow As Long
    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 2
        mCurrentSheet = .Sheet
        .Row = 39
        .Col = .ColLetterToNumber("C")
        .Text = 2
        UpdateCell fps, .Col, .Row, .Text

      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
End Sub

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

Public Sub Prepare_KHBS()
 
    
End Sub


Public Sub UpdateChangeKHBS()

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
    SaveErrorLog "cls_01TBVMT", "getValueTK", Err.number, Err.Description
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
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT_TT28.xml"
    
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "ID") = GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") Then
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
    SaveErrorLog "01_TBVMT", "getMapCT", Err.number, Err.Description
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
    SaveErrorLog "cls_01TBVMT", "getValueCTDC", Err.number, Err.Description
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
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
                Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(2).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(2).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(2).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(2).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v1.Data(2).getElementsByTagName("Cell").length > 31 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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
         For Each xmlCellNodeData In TAX_Utilities_v1.Data(.Sheet - 1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .Col = lCol
               .Row = lRow
               If .CellType = CellTypeNumber Then
                    varTemp = .value
               Else
                    varTemp = .Text
               End If
               UpdateCell fps, lCol, lRow, varTemp
         Next
         
         .Sheet = 2
        mCurrentSheet = .Sheet
        blCheckS2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheckS2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(.Sheet - 1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .Col = lCol
                    .Row = lRow
                    If .CellType = CellTypeNumber Then
                         varTemp = .value
                    Else
                         varTemp = .Text
                    End If
                    UpdateCell fps, lCol, lRow, varTemp
            Next
        End If

        
       ' Update lai sheet KHBS
        .Sheet = .SheetCount - 1
        mCurrentSheet = .Sheet
        blCheckS2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheckS2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(.Sheet - 1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .Col = lCol
                     .Row = lRow
                     If .CellType = CellTypeNumber Then
                          varTemp = .value
                     Else
                          varTemp = .Text
                     End If
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

' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 16
    denRowDL = 21
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "21_31_2~20_30_3"
    ' end
End Sub


' Lay ve ky tu ngan cach so thap phan va so hang nghin
Private Sub test()
    Dim LocaleDecimal As String
    Dim LocaleComma As String
    LocaleDecimal = Mid$(CStr(11 / 10), 2, 1)
    LocaleComma = Chr$(90 - Asc(LocaleDecimal))
End Sub



Public Sub fillCQT()
    Dim CQT_CAPCUC    As Variant
    Dim CQT_CAPCHICUC  As Variant
    Dim tCQT_CAPCUC   As String
    Dim tCQT_CAPCHICUC As String
    With fps
        .Sheet = 2
        If .SheetVisible = True Then
                .Row = 35
                Do
                     .GetText .ColLetterToNumber("Q"), .Row, CQT_CAPCUC
                     .GetText .ColLetterToNumber("R"), .Row, CQT_CAPCHICUC
                     GetCQTByMaCQT CQT_CAPCUC, tCQT_CAPCUC
                     GetCQTByMaCQT CQT_CAPCHICUC, tCQT_CAPCHICUC
                     .Col = .ColLetterToNumber("M")
                     .Text = tCQT_CAPCUC
                     UpdateCell fps, .Col, .Row, .Text
                     
                     .Col = .ColLetterToNumber("O")
                     .Text = tCQT_CAPCHICUC
                     UpdateCell fps, .Col, .Row, .Text
                     
                     .Col = .ColLetterToNumber("C")
                    .Row = .Row + 1
                Loop Until .Text = "aa"
        End If
    End With
End Sub



' ham set thong tin sau khi tai XML
Public Sub finishImportXML()
    Dim maNhom    As Variant
    Dim tenNhom  As String
    With fps
        .EventEnabled(EventAllEvents) = False
          .Sheet = 1
          mCurrentSheet = .Sheet
            .Row = 44
            Do
                .GetText .ColLetterToNumber("AC"), .Row, maNhom
                 getNhomHH maNhom, tenNhom
                .Col = .ColLetterToNumber("C")
                .Text = tenNhom
                UpdateCell fps, .Col, .Row, .Text
                
                 .Col = .ColLetterToNumber("B")
                 .Row = .Row + 1
            Loop Until .Text = "aa"
        .EventEnabled(EventAllEvents) = True
    End With
End Sub


'
Public Sub getNhomHH(ByVal maNhom As String, ByRef tenNhom As String)
    Dim arrDanhsach() As String
    Dim strDataFileName As String
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode

       strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_DM_thue_BVMT.xml")
    
       If xmlDomData.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If maNhom = arrDanhsach(0) Then
                            tenNhom = arrDanhsach(1)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub
