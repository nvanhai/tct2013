VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_16TH_DKNPT"
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

Public flgLoadData16TH As Boolean

Private flgErr As Boolean

' Const ctieu
Private Const Col_D = "D"
Private Const Col_E = "E"
Private Const Col_F = "F"
Private Const Col_G = "G"

Private Const sheetTk = 1

' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

Private rowData As Long

Private arrStrXMLFileNames() As String
Private arrData05_3() As String
Public rowGroup1 As Long
Public rowGroup2 As Long
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    Dim vCell As Variant
    With fps
    
        SetDateFormat fps, sheetTk, 18, .ColLetterToNumber("Z"), DDMMYYYY
        ' PL 16TH
        SetDateFormat fps, sheetTk, 40, .ColLetterToNumber("W"), MMYYYY
        SetDateFormat fps, sheetTk, 40, .ColLetterToNumber("X"), MMYYYY
        SetDateFormat fps, sheetTk, 40, .ColLetterToNumber("F"), DDMMYYYY
        
        SetDateFormat fps, sheetTk, 49, .ColLetterToNumber("W"), MMYYYY
        SetDateFormat fps, sheetTk, 49, .ColLetterToNumber("X"), MMYYYY
        SetDateFormat fps, sheetTk, 49, .ColLetterToNumber("F"), DDMMYYYY
        
        .Sheet = sheetTk
        .Col = .ColLetterToNumber("Z")
        .Row = 18
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        .Sheet = sheetTk
        ' Lay danh muc quan he voi NNT
        Dim xmlDocument As New MSXML.DOMDocument
        Dim xmlNode As MSXML.IXMLDOMNode
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_QHNNT.xml"))
        Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
        Dim strQuanHeNNTID As String
        Dim strQuanHeNNT As String
        
        Dim strQuocTich As String
        Dim strMaQuocTich As String
        Dim strTenTinh As String
        Dim strTenHuyen As String
        Dim strTenXa As String
        Dim strMaTinh As String
        Dim strMaHuyen  As String
        Dim strMaXa As String
        Dim temValue As String
        Dim arrValue() As String
        
        ' Chuoi gia tri ve quan he NNT bao gom:
        ' 01-Con, 02-Vo, 03-Chong, 04-Cha, 05-Me, 06-Chau Ruot, 07-Ho Hang, 99-Khac
        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListMap
            arrValue = Split(GetAttribute(xmlNode, "value"), "###")
            strQuanHeNNT = strQuanHeNNT + arrValue(1) + Chr$(9)
            strQuanHeNNTID = strQuanHeNNTID + arrValue(0) + Chr$(9)
        Next
        
        .Row = 40
        .Col = .ColLetterToNumber("L")
        .TypeComboBoxList = strQuanHeNNT
        
        .Col = .ColLetterToNumber("K")
        .TypeComboBoxList = strQuanHeNNTID
        
        
        .Row = 49
        .Col = .ColLetterToNumber("L")
        .TypeComboBoxList = strQuanHeNNT
        
        .Col = .ColLetterToNumber("K")
        .TypeComboBoxList = strQuanHeNNTID
        
        ' lay du lieu danh muc dia ban tinh
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Tinh.xml"))
        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListMap
            temValue = GetAttribute(xmlNode, "Value")
            arrValue = Split(temValue, "###")
            'If arrValue(0) <> "1" Then
                strTenTinh = strTenTinh + arrValue(2) + Chr$(9)
                strMaTinh = strMaTinh + arrValue(1) + Chr$(9)
            'End If
        Next
        
        ' lay danh sach huyen
        Set xmlDocument = Nothing
        Set xmlNodeListMap = Nothing
        ' lay du lieu danh muc dia ban tinh
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Huyen.xml"))
        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListMap
            temValue = GetAttribute(xmlNode, "Value")
            arrValue = Split(temValue, "###")
            'If arrValue(0) <> "1" Then
                strTenHuyen = strTenHuyen + arrValue(3) + Chr$(9)
                strMaHuyen = strMaHuyen + arrValue(2) + Chr$(9)
            'End If
        Next
        
        ' Lay danh muc xa
        Set xmlDocument = Nothing
        Set xmlNodeListMap = Nothing
        ' lay du lieu danh muc dia ban tinh
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Xa.xml"))
        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListMap
            temValue = GetAttribute(xmlNode, "Value")
            arrValue = Split(temValue, "###")
            'If arrValue(0) <> "1" Then
                strTenXa = strTenXa + arrValue(4) + Chr$(9)
                strMaXa = strMaXa + arrValue(3) + Chr$(9)
            'End If
        Next
        
        .Row = 40
        ' set danh sach tinh
        .Col = .ColLetterToNumber("R")
        .TypeComboBoxList = strTenTinh
        .Col = .ColLetterToNumber("Q")
        .TypeComboBoxList = strMaTinh
        ' set danh sach huyen
        .Col = .ColLetterToNumber("T")
        .TypeComboBoxList = strTenHuyen
        .Col = .ColLetterToNumber("S")
        .TypeComboBoxList = strMaHuyen
        ' set danh sach xa
        .Col = .ColLetterToNumber("V")
        .TypeComboBoxList = strTenXa
        .Col = .ColLetterToNumber("U")
        .TypeComboBoxList = strMaXa
        
        
'        .Row = 49
'        ' set danh sach tinh
'        .Col = .ColLetterToNumber("R")
'        .TypeComboBoxList = strTenTinh
'        .Col = .ColLetterToNumber("Q")
'        .TypeComboBoxList = strMaTinh
'        ' set danh sach huyen
'        .Col = .ColLetterToNumber("T")
'        .TypeComboBoxList = strTenHuyen
'        .Col = .ColLetterToNumber("S")
'        .TypeComboBoxList = strMaHuyen
'        ' set danh sach xa
'        .Col = .ColLetterToNumber("V")
'        .TypeComboBoxList = strTenXa
'        .Col = .ColLetterToNumber("U")
'        .TypeComboBoxList = strMaXa
        
        
        ' lay du lieu dia ban
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Quoc_Gia.xml"))
        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListMap
            temValue = GetAttribute(xmlNode, "Value")
            arrValue = Split(temValue, "###")
            strQuocTich = strQuocTich + arrValue(1) + Chr$(9)
            strMaQuocTich = strMaQuocTich + arrValue(0) + Chr$(9)
        Next
        
        .Row = 40
        ' set danh sach quoc tich
        .Col = .ColLetterToNumber("I")
        .TypeComboBoxList = strQuocTich
        '.typeComboboxCurSel = 0
        .Col = .ColLetterToNumber("H")
        .TypeComboBoxList = strMaQuocTich
        '.typeComboboxCurSel = 0
        
        ' set danh sach quoc gia
        .Col = .ColLetterToNumber("P")
        .TypeComboBoxList = strQuocTich
        '.typeComboboxCurSel = 0
        .Col = .ColLetterToNumber("O")
        .TypeComboBoxList = strMaQuocTich
        '.typeComboboxCurSel = 0
        
        
'        .Row = 49
'        ' set danh sach quoc tich
'        .Col = .ColLetterToNumber("I")
'        .TypeComboBoxList = strQuocTich
'        .typeComboboxCurSel = 0
'        .Col = .ColLetterToNumber("H")
'        .TypeComboBoxList = strMaQuocTich
'        .typeComboboxCurSel = 0
'
'        ' set danh sach quoc gia
'        .Col = .ColLetterToNumber("P")
'        .TypeComboBoxList = strQuocTich
'        .typeComboboxCurSel = 0
'        .Col = .ColLetterToNumber("O")
'        .TypeComboBoxList = strMaQuocTich
'        .typeComboboxCurSel = 0


    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim viTriMSTDL As String
    
     With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        
        UpdateCell fps, .Col, .Row, .Text
        .Col = .ColLetterToNumber("G")
        .Row = 6
        .Text = StrSolanBosung
        UpdateCell fps, .Col, .Row, .Text
                    

        
        'check thong tin ma so thue dai ly
        'set vi tri theo thu tu MSTDL~TENNVDLT~ChungChiHN
        viTriMSTDL = "D_20~AA_16~AA_18"
        updateMSTDL fps, viTriMSTDL
        
        .Sheet = sheetTk
        .ActiveSheet = .Sheet
        .Row = 1
        .Col = .ColLetterToNumber("C")
        .CellType = CellTypeButton
        .TypeButtonTextColor = RGB(0, 0, 255)
        .TypeButtonColor = &H8000000F
        .TypeButtonAlign = TypeButtonAlignRight
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
        
        
        .Col = .ColLetterToNumber("D")
        .CellType = CellTypeButton
        .TypeButtonTextColor = RGB(0, 0, 255)
        .TypeButtonColor = &H8000000F
        .TypeButtonAlign = TypeButtonAlignRight
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonText = GetAttribute(GetMessageCellById("0346"), "Msg")
        .EventEnabled(EventAllEvents) = True
    End With
    
         prepareData05_3
         
    ' set lai format PL16TH
    If flgLoadData16TH = True Then
        Dim idx As Long
        Dim countRowG1 As Long
         ForMatTK16TH
         
         With fps
              .Sheet = 1
            Debug.Print "Tao node XML " & Time
            InsertNode16TH 41, .SearchCol(.ColLetterToNumber("B"), 40, -1, "aa", SearchFlagsNone) - 1, 1, True
            
            ' tinh so dong group 1
            .Sheet = 1
            .Col = .ColLetterToNumber("B")
            For idx = 40 To .MaxRows Step 1
                .Row = idx
                countRowG1 = countRowG1 + 1
                If .Text = "aa" Then
                    Exit For
                End If
            Next
    
            
            InsertNode16TH 48 + countRowG1, .SearchCol(.ColLetterToNumber("B"), 47 + countRowG1, -1, "aa", SearchFlagsNone) - 1, 1, False
            Debug.Print "End Tao node XML " & Time
            UpdateSheets
        End With
    End If
End Sub

Private Function GetValidatedCellIndex(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long) As Integer
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long
    
    Set xmlCellNode = TAX_Utilities_v2.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
    Set xmlCellsNode = xmlCellNode.parentNode
    
    'Get Index of anchor cell
    For lCtrl = 1 To xmlCellsNode.childNodes.length
        If GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID") = GetAttribute(xmlCellNode, "CellID") Then _
            Exit For
    Next
    GetValidatedCellIndex = lCtrl
    
End Function

Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    With fps
        If Col = .ColLetterToNumber("D") And Row = 1 Then
           fillData
           UpdateSheets
        End If
    End With
    
End Sub


Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
Dim intIndexCombo As Integer
    
    Dim varIDTinh As Variant
    Dim varIDHuyen As Variant
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim i As Integer
    
    Dim tenHuyen As String
    Dim maHuyen As String
    Dim tenXa As String
    Dim maXa As String
    Dim arrDanhsach() As String
    
    With fps
        If .ActiveSheet = 1 Then
            .EventEnabled(EventAllEvents) = False
            ' chon quoc tich
            If Col = .ColLetterToNumber("I") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("H")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
            End If
            
            ' chon moi quan he
            If Col = .ColLetterToNumber("L") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("K")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
            End If
            
            ' chon quoc gia
            If Col = .ColLetterToNumber("P") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("O")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
                
                 If .Text = "01" Then
                    .Col = .ColLetterToNumber("Q")
                    .Row = Row
                    .Text = ""
                    .Lock = False
                    
                    .Col = .ColLetterToNumber("R")
                    .Row = Row
                    .Text = ""
                    .Lock = False
                    
                    .Col = .ColLetterToNumber("S")
                    .Row = Row
                    .Text = ""
                    .Lock = False
                    
                    .Col = .ColLetterToNumber("T")
                    .Row = Row
                    .Text = ""
                    .Lock = False
                    
                    .Col = .ColLetterToNumber("U")
                    .Row = Row
                    .Text = ""
                    .Lock = False
                    
                    .Col = .ColLetterToNumber("V")
                    .Row = Row
                    .Text = ""
                    .Lock = False
                Else
                    .Col = .ColLetterToNumber("Q")
                    .Row = Row
                    .Text = "999"
                    .Lock = True
                    
                    .Col = .ColLetterToNumber("R")
                    .Row = Row
                    .Text = "Khác"
                    .Lock = True
                    
                    .Col = .ColLetterToNumber("S")
                    .Row = Row
                    .Text = "99999"
                    .Lock = True
                    
                    .Col = .ColLetterToNumber("T")
                    .Row = Row
                    .Text = "Khác"
                    .Lock = True
                    
                    .Col = .ColLetterToNumber("U")
                    .Row = Row
                    .Text = "9999999"
                    .Lock = True
                    
                    .Col = .ColLetterToNumber("V")
                    .Row = Row
                    .Text = "Khác"
                    .Lock = True
                End If
            End If
            
            ' chon tinh thanh
            If Col = .ColLetterToNumber("R") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("Q")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
                
                .GetText .Col, .Row, varIDTinh
                If varIDTinh <> "" Or varIDTinh <> vbNullString Then
                    
                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Huyen.xml")) Then
                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
                        For Each xmlNode In xmlNodeListCell
                            If GetAttribute(xmlNode, "Value") <> "" Then
                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                                
                                If arrDanhsach(1) = varIDTinh Or arrDanhsach(1) = "999" Then
                                    tenHuyen = tenHuyen + arrDanhsach(3) + Chr$(9)
                                    maHuyen = maHuyen + arrDanhsach(2) + Chr$(9)
                                End If
                            End If
                        Next
                        Set xmlDomData = Nothing
                        Set xmlNodeListCell = Nothing
                        Set xmlNode = Nothing
                    End If
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("T")
                    .TypeComboBoxList = tenHuyen
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("S")
                    .TypeComboBoxList = maHuyen
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                   
                    .Row = Row
                    .Col = .ColLetterToNumber("U")
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("V")
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                                                   
                End If
            End If
            ' chon quan huyen
            If Col = .ColLetterToNumber("T") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("S")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
                
                .GetText .Col, .Row, varIDHuyen
                If varIDHuyen <> "" Or varIDHuyen <> vbNullString Then
                    
                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Xa.xml")) Then
                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
                        For Each xmlNode In xmlNodeListCell
                            If GetAttribute(xmlNode, "Value") <> "" Then
                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                                
                                If arrDanhsach(2) = varIDHuyen Or arrDanhsach(2) = "99999" Then
                                    tenXa = tenXa + arrDanhsach(4) + Chr$(9)
                                    maXa = maXa + arrDanhsach(3) + Chr$(9)
                                End If
                            End If
                        Next
                        Set xmlDomData = Nothing
                        Set xmlNodeListCell = Nothing
                        Set xmlNode = Nothing
                    End If
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("V")
                    .TypeComboBoxList = tenXa
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("U")
                    .TypeComboBoxList = maXa
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                
                End If
            End If
            ' chon xa phuong
            If Col = .ColLetterToNumber("V") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("U")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
            End If
            .EventEnabled(EventAllEvents) = True
        End If
    End With
End Sub

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
'    Dim intIndexCombo As Integer
'
'    Dim varIDTinh As Variant
'    Dim varIDHuyen As Variant
'    Dim xmlDomData As New MSXML.DOMDocument
'    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
'    Dim xmlNode As MSXML.IXMLDOMNode
'    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
'    Dim i As Integer
'
'    Dim tenHuyen As String
'    Dim maHuyen As String
'    Dim tenXa As String
'    Dim maXa As String
'    Dim arrDanhsach() As String
'
'    With fps
'        If .ActiveSheet = 1 Then
'            .EventEnabled(EventAllEvents) = False
'            ' chon quoc tich
'            If Col = .ColLetterToNumber("I") And .Text <> "" Then
'                intIndexCombo = .typeComboboxCurSel
'                 .Row = Row
'                 .Col = .ColLetterToNumber("H")
'                 .typeComboboxCurSel = intIndexCombo
'                UpdateCell fps, .Col, .Row, .value
'            End If
'
'            ' chon moi quan he
'            If Col = .ColLetterToNumber("L") And .Text <> "" Then
'                intIndexCombo = .typeComboboxCurSel
'                 .Row = Row
'                 .Col = .ColLetterToNumber("K")
'                 .typeComboboxCurSel = intIndexCombo
'                UpdateCell fps, .Col, .Row, .value
'            End If
'
'            ' chon quoc gia
'            If Col = .ColLetterToNumber("P") And .Text <> "" Then
'                intIndexCombo = .typeComboboxCurSel
'                 .Row = Row
'                 .Col = .ColLetterToNumber("O")
'                 .typeComboboxCurSel = intIndexCombo
'                UpdateCell fps, .Col, .Row, .value
'
'                 If .Text = "01" Then
'                    .Col = .ColLetterToNumber("Q")
'                    .Row = Row
'                    .Text = ""
'                    .Lock = False
'
'                    .Col = .ColLetterToNumber("R")
'                    .Row = Row
'                    .Text = ""
'                    .Lock = False
'
'                    .Col = .ColLetterToNumber("S")
'                    .Row = Row
'                    .Text = ""
'                    .Lock = False
'
'                    .Col = .ColLetterToNumber("T")
'                    .Row = Row
'                    .Text = ""
'                    .Lock = False
'
'                    .Col = .ColLetterToNumber("U")
'                    .Row = Row
'                    .Text = ""
'                    .Lock = False
'
'                    .Col = .ColLetterToNumber("V")
'                    .Row = Row
'                    .Text = ""
'                    .Lock = False
'                Else
'                    .Col = .ColLetterToNumber("Q")
'                    .Row = Row
'                    .Text = "999"
'                    .Lock = True
'
'                    .Col = .ColLetterToNumber("R")
'                    .Row = Row
'                    .Text = "Khác"
'                    .Lock = True
'
'                    .Col = .ColLetterToNumber("S")
'                    .Row = Row
'                    .Text = "99999"
'                    .Lock = True
'
'                    .Col = .ColLetterToNumber("T")
'                    .Row = Row
'                    .Text = "Khác"
'                    .Lock = True
'
'                    .Col = .ColLetterToNumber("U")
'                    .Row = Row
'                    .Text = "9999999"
'                    .Lock = True
'
'                    .Col = .ColLetterToNumber("V")
'                    .Row = Row
'                    .Text = "Khác"
'                    .Lock = True
'                End If
'            End If
'
'            ' chon tinh thanh
'            If Col = .ColLetterToNumber("R") And .Text <> "" Then
'                intIndexCombo = .typeComboboxCurSel
'                 .Row = Row
'                 .Col = .ColLetterToNumber("Q")
'                 .typeComboboxCurSel = intIndexCombo
'                UpdateCell fps, .Col, .Row, .value
'
'                .GetText .Col, .Row, varIDTinh
'                If varIDTinh <> "" Or varIDTinh <> vbNullString Then
'
'                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Huyen.xml")) Then
'                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
'                        For Each xmlNode In xmlNodeListCell
'                            If GetAttribute(xmlNode, "Value") <> "" Then
'                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
'
'                                If arrDanhsach(1) = varIDTinh Or arrDanhsach(1) = "999" Then
'                                    tenHuyen = tenHuyen + arrDanhsach(3) + Chr$(9)
'                                    maHuyen = maHuyen + arrDanhsach(2) + Chr$(9)
'                                End If
'                            End If
'                        Next
'                        Set xmlDomData = Nothing
'                        Set xmlNodeListCell = Nothing
'                        Set xmlNode = Nothing
'                    End If
'
'                    .Row = Row
'                    .Col = .ColLetterToNumber("T")
'                    .TypeComboBoxList = tenHuyen
'                    .Text = ""
'                     UpdateCell fps, .Col, .Row, .Text
'
'                    .Row = Row
'                    .Col = .ColLetterToNumber("S")
'                    .TypeComboBoxList = maHuyen
'                    .Text = ""
'                     UpdateCell fps, .Col, .Row, .Text
'
'                    .Row = Row
'                    .Col = .ColLetterToNumber("U")
'                    .Text = ""
'                     UpdateCell fps, .Col, .Row, .Text
'
'                    .Row = Row
'                    .Col = .ColLetterToNumber("V")
'                    .Text = ""
'                     UpdateCell fps, .Col, .Row, .Text
'
'                End If
'            End If
'            ' chon quan huyen
'            If Col = .ColLetterToNumber("T") And .Text <> "" Then
'                intIndexCombo = .typeComboboxCurSel
'                 .Row = Row
'                 .Col = .ColLetterToNumber("S")
'                 .typeComboboxCurSel = intIndexCombo
'                UpdateCell fps, .Col, .Row, .value
'
'                .GetText .Col, .Row, varIDHuyen
'                If varIDHuyen <> "" Or varIDHuyen <> vbNullString Then
'
'                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Xa.xml")) Then
'                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
'                        For Each xmlNode In xmlNodeListCell
'                            If GetAttribute(xmlNode, "Value") <> "" Then
'                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
'
'                                If arrDanhsach(2) = varIDHuyen Or arrDanhsach(2) = "99999" Then
'                                    tenXa = tenXa + arrDanhsach(4) + Chr$(9)
'                                    maXa = maXa + arrDanhsach(3) + Chr$(9)
'                                End If
'                            End If
'                        Next
'                        Set xmlDomData = Nothing
'                        Set xmlNodeListCell = Nothing
'                        Set xmlNode = Nothing
'                    End If
'
'                    .Row = Row
'                    .Col = .ColLetterToNumber("V")
'                    .TypeComboBoxList = tenXa
'                    .Text = ""
'                     UpdateCell fps, .Col, .Row, .Text
'
'                    .Row = Row
'                    .Col = .ColLetterToNumber("U")
'                    .TypeComboBoxList = maXa
'                    .Text = ""
'                     UpdateCell fps, .Col, .Row, .Text
'
'                End If
'            End If
'            ' chon xa phuong
'            If Col = .ColLetterToNumber("V") And .Text <> "" Then
'                intIndexCombo = .typeComboboxCurSel
'                 .Row = Row
'                 .Col = .ColLetterToNumber("U")
'                 .typeComboboxCurSel = intIndexCombo
'                UpdateCell fps, .Col, .Row, .value
'            End If
'            .EventEnabled(EventAllEvents) = True
'        End If
'    End With
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Integer, iCol As Long, iRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
    Dim tCol As Long, tRow As Long
    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 1 Then
                .EventEnabled(EventAllEvents) = False
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 40
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = .Row + 1
                         i = i + 1
                    Loop Until .Text = "aa"
                    
                    
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = .Row + 8
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = .Row + 1
                         i = i + 1
                    Loop Until .Text = "aa"
                    
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
    If flgLoadToKhai = True Then Exit Sub
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
       If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("F") And Row >= 40 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    Else
                        '.SetFocus
                        .SetActiveCell Col, Row
                    End If
                Else
                 .SetText Col, Row, ""
                End If
                .Col = Col
                .Row = Row
                UpdateCell fps, Col, Row, .Text
            End If
            
            
            If (Col = .ColLetterToNumber("W") Or Col = .ColLetterToNumber("X")) And Row >= 40 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../...." Then
                    If Format_mmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_mmyyyy(CStr(varTemp))
                    Else
                        '.SetFocus
                        .SetActiveCell Col, Row
                    End If
                Else
                 .SetText Col, Row, ""
                End If
                .Col = Col
                .Row = Row
                UpdateCell fps, Col, Row, .Text
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub



'Calling this function whenever we don't need this class anymore.
'No parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
            
'    If flgLoadToKhai = False Then
'        defaultFormula
'    End If
       
       
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet

        'SetData

        ' Lay cac mang MST trong cac bang ke va phu luc 05A, 05B
'        prepareMSTCMT
'
        'kiemTraDuLieuImport
        
'        .Sheet = 1
'        InsertNode 41, .SearchCol(.ColLetterToNumber("B"), 40, -1, "aa", SearchFlagsNone) - 1, 1
'
        UpdateSheets
        CheckDynamicError 'Set Exception Error on cells of interface

'        If flgErr5A = True And flgErr5B = True Then kiemTraKhac

        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True

    End With
    
    'resetFormula
    
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    'chuoi vitriheader tren to khai, chuoi vi tri tren sheet header de check lai
    Dim viTriHeaderTk As String, viTriSheetHeader As String, mangMessage As String, xmlNode As MSXML.IXMLDOMNode
    Dim MSTDTNT As Variant, tSheet As Integer

    With fps
    
        tSheet = .Sheet
        'If flgLoadToKhai = True Then Exit Sub
        
        'get MST cua DTNT
        .Sheet = 1
        .GetText .ColLetterToNumber("D"), 10, MSTDTNT
        MSTDTNT = Trim(MSTDTNT)
        
        'check thong tin header khong dc trong
        'lay vi tri header tren tk theo thu tu MSTDL, TENDLT, DIACHIDLT,QUAN/HUYEN,TP,SO HDDLT,NGAYDH,Ten NV DLT, CC Hanh Nghe

        viTriHeaderTk = "D_20~D_18~D_22~D_24~G_24~D_28~L_28"
        viTriSheetHeader = "B_58~B_59~B_60~B_61~B_62~B_63~B_64"
        mangMessage = "E_58~E_59~E_60~E_61~E_62~E_63~E_64"
        'check header
        checkErrorHeader fps, viTriHeaderTk, viTriSheetHeader, mangMessage, MSTDTNT
        
        .Sheet = tSheet
    
        CheckErrorMST
       
        If IIf(TAX_Utilities_v2.NodeValidity.childNodes(0).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            CheckDynamicSheet3
        End If
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
        .Row = 3
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
    
End Sub
Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean

End Function

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()

    With fps
        If fps.ActiveSheet = 1 Then
            fps.Row = 6
            fps.Col = fps.ColLetterToNumber("G")
            fps.Text = StrSolanBosung
            
            'check thong tin ma so thue dai ly
            updateMSTDL_16TH
        
            fps.Row = 40
            fps.Col = fps.ColLetterToNumber("B")
            fps.Formula = "B39 + 1"
            
            
            fps.Row = 49
            fps.Col = fps.ColLetterToNumber("B")
            fps.Formula = "B48 + 1"
        End If
        
    End With

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
        mCurrentSheet = .Sheet
        .Sheet = 1
        mCurrentSheet = .Sheet
        For Each xmlCellNodeData In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
            ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
            .GetText lCol, lRow, varTemp
            UpdateCell fps, lCol, lRow, varTemp
            If lCol = .ColLetterToNumber("P") And Trim$(varTemp) = "" Then
                  UpdateCell fps, .ColLetterToNumber("O"), lRow, ""
            ElseIf lCol = .ColLetterToNumber("R") And Trim$(varTemp) = "" Then
                  UpdateCell fps, .ColLetterToNumber("Q"), lRow, ""
            ElseIf lCol = .ColLetterToNumber("T") And Trim$(varTemp) = "" Then
                  UpdateCell fps, .ColLetterToNumber("S"), lRow, ""
            ElseIf lCol = .ColLetterToNumber("V") And Trim$(varTemp) = "" Then
                  UpdateCell fps, .ColLetterToNumber("U"), lRow, ""
            ElseIf lCol = .ColLetterToNumber("I") And Trim$(varTemp) = "" Then
                  UpdateCell fps, .ColLetterToNumber("H"), lRow, ""
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
    Dim ASheet As Integer, SSheet As Integer
    Dim countRowG1 As Long
    Dim idx As Long

    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        
        .Sheet = 1
        Debug.Print "Tao node XML " & Time
        InsertNode16TH 41, .SearchCol(.ColLetterToNumber("B"), 40, -1, "aa", SearchFlagsNone) - 1, 1, True
        
        ' tinh so dong group 1
        .Sheet = 1
        .Col = .ColLetterToNumber("B")
        For idx = 40 To .MaxRows Step 1
            .Row = idx
            countRowG1 = countRowG1 + 1
            If .Text = "aa" Then
                Exit For
            End If
        Next

        
        InsertNode16TH 48 + countRowG1, .SearchCol(.ColLetterToNumber("B"), 47 + countRowG1, -1, "aa", SearchFlagsNone) - 1, 1, False
        Debug.Print "End Tao node XML " & Time
        UpdateSheets
        
'        .Sheet = 1
'        ' Lay danh muc quan he voi NNT
'        Dim xmlDocument As New MSXML.DOMDocument
'        Dim xmlNode As MSXML.IXMLDOMNode
'        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_QHNNT.xml"))
'        Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
'        Dim strQuanHeNNTID As String
'        Dim strQuanHeNNT As String
'
'        Dim strQuocTich As String
'        Dim strMaQuocTich As String
'        Dim strTenTinh As String
'        Dim strTenHuyen As String
'        Dim strTenXa As String
'        Dim strMaTinh As String
'        Dim strMaHuyen  As String
'        Dim strMaXa As String
'        Dim temValue As String
'        Dim arrValue() As String
'
'
'        Dim maTinh As Variant
'        Dim maHuyen As Variant
'        Dim tempArr() As String
'
'        Dim index As Integer
'        ' Chuoi gia tri ve quan he NNT bao gom:
'        ' 01-Con, 02-Vo, 03-Chong, 04-Cha, 05-Me, 06-Chau Ruot, 07-Ho Hang, 99-Khac
'        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
'        For Each xmlNode In xmlNodeListMap
'            arrValue = Split(GetAttribute(xmlNode, "value"), "###")
'            strQuanHeNNT = strQuanHeNNT + arrValue(1) + Chr$(9)
'            strQuanHeNNTID = strQuanHeNNTID + arrValue(0) + Chr$(9)
'        Next
'
'
'        ' lay du lieu danh muc dia ban tinh
'        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Tinh.xml"))
'        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
'        For Each xmlNode In xmlNodeListMap
'            temValue = GetAttribute(xmlNode, "Value")
'            arrValue = Split(temValue, "###")
'            'If arrValue(0) <> "1" Then
'                strTenTinh = strTenTinh + arrValue(2) + Chr$(9)
'                strMaTinh = strMaTinh + arrValue(1) + Chr$(9)
'            'End If
'        Next
'
'        ' lay danh sach huyen
'        Set xmlDocument = Nothing
'        Set xmlNodeListMap = Nothing
'        ' lay du lieu danh muc dia ban tinh
'        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Huyen.xml"))
'        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
'        For Each xmlNode In xmlNodeListMap
'            temValue = GetAttribute(xmlNode, "Value")
'            arrValue = Split(temValue, "###")
'            'If arrValue(0) <> "1" Then
'                strTenHuyen = strTenHuyen + arrValue(3) + Chr$(9)
'                strMaHuyen = strMaHuyen + arrValue(2) + Chr$(9)
'            'End If
'        Next
'
'        ' Lay danh muc xa
'        Set xmlDocument = Nothing
'        Set xmlNodeListMap = Nothing
'        ' lay du lieu danh muc dia ban tinh
'        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Xa.xml"))
'        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
'        For Each xmlNode In xmlNodeListMap
'            temValue = GetAttribute(xmlNode, "Value")
'            arrValue = Split(temValue, "###")
'            'If arrValue(0) <> "1" Then
'                strTenXa = strTenXa + arrValue(4) + Chr$(9)
'                strMaXa = strMaXa + arrValue(3) + Chr$(9)
'            'End If
'        Next
'
'        ' lay du lieu dia ban
'        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Quoc_Gia.xml"))
'        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
'        For Each xmlNode In xmlNodeListMap
'            temValue = GetAttribute(xmlNode, "Value")
'            arrValue = Split(temValue, "###")
'            strQuocTich = strQuocTich + arrValue(1) + Chr$(9)
'            strMaQuocTich = strMaQuocTich + arrValue(0) + Chr$(9)
'        Next
        Debug.Print "Format row " & Time
        ForMatTK16TH
        Debug.Print "End  Format row " & Time
'        .Sheet = 1
'        .Col = .ColLetterToNumber("B")
'        .Row = 40
'        Do
'            .Col = .ColLetterToNumber("K")
'            .TypeComboBoxList = strQuanHeNNTID
'            'index = .typeComboboxCurSel
'
'            .Col = .ColLetterToNumber("L")
'            .TypeComboBoxList = strQuanHeNNT
'            '.typeComboboxCurSel = index
'
'                ' set danh sach tinh
'            .Col = .ColLetterToNumber("Q")
'            .TypeComboBoxList = strMaTinh
'            'index = .typeComboboxCurSel
'
'            .Col = .ColLetterToNumber("R")
'            .TypeComboBoxList = strTenTinh
'            '.typeComboboxCurSel = index
'
'            ' set danh sach huyen
'
'            .GetText .ColLetterToNumber("Q"), .Row, maTinh
'            tempArr = Split(getHuyenByMaTinh(CStr(maTinh)), "###")
'
'            .Col = .ColLetterToNumber("S")
'            .TypeComboBoxList = tempArr(0) 'strMaHuyen
'            'index = .typeComboboxCurSel
'
'            .Col = .ColLetterToNumber("T")
'            .TypeComboBoxList = tempArr(1) 'strTenHuyen
'            '.typeComboboxCurSel = index
'
'            ' set danh sach xa
'            .GetText .ColLetterToNumber("S"), .Row, maHuyen
'            tempArr = Split(getXaByMaHuyen(CStr(maHuyen)), "###")
'            .Col = .ColLetterToNumber("U")
'            .TypeComboBoxList = tempArr(0) ' strMaXa
'            'index = .typeComboboxCurSel
'
'            .Col = .ColLetterToNumber("V")
'            .TypeComboBoxList = tempArr(1) 'strTenXa
'            '.typeComboboxCurSel = index
'
''            ' set danh sach quoc tich
'            .Col = .ColLetterToNumber("I")
'            .TypeComboBoxList = strMaQuocTich
'            'index = .typeComboboxCurSel
'
'            .Col = .ColLetterToNumber("J")
'            .TypeComboBoxList = strQuocTich
'            '.typeComboboxCurSel = index
'
'
'            ' set danh sach quoc gia
'            .Col = .ColLetterToNumber("O")
'            .TypeComboBoxList = strMaQuocTich
'            'index = .typeComboboxCurSel
'
'            .Col = .ColLetterToNumber("P")
'            .TypeComboBoxList = strQuocTich
'            '.typeComboboxCurSel = index
'
'
'             .Col = .ColLetterToNumber("B")
'             .Row = .Row + 1
'        Loop Until .Text = "aa"

        
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
        
    End With
End Sub

Public Sub kiemTraDuLieuImport()

End Sub

'Public Sub InsertNode(ByVal fRow As Long, ByVal lRow As Long, ByVal mSheet As Byte)
'    On Error GoTo ErrorHandle
'
'    Dim xmlNodeCells As MSXML.IXMLDOMNode
'    Dim xmlNodeNewCells As MSXML.IXMLDOMNode
'    Dim xmlNodeNewCell As MSXML.IXMLDOMNode
'    Dim arrTemp() As String
'
'
'    Dim ctlRow As Long
'    Dim cTemplate, cReport As String
'
'    Dim sumRowDel As Long
'    sumRowDel = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Section").Item(0).childNodes.length
'    For ctlRow = 1 To sumRowDel - 1
'        If mSheet = 1 Then
'            Set xmlNodeCells = TAX_Utilities_v2.Data(mSheet - 1).nodeFromID("C" & "_" & 40 + ctlRow & "").parentNode
'            xmlNodeCells.parentNode.removeChild xmlNodeCells
'        End If
'    Next
'
'    If mSheet = 1 Then
'        Set xmlNodeCells = TAX_Utilities_v2.Data(mSheet - 1).nodeFromID("C_40").parentNode
'    End If
'
'    For ctlRow = fRow To lRow
'
'        Set xmlNodeNewCells = xmlNodeCells.cloneNode(True)
'
'        For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
'            cTemplate = Left$(GetAttribute(xmlNodeNewCell, "CellID"), 2)
'            arrTemp = Split(GetAttribute(xmlNodeNewCell, "CellID2"), "_")
'            'cReport = Left$(GetAttribute(xmlNodeNewCell, "CellID2"), 2)
'            cReport = Trim(arrTemp(0)) & "_"
'
'            SetAttribute xmlNodeNewCell, "CellID", cTemplate & ctlRow
'
'            ' Set first cell = 1
'            SetAttribute xmlNodeNewCell, "FirstCell", "1"
'            If mSheet = 1 Then
'                SetAttribute xmlNodeNewCell, "Value", vbNullString
'                SetAttribute xmlNodeNewCell, "CellID2", cReport & (ctlRow - 2)
'            End If
'
'        Next
'        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, Null
'
'
'    Next
''   **********************************
''    added
''   Date: 12/04/06
'    'mAdjustData = True
'    TAX_Utilities_v2.AdjustData(mSheet - 1) = True
''   **********************************
'EXIT_SUB:
'    Set xmlNodeNewCell = Nothing
'    Set xmlNodeNewCells = Nothing
'    Set xmlNodeCells = Nothing
'
'    Exit Sub
'
'ErrorHandle:
'    SaveErrorLog "cls_16TH_DKNPT", "InsertNode 05_TNCN", Err.number, Err.Description
'End Sub

' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 15
    denRowDL = 20
    
End Sub

Public Sub CheckDynamicSheet1()

    
End Sub


Public Sub fillDataPL27()

End Sub





Private Sub CheckDynamicSheet3()
    Dim vHoVaTenNNT, vMSTNNT, vHoVaTenNPT, vNgaySinhNPT, vMSTNPT, vQuocTich, vSoCMNDNPT, vQuanHeNPT, vSo, vQuyenSo, vQuocGia As Variant
    Dim vTinh, vHuyen, vXa, vTuThang, vDenThang As Variant
    Dim vMaQuocGia As Variant, vMaTinh As Variant, vMaHuyen As Variant, vMaXa As Variant, vMaQuocTich As Variant
    
    Dim vErrHoVaTenNNT, vErrMSTNNT, vErrHoVaTenNPT, vErrNgaySinhNPT, vErrMSTNPT, vErrSoCMNDNPT, vErrTuThang, vErrDenThang     As Variant
    Dim vErrEmpty As Variant, vErrEmptyCMNDNPT As Variant
    Dim vErrDuLieuKhongHL As Variant
    
    Dim iFlagHoVaTenNNT As Integer, iFlagMSTNNT As Integer, iFlagHoVaTenNPT As Integer, iFlagNgaySinhNPT As Integer
    Dim iFlagMSTNPT As Integer, iFlagSoCMNDNPT As Integer, iFlagTuThang As Integer, iFlagDenThang As Integer
    Dim iFlagEmpty As Integer
    
    Dim strColReset As String
    Dim arrColReset() As String
    Dim strCheck As String
    Dim i As Integer
    Dim vErrNgaySinhNPT1 As Variant
    
    Dim vMSTDN As Variant
    Dim vErrMstDN As Variant
    Dim iFlagMstDN As Variant
    
    Dim vNgaySinhNPT_Goc As Variant
    
    With fps
        .Sheet = .SheetCount
        
        .GetText .ColLetterToNumber("E"), 71, vErrHoVaTenNNT
        .GetText .ColLetterToNumber("E"), 72, vErrMSTNNT
        .GetText .ColLetterToNumber("E"), 73, vErrHoVaTenNPT
        .GetText .ColLetterToNumber("E"), 74, vErrNgaySinhNPT
        .GetText .ColLetterToNumber("E"), 75, vErrMSTNPT
        
        .GetText .ColLetterToNumber("E"), 76, vErrSoCMNDNPT
        .GetText .ColLetterToNumber("E"), 77, vErrTuThang
        .GetText .ColLetterToNumber("E"), 78, vErrDenThang
        .GetText .ColLetterToNumber("E"), 79, vErrEmpty
        
        .GetText .ColLetterToNumber("E"), 83, vErrEmptyCMNDNPT
        .GetText .ColLetterToNumber("E"), 84, vErrDuLieuKhongHL
        .GetText .ColLetterToNumber("E"), 85, vErrNgaySinhNPT1
        
        .GetText .ColLetterToNumber("E"), 86, vErrMstDN
        
        .Sheet = 1
        .GetText .ColLetterToNumber("D"), 10, vMSTDN
                
        .Row = 40
        Do
            .GetText .ColLetterToNumber("C"), .Row, vHoVaTenNNT
            .GetText .ColLetterToNumber("D"), .Row, vMSTNNT
            .GetText .ColLetterToNumber("E"), .Row, vHoVaTenNPT
            .GetText .ColLetterToNumber("F"), .Row, vNgaySinhNPT
            .GetText .ColLetterToNumber("F"), .Row, vNgaySinhNPT_Goc
            .GetText .ColLetterToNumber("G"), .Row, vMSTNPT
            .GetText .ColLetterToNumber("I"), .Row, vQuocTich
            .GetText .ColLetterToNumber("J"), .Row, vSoCMNDNPT
            .GetText .ColLetterToNumber("L"), .Row, vQuanHeNPT
            .GetText .ColLetterToNumber("M"), .Row, vSo
            .GetText .ColLetterToNumber("N"), .Row, vQuyenSo
            .GetText .ColLetterToNumber("P"), .Row, vQuocGia
            .GetText .ColLetterToNumber("R"), .Row, vTinh
            .GetText .ColLetterToNumber("T"), .Row, vHuyen
            .GetText .ColLetterToNumber("V"), .Row, vXa
            .GetText .ColLetterToNumber("W"), .Row, vTuThang
            .GetText .ColLetterToNumber("X"), .Row, vDenThang
            
            .GetText .ColLetterToNumber("O"), .Row, vMaQuocGia
            .GetText .ColLetterToNumber("Q"), .Row, vMaTinh
            .GetText .ColLetterToNumber("S"), .Row, vMaHuyen
            .GetText .ColLetterToNumber("U"), .Row, vMaXa
            
            .GetText .ColLetterToNumber("H"), .Row, vMaQuocTich
            
            ' reset du lieu loi
            strColReset = "C~D~E~F~I~J~L~M~N~P~R~T~V~W~X"
            arrColReset = Split(strColReset, "~")
            For i = 0 To UBound(arrColReset)
                .Col = .ColLetterToNumber(arrColReset(i))
                .BackColor = mNonErrorColor
                .CellNote = ""
            Next
            If Trim(vHoVaTenNNT) = vbNullString And Trim(vMSTNNT) = vbNullString And Trim(vHoVaTenNPT) = vbNullString And Trim(vNgaySinhNPT) = vbNullString _
            And Trim(vMSTNPT) = vbNullString And Trim(vSoCMNDNPT) = vbNullString And Trim(vQuanHeNPT) = vbNullString _
            And Trim(vSo) = vbNullString And Trim(vQuyenSo) = vbNullString And Trim(vTinh) = vbNullString _
            And Trim(vHuyen) = vbNullString And Trim(vXa) = vbNullString And Trim(vTuThang) = vbNullString And Trim(vDenThang) = vbNullString And Trim$(vQuocTich) = vbNullString Then

            Else
                ' check ho ten
                If Trim(vHoVaTenNNT) = vbNullString Then
                    .Col = .ColLetterToNumber("C")
                    .BackColor = mErrorColor
                    .CellNote = vErrHoVaTenNNT
                    iFlagHoVaTenNNT = 1
                Else
                    .Col = .ColLetterToNumber("C")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If
                
                ' check MST
                'check cau truc mst
                vMSTNNT = Replace$(vMSTNNT, "-", "")
                If Len(Trim(vMSTNNT)) = 10 Then
                     strCheck = CheckTaxCode(Mid(Trim(vMSTNNT), 1, 1), Mid(Trim(vMSTNNT), 2, 1), Mid(Trim(vMSTNNT), 3, 1), Mid(Trim(vMSTNNT), 4, 1), Mid(Trim(vMSTNNT), 5, 1), Mid(Trim(vMSTNNT), 6, 1), Mid(Trim(vMSTNNT), 7, 1), Mid(Trim(vMSTNNT), 8, 1), Mid(Trim(vMSTNNT), 9, 1), Mid(Trim(vMSTNNT), 10, 1), Mid(Trim(vMSTNNT), 11, 1), Mid(Trim(vMSTNNT), 12, 1), Mid(Trim(vMSTNNT), 13, 1))
                     If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMSTNNT) Then
                        .Col = .ColLetterToNumber("D")
                        .CellNote = vErrMSTNNT
                        .BackColor = mErrorColor
                        iFlagMSTNNT = 1
                    Else
                        If Trim(vMSTNNT) = Trim$(vMSTDN) Then
                            .Col = .ColLetterToNumber("D")
                            .CellNote = vErrMstDN
                            .BackColor = mErrorColor
                            iFlagMstDN = 1
                        Else
                            .Col = .ColLetterToNumber("D")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                     End If
                Else
                        .Col = .ColLetterToNumber("D")
                        .CellNote = vErrMSTNNT
                        .BackColor = mErrorColor
                        iFlagMSTNNT = 1
                End If
                
                ' check ho ten NPT
                If Trim(vHoVaTenNPT) = vbNullString Then
                    .Col = .ColLetterToNumber("E")
                    .BackColor = mErrorColor
                    .CellNote = vErrHoVaTenNPT
                    iFlagHoVaTenNPT = 1
                Else
                    .Col = .ColLetterToNumber("E")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If
                
                ' kiem tra ngay sinh khong duoc lon hon ngay hien tai
                If Trim(vNgaySinhNPT) <> "" Then
                    If Format_ddmmyyyy1(CStr(vNgaySinhNPT_Goc)) <> "" Then
                        'If DateDiff("D", ToDate(CStr(vNgaySinhNPT), "dd/mm/yyyy"), Date) > 0 Then
                        vNgaySinhNPT = DateSerial(CInt(Mid$(vNgaySinhNPT, 7, 4)), CInt(Mid$(vNgaySinhNPT, 4, 2)), CInt(Mid$(vNgaySinhNPT, 1, 2)))
                        If vNgaySinhNPT > Date Then
                            .Col = .ColLetterToNumber("F")
                            .BackColor = mErrorColor
                            .CellNote = vErrNgaySinhNPT
                            iFlagNgaySinhNPT = 1
                        Else
                             If Trim(vTuThang) <> "" Then
                                  If Format_ddmmyyyy1("01/" & Trim$(vTuThang)) <> "" Then
                                        If DateDiff("M", vNgaySinhNPT, Format(vTuThang, "mm/yyyy")) < 0 Then
                                            .Col = .ColLetterToNumber("F")
                                            .BackColor = mErrorColor
                                            .CellNote = vErrNgaySinhNPT1
                                            iFlagNgaySinhNPT = 1
                                        Else
                                            .Col = .ColLetterToNumber("F")
                                            .BackColor = mNonErrorColor
                                            .CellNote = ""
                                        End If
                                   End If
                            Else
                                 .Col = .ColLetterToNumber("F")
                                .BackColor = mNonErrorColor
                                .CellNote = ""
                            End If
                            
                        End If
                    Else
                        .Col = .ColLetterToNumber("F")
                        .BackColor = mErrorColor
                        .CellNote = vErrDuLieuKhongHL
                        iFlagNgaySinhNPT = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("F")
                    .BackColor = mErrorColor
                    .CellNote = vErrEmpty
                    iFlagNgaySinhNPT = 1
                End If
                
'                ' Kiem tra MST NPT
'                If Len(vMSTNPT) > 0 Then
'                    vMSTNPT = Replace$(vMSTNPT, "-", "")
'                    If Len(Trim(vMSTNPT)) = 10 Then
'                         strCheck = CheckTaxCode(Mid(Trim(vMSTNPT), 1, 1), Mid(Trim(vMSTNPT), 2, 1), Mid(Trim(vMSTNPT), 3, 1), Mid(Trim(vMSTNPT), 4, 1), Mid(Trim(vMSTNPT), 5, 1), Mid(Trim(vMSTNPT), 6, 1), Mid(Trim(vMSTNPT), 7, 1), Mid(Trim(vMSTNPT), 8, 1), Mid(Trim(vMSTNPT), 9, 1), Mid(Trim(vMSTNPT), 10, 1), Mid(Trim(vMSTNPT), 11, 1), Mid(Trim(vMSTNPT), 12, 1), Mid(Trim(vMSTNPT), 13, 1))
'                         If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMSTNPT) Then
'                            .Col = .ColLetterToNumber("G")
'                            .CellNote = vErrMSTNPT
'                            .BackColor = mErrorColor
'                            iFlagMSTNPT = 1
'                        Else
'                            .Col = .ColLetterToNumber("G")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
'                         End If
'                    Else
'                            .Col = .ColLetterToNumber("G")
'                            .CellNote = vErrMSTNPT
'                            .BackColor = mErrorColor
'                            iFlagMSTNPT = 1
'                    End If
'                Else
'                    .Col = .ColLetterToNumber("G")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                End If
                
                ' check CMND
                If Trim(vSoCMNDNPT) <> vbNullString Then
                    If InStr(1, vSoCMNDNPT, " ", vbTextCompare) <> 0 Then
                        .Col = .ColLetterToNumber("J")
                        .CellNote = vErrSoCMNDNPT
                        .BackColor = mErrorColor
                        iFlagSoCMNDNPT = 1
                    Else
                        .Col = .ColLetterToNumber("J")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                Else
                     ' kiem tra neu tuoi >15 bat buoc nhap CMT
                    If vNgaySinhNPT <> "" Then
                        If Format_ddmmyyyy1(CStr(vNgaySinhNPT_Goc)) <> "" Then
                            If Year(Date) - Year(Format(vNgaySinhNPT, "dd/mm/yyyy")) > 15 And Trim(vSoCMNDNPT) = "" Then
                                 .Col = .ColLetterToNumber("J")
                                .CellNote = vErrEmptyCMNDNPT
                                .BackColor = mErrorColor
                                iFlagEmpty = 1
                            End If
                         Else
                             .Col = .ColLetterToNumber("F")
                            .CellNote = vErrDuLieuKhongHL
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                         End If
                    End If
                End If
                 
                 ' Kiem tra tu thang bat buoc nhap
                 If Trim(vTuThang) <> "" Then
                    ' kiem tra tinh hop le cua tu thang
                    If Format_ddmmyyyy1("01/" & Trim(vTuThang)) <> "" Then
                        .Col = .ColLetterToNumber("W")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    Else
                        .Col = .ColLetterToNumber("W")
                        .CellNote = vErrDuLieuKhongHL
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                 Else
                    .Col = .ColLetterToNumber("W")
                    .CellNote = vErrEmpty
                    .BackColor = mErrorColor
                    iFlagEmpty = 1
                 End If
                 
                 ' kiem tra tinh hop le cua den thang
                 If Trim(vDenThang) <> "" Then
                    If Format_ddmmyyyy1("01/" & Trim(vDenThang)) <> "" Then
                        .Col = .ColLetterToNumber("X")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    Else
                        .Col = .ColLetterToNumber("X")
                        .CellNote = vErrDuLieuKhongHL
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                 End If
                 
                 ' kiem tra den thang bat buoc nhap
'                 If Trim(vDenThang) <> "" Then
'                    .Col = .ColLetterToNumber("X")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                 Else
'                    .Col = .ColLetterToNumber("X")
'                    .CellNote = vErrEmpty
'                    .BackColor = mErrorColor
'                    iFlagEmpty = 1
'                 End If
                 
                ' Kiem tra tu thang, den thang
                If Trim(vTuThang) <> "" And Trim(vDenThang) <> "" Then
                     If Format_ddmmyyyy1("01/" & Trim(vTuThang)) <> "" And Format_ddmmyyyy1("01/" & Trim(vDenThang)) <> "" Then
                            If DateDiff("M", Format(vTuThang, "mm/yyyy"), Format(vDenThang, "mm/yyyy")) < 0 Then
                                    .Col = .ColLetterToNumber("W")
                                    .CellNote = vErrTuThang
                                    .BackColor = mErrorColor
                                    
                                    .Col = .ColLetterToNumber("X")
                                    .CellNote = vErrTuThang
                                    .BackColor = mErrorColor
                                    iFlagTuThang = 1
                            Else
                                    .Col = .ColLetterToNumber("W")
                                    .CellNote = ""
                                    .BackColor = mNonErrorColor
        '
                                    .Col = .ColLetterToNumber("X")
                                    .CellNote = ""
                                    .BackColor = mNonErrorColor
                            End If
                       End If
                 End If
                
                
                ' kiem tra bat buoc nhap cot 14
                If Trim(vQuanHeNPT) = "" Then
                    .Col = .ColLetterToNumber("L")
                    .CellNote = vErrEmpty
                    .BackColor = mErrorColor
                    iFlagEmpty = 1
                End If
                
                 ' kiem tra bat buoc nhap cot so neu tuoi nho hon 15t
                If Trim(vSoCMNDNPT) = "" Then
                    If Trim(vSo) = "" And vMaQuocTich <> "02" Then
                        If vNgaySinhNPT <> "" And Format_ddmmyyyy1(CStr(vNgaySinhNPT_Goc)) <> "" Then
                            If Year(Date) - Year(Format(vNgaySinhNPT, "dd/mm/yyyy")) <= 15 Then
                                .Col = .ColLetterToNumber("M")
                                .CellNote = vErrEmpty
                                .BackColor = mErrorColor
                                iFlagEmpty = 1
                            End If
                        End If
                    Else
                        .Col = .ColLetterToNumber("M")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                Else
                    .Col = .ColLetterToNumber("M")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                 ' kiem tra bat buoc nhap cot quyen so neu tuoi nho hon 15t
                If Trim(vSoCMNDNPT) = "" Then
                    If Trim(vQuyenSo) = "" And vMaQuocTich <> "02" Then
                        If vNgaySinhNPT <> "" And Format_ddmmyyyy1(CStr(vNgaySinhNPT_Goc)) <> "" Then
                            If Year(Date) - Year(Format(vNgaySinhNPT, "dd/mm/yyyy")) <= 15 Then
                                .Col = .ColLetterToNumber("N")
                                .CellNote = vErrEmpty
                                .BackColor = mErrorColor
                                iFlagEmpty = 1
                            End If
                        End If
                    Else
                        .Col = .ColLetterToNumber("N")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                Else
                    .Col = .ColLetterToNumber("N")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                ' kiem tra khong co cot 11 va 13 bat buoc nhap cot 18,19,20
                If Trim(vSoCMNDNPT) = "" And vMaQuocTich <> "02" Then
                    ' bat buoc nhap QG
                    
                    If Trim(vQuocGia) = "" Then
                        .Col = .ColLetterToNumber("P")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    ' bat buoc nhap thong tin tinh
                    If Trim(vTinh) = "" Then
                        .Col = .ColLetterToNumber("R")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    ' bat buoc nhap thong tin huyen
                    If Trim(vHuyen) = "" Then
                        .Col = .ColLetterToNumber("T")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    ' bat buoc nhap thong tin xa
                    If Trim(vXa) = "" Then
                        .Col = .ColLetterToNumber("V")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    
                    ' kiem tra neu chon quoc gia Viet Nam
                    If Trim(vMaQuocGia) = "01" And Trim$(vQuocGia) <> "" Then
                        If Trim$(vMaTinh) = "999" Then
                            .Col = .ColLetterToNumber("R")
                            .CellNote = vErrDuLieuKhongHL
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                        End If
                        
                        If Trim$(vMaHuyen) = "99999" Then
                            .Col = .ColLetterToNumber("T")
                            .CellNote = vErrDuLieuKhongHL
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                        End If
                        
                        If Trim$(vMaXa) = "9999999" Then
                            .Col = .ColLetterToNumber("V")
                            .CellNote = vErrDuLieuKhongHL
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                        End If
                    End If
'                Else
'                ' them phan check neu <=15 tuoi se bat buoc nhap thong tin tren GKS
'                    If vNgaySinhNPT <> "" And Format_ddmmyyyy1(CStr(vNgaySinhNPT_Goc)) <> "" Then
'                            If Year(Date) - Year(Format(vNgaySinhNPT, "dd/mm/yyyy")) <= 15 Then
'                                   ' bat buoc nhap QG
'                               If vMaQuocTich <> "02" Then
'
'                                    If Trim(vQuocGia) = "" Then
'                                        .Col = .ColLetterToNumber("P")
'                                        .CellNote = vErrEmpty
'                                        .BackColor = mErrorColor
'                                        iFlagEmpty = 1
'                                    End If
'
'                                    ' bat buoc nhap thong tin tinh
'                                    If Trim(vTinh) = "" Then
'                                        .Col = .ColLetterToNumber("R")
'                                        .CellNote = vErrEmpty
'                                        .BackColor = mErrorColor
'                                        iFlagEmpty = 1
'                                    End If
'
'                                    ' bat buoc nhap thong tin huyen
'                                    If Trim(vHuyen) = "" Then
'                                        .Col = .ColLetterToNumber("T")
'                                        .CellNote = vErrEmpty
'                                        .BackColor = mErrorColor
'                                        iFlagEmpty = 1
'                                    End If
'
'                                    ' bat buoc nhap thong tin xa
'                                    If Trim(vXa) = "" Then
'                                        .Col = .ColLetterToNumber("V")
'                                        .CellNote = vErrEmpty
'                                        .BackColor = mErrorColor
'                                        iFlagEmpty = 1
'                                    End If
'                                 End If
'
'                                    ' kiem tra neu chon quoc gia Viet Nam
'                                    If Trim(vMaQuocGia) = "01" And Trim$(vQuocGia) <> "" Then
'                                        If Trim$(vMaTinh) = "999" Then
'                                            .Col = .ColLetterToNumber("R")
'                                            .CellNote = vErrDuLieuKhongHL
'                                            .BackColor = mErrorColor
'                                            iFlagEmpty = 1
'                                        End If
'
'                                        If Trim$(vMaHuyen) = "99999" Then
'                                            .Col = .ColLetterToNumber("T")
'                                            .CellNote = vErrDuLieuKhongHL
'                                            .BackColor = mErrorColor
'                                            iFlagEmpty = 1
'                                        End If
'
'                                        If Trim$(vMaXa) = "9999999" Then
'                                            .Col = .ColLetterToNumber("V")
'                                            .CellNote = vErrDuLieuKhongHL
'                                            .BackColor = mErrorColor
'                                            iFlagEmpty = 1
'                                        End If
'                                    End If
'
'                            End If
'                    End If
                 End If
                 
                  If Trim$(vTinh) <> "" Or Trim$(vHuyen) <> "" Or Trim$(vXa) <> "" Then
                    ' bat buoc nhap QG
                    
                    If Trim(vQuocGia) = "" Then
                        .Col = .ColLetterToNumber("P")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    ' bat buoc nhap thong tin tinh
                    If Trim(vTinh) = "" Then
                        .Col = .ColLetterToNumber("R")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    ' bat buoc nhap thong tin huyen
                    If Trim(vHuyen) = "" Then
                        .Col = .ColLetterToNumber("T")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    ' bat buoc nhap thong tin xa
                    If Trim(vXa) = "" Then
                        .Col = .ColLetterToNumber("V")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                End If
            End If
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
        
        ' check group 2
        .Row = .Row + 8
        Do
            .GetText .ColLetterToNumber("C"), .Row, vHoVaTenNNT
            .GetText .ColLetterToNumber("D"), .Row, vMSTNNT
            .GetText .ColLetterToNumber("E"), .Row, vHoVaTenNPT
            .GetText .ColLetterToNumber("F"), .Row, vNgaySinhNPT
            .GetText .ColLetterToNumber("G"), .Row, vMSTNPT
            .GetText .ColLetterToNumber("I"), .Row, vQuocTich
            .GetText .ColLetterToNumber("J"), .Row, vSoCMNDNPT
            .GetText .ColLetterToNumber("L"), .Row, vQuanHeNPT
            .GetText .ColLetterToNumber("M"), .Row, vSo
            .GetText .ColLetterToNumber("N"), .Row, vQuyenSo
            .GetText .ColLetterToNumber("P"), .Row, vQuocGia
            .GetText .ColLetterToNumber("R"), .Row, vTinh
            .GetText .ColLetterToNumber("T"), .Row, vHuyen
            .GetText .ColLetterToNumber("V"), .Row, vXa
            .GetText .ColLetterToNumber("W"), .Row, vTuThang
            .GetText .ColLetterToNumber("X"), .Row, vDenThang
            
            .GetText .ColLetterToNumber("O"), .Row, vMaQuocGia
            .GetText .ColLetterToNumber("Q"), .Row, vMaTinh
            .GetText .ColLetterToNumber("S"), .Row, vMaHuyen
            .GetText .ColLetterToNumber("U"), .Row, vMaXa
            
            ' reset du lieu loi
            strColReset = "C~D~E~G~L~W~X"
            arrColReset = Split(strColReset, "~")
            For i = 0 To UBound(arrColReset)
                .Col = .ColLetterToNumber(arrColReset(i))
                .BackColor = mNonErrorColor
                .CellNote = ""
            Next
            If Trim(vHoVaTenNNT) = vbNullString And Trim(vMSTNNT) = vbNullString And Trim(vHoVaTenNPT) = vbNullString And Trim(vNgaySinhNPT) = vbNullString _
            And Trim(vMSTNPT) = vbNullString And Trim(vSoCMNDNPT) = vbNullString And Trim(vQuanHeNPT) = vbNullString _
            And Trim(vSo) = vbNullString And Trim(vQuyenSo) = vbNullString And Trim(vTinh) = vbNullString _
            And Trim(vHuyen) = vbNullString And Trim(vXa) = vbNullString And Trim(vTuThang) = vbNullString And Trim(vDenThang) = vbNullString Then

            Else
                ' check ho ten
                If Trim(vHoVaTenNNT) = vbNullString Then
                    .Col = .ColLetterToNumber("C")
                    .BackColor = mErrorColor
                    .CellNote = vErrHoVaTenNNT
                    iFlagHoVaTenNNT = 1
                Else
                    .Col = .ColLetterToNumber("C")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If
                
                ' check MST
                'check cau truc mst
                vMSTNNT = Replace$(vMSTNNT, "-", "")
                If Len(Trim(vMSTNNT)) = 10 Then
                     strCheck = CheckTaxCode(Mid(Trim(vMSTNNT), 1, 1), Mid(Trim(vMSTNNT), 2, 1), Mid(Trim(vMSTNNT), 3, 1), Mid(Trim(vMSTNNT), 4, 1), Mid(Trim(vMSTNNT), 5, 1), Mid(Trim(vMSTNNT), 6, 1), Mid(Trim(vMSTNNT), 7, 1), Mid(Trim(vMSTNNT), 8, 1), Mid(Trim(vMSTNNT), 9, 1), Mid(Trim(vMSTNNT), 10, 1), Mid(Trim(vMSTNNT), 11, 1), Mid(Trim(vMSTNNT), 12, 1), Mid(Trim(vMSTNNT), 13, 1))
                     If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMSTNNT) Then
                        .Col = .ColLetterToNumber("D")
                        .CellNote = vErrMSTNNT
                        .BackColor = mErrorColor
                        iFlagMSTNNT = 1
                    Else
                        If Trim(vMSTNNT) = Trim$(vMSTDN) Then
                            .Col = .ColLetterToNumber("D")
                            .CellNote = vErrMstDN
                            .BackColor = mErrorColor
                            iFlagMstDN = 1
                        Else
                            .Col = .ColLetterToNumber("D")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                     End If
                Else
                        .Col = .ColLetterToNumber("D")
                        .CellNote = vErrMSTNNT
                        .BackColor = mErrorColor
                        iFlagMSTNNT = 1
                End If
                
                ' check ho ten NPT
                If Trim(vHoVaTenNPT) = vbNullString Then
                    .Col = .ColLetterToNumber("E")
                    .BackColor = mErrorColor
                    .CellNote = vErrHoVaTenNPT
                    iFlagHoVaTenNPT = 1
                Else
                    .Col = .ColLetterToNumber("E")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If
                
                ' check MST cua NPT
                
                vMSTNPT = Replace$(vMSTNPT, "-", "")
                If Len(Trim(vMSTNPT)) = 10 Then
                     strCheck = CheckTaxCode(Mid(Trim(vMSTNPT), 1, 1), Mid(Trim(vMSTNPT), 2, 1), Mid(Trim(vMSTNPT), 3, 1), Mid(Trim(vMSTNPT), 4, 1), Mid(Trim(vMSTNPT), 5, 1), Mid(Trim(vMSTNPT), 6, 1), Mid(Trim(vMSTNPT), 7, 1), Mid(Trim(vMSTNPT), 8, 1), Mid(Trim(vMSTNPT), 9, 1), Mid(Trim(vMSTNPT), 10, 1), Mid(Trim(vMSTNPT), 11, 1), Mid(Trim(vMSTNPT), 12, 1), Mid(Trim(vMSTNPT), 13, 1))
                     If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMSTNPT) Then
                        .Col = .ColLetterToNumber("G")
                        .CellNote = vErrMSTNNT
                        .BackColor = mErrorColor
                        iFlagMSTNNT = 1
                    Else
                        .Col = .ColLetterToNumber("G")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                     End If
                Else
                        .Col = .ColLetterToNumber("G")
                        .CellNote = vErrMSTNNT
                        .BackColor = mErrorColor
                        iFlagMSTNNT = 1
                End If
                
                ' kiem tra ngay sinh khong duoc lon hon ngay hien tai
'                If Trim(vNgaySinhNPT) <> "" Then
'                    'If DateDiff("D", ToDate(CStr(vNgaySinhNPT), "dd/mm/yyyy"), Date) > 0 Then
'                    vNgaySinhNPT = DateSerial(CInt(Mid$(vNgaySinhNPT, 7, 4)), CInt(Mid$(vNgaySinhNPT, 4, 2)), CInt(Mid$(vNgaySinhNPT, 1, 2)))
'                    If vNgaySinhNPT > Date Then
'                        .Col = .ColLetterToNumber("F")
'                        .BackColor = mErrorColor
'                        .CellNote = vErrNgaySinhNPT
'                        iFlagNgaySinhNPT = 1
'                    Else
'                        .Col = .ColLetterToNumber("F")
'                        .BackColor = mNonErrorColor
'                        .CellNote = ""
'                    End If
'                Else
'                    .Col = .ColLetterToNumber("F")
'                    .BackColor = mErrorColor
'                    .CellNote = vErrEmpty
'                    iFlagNgaySinhNPT = 1
'                End If
                
                ' Kiem tra MST NPT
'                If Len(vMSTNPT) > 0 Then
'                    vMSTNPT = Replace$(vMSTNPT, "-", "")
'                    If Len(Trim(vMSTNPT)) = 10 Then
'                         strCheck = CheckTaxCode(Mid(Trim(vMSTNPT), 1, 1), Mid(Trim(vMSTNPT), 2, 1), Mid(Trim(vMSTNPT), 3, 1), Mid(Trim(vMSTNPT), 4, 1), Mid(Trim(vMSTNPT), 5, 1), Mid(Trim(vMSTNPT), 6, 1), Mid(Trim(vMSTNPT), 7, 1), Mid(Trim(vMSTNPT), 8, 1), Mid(Trim(vMSTNPT), 9, 1), Mid(Trim(vMSTNPT), 10, 1), Mid(Trim(vMSTNPT), 11, 1), Mid(Trim(vMSTNPT), 12, 1), Mid(Trim(vMSTNPT), 13, 1))
'                         If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMSTNPT) Then
'                            .Col = .ColLetterToNumber("G")
'                            .CellNote = vErrMSTNPT
'                            .BackColor = mErrorColor
'                            iFlagMSTNPT = 1
'                        Else
'                            .Col = .ColLetterToNumber("G")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
'                         End If
'                    Else
'                            .Col = .ColLetterToNumber("G")
'                            .CellNote = vErrMSTNPT
'                            .BackColor = mErrorColor
'                            iFlagMSTNPT = 1
'                    End If
'                Else
'                    .Col = .ColLetterToNumber("G")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                End If
                
                ' check CMND
'                If Trim(vSoCMNDNPT) <> vbNullString Then
'                    If InStr(1, vSoCMNDNPT, " ", vbTextCompare) <> 0 Then
'                        .Col = .ColLetterToNumber("J")
'                        .CellNote = vErrSoCMNDNPT
'                        .BackColor = mErrorColor
'                        iFlagSoCMNDNPT = 1
'                    Else
'                        .Col = .ColLetterToNumber("J")
'                        .CellNote = ""
'                        .BackColor = mNonErrorColor
'                    End If
'                Else
'                     ' kiem tra neu tuoi >15 bat buoc nhap CMT
'                    If vNgaySinhNPT <> "" Then
'                        If Year(Date) - Year(Format(vNgaySinhNPT, "dd/mm/yyyy")) > 15 And Trim(vSoCMNDNPT) = "" Then
'                             .Col = .ColLetterToNumber("J")
'                            .CellNote = vErrEmptyCMNDNPT
'                            .BackColor = mErrorColor
'                            iFlagEmpty = 1
'                        End If
'                    End If
'                End If
                 
                 ' Kiem tra tu thang bat buoc nhap
                 If Trim(vTuThang) <> "" Then
                    ' kiem tra tinh hop le cua du lieu
                    If Format_ddmmyyyy1("01/" & Trim$(vTuThang)) <> "" Then
                        .Col = .ColLetterToNumber("W")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                     Else
                        .Col = .ColLetterToNumber("W")
                        .CellNote = vErrDuLieuKhongHL
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                     End If
                 Else
                    .Col = .ColLetterToNumber("W")
                    .CellNote = vErrEmpty
                    .BackColor = mErrorColor
                    iFlagEmpty = 1
                 End If
                 
                 ' kiem tra tinh hop le cua den thang
                 If Trim(vDenThang) <> "" Then
                    If Format_ddmmyyyy1("01/" & Trim$(vDenThang)) <> "" Then
                        .Col = .ColLetterToNumber("X")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    Else
                        .Col = .ColLetterToNumber("X")
                        .CellNote = vErrDuLieuKhongHL
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                 Else
                    .Col = .ColLetterToNumber("X")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                 End If
                 ' kiem tra den thang bat buoc nhap
'                 If Trim(vDenThang) <> "" Then
'                    .Col = .ColLetterToNumber("X")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                 Else
'                    .Col = .ColLetterToNumber("X")
'                    .CellNote = vErrEmpty
'                    .BackColor = mErrorColor
'                    iFlagEmpty = 1
'                 End If
                 
                ' Kiem tra tu thang, den thang
                If Trim(vTuThang) <> "" And Trim(vDenThang) <> "" Then
                      If Format_ddmmyyyy1("01/" & Trim(vTuThang)) <> "" And Format_ddmmyyyy1("01/" & Trim(vDenThang)) <> "" Then
                            If DateDiff("M", Format(vTuThang, "mm/yyyy"), Format(vDenThang, "mm/yyyy")) < 0 Then
                                    .Col = .ColLetterToNumber("W")
                                    .CellNote = vErrTuThang
                                    .BackColor = mErrorColor
                                    
                                    .Col = .ColLetterToNumber("X")
                                    .CellNote = vErrTuThang
                                    .BackColor = mErrorColor
                                    iFlagTuThang = 1
                            Else
                                    .Col = .ColLetterToNumber("W")
                                    .CellNote = ""
                                    .BackColor = mNonErrorColor
        '
                                    .Col = .ColLetterToNumber("X")
                                    .CellNote = ""
                                    .BackColor = mNonErrorColor
                            End If
                      End If
                 End If
                
                
                ' kiem tra bat buoc nhap cot 14
                If Trim(vQuanHeNPT) = "" Then
                    .Col = .ColLetterToNumber("L")
                    .CellNote = vErrEmpty
                    .BackColor = mErrorColor
                    iFlagEmpty = 1
                End If
                
                 ' kiem tra bat buoc nhap cot so
'                If Trim(vSo) = "" Then
'                    .Col = .ColLetterToNumber("M")
'                    .CellNote = vErrEmpty
'                    .BackColor = mErrorColor
'                    iFlagEmpty = 1
'                End If
                
                 ' kiem tra bat buoc nhap cot quyen so
'                If Trim(vQuyenSo) = "" Then
'                    .Col = .ColLetterToNumber("N")
'                    .CellNote = vErrEmpty
'                    .BackColor = mErrorColor
'                    iFlagEmpty = 1
'                End If
                
                ' kiem tra khong co cot 11 va 13 bat buoc nhap cot 18,19,20
'                If Trim(vMaQuocGia) = "01" Then
'                    ' bat buoc nhap thong tin tinh
'                    If Trim(vTinh) = "" Then
'                        .Col = .ColLetterToNumber("R")
'                        .CellNote = vErrEmpty
'                        .BackColor = mErrorColor
'                        iFlagEmpty = 1
'                    End If
'
'                    ' bat buoc nhap thong tin huyen
'                    If Trim(vHuyen) = "" Then
'                        .Col = .ColLetterToNumber("T")
'                        .CellNote = vErrEmpty
'                        .BackColor = mErrorColor
'                        iFlagEmpty = 1
'                    End If
'
'                    ' bat buoc nhap thong tin xa
'                    If Trim(vXa) = "" Then
'                        .Col = .ColLetterToNumber("V")
'                        .CellNote = vErrEmpty
'                        .BackColor = mErrorColor
'                        iFlagEmpty = 1
'                    End If
'                End If
                
                ' kiem tra neu chon quoc gia Viet Nam
'                If Trim(vMaQuocGia) = "01" Then
'                    If Trim$(vMaTinh) = "000" Then
'                        .Col = .ColLetterToNumber("S")
'                        .CellNote = vErrDuLieuKhongHL
'                        .BackColor = mErrorColor
'                        iFlagEmpty = 1
'                    End If
'
'                    If Trim$(vMaHuyen) = "00000" Then
'                        .Col = .ColLetterToNumber("U")
'                        .CellNote = vErrDuLieuKhongHL
'                        .BackColor = mErrorColor
'                        iFlagEmpty = 1
'                    End If
'
'                    If Trim$(vMaXa) = "0000000" Then
'                        .Col = .ColLetterToNumber("W")
'                        .CellNote = vErrDuLieuKhongHL
'                        .BackColor = mErrorColor
'                        iFlagEmpty = 1
'                    End If
'                End If
            End If
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
        
        
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 71, IIf(iFlagHoVaTenNNT = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 72, IIf(iFlagMSTNNT = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 73, IIf(iFlagHoVaTenNPT = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 74, IIf(iFlagNgaySinhNPT = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 75, IIf(iFlagMSTNPT = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 76, IIf(iFlagSoCMNDNPT = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 77, IIf(iFlagTuThang = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 78, IIf(iFlagDenThang = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 79, IIf(iFlagEmpty = 1, "0", "1")
        
        .SetText .ColLetterToNumber("B"), 86, IIf(iFlagMstDN = 1, "0", "1")
    End With
End Sub



Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
 
End Sub


Private Sub prepareData05_3()
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCells As MSXML.IXMLDOMNodeList
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    
    Dim strDataFileName As String
    Dim i, j, k As Long
    
    Dim fso As New FileSystemObject
    Dim fle As File
    Dim countRow As Long
    
    Dim rowGrid As Long
    
    Dim arrFileName() As String
    
    ReDim arrStrXMLFileNames(0)
    LoadXMLFileNames
    
    countRow = 0
    
    For k = 0 To UBound(arrStrXMLFileNames)
        arrFileName = Split(arrStrXMLFileNames(k), "_")
        If UBound(arrFileName) <> -1 Then
            strDataFileName = arrStrXMLFileNames(k)
            If xmlDomData.Load(TAX_Utilities_v2.DataFolder & strDataFileName & ".xml") Then
                Set xmlNodeListCells = xmlDomData.getElementsByTagName("Cells")
                countRow = countRow + xmlNodeListCells.length
            End If
        End If
    Next
    
    If countRow > 0 Then
        ReDim arrData05_3(countRow, 22) As String
        i = 0
        For k = 0 To UBound(arrStrXMLFileNames)
            arrFileName = Split(arrStrXMLFileNames(k), "_")
            If UBound(arrFileName) <> -1 Then
                strDataFileName = arrStrXMLFileNames(k)
                If xmlDomData.Load(TAX_Utilities_v2.DataFolder & strDataFileName & ".xml") Then
                    Set xmlNodeListCells = xmlDomData.getElementsByTagName("Cells")
                    
                     For Each xmlNodeCells In xmlNodeListCells
                        Set xmlNodeListCell = xmlNodeListCells.Item(i).childNodes
                        j = 0
                        For Each xmlNodeCell In xmlNodeListCell
                             arrData05_3(i, j) = GetAttribute(xmlNodeCell, "Value")
                            j = j + 1
                        Next
                        i = i + 1
                     Next
                End If
            End If
        Next
        
        Set xmlDomData = Nothing
        Set xmlNodeListCells = Nothing
        Set xmlNodeListCell = Nothing
        Set xmlNodeCells = Nothing
        Set xmlNodeCell = Nothing
    Else
        ReDim arrData05_3(1, 22) As String
        arrData05_3(0, 0) = ""
        arrData05_3(0, 1) = ""
        arrData05_3(0, 2) = ""
        arrData05_3(0, 3) = ""
        arrData05_3(0, 4) = ""
        arrData05_3(0, 5) = ""
        arrData05_3(0, 6) = ""
        arrData05_3(0, 7) = ""
        arrData05_3(0, 8) = ""
        arrData05_3(0, 9) = ""
        arrData05_3(0, 10) = ""
        arrData05_3(0, 11) = ""
        arrData05_3(0, 12) = ""
        arrData05_3(0, 13) = ""
        arrData05_3(0, 14) = ""
        arrData05_3(0, 15) = ""
        arrData05_3(0, 16) = ""
        arrData05_3(0, 17) = ""
        arrData05_3(0, 18) = ""
        arrData05_3(0, 19) = ""
        arrData05_3(0, 20) = ""
        arrData05_3(0, 21) = ""
    End If
    
    Dim totalRow As Long
    totalRow = UBound(arrData05_3) - 1
    rowGroup2 = 0
    rowGroup1 = 0
    For i = 0 To totalRow
        If Trim(arrData05_3(i, 4)) <> "" Then
            rowGroup2 = rowGroup2 + 1
        Else
            rowGroup1 = rowGroup1 + 1
        End If
    Next
End Sub


Private Sub fillData()
    With fps
        .Refresh
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        .Row = 40

        Dim totalRow As Long
        Dim strCol As String
        Dim ArrCol() As String
        Dim idx As Long
        Dim i As Long
        Dim j As Long
        Dim STT As Long
        Dim isGroup1 As Boolean

        strCol = "C_D_E_F_G_H_I_J_K_L_M_N_O_P_Q_R_S_T_U_V_W_X"
        ArrCol = Split(strCol, "_")
        isGroup1 = False

        totalRow = UBound(arrData05_3) - 1
        .Row = 40
        STT = 1
        For i = 0 To totalRow
            If Trim(arrData05_3(i, 4)) <> "" Then
            Else
                isGroup1 = True
                ' set stt
                .Col = .ColLetterToNumber("B")
                .Text = STT
                
                For j = 0 To 21
                    ' Tu cot Ho va ten cho den cot thue da khau tru (C -> L). Dung cac thong tin theo bang ke 05-3
                    If ArrCol(j) = "G" Then
                        ' cot MST group 1 khong tai du lieu
                    Else
                        .Col = .ColLetterToNumber(ArrCol(j))
                        .Text = arrData05_3(i, j)
                    End If
                Next
                STT = STT + 1
                .Row = .Row + 1
            End If
        Next

        If isGroup1 = True Then
            idx = .Row + 8
            .Row = .Row + 8
        Else
            idx = .Row + 9
            .Row = .Row + 9
        End If
        STT = 1
        For i = 0 To totalRow
            If Trim(arrData05_3(i, 4)) <> "" Then
                 ' set stt
                .Col = .ColLetterToNumber("B")
                .Text = STT
                
                For j = 0 To 21
                    If ArrCol(j) = "F" Or ArrCol(j) = "H" Or ArrCol(j) = "I" Or ArrCol(j) = "J" Or ArrCol(j) = "M" Or ArrCol(j) = "N" _
                    Or ArrCol(j) = "O" Or ArrCol(j) = "P" Or ArrCol(j) = "Q" Or ArrCol(j) = "R" Or ArrCol(j) = "S" Or ArrCol(j) = "T" _
                    Or ArrCol(j) = "U" Or ArrCol(j) = "V" Then
                    Else
                        ' Tu cot Ho va ten cho den cot thue da khau tru (C -> L). Dung cac thong tin theo bang ke 05-3
                        .Col = .ColLetterToNumber(ArrCol(j))
                        .Text = arrData05_3(i, j)
                    End If
                Next
                STT = STT + 1
                .Row = .Row + 1
            End If
        Next

        .EventEnabled(EventAllEvents) = True
        .Refresh
    End With
End Sub



Private Sub LoadXMLFileNames()
    Dim lngIndex As Long
    Dim fso As New FileSystemObject
    Dim fle As File
    
    For Each fle In fso.GetFolder(GetAbsolutePath(TAX_Utilities_v2.DataFolder)).Files
        If Right$(fle.Name, 4) = ".xml" And InStr(fle.Name, "05_3_TNCN_" & TAX_Utilities_v2.Year & "_") > 0 Then
            ReDim Preserve arrStrXMLFileNames(lngIndex)
            arrStrXMLFileNames(lngIndex) = Mid$(fle.Name, 1, Len(fle.Name) - 4)
            lngIndex = lngIndex + 1
        End If
    Next
End Sub



'Private Sub updateData(ByVal Row As Long, ByVal rowData As Long)
'    Dim varTemp As Variant
'    Dim lCol As Long, lRow As Long
'    Dim xmlCellNode As MSXML.IXMLDOMNodeList
'    Dim xmlCellNodeData As MSXML.IXMLDOMNode
'    With fps
'        .Sheet = 1
'        mCurrentSheet = .Sheet
'        .Row = Row
'        ' Cot D la cot chua Ho ten nguoi quyet toan
'        .Col = .ColLetterToNumber("D")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot E la cot MST cua nguoi quyet toan
'        .Col = .ColLetterToNumber("E")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot F la tong thu nhap chiu thue cua nguoi quyet toan
'        .Col = .ColLetterToNumber("F")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot G la thu nhap luong cua nguoi quyet toan
'        .Col = .ColLetterToNumber("G")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot H la cot thu nhap thuong cua nguoi quyet toan
'        .Col = .ColLetterToNumber("H")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot I la cot thu nhap khac cua nguoi quyet toan
'        .Col = .ColLetterToNumber("I")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot J la tong so nguoi phu thuoc cua nguoi quyet toan
'        .Col = .ColLetterToNumber("J")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot K la cot tong so thang giam tru cua nguoi quyet toan
'        .Col = .ColLetterToNumber("K")
'        UpdateCell fps, .Col, rowData, .value
'
'        'Cot L Thue da khau tru cua cua nguoi quyet toan
'        .Col = .ColLetterToNumber("L")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot M la cot thu nhap chiu thue cua nguoi quyet toan
'        .Col = .ColLetterToNumber("M")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot N la cot giam tru cho ban than cua nguoi quyet toan
'        .Col = .ColLetterToNumber("N")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot O la cot giam tru cho nguoi phu thuoc cua nguoi quyet toan
'        .Col = .ColLetterToNumber("O")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot P la cot thu nhap tinh thue cua nguoi quyet toan
'        .Col = .ColLetterToNumber("P")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot Q la cot thue phai nop cua nguoi quyet toan
'        .Col = .ColLetterToNumber("Q")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot R la cot thue mien giam cua nguoi quyet toan
'        .Col = .ColLetterToNumber("R")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot S la cot thu nhap khac cua nguoi quyet toan
'        .Col = .ColLetterToNumber("S")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot T la cot so thue da khau tru cua nguoi quyet toan
'        .Col = .ColLetterToNumber("T")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot T la cot so thue da khau tru cua nguoi quyet toan
'        .Col = .ColLetterToNumber("U")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot V la cot thu nhap do lam viec trong khu kinh te cua nguoi quyet toan
'        .Col = .ColLetterToNumber("V")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot T la cot so thue da khau tru cua nguoi quyet toan
'        .Col = .ColLetterToNumber("W")
'        UpdateCell fps, .Col, rowData, .value
'
'        ' Cot T la cot so thue da khau tru cua nguoi quyet toan
'        .Col = .ColLetterToNumber("X")
'        UpdateCell fps, .Col, rowData, .value
'    End With
'
'    Set xmlCellNodeData = Nothing
'    Set xmlCellNode = Nothing
'End Sub


Public Sub updateData()
    UpdateSheets
End Sub



Private Sub updateMSTDL_16TH()
    Dim lRow As Long, lCol As Long
    Dim mstDL As Variant
    Dim xmlNodeList As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlDom2 As New MSXML.DOMDocument
    Dim value As String
    
    With fps
    
        .Sheet = 1
        If isCheckTTDLT = True Then
            'check MST DaiLy, ten NVDLT, chungchi hanh nghe da ke khai tu lan truoc chua
            xmlDom2.Load TAX_Utilities_v2.DataFolder & "\Header_01.xml"
            Set xmlNodeList = xmlDom2.getElementsByTagName("Cell")
            
            'get MSTDL vu
            .GetText .ColLetterToNumber("D"), 20, mstDL
            If Trim(mstDL) = vbNullString Then
                 Set xmlNode = xmlNodeList.Item(32)
                value = GetAttribute(xmlNode, "Value")
                .SetText lCol, lRow, value
                UpdateCell fps, lCol, lRow, value
            End If
        Else
            Exit Sub
        End If
    End With
End Sub


Private Function getHuyenByMaTinh(maTinh As String) As String
    Dim strHuyen As String
    Dim strTenHuyen As String
    Dim strMaHuyen As String
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
    Dim temValue As String
    Dim arrValue() As String
    Set xmlDocument = Nothing
    Set xmlNodeListMap = Nothing
    ' lay du lieu danh muc dia ban tinh
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Huyen_" & maTinh & ".xml"))
    Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
    For Each xmlNode In xmlNodeListMap
        temValue = GetAttribute(xmlNode, "Value")
        arrValue = Split(temValue, "###")
        If arrValue(1) = maTinh Then
            strTenHuyen = strTenHuyen + arrValue(3) + Chr$(9)
            strMaHuyen = strMaHuyen + arrValue(2) + Chr$(9)
        End If
    Next
    strHuyen = strMaHuyen & "###" & strTenHuyen
    getHuyenByMaTinh = strHuyen
End Function


Private Function getXaByMaHuyen(maHuyen As String) As String
    Dim strXa As String
    Dim strTenXa As String
    Dim strMaXa As String
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
    Dim temValue As String
    Dim arrValue() As String
    Set xmlDocument = Nothing
    Set xmlNodeListMap = Nothing
    ' lay du lieu danh muc dia ban tinh
    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Xa" & "_" & Left$(maHuyen, 3) & ".xml"))
    Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
    For Each xmlNode In xmlNodeListMap
        temValue = GetAttribute(xmlNode, "Value")
        arrValue = Split(temValue, "###")
        temValue = GetAttribute(xmlNode, "Value")
        arrValue = Split(temValue, "###")
        If arrValue(2) = maHuyen Then
            strTenXa = strTenXa + arrValue(4) + Chr$(9)
            strMaXa = strMaXa + arrValue(3) + Chr$(9)
        End If
    Next
    strXa = strMaXa & "###" & strTenXa
    getXaByMaHuyen = strXa
End Function

Private Sub ForMatTK16TH()
    With fps
        .EventEnabled(EventAllEvents) = False
            .Sheet = 1
            ' Lay danh muc quan he voi NNT
            Dim xmlDocument As New MSXML.DOMDocument
            Dim xmlNode As MSXML.IXMLDOMNode
            xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_QHNNT.xml"))
            Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
            Dim strQuanHeNNTID As String
            Dim strQuanHeNNT As String
            Dim i As Long
    
            Dim strQuocTich As String
            Dim strMaQuocTich As String
            Dim strTenTinh As String
            Dim strTenHuyen As String
            Dim strTenXa As String
            Dim strMaTinh As String
            Dim strMaHuyen  As String
            Dim strMaXa As String
            Dim temValue As String
            Dim arrValue() As String
                
            Dim maTinh As Variant
            Dim maHuyen As Variant
            Dim tempArr() As String
    
            Dim index As Integer
            
            Dim lrowCount1 As Double
            Dim lrowCount2 As Double
            Dim rowStartSpread11 As Double
            
            Dim backGroudColor As Variant
            
            rowStartSpread11 = 40
            ' Chuoi gia tri ve quan he NNT bao gom:
            ' 01-Con, 02-Vo, 03-Chong, 04-Cha, 05-Me, 06-Chau Ruot, 07-Ho Hang, 99-Khac
            Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListMap
                arrValue = Split(GetAttribute(xmlNode, "value"), "###")
                strQuanHeNNT = strQuanHeNNT + arrValue(1) + Chr$(9)
                strQuanHeNNTID = strQuanHeNNTID + arrValue(0) + Chr$(9)
            Next
    
    
            ' lay du lieu danh muc dia ban tinh
            xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Tinh.xml"))
            Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListMap
                temValue = GetAttribute(xmlNode, "Value")
                arrValue = Split(temValue, "###")
                'If arrValue(0) <> "1" Then
                    strTenTinh = strTenTinh + arrValue(2) + Chr$(9)
                    strMaTinh = strMaTinh + arrValue(1) + Chr$(9)
                'End If
            Next
    
            ' lay danh sach huyen
            Set xmlDocument = Nothing
            Set xmlNodeListMap = Nothing
            ' lay du lieu danh muc dia ban tinh
            xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Huyen.xml"))
            Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListMap
                temValue = GetAttribute(xmlNode, "Value")
                arrValue = Split(temValue, "###")
                'If arrValue(0) <> "1" Then
                    strTenHuyen = strTenHuyen + arrValue(3) + Chr$(9)
                    strMaHuyen = strMaHuyen + arrValue(2) + Chr$(9)
                'End If
            Next
    
            ' Lay danh muc xa
            Set xmlDocument = Nothing
            Set xmlNodeListMap = Nothing
            ' lay du lieu danh muc dia ban tinh
            xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Xa.xml"))
            Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListMap
                temValue = GetAttribute(xmlNode, "Value")
                arrValue = Split(temValue, "###")
                'If arrValue(0) <> "1" Then
                    strTenXa = strTenXa + arrValue(4) + Chr$(9)
                    strMaXa = strMaXa + arrValue(3) + Chr$(9)
                'End If
            Next
    
            ' lay du lieu dia ban
            xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Quoc_Gia.xml"))
            Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
            strMaQuocTich = ""
            For Each xmlNode In xmlNodeListMap
                temValue = GetAttribute(xmlNode, "Value")
                arrValue = Split(temValue, "###")
                strQuocTich = strQuocTich + arrValue(1) + Chr$(9)
                strMaQuocTich = strMaQuocTich + arrValue(0) + Chr$(9)
            Next
    
            
    
            .Sheet = 1
            ' tinh so row group1, group2
            For lrowCount2 = .MaxRows To 40 Step -1
                .Row = lrowCount2
                .Col = .ColLetterToNumber("B")
                If .Text = "aa" Then
                    Exit For
                End If
            Next
                            
            For lrowCount1 = lrowCount2 - 1 To 40 Step -1
                .Row = lrowCount1
                .Col = .ColLetterToNumber("B")
                If .Text = "aa" Then
                    Exit For
                End If
            Next
                            
         ' set format group 1
         .SetCellBorder .ColLetterToNumber("B"), 40, .ColLetterToNumber("Y"), lrowCount1 - 1, 15, &O0, CellBorderStyleSolid
                            
         .Row = 40
         .Col = .ColLetterToNumber("G")
          backGroudColor = .BackColor
                            
         .Row = 40
        .Col = .ColLetterToNumber("B")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("B")
        .BlockMode = True
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
            'format chi tieu [7],[C]
        .Row = 40
        .Col = .ColLetterToNumber("C")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("C")
        .BlockMode = True
        .TypeMaxEditLen = 100
        .BlockMode = False
        
        'format chi tieu [8],[D]
        .Row = 40
        .Col = .ColLetterToNumber("D")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("D")
        .BlockMode = True
        .TypeMaxEditLen = 10
        .BlockMode = False
        
        'format chi tieu [9],[E]
        .Row = 40
        .Col = .ColLetterToNumber("E")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("E")
        .BlockMode = True
        .TypeMaxEditLen = 100
        .BlockMode = False
        
        'format chi tieu [10],[F]
        .Row = 40
        .Col = .ColLetterToNumber("F")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("F")
        .BlockMode = True
        .TypeMaxEditLen = 10
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
        'format chi tieu [10],[G]
        .Row = 40
        .Col = .ColLetterToNumber("G")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("G")
        .BlockMode = True
        .BackColor = backGroudColor
        .BlockMode = False
        
        'format chi tieu [13],[J]
        .Row = 40
        .Col = .ColLetterToNumber("J")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("J")
        .BlockMode = True
        .TypeMaxEditLen = 50
        .BlockMode = False
        
        
        'format chi tieu [15],[M]
        .Row = 40
        .Col = .ColLetterToNumber("M")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("M")
        .BlockMode = True
        .TypeMaxEditLen = 8
        .BlockMode = False
        
        'format chi tieu [16],[N]
        .Row = 40
        .Col = .ColLetterToNumber("N")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("N")
        .BlockMode = True
        .TypeMaxEditLen = 7
        .BlockMode = False
        
        'format chi tieu [21],[X]
        .Row = 40
        .Col = .ColLetterToNumber("W")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("X")
        .BlockMode = True
        .TypeMaxEditLen = 7
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
        
        'format chi tieu [12],[I]
        .Row = 40
        .Col = .ColLetterToNumber("I")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("I")
        .BlockMode = True
        .CellType = CellTypeComboBox
        .BlockMode = False


        'format chi tieu ma quoc tich [H]
        .Row = 40
        .Col = .ColLetterToNumber("H")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("H")
        .BlockMode = True
        .CellType = CellTypeComboBox
        .BlockMode = False
        
        'format chi tieu [14],[L]
        .Row = 40
        .Col = .ColLetterToNumber("L")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("L")
        .BlockMode = True
        .CellType = CellTypeComboBox
        .BlockMode = False
        
        
        'format chi tieu ma quan he [K]
        .Row = 40
        .Col = .ColLetterToNumber("K")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("K")
        .BlockMode = True
        .CellType = CellTypeComboBox
        .BlockMode = False
        
        'format chi tieu [14],[P]
        .Row = 40
        .Col = .ColLetterToNumber("O")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("V")
        .BlockMode = True
        .CellType = CellTypeComboBox
        .BlockMode = False
        
        
        .Row = 40
        .Col = .ColLetterToNumber("C")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("X")
        .BlockMode = True
        .FontSize = 8
        .Lock = False
        .BlockMode = False
        
        .Row = 40
        .Col = .ColLetterToNumber("G")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("G")
        .BlockMode = True
        .FontSize = 8
        .Lock = True
        .BlockMode = False
                                
        .Row = 40
        .Col = .ColLetterToNumber("B")
        .Row2 = lrowCount1 - 1
        .col2 = .ColLetterToNumber("B")
        .BlockMode = True
        .FontSize = 8
        .Lock = True
        .BlockMode = False
                                
         ' end group 1
         .SetCellBorder .ColLetterToNumber("B"), lrowCount1 + 8, .ColLetterToNumber("Y"), lrowCount2, 15, &O0, CellBorderStyleSolid
         ' set format group 2
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("B")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("B")
        .BlockMode = True
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
'       format chi tieu [7],[C]
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("C")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("C")
        .BlockMode = True
        .TypeMaxEditLen = 100
        .BlockMode = False
'
'        format chi tieu [8],[D]
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("D")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("D")
        .BlockMode = True
        .TypeMaxEditLen = 10
        .BlockMode = False
'
'        format chi tieu [9],[E]
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("E")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("E")
        .BlockMode = True
        .TypeMaxEditLen = 100
        .BlockMode = False
        
        
'        format chi tieu [9],[F]
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("F")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("F")
        .BlockMode = True
        .BackColor = backGroudColor
        .BlockMode = False
        
'        format chi tieu [9],[I], [J]
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("I")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("J")
        .BlockMode = True
        .BackColor = backGroudColor
        .BlockMode = False
        
'        format chi tieu [9],[M], [V]
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("M")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("V")
        .BlockMode = True
        .BackColor = backGroudColor
        .BlockMode = False
        
        'format chi tieu [10],[G]
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("G")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("G")
        .BlockMode = True
        .TypeMaxEditLen = 10
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
         'format chi tieu [21],[X]
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("W")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("X")
        .BlockMode = True
        .TypeMaxEditLen = 7
        .TypeHAlign = TypeHAlignCenter
        .BlockMode = False
        
'        format chi tieu [12],[L]
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("L")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("L")
        .BlockMode = True
        .CellType = CellTypeComboBox
        .BlockMode = False
                
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("C")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("E")
        .BlockMode = True
        .FontSize = 8
        .Lock = False
        .BlockMode = False
        
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("G")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("G")
        .BlockMode = True
        .FontSize = 8
        .Lock = False
        .BlockMode = False
        
        
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("L")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("L")
        .BlockMode = True
        .FontSize = 8
        .Lock = False
        .BlockMode = False
        
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("W")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("X")
        .BlockMode = True
        .FontSize = 8
        .Lock = False
        .BlockMode = False
        
        .Row = lrowCount1 + 8
        .Col = .ColLetterToNumber("B")
        .Row2 = lrowCount2 - 1
        .col2 = .ColLetterToNumber("B")
        .BlockMode = True
        .FontSize = 8
        .Lock = True
        .BlockMode = False
        copyFormulasSheet1 lrowCount2 + lrowCount1 + 9, fps, 40
            
            ' group 1
            i = 1
            .Col = .ColLetterToNumber("B")
            .Row = 40
            .EventEnabled(EventAllEvents) = False
            .AutoCalc = False
            Do
                .RowHeight(-2) = 16.5

                .Col = .ColLetterToNumber("K")
                .TypeComboBoxList = strQuanHeNNTID

                .Col = .ColLetterToNumber("L")
                .TypeComboBoxList = strQuanHeNNT

                    ' set danh sach tinh
                .Col = .ColLetterToNumber("Q")
                .TypeComboBoxList = strMaTinh

                .Col = .ColLetterToNumber("R")
                .TypeComboBoxList = strTenTinh

                ' set danh sach huyen
                .GetText .ColLetterToNumber("Q"), .Row, maTinh
                tempArr = Split(getHuyenByMaTinh(CStr(maTinh)), "###")

                .Col = .ColLetterToNumber("S")
                .TypeComboBoxList = tempArr(0)

                .Col = .ColLetterToNumber("T")
                .TypeComboBoxList = tempArr(1)

                ' set danh sach xa
                .GetText .ColLetterToNumber("S"), .Row, maHuyen
                tempArr = Split(getXaByMaHuyen(CStr(maHuyen)), "###")
                .Col = .ColLetterToNumber("U")
                .TypeComboBoxList = tempArr(0)

                .Col = .ColLetterToNumber("V")
                .TypeComboBoxList = tempArr(1)

                 'set danh sach quoc tich
                .Col = .ColLetterToNumber("H")
                .TypeComboBoxList = strMaQuocTich

                .Col = .ColLetterToNumber("I")
                .TypeComboBoxList = strQuocTich


                ' set danh sach quoc gia
                .Col = .ColLetterToNumber("O")
                .TypeComboBoxList = strMaQuocTich

                .Col = .ColLetterToNumber("P")
                .TypeComboBoxList = strQuocTich
'
                 .Col = .ColLetterToNumber("B")
                 '.Formula = "B" & .Row - 1 & "+1"
                 .Text = i
                 i = i + 1
                 .Row = .Row + 1
            Loop Until .Text = "aa"
            
            ' group 2
            i = 1
            .Col = .ColLetterToNumber("B")
            .Row = lrowCount1 + 8
            .EventEnabled(EventAllEvents) = False
            .AutoCalc = False
            Do
                .RowHeight(-2) = 16.5

                .Col = .ColLetterToNumber("K")
                .TypeComboBoxList = strQuanHeNNTID

                .Col = .ColLetterToNumber("L")
                .TypeComboBoxList = strQuanHeNNT

'                    ' set danh sach tinh
'                .Col = .ColLetterToNumber("Q")
'                .TypeComboBoxList = strMaTinh
'
'                .Col = .ColLetterToNumber("R")
'                .TypeComboBoxList = strTenTinh
'
'                ' set danh sach huyen
'                .GetText .ColLetterToNumber("Q"), .Row, maTinh
'                tempArr = Split(getHuyenByMaTinh(CStr(maTinh)), "###")
'
'                .Col = .ColLetterToNumber("S")
'                .TypeComboBoxList = tempArr(0)
'
'                .Col = .ColLetterToNumber("T")
'                .TypeComboBoxList = tempArr(1)
'
'                ' set danh sach xa
'                .GetText .ColLetterToNumber("S"), .Row, maHuyen
'                tempArr = Split(getXaByMaHuyen(CStr(maHuyen)), "###")
'                .Col = .ColLetterToNumber("U")
'                .TypeComboBoxList = tempArr(0)
'
'                .Col = .ColLetterToNumber("V")
'                .TypeComboBoxList = tempArr(1)
'
'                 'set danh sach quoc tich
'                .Col = .ColLetterToNumber("H")
'                .TypeComboBoxList = strMaQuocTich
'
'                .Col = .ColLetterToNumber("I")
'                .TypeComboBoxList = strQuocTich
'
'
'                ' set danh sach quoc gia
'                .Col = .ColLetterToNumber("O")
'                .TypeComboBoxList = strMaQuocTich
'
'                .Col = .ColLetterToNumber("P")
'                .TypeComboBoxList = strQuocTich
'
                 .Col = .ColLetterToNumber("B")
                 '.Formula = "B" & .Row - 1 & "+1"
                 .Text = i
                 i = i + 1
                 .Row = .Row + 1
            Loop Until .Text = "aa"
            
            
             .AutoCalc = True
            .EventEnabled(EventAllEvents) = True
      End With
End Sub


Private Sub copyFormulasSheet1(numRow As Long, fps As fpSpread, rowStart As Long)
    Dim a As Long
    a = 0

    With fps

        .Sheet = 1
            
        'truong hop so ban ghe len hon 10000
        If numRow >= 2000 Then

            Do While a * 2 <= 1024

                If a = 0 Then
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a, .ColLetterToNumber("A"), (rowStart + a + 1)
                    '.CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a, .ColLetterToNumber("B"), (rowStart + a + 1)
                    '.CopyRange .ColLetterToNumber("G"), rowStart, .ColLetterToNumber("G"), rowStart + a, .ColLetterToNumber("G"), (rowStart + a + 1)
                    '.CopyRange .ColLetterToNumber("I"), rowStart, .ColLetterToNumber("M"), rowStart + a, .ColLetterToNumber("I"), (rowStart + a + 1)
                    '.CopyRange .ColLetterToNumber("P"), rowStart, .ColLetterToNumber("Y"), rowStart + a, .ColLetterToNumber("P"), (rowStart + a + 1)
                    '.CopyRange .ColLetterToNumber("X"), rowStart, .ColLetterToNumber("Y"), rowStart + a, .ColLetterToNumber("X"), (rowStart + a + 1)
                    '.CopyRange .ColLetterToNumber("AH"), rowStart, .ColLetterToNumber("AH"), rowStart + a, .ColLetterToNumber("AH"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("AI"), rowStart, .ColLetterToNumber("AI"), rowStart + a, .ColLetterToNumber("AI"), (rowStart + a + 1)
                    a = a + 2
                ElseIf a <> 0 Then
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a - 1, .ColLetterToNumber("A"), rowStart + a
                    '.CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a - 1, .ColLetterToNumber("B"), rowStart + a
                    '.CopyRange .ColLetterToNumber("G"), rowStart, .ColLetterToNumber("G"), rowStart + a - 1, .ColLetterToNumber("G"), rowStart + a
                    '.CopyRange .ColLetterToNumber("I"), rowStart, .ColLetterToNumber("M"), rowStart + a - 1, .ColLetterToNumber("I"), rowStart + a
                    '.CopyRange .ColLetterToNumber("P"), rowStart, .ColLetterToNumber("Y"), rowStart + a - 1, .ColLetterToNumber("P"), rowStart + a
                    '.CopyRange .ColLetterToNumber("X"), rowStart, .ColLetterToNumber("Y"), rowStart + a - 1, .ColLetterToNumber("X"), rowStart + a
                    '.CopyRange .ColLetterToNumber("AH"), rowStart, .ColLetterToNumber("AH"), rowStart + a - 1, .ColLetterToNumber("AH"), rowStart + a
                    .CopyRange .ColLetterToNumber("AI"), rowStart, .ColLetterToNumber("AI"), rowStart + a - 1, .ColLetterToNumber("AI"), rowStart + a
                    a = a * 2
                End If

            Loop
                 
            a = 1
            Dim dem As Long
            Dim du  As Long
            dem = numRow \ 1024
            du = numRow Mod 1024

            If dem > 0 Then

                Do While a < dem
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), 1024 + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * a
                    '.CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), 1024 + rowStart - 1, .ColLetterToNumber("B"), rowStart + 1024 * a
                    '.CopyRange .ColLetterToNumber("G"), rowStart, .ColLetterToNumber("G"), 1024 + rowStart - 1, .ColLetterToNumber("G"), rowStart + 1024 * a
                    '.CopyRange .ColLetterToNumber("I"), rowStart, .ColLetterToNumber("M"), 1024 + rowStart - 1, .ColLetterToNumber("I"), rowStart + 1024 * a
                    '.CopyRange .ColLetterToNumber("P"), rowStart, .ColLetterToNumber("Y"), 1024 + rowStart - 1, .ColLetterToNumber("P"), rowStart + 1024 * a
                    '.CopyRange .ColLetterToNumber("X"), rowStart, .ColLetterToNumber("Y"), 1024 + rowStart - 1, .ColLetterToNumber("X"), rowStart + 1024 * a
                    '.CopyRange .ColLetterToNumber("AH"), rowStart, .ColLetterToNumber("AH"), 1024 + rowStart - 1, .ColLetterToNumber("AH"), rowStart + 1024 * a
                    .CopyRange .ColLetterToNumber("AI"), rowStart, .ColLetterToNumber("AI"), 1024 + rowStart - 1, .ColLetterToNumber("AI"), rowStart + 1024 * a
                    a = a + 1
                Loop

                .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), du + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * a
                '.CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), du + rowStart - 1, .ColLetterToNumber("B"), rowStart + 1024 * a
                '.CopyRange .ColLetterToNumber("G"), rowStart, .ColLetterToNumber("G"), du + rowStart - 1, .ColLetterToNumber("G"), rowStart + 1024 * a
                '.CopyRange .ColLetterToNumber("I"), rowStart, .ColLetterToNumber("M"), du + rowStart - 1, .ColLetterToNumber("I"), rowStart + 1024 * a
                '.CopyRange .ColLetterToNumber("P"), rowStart, .ColLetterToNumber("Y"), du + rowStart - 1, .ColLetterToNumber("P"), rowStart + 1024 * a
                '.CopyRange .ColLetterToNumber("X"), rowStart, .ColLetterToNumber("Y"), du + rowStart - 1, .ColLetterToNumber("X"), rowStart + 1024 * a
                '.CopyRange .ColLetterToNumber("AH"), rowStart, .ColLetterToNumber("AH"), du + rowStart - 1, .ColLetterToNumber("AH"), rowStart + 1024 * a
                .CopyRange .ColLetterToNumber("AI"), rowStart, .ColLetterToNumber("AI"), du + rowStart - 1, .ColLetterToNumber("AI"), rowStart + 1024 * a
            Else
                .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), du + rowStart - 1, .ColLetterToNumber("A"), rowStart + 1024 * (a - 1)
                '.CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), du + rowStart - 1, .ColLetterToNumber("B"), rowStart + 1024 * (a - 1)
                '.CopyRange .ColLetterToNumber("G"), rowStart, .ColLetterToNumber("G"), du + rowStart - 1, .ColLetterToNumber("G"), rowStart + 1024 * (a - 1)
                '.CopyRange .ColLetterToNumber("I"), rowStart, .ColLetterToNumber("M"), du + rowStart - 1, .ColLetterToNumber("I"), rowStart + 1024 * (a - 1)
                '.CopyRange .ColLetterToNumber("P"), rowStart, .ColLetterToNumber("Y"), du + rowStart - 1, .ColLetterToNumber("P"), rowStart + 1024 * (a - 1)
                '.CopyRange .ColLetterToNumber("X"), rowStart, .ColLetterToNumber("Y"), du + rowStart - 1, .ColLetterToNumber("X"), rowStart + 1024 * (a - 1)
                '.CopyRange .ColLetterToNumber("AH"), rowStart, .ColLetterToNumber("AH"), du + rowStart - 1, .ColLetterToNumber("AH"), rowStart + 1024 * (a - 1)
                .CopyRange .ColLetterToNumber("AI"), rowStart, .ColLetterToNumber("AI"), du + rowStart - 1, .ColLetterToNumber("AI"), rowStart + 1024 * (a - 1)
            End If

            ' truong hop nho hon 10000
        ElseIf numRow < 2000 Then
            a = 0

            Do While a * 2 < numRow

                If a = 0 Then
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a, .ColLetterToNumber("A"), (rowStart + a + 1)
                    '.CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a, .ColLetterToNumber("B"), (rowStart + a + 1)
                    '.CopyRange .ColLetterToNumber("G"), rowStart, .ColLetterToNumber("G"), rowStart + a, .ColLetterToNumber("G"), (rowStart + a + 1)
                    '.CopyRange .ColLetterToNumber("I"), rowStart, .ColLetterToNumber("M"), rowStart + a, .ColLetterToNumber("I"), (rowStart + a + 1)
                    '.CopyRange .ColLetterToNumber("P"), rowStart, .ColLetterToNumber("Y"), rowStart + a, .ColLetterToNumber("P"), (rowStart + a + 1)
                    '.CopyRange .ColLetterToNumber("X"), rowStart, .ColLetterToNumber("Y"), rowStart + a, .ColLetterToNumber("X"), (rowStart + a + 1)
                    '.CopyRange .ColLetterToNumber("AH"), rowStart, .ColLetterToNumber("AH"), rowStart + a, .ColLetterToNumber("AH"), (rowStart + a + 1)
                    .CopyRange .ColLetterToNumber("AI"), rowStart, .ColLetterToNumber("AI"), rowStart + a, .ColLetterToNumber("AI"), (rowStart + a + 1)
                    a = a + 2
                ElseIf a <> 0 Then
                    .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + a - 1, .ColLetterToNumber("A"), rowStart + a
                    '.CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + a - 1, .ColLetterToNumber("B"), rowStart + a
                    '.CopyRange .ColLetterToNumber("G"), rowStart, .ColLetterToNumber("G"), rowStart + a - 1, .ColLetterToNumber("G"), rowStart + a
                    '.CopyRange .ColLetterToNumber("I"), rowStart, .ColLetterToNumber("M"), rowStart + a - 1, .ColLetterToNumber("I"), rowStart + a
                    '.CopyRange .ColLetterToNumber("P"), rowStart, .ColLetterToNumber("Y"), rowStart + a - 1, .ColLetterToNumber("P"), rowStart + a
                    '.CopyRange .ColLetterToNumber("X"), rowStart, .ColLetterToNumber("Y"), rowStart + a - 1, .ColLetterToNumber("X"), rowStart + a
                    '.CopyRange .ColLetterToNumber("AH"), rowStart, .ColLetterToNumber("AH"), rowStart + a - 1, .ColLetterToNumber("AH"), rowStart + a
                    .CopyRange .ColLetterToNumber("AI"), rowStart, .ColLetterToNumber("AI"), rowStart + a - 1, .ColLetterToNumber("AI"), rowStart + a
                    a = a * 2
                End If

            Loop
                
            .CopyRange .ColLetterToNumber("A"), rowStart, .ColLetterToNumber("A"), rowStart + (numRow - a - 1), .ColLetterToNumber("A"), rowStart + a
            '.CopyRange .ColLetterToNumber("B"), rowStart, .ColLetterToNumber("B"), rowStart + (numRow - a - 1), .ColLetterToNumber("B"), rowStart + a
            '.CopyRange .ColLetterToNumber("G"), rowStart, .ColLetterToNumber("G"), rowStart + (numRow - a - 1), .ColLetterToNumber("G"), rowStart + a
            '.CopyRange .ColLetterToNumber("I"), rowStart, .ColLetterToNumber("M"), rowStart + (numRow - a - 1), .ColLetterToNumber("I"), rowStart + a
            '.CopyRange .ColLetterToNumber("P"), rowStart, .ColLetterToNumber("Y"), rowStart + (numRow - a - 1), .ColLetterToNumber("P"), rowStart + a
            '.CopyRange .ColLetterToNumber("X"), rowStart, .ColLetterToNumber("Y"), rowStart + (numRow - a - 1), .ColLetterToNumber("X"), rowStart + a
            '.CopyRange .ColLetterToNumber("AH"), rowStart, .ColLetterToNumber("AH"), rowStart + (numRow - a - 1), .ColLetterToNumber("AH"), rowStart + a
            .CopyRange .ColLetterToNumber("AI"), rowStart, .ColLetterToNumber("AI"), rowStart + (numRow - a - 1), .ColLetterToNumber("AI"), rowStart + a
        End If
    End With
End Sub



Private Sub InsertNode16TH(ByVal fRow As Long, ByVal lRow As Long, ByVal mSheet As Byte, ByVal isGroupI As Boolean)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeCells2 As MSXML.IXMLDOMNode
    
    Dim xmlNodeNewCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCells2 As MSXML.IXMLDOMNode
    
    Dim xmlNodeNewCell As MSXML.IXMLDOMNode
    Dim xmlNodeNewCell2 As MSXML.IXMLDOMNode
    
    Dim arrTemp() As String
    Dim xmlSecionNode1 As MSXML.IXMLDOMNode
    Dim xmlSecionNode2 As MSXML.IXMLDOMNode
    
    Dim ctlRow As Long
    Dim cTemplate, cReport As String
    
    Dim sumRowDel As Long
    Dim sumRowDel2 As Long
    
    Dim tempRow As Long
    
    Dim tempCol As Long
    
    If isGroupI = True Then
    
        Set xmlSecionNode1 = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Section")(2)
        sumRowDel = xmlSecionNode1.childNodes.length
        
        Set xmlSecionNode2 = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Section")(3)
        sumRowDel2 = xmlSecionNode2.childNodes.length
        ' del group 1
        For ctlRow = 1 To sumRowDel - 1
                Set xmlNodeCells = TAX_Utilities_v2.Data(mSheet - 1).nodeFromID("C" & "_" & 40 + ctlRow & "").parentNode
                xmlNodeCells.parentNode.removeChild xmlNodeCells
        Next
        ' del group 2
        For ctlRow = 1 To sumRowDel2 - 1
                Set xmlNodeCells = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Section")(3).childNodes(1)
                xmlNodeCells.parentNode.removeChild xmlNodeCells
                
                If ctlRow = sumRowDel2 - 1 Then
                    Set xmlNodeCells = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Section")(3).childNodes(0)
                    
                    For Each xmlNodeNewCell In xmlNodeCells.childNodes
                        cTemplate = Left$(GetAttribute(xmlNodeNewCell, "CellID"), 2)
                        arrTemp = Split(GetAttribute(xmlNodeNewCell, "CellID2"), "_")
                        'cReport = Left$(GetAttribute(xmlNodeNewCell, "CellID2"), 2)
                        cReport = Trim(arrTemp(0)) & "_"
                        
                        SetAttribute xmlNodeNewCell, "CellID", cTemplate & "49"
                                
                        ' Set first cell = 1
                        SetAttribute xmlNodeNewCell, "FirstCell", "0"
                        SetAttribute xmlNodeNewCell, "Value", vbNullString
                        SetAttribute xmlNodeNewCell, "CellID2", cReport & "47"
                    Next
                End If
        Next
        
        
        Set xmlNodeCells = TAX_Utilities_v2.Data(mSheet - 1).nodeFromID("C_40").parentNode
        
        'Set xmlNodeCells2 = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Section")(3).firstChild
        
    ' set row cua group 2
        Set xmlNodeCells2 = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Section")(3).childNodes(0)
                
        For Each xmlNodeNewCell In xmlNodeCells2.childNodes
            cTemplate = Left$(GetAttribute(xmlNodeNewCell, "CellID"), 2)
            arrTemp = Split(GetAttribute(xmlNodeNewCell, "CellID2"), "_")
            'cReport = Left$(GetAttribute(xmlNodeNewCell, "CellID2"), 2)
            cReport = Trim(arrTemp(0)) & "_"
            
            SetAttribute xmlNodeNewCell, "CellID", cTemplate & "100002"
                    
            ' Set first cell = 1
            SetAttribute xmlNodeNewCell, "FirstCell", "0"
            SetAttribute xmlNodeNewCell, "Value", vbNullString
            SetAttribute xmlNodeNewCell, "CellID2", cReport & "100000"
        Next
        ' end set
        
        For ctlRow = fRow To lRow
                                    
            'IncreaseRowInDOM16TH fps, TAX_Utilities_v2.Data(mSheet - 1), ctlRow, 1, 1
                        
            Set xmlNodeNewCells = xmlNodeCells.cloneNode(True)
            
            For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
                cTemplate = Left$(GetAttribute(xmlNodeNewCell, "CellID"), 2)
                arrTemp = Split(GetAttribute(xmlNodeNewCell, "CellID2"), "_")
                'cReport = Left$(GetAttribute(xmlNodeNewCell, "CellID2"), 2)
                cReport = Trim(arrTemp(0)) & "_"
                
                SetAttribute xmlNodeNewCell, "CellID", cTemplate & ctlRow
                        
                ' Set first cell = 1
                SetAttribute xmlNodeNewCell, "FirstCell", "1"
                SetAttribute xmlNodeNewCell, "Value", vbNullString
                SetAttribute xmlNodeNewCell, "CellID2", cReport & (ctlRow - 2)
            Next
            
            xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, Null
            
        
        Next
    Else
'        Set xmlSecionNode1 = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Section")(3)
'        sumRowDel = xmlSecionNode1.childNodes.length
'        ' del group 2
'        For ctlRow = 1 To sumRowDel - 1
'                Set xmlNodeCells = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Section")(3).childNodes(1)
'                xmlNodeCells.parentNode.removeChild xmlNodeCells
'        Next
        ' lay ve row cuoi group 1
        Set xmlSecionNode1 = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Section")(2).lastChild
        ParserCellID fps, GetAttribute(xmlSecionNode1.childNodes(0), "CellID"), tempCol, tempRow
        
        ' set row cua group 2 theo gia tri thuc
        Set xmlNodeCells2 = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Section")(3).childNodes(0)
                
        For Each xmlNodeNewCell In xmlNodeCells2.childNodes
            cTemplate = Left$(GetAttribute(xmlNodeNewCell, "CellID"), 2)
            arrTemp = Split(GetAttribute(xmlNodeNewCell, "CellID2"), "_")
            'cReport = Left$(GetAttribute(xmlNodeNewCell, "CellID2"), 2)
            cReport = Trim(arrTemp(0)) & "_"
            
            SetAttribute xmlNodeNewCell, "CellID", cTemplate & tempRow + 9
                    
            ' Set first cell = 1
            SetAttribute xmlNodeNewCell, "FirstCell", "0"
            SetAttribute xmlNodeNewCell, "Value", vbNullString
            SetAttribute xmlNodeNewCell, "CellID2", cReport & tempRow + 7
        Next
        ' end set
        
        
        
        Set xmlNodeCells = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Section")(3).firstChild
        
        For ctlRow = fRow To lRow
            
            Set xmlNodeNewCells = xmlNodeCells.cloneNode(True)
            
            For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
                cTemplate = Left$(GetAttribute(xmlNodeNewCell, "CellID"), 2)
                arrTemp = Split(GetAttribute(xmlNodeNewCell, "CellID2"), "_")
                'cReport = Left$(GetAttribute(xmlNodeNewCell, "CellID2"), 2)
                cReport = Trim(arrTemp(0)) & "_"
                
                SetAttribute xmlNodeNewCell, "CellID", cTemplate & ctlRow
                        
                ' Set first cell = 1
                SetAttribute xmlNodeNewCell, "FirstCell", "1"
                SetAttribute xmlNodeNewCell, "Value", vbNullString
                SetAttribute xmlNodeNewCell, "CellID2", cReport & (ctlRow - 2)
            Next
            xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, Null
        
        Next
    End If
'   **********************************
'    added
'   Date: 12/04/06
    'mAdjustData = True
    TAX_Utilities_v2.AdjustData(mSheet - 1) = True
'   **********************************
EXIT_SUB:
    Set xmlNodeNewCell = Nothing
    Set xmlNodeNewCells = Nothing
    Set xmlNodeCells = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_16TH_DKNPT", "InsertNode16TH", Err.number, Err.Description
End Sub


Private Sub IncreaseRowInDOM16TH(fpSpread1 As fpSpread, xmlDomData As MSXML.DOMDocument, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, i As Long
        
    If xmlDomData Is Nothing Then Exit Sub
    Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
    
    For i = xmlNodeListCell.length - 1 To xmlNodeListCell.length - 24 Step -1
        ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
        If lRow >= pRow Then
            ' Increase value of row attribute + 1 (CellID)
            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fpSpread1, lCol, lRow + lRows)
            
            ' Increase value of row attribute + 1 (CellID2)
            ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fpSpread1, lCol, lRow + lRow2s)
        End If
    Next
        
    Set xmlNodeListCell = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_16TH_DKNPT", "IncreaseRowInDOM16TH", Err.number, Err.Description
End Sub
