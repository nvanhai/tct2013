VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_05TNCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "To khai quyet toan thu nhap ca nhan mau 04-TNCN" interface sheets
'this Class is belong to TAX_Business_v1 project which will be compline to DLL

Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrSolanBosung As String

Private arrMST5A() As String
Private arrMST5B() As String

Private arrCMT5A() As String
Private arrCMT5B() As String

Public flgLoadToKhai As Boolean

Private totalRow5A, totalRow5B As Long

Private flgErr As Boolean
Private flgErr5A As Boolean
Private flgErr5B As Boolean
Private flgErr5AB As Boolean

' Const ctieu
Private Const Col_D = "D"
Private Const Col_E = "E"
Private Const Col_F = "F"
Private Const Col_G = "G"

Private Const sheetTk = 1
Private Const phuLuc05_1 = 2
Private Const phuLuc05_2 = 3
Private Const phuLuc05_3 = 4
Private Const phuLuc05_27MT = 5

' Lay du lieu BK 05A ca nhan uy quyen QT thuoc b1
Private arrData5A() As String


' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

Private rowData As Long
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    Dim vCell As Variant
    With fps
        SetDateFormat fps, 1, 71, .ColLetterToNumber("M"), DDMMYYYY
        
        .Sheet = 1 'To khai 05/TNCN
        .Row = 71
        .Col = .ColLetterToNumber("M")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        'resetFormula
        ' PL 05-3/TNCN
        SetDateFormat fps, phuLuc05_3, 22, .ColLetterToNumber("X"), MMYYYY
        SetDateFormat fps, phuLuc05_3, 22, .ColLetterToNumber("Y"), MMYYYY
        SetDateFormat fps, phuLuc05_3, 22, .ColLetterToNumber("G"), DDMMYYYY
        
        .Sheet = phuLuc05_3
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
        
        .Row = 22
        .Col = .ColLetterToNumber("M")
        .TypeComboBoxList = strQuanHeNNT
        
        .Col = .ColLetterToNumber("L")
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
        
        .Row = 22
        ' set danh sach tinh
        .Col = .ColLetterToNumber("S")
        .TypeComboBoxList = strTenTinh
        .Col = .ColLetterToNumber("R")
        .TypeComboBoxList = strMaTinh
        ' set danh sach huyen
        .Col = .ColLetterToNumber("U")
        .TypeComboBoxList = strTenHuyen
        .Col = .ColLetterToNumber("T")
        .TypeComboBoxList = strMaHuyen
        ' set danh sach xa
        .Col = .ColLetterToNumber("W")
        .TypeComboBoxList = strTenXa
        .Col = .ColLetterToNumber("V")
        .TypeComboBoxList = strMaXa
        
        ' lay du lieu dia ban
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Quoc_Gia.xml"))
        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListMap
            temValue = GetAttribute(xmlNode, "Value")
            arrValue = Split(temValue, "###")
            strQuocTich = strQuocTich + arrValue(1) + Chr$(9)
            strMaQuocTich = strMaQuocTich + arrValue(0) + Chr$(9)
        Next
        
        .Row = 22
        ' set danh sach quoc tich
        .Col = .ColLetterToNumber("J")
        .TypeComboBoxList = strQuocTich
        .typeComboboxCurSel = 0
        .Col = .ColLetterToNumber("I")
        .TypeComboBoxList = strMaQuocTich
        .typeComboboxCurSel = 0
        
        ' set danh sach quoc gia
        .Col = .ColLetterToNumber("Q")
        .TypeComboBoxList = strQuocTich
        .typeComboboxCurSel = 0
        .Col = .ColLetterToNumber("P")
        .TypeComboBoxList = strMaQuocTich
        .typeComboboxCurSel = 0
        
        
        .Sheet = 1
        .Row = 4
        .Col = .ColLetterToNumber("Q")
        .Text = DateDiff("M", Format(TAX_Utilities_v1.FirstDay, "mm/yyyy"), Format(TAX_Utilities_v1.LastDay, "mm/yyyy")) + 1
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim viTriMSTDL As String
    Dim vRowLock As String
    Dim arrRow() As String
    Dim i As Integer
    
     With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        
        If Strloaitk = "TKCT" And StrSolanBosung = "" Then
            .Col = .ColLetterToNumber("C")
            .Row = 67
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 67
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 67
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKCT" And Val(StrSolanBosung) > 0 Then
            .Col = .ColLetterToNumber("C")
            .Row = 67
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 67
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 67
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
            
            ' Neu la truong hop bo sung thi hien thi cot B trong bang ke 5A (Có dieu chinh so lieu)
            .Sheet = 2
            .Col = .ColLetterToNumber("C")
            .ColHidden = False
            
            ' Neu la truong hop bo sung thi hien thi cot W trong bang ke 5B (Có dieu chinh so lieu)
            .Sheet = 3
            .Col = .ColLetterToNumber("Z")
            .ColHidden = False
        End If
        
        ' set chi tieu tu thang - den thang
        .Row = 67
        .Col = .ColLetterToNumber("M")
        .Text = TAX_Utilities_v1.FirstDay
        UpdateCell fps, .Col, .Row, .Text
        
        .Row = 67
        .Col = .ColLetterToNumber("N")
        .Text = TAX_Utilities_v1.LastDay
        UpdateCell fps, .Col, .Row, .Text
        
        
        
            vRowLock = "43_44_45_47_48_49_51_52_53_62"
            arrRow = Split(vRowLock, "_")
        If DateDiff("M", Format(TAX_Utilities_v1.FirstDay, "mm/yyyy"), Format(TAX_Utilities_v1.LastDay, "mm/yyyy")) + 1 < 12 Then
            ' chi tieu 28,29,30,32,33,34,36,37,38,44
'            For i = 0 To UBound(arrRow)
'                .Row = arrRow(i)
'                .Col = .ColLetterToNumber("I")
'                .Lock = False
'            Next i
'            ' chi tieu 36 lock
'            .Row = 51
'            .Col = .ColLetterToNumber("I")
'            .Lock = False
'            ' chi tieu 37 lock
'            .Row = 52
'            .Col = .ColLetterToNumber("I")
'            .Lock = False
'            ' chi tieu 38 lock
'            .Row = 53
'            .Col = .ColLetterToNumber("I")
'            .Lock = False
'            ' chi tieu 44
'            .Row = 62
'            .Col = .ColLetterToNumber("I")
'            .Lock = False
            
            .Row = 4
            .Col = .ColLetterToNumber("D")
            .Lock = False
        Else
'            For i = 0 To UBound(arrRow)
'                .Row = arrRow(i)
'                .Col = .ColLetterToNumber("I")
'                .Lock = True
'            Next i
'            ' chi tieu 36 lock
'            .Row = 51
'            .Col = .ColLetterToNumber("I")
'            .Lock = True
'            ' chi tieu 37 lock
'            .Row = 52
'            .Col = .ColLetterToNumber("I")
'            .Lock = True
'            ' chi tieu 38 lock
'            .Row = 53
'            .Col = .ColLetterToNumber("I")
'            .Lock = True
'            ' chi tieu 44
'            .Row = 62
'            .Col = .ColLetterToNumber("I")
'            .Lock = True
            
            .Row = 4
            .Col = .ColLetterToNumber("D")
            .Lock = True
        End If
        
        'check thong tin ma so thue dai ly
        'set vi tri theo thu tu MSTDL~TENNVDLT~ChungChiHN
        viTriMSTDL = "D_20~D_69~D_71"
        updateMSTDL fps, viTriMSTDL
        
        .Sheet = 2
        .ActiveSheet = .Sheet
        .Row = 2
        .Col = .ColLetterToNumber("D")
        .CellType = CellTypeButton
        .TypeButtonTextColor = RGB(0, 0, 255)
        .TypeButtonColor = &H8000000F
        .TypeButtonAlign = TypeButtonAlignRight
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
        
        ' nut check chon
        .Row = 20
        .Col = .ColLetterToNumber("G")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\Check.bmp"))
        .TypeButtonText = ""
        
        .Row = 20
        .Col = .ColLetterToNumber("C")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\Check.bmp"))
        .TypeButtonText = ""
        ' Sap xep theo cot Ho va ten
        .Row = 17
        .Col = .ColLetterToNumber("D")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = GetAttribute(GetMessageCellById("0154"), "Msg")
        
        ' Sap xep theo cot thu nhap chiu thue
        .Row = 19
        .Col = .ColLetterToNumber("H")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = GetAttribute(GetMessageCellById("0155"), "Msg")
        
        ' Sap xep theo cot thue thu nhap ca nhan da khau tru
        .Row = 19
        .Col = .ColLetterToNumber("P")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = GetAttribute(GetMessageCellById("0159"), "Msg")
        
        
        .Sheet = 3
        .ActiveSheet = .Sheet
        .Row = 2
        .Col = .ColLetterToNumber("C")
        .CellType = CellTypeButton
        .TypeButtonTextColor = RGB(0, 0, 255)
        .TypeButtonColor = &H8000000F
        .TypeButtonAlign = TypeButtonAlignRight
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
        'net check chon
        .Row = 20
        .Col = .ColLetterToNumber("Y")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\Check.bmp"))
        .TypeButtonText = ""
        
        ' Sap xep theo cot Ho va ten
        .Row = 18
        .Col = .ColLetterToNumber("C")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = GetAttribute(GetMessageCellById("0154"), "Msg")
        
        ' Sap xep theo cot thu nhap chiu thue
        .Row = 19
        .Col = .ColLetterToNumber("G")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = GetAttribute(GetMessageCellById("0155"), "Msg")
        
        ' Sap xep theo cot thue thu nhap ca nhan da khau tru
        .Row = 19
        .Col = .ColLetterToNumber("J")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = GetAttribute(GetMessageCellById("0156"), "Msg")
        
        .Sheet = 4
        .ActiveSheet = .Sheet
        .Row = 2
        .Col = .ColLetterToNumber("D")
        .CellType = CellTypeButton
        .TypeButtonTextColor = RGB(0, 0, 255)
        .TypeButtonColor = &H8000000F
        .TypeButtonAlign = TypeButtonAlignRight
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
        
        .EventEnabled(EventAllEvents) = True
    End With
    
'    ' Kiem tra xem phien ban 3.1.3 hay phien ban 3.1.2 tro xuong
'    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
'    Dim isDFOld05A As Boolean, isDFOld05B As Boolean
'    ' end
'
'    Dim i As Long
'    With fps
'        ' test
'        ' Kiem tra 05A
'        Set xmlCellNode = TAX_Utilities_v1.Data(1).nodeFromID(GetCellID(fps, .ColLetterToNumber("D"), 22))
'        Set xmlCellNode = xmlCellNode.parentNode
'        Set xmlCellNode = xmlCellNode.lastChild
'        If Left$(GetAttribute(xmlCellNode, "CellID"), 1) = "C" Then
'            isDFOld05A = False
'        Else
'            isDFOld05A = True
'        End If
'        ' Kiem tra 05B
'        Set xmlCellNode = TAX_Utilities_v1.Data(2).nodeFromID(GetCellID(fps, .ColLetterToNumber("C"), 22))
'        Set xmlCellNode = xmlCellNode.parentNode
'        Set xmlCellNode = xmlCellNode.lastChild
'        If Left$(GetAttribute(xmlCellNode, "CellID"), 1) = "Y" Then
'            isDFOld05B = False
'        Else
'            isDFOld05B = True
'        End If
'
'
'        ' end test
'
'        .Sheet = 2
'        .EventEnabled(EventAllEvents) = False
'        i = 1
'        .Col = .ColLetterToNumber("B")
'        .Row = 22
'        Do
''             ' Lay gia tri cua cot MST de kiem tra, vi neu ko co MST thi ko duoc quyet toan tai CQCT
''             .Col = .ColLetterToNumber("E")
''             ' Kiem tra xem co null hay ko? Neu la null thi ko duoc check
''             If Trim(.Text) = "" Or Trim(.Text) = vbNullString Then
''                 .Col = .ColLetterToNumber("N")
''                 .Text = "0"
''                 .Lock = True
''             End If
'             If isDFOld05A = True Then
'                .Col = .ColLetterToNumber("C")
'                .Text = "0"
'             End If
'
'             .Col = .ColLetterToNumber("B")
'             .Row = i + 22
'             totalRow5A = i
'             i = i + 1
'        Loop Until .Text = "aa"
'
'        .Sheet = 3
'        i = 1
'        .Col = .ColLetterToNumber("B")
'        .Row = 22
'        Do
'             If isDFOld05B = True Then
'                .Col = .ColLetterToNumber("Y")
'                .Text = "0"
'            End If
'             .Col = .ColLetterToNumber("B")
'             .Row = i + 22
'             totalRow5B = i
'             i = i + 1
'        Loop Until .Text = "aa"
'
'        .EventEnabled(EventAllEvents) = True
'    End With
    ' Fill du lieu vao PL 27
    If TAX_Utilities_v1.Year = 2012 Then
        fillDataPL27
    End If
End Sub

Private Function GetValidatedCellIndex(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long) As Integer
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long
    
    Set xmlCellNode = TAX_Utilities_v1.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
    Set xmlCellsNode = xmlCellNode.parentNode
    
    'Get Index of anchor cell
    For lCtrl = 1 To xmlCellsNode.childNodes.length
        If GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID") = GetAttribute(xmlCellNode, "CellID") Then _
            Exit For
    Next
    GetValidatedCellIndex = lCtrl
    
End Function

Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim Y As Long
    Dim strCheck As Variant
    Dim vTempValue As Variant
    
    Dim vRowLock As String
    Dim arrRow() As String
    Dim i As Integer
    
    
    Static so As Long
    If so = 0 Then
        so = 1
    End If
    With fps
        .EventEnabled(EventAllEvents) = False
            If .ActiveSheet = 1 Then
                .Sheet = 1
                If Col = .ColLetterToNumber("D") And Row = 4 Then
                    .Col = .ColLetterToNumber("D")
                    .Row = 4
                    vTempValue = .Text
                    vRowLock = "43_44_45_47_48_49_51_52_53_62"
                    arrRow = Split(vRowLock, "_")
                    If vTempValue = "1" Or UCase$(vTempValue) = "X" Then
                        For i = 0 To UBound(arrRow)
                            .Row = arrRow(i)
                            .Col = .ColLetterToNumber("I")
                            .Lock = False
                        Next i
                    Else
                        For i = 0 To UBound(arrRow)
                            .Row = arrRow(i)
                            .Col = .ColLetterToNumber("I")
                            .Lock = True
                        Next i
                    End If
                End If
            End If
            If .ActiveSheet = 2 Then
               .Sheet = 2
                
                If Col = .ColLetterToNumber("D") And Row = 17 Then
                    Y = Col
                            
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("BH"), (21 + totalRow5A), SortByRow, Y, so
                    .Refresh
                    
                    ' Chieu sap xep
                    If so = 1 Then
                        so = 2 ' Sap xep giam dan
                    Else
                        so = 1 ' Sap xep tang dan
                    End If
                End If
                
                If Col = .ColLetterToNumber("H") And Row = 18 Then
                    Y = Col
                            
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("BH"), (21 + totalRow5A), SortByRow, Y, so
                    .Refresh
                    
                    ' Chieu sap xep
                    If so = 1 Then
                        so = 2 ' Sap xep giam dan
                    Else
                        so = 1 ' Sap xep tang dan
                    End If
                End If
                
                If Col = .ColLetterToNumber("P") And Row = 19 Then
                    Y = Col
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("BH"), (21 + totalRow5A), SortByRow, Y, so
                    .Refresh
                    
                    ' Chieu sap xep
                    If so = 1 Then
                        so = 2 ' Sap xep giam dan
                    Else
                        so = 1 ' Sap xep tang dan
                    End If
                End If
                
                
                 ' neu check se ho tro tinh gia tri thue khau tru theo cong thuc
                If Col = .ColLetterToNumber("G") And Row >= 22 And Row <= .MaxRows - 8 Then
                    .GetText .ColLetterToNumber("G"), Row, strCheck
                    If Trim(strCheck) = "x" Or Trim(strCheck) = "X" Or Trim(strCheck) = "1" Then
                         'set lai CT
                        .Col = .ColLetterToNumber("S")
                        '.Formula = "ROUND(IF(OR(G22=""1"",G22=""x""),AN22-AO22,0),0)"
                        '.Formula = "ROUND(IF(OR(G22=""1"",G22=""x""),IF($G$5 = ""2009"", BA22, AN22-AO22),0),0)"
                        '.Formula = "ROUND(IF(OR(G" & Row & "=""1"",G" & Row & "=""x"",G" & Row & "=""X"" ),IF($G$5 = ""2009"", BC" & Row & "-AQ" & Row & ", IF(AND($G$5 = ""2012"",AO" & Row & " <=5000000),(AP" & Row & "-AQ" & Row & ")/2,AP" & Row & "-AQ" & Row & ")),0),0)"
                        .Formula = "ROUND(IF(OR(G" & Row & "=""1"",G" & Row & "=""x"",G" & Row & "=""X"" ),IF($G$5 = ""2009"", BC" & Row & "-Q" & Row & ", IF(AND($G$5 = ""2012"",AO" & Row & " <=5000000),(AP" & Row & "-Q" & Row & ")/2,AP" & Row & "-Q" & Row & ")),0),0)"
                        ' End set lai
                    Else
                        .Col = .ColLetterToNumber("S")
                        .Formula = ""
                        .value = "0"
                    End If
                End If
                
            End If
            'Sheet 3
            If .ActiveSheet = 3 Then
               .Sheet = 3
                
                If Col = .ColLetterToNumber("C") And Row = 17 Then
                    Y = Col
                            
                    .Sort .ColLetterToNumber("C"), 22, .ColLetterToNumber("AP"), (21 + totalRow5B), SortByRow, Y, so
                    .Refresh
                    
                    ' Chieu sap xep
                    If so = 1 Then
                        so = 2 ' Sap xep giam dan
                    Else
                        so = 1 ' Sap xep tang dan
                    End If
                End If
                
                If Col = .ColLetterToNumber("G") And Row = 18 Then
                    Y = Col
                            
                    .Sort .ColLetterToNumber("C"), 22, .ColLetterToNumber("AP"), (21 + totalRow5B), SortByRow, Y, so
                    .Refresh
                    
                    ' Chieu sap xep
                    If so = 1 Then
                        so = 2 ' Sap xep giam dan
                    Else
                        so = 1 ' Sap xep tang dan
                    End If
                End If
                
                If Col = .ColLetterToNumber("J") And Row = 19 Then
                    Y = Col
                    .Sort .ColLetterToNumber("C"), 22, .ColLetterToNumber("AP"), (21 + totalRow5B), SortByRow, Y, so
                    .Refresh
                    
                    ' Chieu sap xep
                    If so = 1 Then
                        so = 2 ' Sap xep giam dan
                    Else
                        so = 1 ' Sap xep tang dan
                    End If
                End If
                
                ' neu check se ho tro tinh gia tri thue khau tru theo cong thuc
                If Col = .ColLetterToNumber("F") And Row >= 22 And Row <= .MaxRows - 8 Then
                    .GetText .ColLetterToNumber("F"), Row, strCheck
                    If Trim(strCheck) = "x" Or Trim(strCheck) = "X" Or Trim(strCheck) = "1" Then
                        .Col = .ColLetterToNumber("K")
                        .Row = Row
                        '.Formula = "AJ" & Row
                        '.Formula = "ROUND(G" & Row & "*20/100,0)"
                        .Formula = "ROUND(H" & Row & "*20/100*50/100,0)"
                        ' lay xong gia tri set lai ct =""
                        .Formula = ""
                    Else
                        .Col = .ColLetterToNumber("K")
                        .Row = Row
                        .Formula = ""
                        .value = "0"
                    End If
                End If
            End If
            
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    'UpdateSheets
    ' Test ct
    Dim strCheck As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
            If .ActiveSheet = 2 Then
                If Col = .ColLetterToNumber("H") Or Col = .ColLetterToNumber("I") Or Col = .ColLetterToNumber("J") Or Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("L") Or Col = .ColLetterToNumber("M") Or Col = .ColLetterToNumber("N") Then
                     If Row >= 22 And Row <= .MaxRows - 8 Then
                         'set lai CT
                        .Col = .ColLetterToNumber("O")
                        .Formula = "IF(H" & Row & "-J" & Row & "-K" & Row & "-L" & Row & "-M" & Row & "-N" & Row & ">0,H" & Row & "-J" & Row & "-K" & Row & "-L" & Row & "-M" & Row & "-N" & Row & ",0)"
                        .Col = .ColLetterToNumber("S")
                        '.Formula = "ROUND(IF(OR(G" & Row & "=""1"",G" & Row & "=""x"",G" & Row & "=""X""),IF($G$5 = ""2009"", BA" & Row & "-AO" & Row & ", AN" & Row & "-AO" & Row & "),0),0)"
                        '.Formula = "ROUND(IF(OR(G" & Row & "=""1"",G" & Row & "=""x"",G" & Row & "=""X"" ),IF($G$5 = ""2009"", BA" & Row & "-AO" & Row & ", IF($G$5 = ""2012"",(AN" & Row & "-AO" & Row & ")/2,AN" & Row & "-AO" & Row & "),0),0)"
                        '.Formula = "ROUND(IF(OR(G" & Row & "=""1"",G" & Row & "=""x"",G" & Row & "=""X"" ),IF($G$5 = ""2009"", BC" & Row & "-AQ" & Row & ", IF(AND($G$5 = ""2012"",AO" & Row & " <=5000000),(AP" & Row & "-AQ" & Row & ")/2,AP" & Row & "-AQ" & Row & ")),0),0)"
                        .Formula = "ROUND(IF(OR(G" & Row & "=""1"",G" & Row & "=""x"",G" & Row & "=""X"" ),IF($G$5 = ""2009"", BC" & Row & "-Q" & Row & ", IF(AND($G$5 = ""2012"",AO" & Row & " <=5000000),(AP" & Row & "-Q" & Row & ")/2,AP" & Row & "-Q" & Row & ")),0),0)"
                        '.Formula = "ROUND(IF(OR(G" & Row & "=""1"",G" & Row & "=""x"",G" & Row & "=""X"" ),IF($G$5 = ""2009"", BA" & Row & "-AO" & Row & ", IF(AND($G$5 = ""2012"",AM" & Row & " <=5000000),(AN" & Row & "-AO" & Row & ")/2,AN" & Row & "-AO" & Row & ")),0),0)"
                        ' End set lai
                    End If
                End If
                
                CellChange Col, Row
            End If
        
            If .ActiveSheet = 3 Then
                If Col = .ColLetterToNumber("G") Then
                    If Row >= 22 And Row <= .MaxRows - 8 Then
                        .GetText .ColLetterToNumber("F"), Row, strCheck
                        If Trim$(strCheck) = "x" Or Trim$(strCheck) = "X" Or Trim$(strCheck) = "1" Then
                            .Col = .ColLetterToNumber("K")
                            .Row = Row
                            '.Formula = "ROUND(G" & Row & "*20/100,0)"
                            .Formula = "ROUND(H" & Row & "*20/100*50/100,0)"
                            'ROUND(G22*20/100,0)
                            ' lay xong gia tri set lai ct =""
                            .Formula = ""
                        Else
                            .Col = .ColLetterToNumber("K")
                            .Row = Row
                            .Formula = ""
                            .value = "0"
                        End If
                    End If
                End If
                
                CellChange Col, Row
             End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
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
        If .ActiveSheet = 4 Then
            ' chon quoc tich
            If Col = .ColLetterToNumber("J") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("I")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
            End If
            
            ' chon moi quan he
            If Col = .ColLetterToNumber("M") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("L")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
            End If
            
            ' chon quoc gia
            If Col = .ColLetterToNumber("Q") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("P")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
                
                 If .Text = "01" Then
                    .Col = .ColLetterToNumber("S")
                    .Row = Row
                    .Text = ""
                    .Lock = False
                    
                    .Col = .ColLetterToNumber("U")
                    .Row = Row
                    .Text = ""
                    .Lock = False
                    
                    .Col = .ColLetterToNumber("W")
                    .Row = Row
                    .Text = ""
                    .Lock = False
                    
                    .Col = .ColLetterToNumber("R")
                    .Row = Row
                    .Text = ""
                    
                    .Col = .ColLetterToNumber("T")
                    .Row = Row
                    .Text = ""
                    
                    .Col = .ColLetterToNumber("V")
                    .Row = Row
                    .Text = ""
                Else
                    .Col = .ColLetterToNumber("S")
                    .Row = Row
                    .typeComboboxCurSel = .TypeComboBoxCount - 1
                    .Lock = True
                    
                    .Col = .ColLetterToNumber("U")
                    .Row = Row
                    .Text = "Khác"
                    .Lock = True
                    
                    .Col = .ColLetterToNumber("W")
                    .Row = Row
                    .Text = "Khác"
                    .Lock = True
                    
                    .Col = .ColLetterToNumber("R")
                    .Row = Row
                    .typeComboboxCurSel = .TypeComboBoxCount - 1
                    
                    .Col = .ColLetterToNumber("R")
                    .Row = Row
                    .Text = "999"
                    
                    .Col = .ColLetterToNumber("T")
                    .Row = Row
                    .Text = "99999"
                    
                    .Col = .ColLetterToNumber("V")
                    .Row = Row
                    .Text = "9999999"
                End If
            End If
            
            ' chon tinh thanh
            If Col = .ColLetterToNumber("S") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("R")
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
                    .Col = .ColLetterToNumber("U")
                    .TypeComboBoxList = tenHuyen
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("T")
                    .TypeComboBoxList = maHuyen
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                   
                    .Row = Row
                    .Col = .ColLetterToNumber("V")
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("W")
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                                                   
                End If
            End If
            ' chon quan huyen
            If Col = .ColLetterToNumber("U") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("T")
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
                    .Col = .ColLetterToNumber("W")
                    .TypeComboBoxList = tenXa
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("V")
                    .TypeComboBoxList = maXa
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                
                End If
            End If
            ' chon xa phuong
            If Col = .ColLetterToNumber("W") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("V")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
            End If
            
        End If
    End With
End Sub

Private Sub fps_EditChange(ByVal Col As Long, ByVal Row As Long)
    With fps
        .EventEnabled(EventAllEvents) = False
            If .ActiveSheet = 2 Then
                If Col = .ColLetterToNumber("G") Then
                    flgLoadToKhai = False
                End If
            End If
        .EventEnabled(EventAllEvents) = True
    End With
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
            If .ActiveSheet = 2 Then
                    .EventEnabled(EventAllEvents) = False
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 22
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 22
                         totalRow5A = i
                         i = i + 1
                    Loop Until .Text = "aa"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    
                    .SetActiveCell iCol, iRow
            End If
            If .ActiveSheet = 3 Then
                    .EventEnabled(EventAllEvents) = False
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 22
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 22
                         totalRow5B = i
                         i = i + 1
                    Loop Until .Text = "aa"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                     
                    .SetActiveCell iCol, iRow
            End If
            
            If .ActiveSheet = 4 Then
                .EventEnabled(EventAllEvents) = False
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 22
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 22
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
    Dim varHoVaTen, varMST, varCMT  As Variant
    Dim varThuNhap, varThuNhapTinhThue, varThuNhapTinhThueThang, varBaoHiem, varSoNguoiGiamTru, varSoThangGiamTru As Variant
    Dim varTuThien, varThuNhapGiamThue, varSoThueDuocGiam, varThuePhaiNop, varThuePhaiNopThang, varSoThueTamKhauTru, varSoThueKhauTru As Variant
    
    Dim strCheck As String
        
    flgErr = True
    If flgLoadToKhai = True Then Exit Sub
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            'check date
            If Col = .ColLetterToNumber("M") And Row = 71 Then
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
        ElseIf .ActiveSheet = 2 And Row >= 22 Then
            ' Cot Ho va ten
'            If Col = .ColLetterToNumber("D") Then
'                .GetText .ColLetterToNumber("D"), Row, varHoVaTen
'                .GetText .ColLetterToNumber("N"), Row, varSoThueKhauTru
'                ' Cu them mot lao dong thi tinh la mot nguoi
'                If Trim(varHoVaTen) <> vbNullString And Val(varSoThueKhauTru) > 0 Then
'                    ' Them mot nguoi vao cot N
'                    .Col = .ColLetterToNumber("T")
'                    .Text = "1"
'                Else ' Neu ko thi ko tinh nguoi do
'                    flgErr = False
'                    .Col = .ColLetterToNumber("T")
'                    .Text = "0"
'                End If
'            End If
        ElseIf .ActiveSheet = 3 Then
'            If Col = .ColLetterToNumber("C") Then
'                .GetText .ColLetterToNumber("C"), Row, varTemp
'                ' Cu them mot lao dong thi tinh la mot nguoi
'                If Trim(varTemp) <> vbNullString Then
'                    .Col = .ColLetterToNumber("I")
'                    .Text = "1"
'                Else ' Neu ko thi ko tinh nguoi do
'                    .Col = .ColLetterToNumber("I")
'                    .Text = "0"
'                End If
'            End If
            ' Lay so thue da khau tru cua lao dong do
'            If Col = .ColLetterToNumber("H") Then
'                .GetText .ColLetterToNumber("H"), Row, varTemp
'                If Val(varTemp) > 0 Then
'                    ' Va tinh la lao dong co thue khau tru
'                    .Col = .ColLetterToNumber("J")
'                    .Text = "1"
'                    ' Tinh tong thu nhap cua rieng lao dong co thue khau tru nay
'                    .GetText .ColLetterToNumber("F"), Row, varTemp
'                    .Col = .ColLetterToNumber("K")
'                    .Text = varTemp
'                Else
'                    ' Va tinh la lao dong ko co thue khau tru
'                    .Col = .ColLetterToNumber("J")
'                    .Text = "0"
'                    ' Khong tinh tong thu nhap cua lao dong nay
'                    .Col = .ColLetterToNumber("K")
'                    .Text = "0"
'                End If
'            End If
         ElseIf .ActiveSheet = 4 Then
            If Col = .ColLetterToNumber("G") And Row >= 22 Then
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
            
            If Col = .ColLetterToNumber("X") And Row >= 22 Then
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
            
            If Col = .ColLetterToNumber("Y") And Row >= 22 Then
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

        SetData
        
        If TAX_Utilities_v1.Year = "2012" Then
            fillDataPL27
        End If
        ' Lay cac mang MST trong cac bang ke va phu luc 05A, 05B
'        prepareMSTCMT
'
        kiemTraDuLieuImport

        .Sheet = 2
        InsertNode 23, .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone) - 1, 2
        .Sheet = 3
        InsertNode 23, .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone) - 1, 3
        
        .Sheet = 4
        InsertNode 23, .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone) - 1, 4
'
'
        
        
        UpdateSheets
        CheckDynamicError 'Set Exception Error on cells of interface

''        If flgErr5A = True And flgErr5B = True Then kiemTraKhac


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
    
    Dim vQTTronNam As Variant
    Dim vLyDo As Variant
    Dim iFlagLyDo As Integer
    Dim vErrLyDo As Variant
    
    iFlagLyDo = 0
    
    With fps
    
        tSheet = .Sheet
        'If flgLoadToKhai = True Then Exit Sub
        
        'get MST cua DTNT
        .Sheet = 1
        .GetText .ColLetterToNumber("D"), 10, MSTDTNT
        MSTDTNT = Trim(MSTDTNT)
        
        'check thong tin header khong dc trong
        'lay vi tri header tren tk theo thu tu MSTDL, TENDLT, DIACHIDLT,QUAN/HUYEN,TP,SO HDDLT,NGAYDH,Ten NV DLT, CC Hanh Nghe

        viTriHeaderTk = "D_20~D_18~D_22~D_24~L_24~D_28~L_28~D_69~D_71"
        viTriSheetHeader = "B_58~B_59~B_60~B_61~B_62~B_63~B_64~B_65~B_66"
        mangMessage = "E_58~E_59~E_60~E_61~E_62~E_63~E_64~E_65~E_66"
        'check header
        checkErrorHeader fps, viTriHeaderTk, viTriSheetHeader, mangMessage, MSTDTNT
        
        .Sheet = tSheet
    
        CheckErrorMST
        
        ' check tren TK
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 85, vErrLyDo
        .Sheet = 1
        .GetText .ColLetterToNumber("D"), 4, vQTTronNam
        .GetText .ColLetterToNumber("D"), 6, vLyDo
        
        
        
        If vQTTronNam = "1" Or vQTTronNam = "x" Or vQTTronNam = "X" Then
            If Trim$(vLyDo) = "" Then
                .Col = .ColLetterToNumber("D")
                .Row = 6
                .CellNote = vErrLyDo
                .BackColor = mErrorColor
                iFlagLyDo = 1
            Else
                .Col = .ColLetterToNumber("D")
                .Row = 6
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
        Else
            .Col = .ColLetterToNumber("D")
            .Row = 6
            .CellNote = ""
            .BackColor = mNonErrorColor
        End If
        
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 85, IIf(iFlagLyDo = 1, "0", "1")
        
        ' Neu ton tai file data bang ke 05A thi kiem tra tren bang ke 5A
        If IIf(TAX_Utilities_v1.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            CheckDynamicSheet1
        End If
        If IIf(TAX_Utilities_v1.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            CheckDynamicSheet2
        End If
        If IIf(TAX_Utilities_v1.NodeValidity.childNodes(3).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            CheckDynamicSheet3
        End If
    End With
End Sub

'Private Sub checkSheet1()
'    Dim vErrMsgHoVaTen, vErrMsgMST, vErrMsgMST1, vErrMsgMST2, vErrMsgCMT, vErrMsgCMT1, vErrMsgMSTCMT As Variant
'    Dim vErrMsgBaoHiem, vErrMsgThangGiamTru, vErrMsgThuNhapGiamThue, vErrMsgSoThueDuocGiam, vErrMsgSoThueKhauTru, vErrMsgTongGiamTruGiaCanh As Variant
'    Dim vErrMsgQTCQCT As Variant
'    Dim vErrSoThueNNT As Variant
'    Dim vLastRow As Variant
'    Dim varTemp As Variant, varTemp1 As Variant
'    Dim i, j As Integer
'    Dim iFocusFlag As Integer
'    Dim strFocusSheetName, strFocusCol, strFocusRow As String
'    Dim flgQTCQCT As Boolean
'
'    Dim vErrMsgMSTAB As Variant
'    Dim flagTrungMSTAB As Integer
'
'    flagTrungMSTAB = 0
'    flgQTCQCT = True
'
'    With fps
'        .Sheet = .SheetCount
'
'        .GetText .ColLetterToNumber("E"), 21, vErrMsgHoVaTen
'        .GetText .ColLetterToNumber("E"), 22, vErrMsgMST
'        .GetText .ColLetterToNumber("E"), 23, vErrMsgMST1
'        .GetText .ColLetterToNumber("E"), 24, vErrMsgMST2
'        .GetText .ColLetterToNumber("E"), 25, vErrMsgCMT
'        .GetText .ColLetterToNumber("E"), 26, vErrMsgCMT1
'        .GetText .ColLetterToNumber("E"), 27, vErrMsgMSTCMT
'
'        .GetText .ColLetterToNumber("E"), 28, vErrMsgBaoHiem
'        .GetText .ColLetterToNumber("E"), 29, vErrMsgThangGiamTru
'        .GetText .ColLetterToNumber("E"), 30, vErrMsgThuNhapGiamThue
'        .GetText .ColLetterToNumber("E"), 31, vErrMsgSoThueKhauTru
'
'        .GetText .ColLetterToNumber("E"), 47, vErrMsgMSTAB
'
'        .GetText .ColLetterToNumber("E"), 51, vErrMsgQTCQCT
'        .GetText .ColLetterToNumber("E"), 55, vErrSoThueNNT
'        .GetText .ColLetterToNumber("E"), 57, vErrMsgTongGiamTruGiaCanh
'
'
'        .Sheet = 2 'Check tinh xac thuc cua cac thong tin tren bang ke 05A/TNCN
'
'        ' Cot check loi la AD24 duoc danh dau xem bang ke co con bi sai hay khong. Neu sai thi moi check loi tiep
''        .GetText .ColLetterToNumber("AG"), 16, varTemp
''        If Val(varTemp) = 0 Then Exit Sub
'
'        For j = 22 To .MaxRows
'            .Row = j
'
'            .GetText .ColLetterToNumber("B"), .Row, vLastRow
'            If UCase(vLastRow) = "AA" Then Exit For ' Check cho den het cac dong co du lieu thi thoi
'            ' Kiem tra neu co check QTT thi bat buoc nhap MST
'            ' GD2
'            .GetText .ColLetterToNumber("E"), .Row, varTemp
'            .GetText .ColLetterToNumber("G"), .Row, varTemp1
'            If (Trim(varTemp) = "" Or Trim(varTemp) = vbNullString) Then
'                If (Trim(varTemp1) = "1" Or Trim(varTemp1) = "x") Then
'                    flgQTCQCT = False
'                    .Col = .ColLetterToNumber("G")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgQTCQCT
'                    .BackColor = mErrorColor
'                Else
'                    .Col = .ColLetterToNumber("G")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                End If
'            Else
'                .Col = .ColLetterToNumber("G")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'            End If
'
'            '
'            .GetText .ColLetterToNumber("AJ"), .Row, varTemp
'            If Val(varTemp) = 1 Then
'                .Col = .ColLetterToNumber("D")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("E")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("F")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("H")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("I")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("J")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("K")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("L")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("M")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("N")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("O")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                ' Check ho ten
'                .GetText .ColLetterToNumber("X"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("D")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgHoVaTen
'                    .BackColor = mErrorColor
'                End If
'                'Check mst
'                .GetText .ColLetterToNumber("Y"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("E")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgMST
'                    .BackColor = mErrorColor
'                End If
'                ' Check mst ca nhan trung voi ma so thue CQCT
'                .GetText .ColLetterToNumber("Z"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("E")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgMST1
'                    .BackColor = mErrorColor
'                End If
'                ' Check mst trung nhau
'                .GetText .ColLetterToNumber("AA"), .Row, varTemp
'                If Trim(varTemp) = "1" Then
'                    .Col = .ColLetterToNumber("E")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgMST2
'                    .BackColor = mErrorColor
'                End If
'                ' Check CMT trung nhau
'                .GetText .ColLetterToNumber("AB"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("F")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgCMT
'                    .BackColor = mErrorColor
'                End If
'                ' Check CMT sai
'                .GetText .ColLetterToNumber("AC"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("F")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgCMT1
'                    .BackColor = mErrorColor
'                End If
'                ' Check MST rong, check CMT rong
'                .GetText .ColLetterToNumber("AD"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("E")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgMSTCMT
'                    .BackColor = mErrorColor
'
'                    .Col = .ColLetterToNumber("F")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgMSTCMT
'                    .BackColor = mErrorColor
'                End If
'                ' Check so tien giam tru BH < tong thu nhap
'                .GetText .ColLetterToNumber("AE"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("L")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgBaoHiem
'                    .BackColor = mErrorColor
'                End If
'
''                .GetText .ColLetterToNumber("AD"), .Row, varTemp
''                If Val(varTemp) = 1 Then
''                    .Col = .ColLetterToNumber("I")
''                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                    .CellNote = vErrMsgThangGiamTru
''                    .BackColor = mErrorColor
''                End If
'                ' Check thu nhap tinh lam can cu giam thue < thu nhap
'                .GetText .ColLetterToNumber("AG"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("I")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgThuNhapGiamThue
'                    .BackColor = mErrorColor
'                End If
'                ' Check thu da khau tru voi tong thu nhap
'                .GetText .ColLetterToNumber("AI"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("N")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgSoThueKhauTru
'                    '.BackColor = mErrorColor
'                    .BackColor = mAlertColor
'                End If
'                ' Check so thu da nop thua va so thue con phai khau tru them
'                .GetText .ColLetterToNumber("AT"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("R")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrSoThueNNT
'                    .BackColor = mErrorColor
'
'                    .Col = .ColLetterToNumber("S")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrSoThueNNT
'                    .BackColor = mErrorColor
'                End If
'                ' Kiem tra tong so tien giam tru gia canh phai chan tram nghin
'                .GetText .ColLetterToNumber("AY"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("J")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgTongGiamTruGiaCanh
'                    .BackColor = mErrorColor
'                End If
'
'            Else
'                .Col = .ColLetterToNumber("D")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("E")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("F")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
''                .Col = .ColLetterToNumber("G")
''                .CellNote = ""
''                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("H")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("I")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("J")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("K")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("L")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("M")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
''                .Col = .ColLetterToNumber("P")
''                .CellNote = ""
''                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("N")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("O")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("Q")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("R")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("S")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'
'            End If
'
'            .GetText .ColLetterToNumber("AU"), .Row, varTemp
'            If Val(varTemp) = 1 Then
'                flagTrungMSTAB = 1
'                .Col = .ColLetterToNumber("E")
'                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                .CellNote = vErrMsgMSTAB
'                .BackColor = mErrorColor
''            Else
''                .Col = .ColLetterToNumber("D")
''                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                .CellNote = ""
''                .BackColor = mNonErrorColor
'            End If
'
'        Next
'        .Sheet = .SheetCount
'        .Col = .ColLetterToNumber("B")
'        .Row = 51
'        If flgQTCQCT = False Then
'            .Text = "0"
'        Else
'            .Text = "1"
'        End If
'        .Row = 47
'        If flagTrungMSTAB = 1 Then
'            .Text = "0"
'        Else
'            .Text = "1"
'        End If
'
'    End With
'End Sub

'Private Sub checkSheet2()
'    Dim vErrMsgHoVaTen, vErrMsgMST, vErrMsgMST1, vErrMsgMST2, vErrMsgCMT, vErrMsgCMT1, vErrMsgMSTCMT As Variant
'    Dim vErrMsgThuNhapGiamThue, vErrMsgSoThueDuocGiam, vErrMsgSoThueKhauTru As Variant
'    Dim vLastRow As Variant
'    Dim varTemp As Variant
'    Dim i, j As Integer
'    Dim iFocusFlag As Integer
'    Dim strFocusSheetName, strFocusCol, strFocusRow As String
'
'    Dim vErrMsgMSTAB As Variant
'    Dim flagTrungMSTAB As Integer
'
'    flagTrungMSTAB = 0
'
'    With fps
'        .Sheet = .SheetCount
'
'        .GetText .ColLetterToNumber("E"), 37, vErrMsgHoVaTen
'        .GetText .ColLetterToNumber("E"), 38, vErrMsgMST
'        .GetText .ColLetterToNumber("E"), 39, vErrMsgMST1
'        .GetText .ColLetterToNumber("E"), 40, vErrMsgMST2
'        .GetText .ColLetterToNumber("E"), 41, vErrMsgCMT
'        .GetText .ColLetterToNumber("E"), 42, vErrMsgCMT1
'        .GetText .ColLetterToNumber("E"), 43, vErrMsgMSTCMT
'
'        .GetText .ColLetterToNumber("E"), 44, vErrMsgThuNhapGiamThue
'        .GetText .ColLetterToNumber("E"), 45, vErrMsgSoThueKhauTru
'
'        .GetText .ColLetterToNumber("E"), 47, vErrMsgMSTAB
'        .Sheet = 3 'Check tinh xac thuc cua cac thong tin tren bang ke 05A/TNCN
'
'        ' Cot check loi la AD24 duoc danh dau xem bang ke co con bi sai hay khong. Neu sai thi moi check loi tiep
''        .GetText .ColLetterToNumber("Z"), 16, varTemp
''        If Val(varTemp) = 0 Then Exit Sub
'
'        For j = 22 To .MaxRows
'            .Row = j
'            .GetText .ColLetterToNumber("B"), .Row, vLastRow
'            If UCase(vLastRow) = "AA" Then Exit For ' Check cho den het cac dong co du lieu thi thoi
'
'            .GetText .ColLetterToNumber("X"), .Row, varTemp
'            If Val(varTemp) = 1 Then
'
'                .Col = .ColLetterToNumber("C")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("D")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("E")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
''                .Col = .ColLetterToNumber("F")
''                .CellNote = ""
''                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("G")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("H")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("I")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("J")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                ' Check ho ten
'                .GetText .ColLetterToNumber("N"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("C")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgHoVaTen
'                    .BackColor = mErrorColor
'                End If
'                ' Check ma so thue ca nhan
'                .GetText .ColLetterToNumber("O"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("D")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgMST
'                    .BackColor = mErrorColor
'                End If
'                ' check MST trung voi MST co quan CT
'                .GetText .ColLetterToNumber("P"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("D")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgMST1
'                    .BackColor = mErrorColor
'                End If
'                ' Check MST trung nhau
'                .GetText .ColLetterToNumber("Q"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("D")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgMST2
'                    .BackColor = mErrorColor
'                End If
'                ' Check CMT trung nhau
'                .GetText .ColLetterToNumber("R"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("E")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgCMT
'                    .BackColor = mErrorColor
'                End If
'                ' check CMT sai
'                .GetText .ColLetterToNumber("S"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("E")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgCMT1
'                    .BackColor = mErrorColor
'                End If
'                ' Check khong MST, khong CMT
'                .GetText .ColLetterToNumber("T"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("D")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgMSTCMT
'                    .BackColor = mErrorColor
'
'                    .Col = .ColLetterToNumber("E")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgMSTCMT
'                    .BackColor = mErrorColor
'                End If
'                ' THu nhap lam can cu giam thue <= thu nhap
'                .GetText .ColLetterToNumber("U"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("H")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgThuNhapGiamThue
'                    .BackColor = mErrorColor
'                End If
'                'Thue da khau tru >= thu nhap
'                .GetText .ColLetterToNumber("W"), .Row, varTemp
'                If Val(varTemp) = 1 Then
'                    .Col = .ColLetterToNumber("I")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgSoThueKhauTru
'                    '.BackColor = mErrorColor
'                    .BackColor = mAlertColor
'                End If
'            Else
'                .Col = .ColLetterToNumber("C")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("D")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("E")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
''                .Col = .ColLetterToNumber("F")
''                .CellNote = ""
''                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("G")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("H")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("I")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'                .Col = .ColLetterToNumber("J")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'            End If
'
'            .GetText .ColLetterToNumber("Z"), .Row, varTemp
'            If Val(varTemp) = 1 Then
'                flagTrungMSTAB = 1
'                .Col = .ColLetterToNumber("D")
'                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                .CellNote = vErrMsgMSTAB
'                .BackColor = mErrorColor
''            Else
''                .Col = .ColLetterToNumber("C")
''                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                .CellNote = ""
''                .BackColor = mNonErrorColor
'            End If
'
'        Next
'
'        .Sheet = .SheetCount
'        .Col = .ColLetterToNumber("B")
'        .Row = 47
'        If flagTrungMSTAB = 1 Then
'            .Text = "0"
'        Else
'            .Text = "1"
'        End If
'
'
'    End With
'End Sub

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
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "05-1"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in E column and F column
                    If lCol = fps.ColLetterToNumber("D") Or _
                       lCol = fps.ColLetterToNumber("H") Then
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

             Case "05-2"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_v1.Data(2).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in E column and F column
                    If lCol = fps.ColLetterToNumber("C") Or _
                       lCol = fps.ColLetterToNumber("G") Then
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
               Case "05-3"
                    blnNullValue = True
                    'Check the third sheet
                   For Each xmlCellNode In TAX_Utilities_v1.Data(3).getElementsByTagName("Cell")
                        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                        'Check value in E column and F column
                        If lCol = fps.ColLetterToNumber("D") Or _
                           lCol = fps.ColLetterToNumber("S") Or lCol = fps.ColLetterToNumber("T") Then
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
               Case "05MT"
                    If TAX_Utilities_v1.Year = 2012 Then
                        blnNullValue = True
                        'Check the third sheet
                       For Each xmlCellNode In TAX_Utilities_v1.Data(4).getElementsByTagName("Cell")
                            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                            'Check value in E column and F column
                            If lCol = fps.ColLetterToNumber("C") Or _
                               lCol = fps.ColLetterToNumber("F") Then
                                strValue = GetAttribute(xmlCellNode, "Value")
                                If Not IsNullValue(strValue) Then
                                    blnNullValue = False
                                    Exit For
                                End If
                            End If
                        Next

                        If blnNullValue Then
                            SetAttribute xmlNodeValid, "Active", "0"
                        Else
                            SetAttribute xmlNodeValid, "Active", "1"
                        End If
                    End If
        End Select
    Next
    
    Set xmlCellNode = Nothing
    Set xmlNodeValid = Nothing
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

        If fps.ActiveSheet = 2 Then
            fps.Row = 22
            fps.Col = fps.ColLetterToNumber("B")
            fps.Formula = "B21 + 1"
        ElseIf fps.ActiveSheet = 3 Then
            fps.Row = 22
            fps.Col = fps.ColLetterToNumber("B")
            fps.Formula = "B21 + 1"
        ElseIf fps.ActiveSheet = 4 Then
            fps.Row = 22
            fps.Col = fps.ColLetterToNumber("B")
            fps.Formula = "B21 + 1"
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
    ' phuc vu in PL27
    Dim vThuNhapTT, vSoThueKt, vChiTieu12, vChiTieu13, vChiTieu14, vTongSoNguoi As Variant
    Dim i As Integer
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
        If (GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(1), "Active") <> "0") Then
        .Sheet = 2
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               ' Update gia tri cua truong check la "x" thi khi hien thi moi dung dinh dang
               If (lCol = .ColLetterToNumber("G") Or lCol = .ColLetterToNumber("R") Or lCol = .ColLetterToNumber("C")) And (varTemp = "1" Or varTemp = "x") Then
                    UpdateCell fps, lCol, lRow, "x"
               ElseIf (lCol = .ColLetterToNumber("G") Or lCol = .ColLetterToNumber("R") Or lCol = .ColLetterToNumber("C")) And (varTemp = "0" Or varTemp = "") Then
                    UpdateCell fps, lCol, lRow, ""
               Else
                    If lCol >= 8 Then
                        If Trim(varTemp) <> vbNullString Or Trim(varTemp) <> "" Then
                            UpdateCell fps, lCol, lRow, varTemp
                        End If
                    Else
                        UpdateCell fps, lCol, lRow, varTemp
                    End If
               End If
            Next
        End If
        If (GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(2), "Active") <> "0") Then
        .Sheet = 3
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(2).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                If (lCol = .ColLetterToNumber("F") Or lCol = .ColLetterToNumber("Y")) And (varTemp = "1" Or varTemp = "x") Then
                    UpdateCell fps, lCol, lRow, "x"
                ElseIf (lCol = .ColLetterToNumber("F") Or lCol = .ColLetterToNumber("Y")) And (varTemp = "0" Or varTemp = "") Then
                    UpdateCell fps, lCol, lRow, ""
                Else
                    If lCol >= 7 Then
                        If Trim(varTemp) <> vbNullString Or Trim(varTemp) <> "" Then
                            UpdateCell fps, lCol, lRow, varTemp
                        End If
                    Else
                        UpdateCell fps, lCol, lRow, varTemp
                    End If
               End If
             Next
       End If
       
        If (GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(3), "Active") <> "0") Then
        .Sheet = 4
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(3).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
             Next
       End If
       
       ' doi voi to khai 2012 se ghi dong tong de in PL
       If TAX_Utilities_v1.Year = 2012 Then
            .Sheet = 5
            mCurrentSheet = .Sheet
            
            fillDataPL27
    
            
            .EventEnabled(EventAllEvents) = False
            i = 1
            .Col = .ColLetterToNumber("B")
            .Row = 22
            Do
                 .Text = str(i)
                 .Col = .ColLetterToNumber("B")
                 .Row = i + 22
                 i = i + 1
            Loop Until .Text = "aa"
            ' Lay gia tri dong tong ghi vao xml in
            .Row = .Row + 1
            .GetText .ColLetterToNumber("C"), .Row, vTongSoNguoi
            .GetText .ColLetterToNumber("F"), .Row, vThuNhapTT
            .GetText .ColLetterToNumber("G"), .Row, vSoThueKt
            .GetText .ColLetterToNumber("H"), .Row, vChiTieu12
            .GetText .ColLetterToNumber("I"), .Row, vChiTieu13
            .GetText .ColLetterToNumber("J"), .Row, vChiTieu14
            
            UpdateCell fps, .ColLetterToNumber("C"), 24, vTongSoNguoi
            UpdateCell fps, .ColLetterToNumber("F"), 24, vThuNhapTT
            UpdateCell fps, .ColLetterToNumber("G"), 24, vSoThueKt
            UpdateCell fps, .ColLetterToNumber("H"), 24, vChiTieu12
            UpdateCell fps, .ColLetterToNumber("I"), 24, vChiTieu13
            UpdateCell fps, .ColLetterToNumber("J"), 24, vChiTieu14
                    
       End If
       
      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing

End Sub

'dhdang sua loi tai to khai lam thay doi du lieu
'ngay 08-06-2010
Public Sub FinishImport()
Dim ASheet As Integer, SSheet As Integer
            
'    If flgLoadToKhai = False Then
'        defaultFormula
'    End If
'
'
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
'
        'SetData
        
        ' Lay cac mang MST trong cac bang ke va phu luc 05A, 05B
'        prepareMSTCMT
'
'        kiemTraDuLieuImport
     
        .Sheet = 2
        InsertNode 23, .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone) - 1, 2
        ' Tinh so dong DL trong BK (Phuc vu nut sap xep)
        totalRow5A = .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone) - 22
        
        .Sheet = 3
        InsertNode 23, .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone) - 1, 3
        ' Tinh so dong DL trong BK (Phuc vu nut sap xep)
        totalRow5B = .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone)
        ' test tai khong kiem tra de tang toc do
        .Sheet = 4
        InsertNode 23, .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone) - 1, 4
        
        ' set lai cong thuc
        setFormulaImport
        
        If .ActiveSheet = 4 Then
        
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
    
            Dim index As Integer
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
    
            .Sheet = 4
            i = 1
            .Col = .ColLetterToNumber("B")
            .Row = 22
            .EventEnabled(EventAllEvents) = False
            Do
                .Col = .ColLetterToNumber("L")
                .TypeComboBoxList = strQuanHeNNTID
                index = .typeComboboxCurSel
    
                .Col = .ColLetterToNumber("M")
                .TypeComboBoxList = strQuanHeNNT
                .typeComboboxCurSel = index
    
                    ' set danh sach tinh
                .Col = .ColLetterToNumber("R")
                .TypeComboBoxList = strMaTinh
                index = .typeComboboxCurSel
    
                .Col = .ColLetterToNumber("S")
                .TypeComboBoxList = strTenTinh
                .typeComboboxCurSel = index
    
                ' set danh sach huyen
                .Col = .ColLetterToNumber("T")
                .TypeComboBoxList = strMaHuyen
                index = .typeComboboxCurSel
    
                .Col = .ColLetterToNumber("U")
                .TypeComboBoxList = strTenHuyen
                .typeComboboxCurSel = index
    
                ' set danh sach xa
                .Col = .ColLetterToNumber("V")
                .TypeComboBoxList = strMaXa
                index = .typeComboboxCurSel
    
                .Col = .ColLetterToNumber("W")
                .TypeComboBoxList = strTenXa
                .typeComboboxCurSel = index
    
                ' set danh sach quoc tich
                .Col = .ColLetterToNumber("I")
                .TypeComboBoxList = strMaQuocTich
                index = .typeComboboxCurSel
    
                .Col = .ColLetterToNumber("J")
                .TypeComboBoxList = strQuocTich
                .typeComboboxCurSel = index
    
    
                ' set danh sach quoc gia
                .Col = .ColLetterToNumber("P")
                .TypeComboBoxList = strMaQuocTich
                index = .typeComboboxCurSel
    
                .Col = .ColLetterToNumber("Q")
                .TypeComboBoxList = strQuocTich
                .typeComboboxCurSel = index
    
    
                 .Col = .ColLetterToNumber("B")
                 .Text = i
                 i = i + 1
                 .Row = .Row + 1
            Loop Until .Text = "aa"
    
    '        CheckDynamicError 'Set Exception Error on cells of interface
    '        UpdateSheets
            'end
    '        'CheckDynamicError 'Set Exception Error on cells of interface
    '
    ''        If flgErr5A = True And flgErr5B = True Then kiemTraKhac
            
            .ActiveSheet = ASheet
            .Sheet = SSheet
            .EventEnabled(EventAllEvents) = True
        End If
    End With
    
    'resetFormula
End Sub

'Private Sub prepareMSTCMT()
'    Dim countMST As Long, i As Integer
'    Dim varMST, varCMT, STT As Variant
'    With fps
'        .EventEnabled(EventAllEvents) = False
'        .Sheet = 2 'Bang ke 5A
'        .Row = 22
'
'
'        STT = 1
'        Do
'            'dhdang sua loi danh sô tt
'            'ngay 01/03/2011
'            .Row = .Row
'            .Col = .ColLetterToNumber("B")
'            .Text = STT
'            STT = STT + 1
'
'            .Row = .Row + 1
'            .Col = .ColLetterToNumber("B")
'        Loop Until .Text = "aa"
'
'        .Sheet = 3 'Bang ke 5B
'        .Row = 22
'
'        STT = 1
'        Do
'            'dhdang sua loi danh sô tt
'            'ngay 01/03/2011
'            .Row = .Row
'            .Col = .ColLetterToNumber("B")
'            .Text = STT
'            STT = STT + 1
'
'            .Row = .Row + 1
'            .Col = .ColLetterToNumber("B")
'        Loop Until .Text = "aa"
'
'
'        .EventEnabled(EventAllEvents) = True
'    End With
'End Sub

Public Sub kiemTraDuLieuImport()
'    Debug.Print " batdau kiemtra5A " & Time
'    kiemtra5A
'    Debug.Print " ket thuc kiemtra5A " & Time
'    kiemtra5B
'    prepareMSTCMT
End Sub

'Private Sub kiemtra5A()
'    Dim varTemp As Variant
'    Dim varHoVaTen, varMST, varCMT  As Variant
'    Dim varThuNhap, varBaoHiem, varSoNguoiGiamTru, varSoThangGiamTru, varThuNhapGiamThue, varSoThueDuocGiam, varSoThueKhauTru As Variant
'    Dim varSoTienKH, varSoTienBH, varSoTienGiamThue, varSoThueKT As Variant
'    Dim strCheck As String
'
'    Dim varMSTCQCT As Variant
'
'    Dim startRowMST As Long
'    Dim rowCountMST As Long
'    Dim totalRowMST As Long
'    Dim oldRowMST As Long
'    Dim ret2, tam As Long
'
'    Dim checkQTT As Variant
'    Dim result As Boolean
'    Dim currRow As Long
'
'    Dim i As Integer
'
'    'demo
'    Dim flagCheck As Boolean
'    flagCheck = True
'
'    flgErr5A = True
'    totalRowMST = UBound(arrMST5A)
'
'    ' reset trang thai check trung sheet 2,3 MST
'    Debug.Print "bat dau reset checkTrungAB " & Time
'    resetCheckTrungAB
'    Debug.Print "ket thuc reset checkTrungAB " & Time
''    Debug.Print "bat dau resetcheck " & Time
''    resetCheck
''    Debug.Print "ket thuc resetcheck " & Time
'    ' end
'
'    With fps
'        .EventEnabled(EventAllEvents) = False
'        .Sheet = 2 'Bang ke 5A
'        ' MST co quan chi tra
'        .GetText .ColLetterToNumber("E"), 9, varMSTCQCT
'
'        .Row = 22
'        Do
'            i = .Row
'            ' GD2
'            ' Lam tron so cot thu nhap
'            .GetText .ColLetterToNumber("H"), .Row, varThuNhap
'
'            ' Clear ct
'            If .Row <= .MaxRows - 8 Then
'                .Col = .ColLetterToNumber("M")
'                .Formula = ""
'                .Col = .ColLetterToNumber("Q")
'                .Formula = ""
'            End If
'            ' End clear
'
'            .Col = .ColLetterToNumber("H")
'            .Text = Round(Val(varThuNhap), 0)
'
'            ' Lam tron cot tu thien, khuyen hoc
'            .GetText .ColLetterToNumber("K"), .Row, varSoTienKH
'            .Col = .ColLetterToNumber("K")
'            .Text = Round(Val(varSoTienKH), 0)
'
'            'Lam tron cot bao hiem bat buoc
'            .GetText .ColLetterToNumber("L"), .Row, varSoTienBH
'            .Col = .ColLetterToNumber("L")
'            .Text = Round(Val(varSoTienBH), 0)
'
'            'Lam tron cot can cu giam thue
'            .GetText .ColLetterToNumber("I"), .Row, varSoTienGiamThue
'            .Col = .ColLetterToNumber("I")
'            .Text = Round(Val(varSoTienGiamThue), 0)
'
'            ' Lam tron cot so tien khau tru
'            .GetText .ColLetterToNumber("N"), .Row, varSoThueKT
'            .Col = .ColLetterToNumber("N")
'            .Text = Round(Val(varSoThueKT), 0)
'
'
'            'set lai CT
''            .Col = .ColLetterToNumber("M")
''            .Formula = "IF(H22-J22-K22-L22>0,H22-J22-K22-L22,0)"
''            .Col = .ColLetterToNumber("Q")
''            .Formula = "ROUND(IF(OR(G22=""1"",G22=""x""),AN22-AO22,0),0)"
'            ' End set lai
'
'            ' Cot Ho va ten
'            ' GD2
'            .GetText .ColLetterToNumber("D"), .Row, varHoVaTen
'            ' Cu them mot lao dong thi tinh la mot nguoi
'            If Trim(varHoVaTen) <> vbNullString Then
'                ' Va danh dau cot Ho va ten nay dung de sau nay ko bat loi
'                .Col = .ColLetterToNumber("X")
'                .Text = "0"
'            Else ' Neu ko thi ko tinh nguoi do
'                flgErr5A = False
'                flagCheck = False
'                ' Va danh dau cot Ho va ten nay sai de sau nay bat loi
'                .Col = .ColLetterToNumber("X")
'                .Text = "1"
'            End If
'
'            'reset
'
''                .Col = .ColLetterToNumber("X")
''                .Text = "0"
'                .Col = .ColLetterToNumber("Y")
'                .Text = "0"
'                .Col = .ColLetterToNumber("Z")
'                .Text = "0"
'                .Col = .ColLetterToNumber("AA")
'                .Text = "0"
'                ' MST & CMT
'                .Col = .ColLetterToNumber("AB")
'                .Text = "0"
'                ' Cot check loi
'                .Col = .ColLetterToNumber("AC")
'                .Text = "0"
'
'                .Col = .ColLetterToNumber("AD")
'                .Text = "0"
'
'            ' Cot MST
'            ' GD2
'            .GetText .ColLetterToNumber("E"), .Row, varMST
'            If Trim(varMST) = vbNullString Then
'                ' MST ko bat buoc nhap nua, nen se danh dau MST nay luon dung
''                .Col = .ColLetterToNumber("AD")
''                .Text = "0"
''                ' Cac cot lien quan den MST con lai la dung
''                .Col = .ColLetterToNumber("Y")
''                .Text = "0"
''                .Col = .ColLetterToNumber("Z")
''                .Text = "0"
''                .Col = .ColLetterToNumber("AA")
''                .Text = "0"
'            Else
'                ' Ma > 10 so la sai
'                If Len(Trim(varMST)) <> 10 Then
'                    flgErr5A = False
'                    ' Danh dau cot MST nay sai de sau nay bat loi
'                    flagCheck = False
'                    .Col = .ColLetterToNumber("Y")
'                    .Text = "1"
'                    ' Cac cot lien quan den MST con lai la dung
''                    .Col = .ColLetterToNumber("AD")
''                    .Text = "0"
''                    .Col = .ColLetterToNumber("Z")
''                    .Text = "0"
''                    .Col = .ColLetterToNumber("AA")
''                    .Text = "0"
'                End If
'                ' Ma 10 so
'                If Len(Trim(varMST)) = 10 Then
'                    strCheck = CheckTaxCode(Mid(Trim(varMST), 1, 1), Mid(Trim(varMST), 2, 1), Mid(Trim(varMST), 3, 1), Mid(Trim(varMST), 4, 1), Mid(Trim(varMST), 5, 1), Mid(Trim(varMST), 6, 1), Mid(Trim(varMST), 7, 1), Mid(Trim(varMST), 8, 1), Mid(Trim(varMST), 9, 1), Mid(Trim(varMST), 10, 1), Mid(Trim(varMST), 11, 1), Mid(Trim(varMST), 12, 1), Mid(Trim(varMST), 13, 1))
'                    If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(varMST) Then
'                        flgErr5A = False
'                        ' Danh dau cot MST nay sai de sau nay bat loi
'                        flagCheck = False
'                        .Col = .ColLetterToNumber("Y")
'                        .Text = "1"
''                        ' Cac cot lien quan den MST con lai la dung
''                        .Col = .ColLetterToNumber("AD")
''                        .Text = "0"
''                        .Col = .ColLetterToNumber("Z")
''                        .Text = "0"
''                        .Col = .ColLetterToNumber("AA")
''                        .Text = "0"
'                    Else
'                        ' Kiem tra trung voi MST cqct
'                        If Trim(varMST) = Trim(varMSTCQCT) Then
'                           flgErr5A = False
'                            flagCheck = False
'                           .Col = .ColLetterToNumber("Z")
'                           .Text = "1"
''                           .Col = .ColLetterToNumber("Y")
''                           .Text = "0"
''                           ' Cac cot lien quan den MST con lai la dung
''                           .Col = .ColLetterToNumber("AD")
''                           .Text = "0"
''                           .Col = .ColLetterToNumber("AA")
''                           .Text = "0"
'                        Else
''                            .Col = .ColLetterToNumber("Z")
''                            .Text = "0"
''                                  .Col = .ColLetterToNumber("Y")
''                           .Text = "0"
''                           ' Cac cot lien quan den MST con lai la dung
''                           .Col = .ColLetterToNumber("AD")
''                           .Text = "0"
''                           .Col = .ColLetterToNumber("AA")
''                           .Text = "0"
'                        End If
'                    End If
'                End If
'
'                'dhdang check trung MST
'                ret2 = .SearchCol(.ColLetterToNumber("E"), 21, (21 + .MaxRows), varMST, SearchFlagsNone)
'                If ret2 > -1 And ret2 <> .Row Then
'                    flgErr5A = False
'                    flagCheck = False
'                    ' Danh dau cot MST nay sai de sau nay bat loi
''                    .Col = .ColLetterToNumber("Y")
''                    .Text = "0"
''                    ' Cac cot lien quan den MST con lai la dung
''                    .Col = .ColLetterToNumber("AD")
''                    .Text = "0"
''                    .Col = .ColLetterToNumber("Z")
''                    .Text = "0"
'                    .Col = .ColLetterToNumber("AA")
'                    tam = .Row
'                    .Row = ret2
'                    .Text = "1"
'                    .Row = tam
'                    .Text = "1"
'                End If
'            End If
'
'            ' Cot CMT
'            ' GD2
'            .Row = i
'            .GetText .ColLetterToNumber("E"), .Row, varMST
'            .GetText .ColLetterToNumber("F"), .Row, varCMT
'            If Trim(varMST) = vbNullString And Trim(varCMT) = vbNullString Then
'                    'flgErr5A = False
'                    ' Danh dau ca cot MST va CMT deu khong co
'                    flagCheck = False
'                    .Col = .ColLetterToNumber("AD")
'                    .Text = "1"
'''                    .Text = "0"
''                    ' Danh dau cac cot con lai la dung
''                    .Col = .ColLetterToNumber("Y")
''                    .Text = "0"
''                    .Col = .ColLetterToNumber("Z")
''                    .Text = "0"
''                    .Col = .ColLetterToNumber("AA")
''                    .Text = "0"
''                    .Col = .ColLetterToNumber("AB")
''                    .Text = "0"
''                    .Col = .ColLetterToNumber("AC")
''                    .Text = "0"
''            ElseIf Trim(varCMT) = vbNullString Or Trim(varCMT) = "" Then
''                .Col = .ColLetterToNumber("AD")
''                .Text = "0"
'            ElseIf Trim(varCMT) <> vbNullString Then
'                If InStr(1, Trim(varCMT), " ", vbTextCompare) <> 0 Then
'                     flgErr5A = False
'                    ' Gio ko bat nhap CMT nua nen dang dau cot nay luon dung
'                    flagCheck = False
'                    .Col = .ColLetterToNumber("AC")
'                    .Text = "1"
'                    ' Danh dau cot CMT nay sai de sau nay bat loi
''                    .Col = .ColLetterToNumber("AB")
''                    .Text = "0"
''                    .Col = .ColLetterToNumber("AD")
''                    .Text = "0"
''                Else
''                    ' Danh dau cac cot con lai la dung
''                    .Col = .ColLetterToNumber("AC")
''                    .Text = "0"
''                    .Col = .ColLetterToNumber("AD")
''                    .Text = "0"
''                    .Col = .ColLetterToNumber("AB")
''                    .Text = "0"
'                End If
'
'                'dhdang check trung CMT
'                ret2 = .SearchCol(.ColLetterToNumber("F"), 21, (21 + .MaxRows), varCMT, SearchFlagsNone)
'                If ret2 > -1 And ret2 <> .Row Then
'                    flgErr5A = False
'                    flagCheck = False
'                    ' Danh dau cot CMT nay sai de sau nay bat loi
'                    ' Cac cot lien quan den CMT con lai la dung
''                    .Col = .ColLetterToNumber("AC")
''                    .Text = "0"
''                    .Col = .ColLetterToNumber("AD")
''                    .Text = "0"
'                    .Col = .ColLetterToNumber("AB")
'                    tam = .Row
'                    .Row = ret2
'                    .Text = "1"
'                    .Row = tam
'                    .Text = "1"
'                End If
'            End If
'
'
'            ' Neu co bat cu mot chi tieu nao do bi sai thi danh dau vao de con bat loi
''            If flgErr5A = False Then
''                ' Danh dau cot check loi nay sai sau nay bat loi
''                .Col = .ColLetterToNumber("AH")
''                .Text = "1"
''            Else
''                .Col = .ColLetterToNumber("AH")
''                .Text = "0"
''            End If
'
'            ' Trong truong hop cot Ho ten, MST, CMT, ThuNhap khong co gia tri thi ko kiem tra
'            ' GD2
'            If (Trim(varHoVaTen) = vbNullString And Val(varThuNhap) = 0) Then
'                ' Va danh dau cac cot sau de sau nay ko bat loi
'                ' Ho va ten
''                .Col = .ColLetterToNumber("V")
''                .Text = "0"
''                ' MST
''                .Col = .ColLetterToNumber("W")
''                .Text = "0"
'                .Col = .ColLetterToNumber("X")
'                .Text = "0"
'                .Col = .ColLetterToNumber("Y")
'                .Text = "0"
'                ' CMT
'                .Col = .ColLetterToNumber("Z")
'                .Text = "0"
'                .Col = .ColLetterToNumber("AA")
'                .Text = "0"
'                ' MST & CMT
'                .Col = .ColLetterToNumber("AB")
'                .Text = "0"
'                ' Cot check loi
'                .Col = .ColLetterToNumber("AC")
'                .Text = "0"
'
'                .Col = .ColLetterToNumber("AD")
'                .Text = "0"
'            End If
'
'            If flagCheck = False Then
'                .Col = .ColLetterToNumber("AJ")
'                .Text = "1"
'            End If
'
'            ' kiem tra MST trung voi 5B trong truong hop check quyet toan
'            result = False
'            .GetText .ColLetterToNumber("G"), .Row, checkQTT
'            currRow = .Row
'            If CStr(checkQTT) = "1" Then
'                'result = kiemtra5AB(varMST)
'
'                .Sheet = 3 'Bang ke 5A
'                .Row = 22
'                Do
'                    i = .Row
'                    ' Ma la null
'                        'dhdang check trung MST
'                        If varMST <> "" And varMST <> vbNull Then
'                            ret2 = .SearchCol(.ColLetterToNumber("D"), 21, (21 + .MaxRows), varMST, SearchFlagsNone)
'                            If ret2 > -1 Then
'                                flgErr5AB = False
'                                .Col = .ColLetterToNumber("Z")
'                                tam = .Row
'                                .Row = ret2
'                                .Text = "1"
'                                .Row = tam
'                                result = True
'                            End If
'                        End If
'                    .Row = .Row + 1
'                    .Col = .ColLetterToNumber("B")
'                Loop Until .Text = "aa"
'
'                .Sheet = 2
'                .Row = currRow
'                If result = True Then
'                    .Col = .ColLetterToNumber("AU")
'                    .Text = "1"
'                End If
'            End If
'            'end
'
'
'
'
'            .Row = .Row + 1
'            .Col = .ColLetterToNumber("B")
'        Loop Until .Text = "aa"
'        .EventEnabled(EventAllEvents) = True
'    End With
'
'End Sub

'Private Sub kiemtra5B()
'    Dim varTemp As Variant
'    Dim varHoVaTen, varMST, varCMT  As Variant
'    Dim varThuNhap, varBaoHiem, varSoNguoiGiamTru, varSoThangGiamTru, varThuNhapGiamThue, varSoThueDuocGiam, varSoThueKhauTru As Variant
'    Dim varSoTienGiamThue, varSoTienKhauTru As Variant
'    Dim strCheck As String
'
'    Dim varMSTCQCT As Variant
'
'    Dim startRowMST As Long
'    Dim rowCountMST As Long
'    Dim totalRowMST As Long
'    Dim oldRowMST As Long
'    Dim ret2, tam As Long
'
'    flgErr5B = True
''    totalRowMST = UBound(arrMST5B)
'    Dim i As Integer
'
'    With fps
'        .EventEnabled(EventAllEvents) = False
'        .Sheet = 3 'Bang ke 5B
'        ' MST co quan chi tra
'        .GetText .ColLetterToNumber("D"), 8, varMSTCQCT
'
'        .Row = 22
'        Do
'            i = .Row
'            ' Cot Ho va ten
'            ' GD2
'            .GetText .ColLetterToNumber("C"), .Row, varHoVaTen
'            ' Cu them mot lao dong thi tinh la mot nguoi
'            If Trim(varHoVaTen) <> vbNullString Then
'                ' Them mot nguoi vao cot
'                .Col = .ColLetterToNumber("K")
'                .Text = "1"
'                ' Va danh dau cot Ho va ten nay dung de sau nay ko bat loi
'                .Col = .ColLetterToNumber("N")
'                .Text = "0"
'            Else ' Neu ko thi ko tinh nguoi do
'                flgErr5B = False
'                .Col = .ColLetterToNumber("K")
'                .Text = "0"
'                ' Va danh dau cot Ho va ten nay sai de sau nay bat loi
'                .Col = .ColLetterToNumber("N")
'                .Text = "1"
'            End If
'
'            ' Cot MST
'
'            .GetText .ColLetterToNumber("D"), .Row, varMST
'            ' Ma la null
'            If Trim(varMST) = vbNullString Then
'                'flgErr5B = False
'                ' Danh dau cot MST nay sai de sau nay bat loi
''                .Col = .ColLetterToNumber("R")
''                .Text = "1"
'                ' MST ko con bat rang buoc nua nen danh dau cot nay luon dung
'                .Col = .ColLetterToNumber("T")
'                .Text = "0"
'                ' Cac cot lien quan den MST con lai la dung
'                .Col = .ColLetterToNumber("O")
'                .Text = "0"
'                .Col = .ColLetterToNumber("P")
'                .Text = "0"
'                .Col = .ColLetterToNumber("Q")
'                .Text = "0"
'            Else
'                ' Ma > 10 so la sai
'                If Len(Trim(varMST)) <> 10 Then
'                    flgErr5B = False
'                    ' Danh dau cot MST nay sai de sau nay bat loi
'                    .Col = .ColLetterToNumber("O")
'                    .Text = "1"
'                    ' Cac cot lien quan den MST con lai la dung
'                    .Col = .ColLetterToNumber("T")
'                    .Text = "0"
'                    .Col = .ColLetterToNumber("P")
'                    .Text = "0"
'                    .Col = .ColLetterToNumber("Q")
'                    .Text = "0"
'                End If
'                ' Ma 10 so
'                If Len(Trim(varMST)) = 10 Then
'                    strCheck = CheckTaxCode(Mid(Trim(varMST), 1, 1), Mid(Trim(varMST), 2, 1), Mid(Trim(varMST), 3, 1), Mid(Trim(varMST), 4, 1), Mid(Trim(varMST), 5, 1), Mid(Trim(varMST), 6, 1), Mid(Trim(varMST), 7, 1), Mid(Trim(varMST), 8, 1), Mid(Trim(varMST), 9, 1), Mid(Trim(varMST), 10, 1), Mid(Trim(varMST), 11, 1), Mid(Trim(varMST), 12, 1), Mid(Trim(varMST), 13, 1))
'                    If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(varMST) Then
'                        flgErr5B = False
'                        ' Danh dau cot MST nay sai de sau nay bat loi
'                        .Col = .ColLetterToNumber("O")
'                        .Text = "1"
'                        ' Cac cot lien quan den MST con lai la dung
'                        .Col = .ColLetterToNumber("T")
'                        .Text = "0"
'                        .Col = .ColLetterToNumber("P")
'                        .Text = "0"
'                        .Col = .ColLetterToNumber("Q")
'                        .Text = "0"
'                    Else
''                        .Col = .ColLetterToNumber("M")
''                        .Text = "0"
'                        ' Kiem tra trung voi MST cqct
'                        If Trim(varMST) = Trim(varMSTCQCT) Then
'                           flgErr5B = False
'                           .Col = .ColLetterToNumber("P")
'                           .Text = "1"
'                           .Col = .ColLetterToNumber("O")
'                           .Text = "0"
'                           ' Cac cot lien quan den MST con lai la dung
'                           .Col = .ColLetterToNumber("Q")
'                           .Text = "0"
'                           .Col = .ColLetterToNumber("T")
'                           .Text = "0"
'                        Else
'                            .Col = .ColLetterToNumber("P")
'                            .Text = "0"
'                              .Col = .ColLetterToNumber("O")
'                           .Text = "0"
'                           ' Cac cot lien quan den MST con lai la dung
'                           .Col = .ColLetterToNumber("Q")
'                           .Text = "0"
'                           .Col = .ColLetterToNumber("T")
'                           .Text = "0"
'                         End If
'                    End If
'                End If
'                'dhdang check trung MST
'
'                ret2 = .SearchCol(.ColLetterToNumber("D"), 21, (21 + .MaxRows), varMST, SearchFlagsNone)
'                If ret2 > -1 And ret2 <> .Row Then
'                     flgErr5B = False
'                    ' Danh dau cot MST nay sai de sau nay bat loi
'                    .Col = .ColLetterToNumber("O")
'                    .Text = "0"
'                    ' Cac cot lien quan den MST con lai la dung
'                    .Col = .ColLetterToNumber("T")
'                    .Text = "0"
'                    .Col = .ColLetterToNumber("P")
'                    .Text = "0"
'                    .Col = .ColLetterToNumber("Q")
'                    tam = .Row
'                    .Row = ret2
'                    .Text = "1"
'                    .Row = tam
'                    .Text = "1"
'                End If
'
'            End If
'
'            ' Cot CMT
'            .Row = i
'            .GetText .ColLetterToNumber("D"), .Row, varMST
'            .GetText .ColLetterToNumber("E"), .Row, varCMT
'             If Trim(varMST) = vbNullString And Trim(varCMT) = vbNullString Then
'            'If Trim(varCMT) = "" And Trim(varCMT) = vbNullString Then
'                'flgErr5B = False
'                ' Danh dau ca cot MST va CMT deu khong co
'                ' MST ko con bat rang buoc nua nen danh dau cot nay luon dung
'                .Col = .ColLetterToNumber("T")
'                '.Text = "1"
'                .Text = "0"
'                ' Danh dau cac cot con lai la dung
'                .Col = .ColLetterToNumber("O")
'                .Text = "0"
'                .Col = .ColLetterToNumber("P")
'                .Text = "0"
'                .Col = .ColLetterToNumber("Q")
'                .Text = "0"
'                .Col = .ColLetterToNumber("R")
'                .Text = "0"
'                .Col = .ColLetterToNumber("S")
'                .Text = "0"
'            ElseIf Trim(varCMT) = vbNullString Then
'                .Col = .ColLetterToNumber("T")
'                .Text = "0"
'            ElseIf Trim(varCMT) <> vbNullString Then
'                If InStr(1, Trim(varCMT), " ", vbTextCompare) <> 0 Then
'                     flgErr5B = False
'                    ' Ko bat buoc phai nhap CMT nua nen set gia tri truong nay luon dung
'                    .Col = .ColLetterToNumber("S")
'                    .Text = "1"
'                    ' Danh dau cot CMT nay sai de sau nay bat loi
'                    .Col = .ColLetterToNumber("R")
'                    .Text = "0"
'                    .Col = .ColLetterToNumber("T")
'                    .Text = "0"
'                Else
'                    ' If Trim(varMST) = vbNullString Then flgErr5B = True
'                    ' Danh dau cac cot con lai la dung
'                    .Col = .ColLetterToNumber("S")
'                    .Text = "0"
'                    .Col = .ColLetterToNumber("R")
'                    .Text = "0"
'                    .Col = .ColLetterToNumber("T")
'                    .Text = "0"
'
'
'                    ' Neu ko co loi nao kiem tra xem co trung voi CMT nao tren bang ke khong
'                    ' .Row = i
'
'                End If
'                 'dhdang check trung CMT
'
'                ret2 = .SearchCol(.ColLetterToNumber("E"), 21, (21 + .MaxRows), varCMT, SearchFlagsNone)
'                If ret2 > -1 And ret2 <> .Row Then
'                     flgErr5B = False
'                    .Col = .ColLetterToNumber("R")
'                    tam = .Row
'                    .Row = ret2
'                    .Text = "1"
'                    .Row = tam
'                    .Text = "1"
'                End If
'
'            End If
'
'            ' Cot thu nhap chiu thue
'
'            .GetText .ColLetterToNumber("G"), .Row, varThuNhap
'            .Col = .ColLetterToNumber("G")
'            .Text = Round(Val(varThuNhap), 0)
'
'            ' Lam tron cot can cu giam thue
'            .GetText .ColLetterToNumber("H"), .Row, varSoTienGiamThue
'            .Col = .ColLetterToNumber("H")
'            .Text = Round(Val(varSoTienGiamThue), 0)
'
'
'            ' Lam tron cot so theu da khau tru
'            .GetText .ColLetterToNumber("I"), .Row, varSoTienKhauTru
'            .Col = .ColLetterToNumber("I")
'            .Text = Round(Val(varSoTienKhauTru), 0)
'
'
'
'            .GetText .ColLetterToNumber("H"), .Row, varThuNhapGiamThue
'            If Val(varThuNhapGiamThue) > Val(varThuNhap) Then
'                flgErr5B = False
'                ' Danh dau cot thu nhap giam thue nay sai sau nay bat loi
'                .Col = .ColLetterToNumber("U")
'                .Text = "1"
'            Else
'                ' Danh dau cac cot con lai la dung
'                .Col = .ColLetterToNumber("U")
'                .Text = "0"
'            End If
'
'
'            ' Cot Thue da khau tru
'
'            .GetText .ColLetterToNumber("I"), .Row, varSoThueKhauTru
'            .GetText .ColLetterToNumber("G"), .Row, varThuNhap
'            If Val(varSoThueKhauTru) > 0 And Val(varSoThueKhauTru) >= Val(varThuNhap) Then
'                flgErr5B = False
'                ' Danh dau cot so thue khau tru nay sai sau nay bat loi
'                .Col = .ColLetterToNumber("W")
'                .Text = "1"
'            Else
'                ' Danh dau cac cot con lai la dung
'                .Col = .ColLetterToNumber("W")
'                .Text = "0"
'            End If
'
'
'            ' Neu co bat cu mot chi tieu nao do bi sai thi danh dau vao de con bat loi
''            If flgErr5B = False Then
''                ' Danh dau cot check loi nay sai sau nay bat loi
''                .Col = .ColLetterToNumber("V")
''                .Text = "1"
''            Else
''                .Col = .ColLetterToNumber("V")
''                .Text = "0"
''            End If
'
'            ' Lay so thue da khau tru cua lao dong do
'
''            .GetText .ColLetterToNumber("H"), .Row, varSoThueKhauTru
''            If Val(varSoThueKhauTru) > 0 Then
''                ' Va tinh la lao dong co thue khau tru
''                .Col = .ColLetterToNumber("J")
''                .Text = "1"
''                ' Tinh tong thu nhap cua rieng lao dong co thue khau tru nay
''                .GetText .ColLetterToNumber("F"), .Row, varThuNhap
''                .Col = .ColLetterToNumber("K")
''                .Text = varThuNhap
''            Else
''                ' Va tinh la lao dong ko co thue khau tru
''                .Col = .ColLetterToNumber("J")
''                .Text = "0"
''                ' Khong tinh tong thu nhap cua lao dong nay
''                .Col = .ColLetterToNumber("K")
''                .Text = "0"
''            End If
'
'            ' Trong truong hop cot Ho ten, MST, CMT, ThuNhap khong co gia tri thi ko kiem tra
'            If (Trim(varHoVaTen) = vbNullString And (Trim(varMST) = vbNullString Or Trim(varCMT) = vbNullString) And Val(varThuNhap) = 0) Then
'                ' Va danh dau cac cot sau de sau nay ko bat loi
'                ' Ho va ten
'                .Col = .ColLetterToNumber("N")
'                .Text = "0"
'                ' MST
'                .Col = .ColLetterToNumber("O")
'                .Text = "0"
'                .Col = .ColLetterToNumber("P")
'                .Text = "0"
'                .Col = .ColLetterToNumber("Q")
'                .Text = "0"
'                ' CMT
'                .Col = .ColLetterToNumber("R")
'                .Text = "0"
'                .Col = .ColLetterToNumber("S")
'                .Text = "0"
'                ' MST & CMT
'                .Col = .ColLetterToNumber("T")
'                .Text = "0"
'                ' Cot check loi
''                .Col = .ColLetterToNumber("V")
''                .Text = "0"
'            End If
'            .Row = .Row + 1
'            .Col = .ColLetterToNumber("B")
'        Loop Until .Text = "aa"
'        .EventEnabled(EventAllEvents) = True
'    End With
'End Sub


Private Sub resetCheckTrungAB()
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 3
        .Row = 22
        Do
            .Col = .ColLetterToNumber("AA")
            .Text = "0"
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

'Private Sub resetFormula()
'    Dim i As Integer
'    With fps
'        .Sheet = 1
'        .Col = .ColLetterToNumber("I")
'        .Row = 37
'        .Formula = ""
'        .Row = 38
'        .Formula = ""
'
'        .Row = 40
'        .Formula = ""
'        .Row = 41
'        .Formula = ""
'        .Row = 42
'        .Formula = ""
'
'        .Row = 44
'        .Formula = ""
'        .Row = 45
'        .Formula = ""
'        .Row = 46
'        .Formula = ""
'
'        .Row = 48
'        .Formula = ""
'        .Row = 49
'        .Formula = ""
'        .Row = 50
'        .Formula = ""
'
'        .Row = 61
'        .Formula = ""
'        .Row = 62
'        .Formula = ""
'        .Row = 63
'        .Formula = ""
'    End With
'End Sub

Private Sub defaultFormula()
    Dim i As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        .Col = .ColLetterToNumber("I")
        .Row = 36
        .Formula = "Q36"
        .Row = 37
        .Formula = "Q37"
        
        .Row = 39
        .Formula = "Q39"
        '.Row = 40
'        .Formula = "Q40"
'        .Row = 41
'        .Formula = "Q41"
        
        .Row = 43
        .Formula = "Q43"
        .Row = 44
        .Formula = "Q44"
        .Row = 45
        .Formula = "Q45"
        
        
        .Row = 47
        .Formula = "Q47"
        .Row = 48
        .Formula = "Q48"
        .Row = 49
        .Formula = "Q49"
        
        
        .Row = 51
        .Formula = "Q51"
        .Row = 52
        .Formula = "Q52"
        .Row = 53
        .Formula = "Q53"
        
        
        .Row = 61
        .Formula = "Q61"
        .Row = 62
        .Formula = "Q62"
        .Row = 63
        .Formula = "Q63"
        
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub


Public Sub InsertNode(ByVal fRow As Long, ByVal lRow As Long, ByVal mSheet As Byte)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCell As MSXML.IXMLDOMNode
    Dim arrTemp() As String
    
    
    Dim ctlRow As Long
    Dim cTemplate, cReport As String
    
    Dim sumRowDel As Long
    sumRowDel = TAX_Utilities_v1.Data(mSheet - 1).getElementsByTagName("Cells").length
    For ctlRow = 1 To sumRowDel - 1
        If mSheet = 2 Then
            Set xmlNodeCells = TAX_Utilities_v1.Data(mSheet - 1).nodeFromID("D" & "_" & 22 + ctlRow & "").parentNode
            xmlNodeCells.parentNode.removeChild xmlNodeCells
        ElseIf mSheet = 3 Then
            Set xmlNodeCells = TAX_Utilities_v1.Data(mSheet - 1).nodeFromID("C" & "_" & 22 + ctlRow & "").parentNode
            xmlNodeCells.parentNode.removeChild xmlNodeCells
        ElseIf mSheet = 4 Then
            Set xmlNodeCells = TAX_Utilities_v1.Data(mSheet - 1).nodeFromID("D" & "_" & 22 + ctlRow & "").parentNode
            xmlNodeCells.parentNode.removeChild xmlNodeCells
        ElseIf mSheet = 5 Then
            Set xmlNodeCells = TAX_Utilities_v1.Data(mSheet - 1).nodeFromID("C" & "_" & 22 + ctlRow & "").parentNode
            xmlNodeCells.parentNode.removeChild xmlNodeCells
        End If
    Next
    
    If mSheet = 2 Then
        Set xmlNodeCells = TAX_Utilities_v1.Data(mSheet - 1).nodeFromID("D_22").parentNode
    ElseIf mSheet = 3 Then
        Set xmlNodeCells = TAX_Utilities_v1.Data(mSheet - 1).nodeFromID("C_22").parentNode
    ElseIf mSheet = 4 Then
        Set xmlNodeCells = TAX_Utilities_v1.Data(mSheet - 1).nodeFromID("D_22").parentNode
    ElseIf mSheet = 5 Then
        Set xmlNodeCells = TAX_Utilities_v1.Data(mSheet - 1).nodeFromID("C_22").parentNode
    End If
    
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
            If mSheet = 2 Then
                If Trim(cTemplate) = "D_" Or Trim(cTemplate) = "E_" Or Trim(cTemplate) = "F_" Then
                    SetAttribute xmlNodeNewCell, "Value", vbNullString
                Else
                    SetAttribute xmlNodeNewCell, "Value", "0"
                End If
                SetAttribute xmlNodeNewCell, "CellID2", cReport & (ctlRow - 7)
            ElseIf mSheet = 3 Then
                If Trim(cTemplate) = "C_" Or Trim(cTemplate) = "D_" Or Trim(cTemplate) = "E_" Then
                    SetAttribute xmlNodeNewCell, "Value", vbNullString
                Else
                    SetAttribute xmlNodeNewCell, "Value", "0"
                End If
                SetAttribute xmlNodeNewCell, "CellID2", cReport & (ctlRow - 7)
            ElseIf mSheet = 4 Then
                SetAttribute xmlNodeNewCell, "Value", vbNullString
                SetAttribute xmlNodeNewCell, "CellID2", cReport & (ctlRow - 7)
            ElseIf mSheet = 5 Then
                If Trim(cTemplate) = "C_" Or Trim(cTemplate) = "D_" Or Trim(cTemplate) = "E_" Then
                    SetAttribute xmlNodeNewCell, "Value", vbNullString
                Else
                    SetAttribute xmlNodeNewCell, "Value", "0"
                End If
                SetAttribute xmlNodeNewCell, "CellID2", cReport & (ctlRow - 4)
            End If
            
        Next
        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, Null
        
    
    Next
'   **********************************
'    added
'   Date: 12/04/06
    'mAdjustData = True
    TAX_Utilities_v1.AdjustData(mSheet - 1) = True
'   **********************************
EXIT_SUB:
    Set xmlNodeNewCell = Nothing
    Set xmlNodeNewCells = Nothing
    Set xmlNodeCells = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_05TNCN", "InsertNode 05_TNCN", Err.number, Err.Description
End Sub

' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 15
    denRowDL = 20
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "7_8_2~7_8_3~7_8_4~10_11_5"
    ' end
End Sub

'Private Sub resetCheck()
'    With fps
'        .EventEnabled(EventAllEvents) = False
'        .Sheet = 2
'        .Row = 22
'        Do
'            .Col = .ColLetterToNumber("X")
'            .Text = "0"
'            .Col = .ColLetterToNumber("Y")
'            .Text = "0"
'            .Col = .ColLetterToNumber("Z")
'            .Text = "0"
'            .Col = .ColLetterToNumber("AA")
'            .Text = "0"
'            .Col = .ColLetterToNumber("AB")
'            .Text = "0"
'            .Col = .ColLetterToNumber("AC")
'            .Text = "0"
'            .Col = .ColLetterToNumber("AD")
'            .Text = "0"
'            .Row = .Row + 1
'            .Col = .ColLetterToNumber("B")
'        Loop Until .Text = "aa"
'        .EventEnabled(EventAllEvents) = True
'    End With
'End Sub


Public Sub CheckDynamicSheet1()
    Dim vErrMsgHoVaTen, vErrMsgMST, vErrMsgMST1, vErrMsgMST2, vErrMsgCMT, vErrMsgCMT1, vErrMsgMSTCMT As Variant
    Dim vErrMsgBaoHiem, vErrMsgThangGiamTru, vErrMsgThuNhapGiamThue, vErrMsgSoThueDuocGiam, vErrMsgSoThueKhauTru, vErrMsgTongGiamTruGiaCanh As Variant
    Dim vErrMsgGiamThueTheoHD As Variant
    Dim vErrMsgQTCQCT As Variant
    Dim vErrSoThueNNT As Variant
    Dim vLastRow As Variant
    Dim varTemp As Variant, varTemp1 As Variant
    Dim i, j As Integer
    Dim iFocusFlag As Integer
    Dim strFocusSheetName, strFocusCol, strFocusRow As String
    Dim flgQTCQCT As Boolean
    
    Dim vErrMsgMSTAB As Variant
    Dim flagTrungMSTAB As Integer
    
    
    
    'dntai 22/02/2012 them bien
    Dim vHoVaTen As Variant, vMST As Variant, vCMTND As Variant, chkUyQuyen As Variant, vTongSo As Variant, vTNCNLamCanCuMG As Variant
    Dim vTongTienGiamTru, vTuThienND, vBHBatBuoc, vTNTinhThue, vTNCNDaKT, vTNCNPhaiKT, vTongThuePhaiNop, vThueNopThua, vThuePhaiKTThem As Variant
    
    Dim vGiamTruTheoHD, vQuyHuuTri As Variant
    Dim vMSTCQCT As Variant
    Dim vErrMsg1213 As Variant
    Dim vErrMsgHuuTriKT As Variant
    
    Dim iFlagTrongHoTen As Integer, iFlagLoiMST As Integer, iFlagLoiMST1 As Integer, iFlagLoiMST2 As Integer, iFlagMSTTrung As Integer
    Dim iFlagLoiCMT As Integer, iFlagLoiCMT1 As Integer, iFlagCheckCT12CT11 As Integer, iFlagCheckCT17CT11 As Integer, iFlagMSTCMT As Integer, iFlagQTCQCT As Integer
    Dim iFlagCheckCT13CT11 As Integer
    Dim iFlagLamTronTongGiamTruGiaCanh As Integer
    Dim countRow As Integer
    Dim strCheck As String, CountRow1 As Long, rowDes As Long
    Dim tam As Long
    
    flgQTCQCT = True
    
    ' Tinh so nguoi dc MTb1
    Dim soNguoiDuocMTB1 As Long
    Dim rowTemp As Integer
    Dim totalRow As Integer
    Dim countRowDel As Integer
    
    soNguoiDuocMTB1 = 0
    countRowDel = 0
    
    'reset cell danh dau trung mst tren 2 phu luc 05A va 05B
    resetCheckTrungAB
    
    With fps
        .EventEnabled(EventAllEvents) = False
        
        .Sheet = .SheetCount
                
        .GetText .ColLetterToNumber("E"), 21, vErrMsgHoVaTen
        .GetText .ColLetterToNumber("E"), 22, vErrMsgMST
        .GetText .ColLetterToNumber("E"), 23, vErrMsgMST1
        .GetText .ColLetterToNumber("E"), 24, vErrMsgMST2
        .GetText .ColLetterToNumber("E"), 25, vErrMsgCMT
        .GetText .ColLetterToNumber("E"), 26, vErrMsgCMT1
        .GetText .ColLetterToNumber("E"), 27, vErrMsgMSTCMT
        
        .GetText .ColLetterToNumber("E"), 28, vErrMsgBaoHiem
        .GetText .ColLetterToNumber("E"), 29, vErrMsgThangGiamTru
        .GetText .ColLetterToNumber("E"), 30, vErrMsgThuNhapGiamThue
        .GetText .ColLetterToNumber("E"), 31, vErrMsgSoThueKhauTru
        
        .GetText .ColLetterToNumber("E"), 47, vErrMsgMSTAB
        
        .GetText .ColLetterToNumber("E"), 51, vErrMsgQTCQCT
        .GetText .ColLetterToNumber("E"), 55, vErrSoThueNNT
        .GetText .ColLetterToNumber("E"), 57, vErrMsgTongGiamTruGiaCanh
        .GetText .ColLetterToNumber("E"), 67, vErrMsgGiamThueTheoHD
        .GetText .ColLetterToNumber("E"), 68, vErrMsgHuuTriKT
        
        .GetText .ColLetterToNumber("E"), 80, vErrMsg1213
        
        
        
    .Sheet = 2 'Check tinh xac thuc cua cac thong tin tren bang ke 05A/TNCN
                        
        ' Cot check loi la AD24 duoc danh dau xem bang ke co con bi sai hay khong. Neu sai thi moi check loi tiep
'        .GetText .ColLetterToNumber("AG"), 16, varTemp
'        If Val(varTemp) = 0 Then Exit Sub
        
        
    
           
    'get MSTCQCT
    .GetText .ColLetterToNumber("E"), 9, vMSTCQCT
           
    countRow = 22
    .Row = countRow
    
    Do
    
'        ' lam tron du lieu
'         lamTron .Row, .Sheet
         
        .GetText .ColLetterToNumber("D"), .Row, vHoVaTen
        .GetText .ColLetterToNumber("E"), .Row, vMST
        .GetText .ColLetterToNumber("F"), .Row, vCMTND
        .GetText .ColLetterToNumber("G"), .Row, chkUyQuyen
        .GetText .ColLetterToNumber("H"), .Row, vTongSo
        .GetText .ColLetterToNumber("I"), .Row, vTNCNLamCanCuMG
        .GetText .ColLetterToNumber("J"), .Row, vGiamTruTheoHD
        
        .GetText .ColLetterToNumber("K"), .Row, vTongTienGiamTru
        .GetText .ColLetterToNumber("L"), .Row, vTuThienND
        .GetText .ColLetterToNumber("M"), .Row, vBHBatBuoc
        .GetText .ColLetterToNumber("N"), .Row, vQuyHuuTri
        
        .GetText .ColLetterToNumber("O"), .Row, vTNTinhThue
        .GetText .ColLetterToNumber("P"), .Row, vTNCNDaKT
        
        If Trim(vHoVaTen) = vbNullString And Trim(vMST) = vbNullString And Trim(vCMTND) = vbNullString And (Trim(chkUyQuyen) = "0" Or Trim(chkUyQuyen) = "") And Trim(vTongSo) = "0" _
           And Trim(vTNCNLamCanCuMG) = "0" And Trim(vTongTienGiamTru) = "0" Then
            'reset color and cellnote
            .Col = .ColLetterToNumber("D")
            .BackColor = mNonErrorColor
            .CellNote = ""
            .Col = .ColLetterToNumber("E")
            .BackColor = mNonErrorColor
            .CellNote = ""
            .Col = .ColLetterToNumber("F")
            .BackColor = mNonErrorColor
            .CellNote = ""
            .Col = .ColLetterToNumber("G")
            .BackColor = mNonErrorColor
            .CellNote = ""
            .Col = .ColLetterToNumber("I")
            .BackColor = mNonErrorColor
            .CellNote = ""
            .Col = .ColLetterToNumber("P")
            .BackColor = mNonErrorColor
            .CellNote = ""
        Else
            'check hovaten khong duoc de trong
            If Trim(vHoVaTen) = vbNullString Then
                .Col = .ColLetterToNumber("D")
                .CellNote = vErrMsgHoVaTen
                .BackColor = mErrorColor
                iFlagTrongHoTen = 1
                
            Else
                .Col = .ColLetterToNumber("D")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            
            'check mst
            'set mst thanh khong loi
            .Col = .ColLetterToNumber("E")
            .CellNote = ""
            .BackColor = mNonErrorColor
            'set cell check thanh khong loi
            .Col = .ColLetterToNumber("G")
            .CellNote = ""
            .BackColor = mNonErrorColor
            
            If Trim(chkUyQuyen) = "1" Or UCase(Trim(chkUyQuyen)) = "X" Then
                'check mst khong duoc de trong neu check uy quyen
                If Trim(vMST) = vbNullString Then
                    .Col = .ColLetterToNumber("E")
                    .CellNote = vErrMsgMSTCMT
                    .BackColor = mErrorColor
                    iFlagLoiMST = 1
                    .Col = .ColLetterToNumber("G")
                    .BackColor = mErrorColor
                    .CellNote = vErrMsgQTCQCT
                    iFlagQTCQCT = 1
                End If
            End If
            
            If Len(Trim(vMST)) > 0 Then
                If Len(Trim(vMST)) = 10 Then

                    ' check cau truc mst
                     strCheck = CheckTaxCode(Mid(Trim(vMST), 1, 1), Mid(Trim(vMST), 2, 1), Mid(Trim(vMST), 3, 1), Mid(Trim(vMST), 4, 1), Mid(Trim(vMST), 5, 1), Mid(Trim(vMST), 6, 1), Mid(Trim(vMST), 7, 1), Mid(Trim(vMST), 8, 1), Mid(Trim(vMST), 9, 1), Mid(Trim(vMST), 10, 1), Mid(Trim(vMST), 11, 1), Mid(Trim(vMST), 12, 1), Mid(Trim(vMST), 13, 1))
                    If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                            .Col = .ColLetterToNumber("E")
                            .CellNote = vErrMsgMST
                            .BackColor = mErrorColor
                            iFlagLoiMST1 = 1
                    Else
                        ' check mst trung voi mst cqct
                        If Trim(vMST) = Trim(vMSTCQCT) Then
                            .Col = .ColLetterToNumber("E")
                            .CellNote = vErrMsgMST1
                            .BackColor = mErrorColor
                            iFlagLoiMST2 = 1
                        Else
                            'check MST trung nhau
                             CountRow1 = .SearchCol(.ColLetterToNumber("E"), -1, -1, vMST, SearchFlagsNone)
                             rowDes = .SearchCol(.ColLetterToNumber("E"), CountRow1, -1, vMST, SearchFlagsNone)
                            If rowDes > -1 Then
                                .Col = .ColLetterToNumber("E")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrMsgMST2       'static
                                .BackColor = mErrorColor
                                iFlagMSTTrung = 1
                            End If
                        End If
                    
                    End If
                    
                Else
                    .Col = .ColLetterToNumber("E")
                    .CellNote = vErrMsgMST
                    .BackColor = mErrorColor
                    iFlagLoiMST1 = 1
                End If
            End If
            
            'set lai cot cmtnd
            .Col = .ColLetterToNumber("F")
            .CellNote = ""
            .BackColor = mNonErrorColor
            
            'ktra neu check vao check uy quyen thi ktra mst va cmtnd ko dc de trong
'            If Trim(chkUyQuyen) = "1" Or UCase(Trim(chkUyQuyen)) = "X" Then
                    
           If Trim(vMST) = vbNullString And Trim(vCMTND) = vbNullString Then
                .Col = .ColLetterToNumber("E")
                .CellNote = vErrMsgMSTCMT
                .BackColor = mErrorColor
                .Col = .ColLetterToNumber("F")
                .CellNote = vErrMsgMSTCMT
                .BackColor = mErrorColor
                iFlagLoiMST = 1
                
'            End If
                
            End If
                       
            If Trim(vCMTND) <> vbNullString Then
                If InStr(1, vCMTND, " ", vbTextCompare) <> 0 Then
                    .Col = .ColLetterToNumber("F")
                    .CellNote = vErrMsgCMT1
                    .BackColor = mErrorColor
                    iFlagLoiCMT1 = 1
                Else
                    'check CMTND trung nhau
                     CountRow1 = .SearchCol(.ColLetterToNumber("F"), -1, -1, vCMTND, SearchFlagsNone)
                     rowDes = .SearchCol(.ColLetterToNumber("F"), CountRow1, -1, vCMTND, SearchFlagsNone)
                    If rowDes > -1 Then
                        .Col = .ColLetterToNumber("F")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgCMT        'static
                        .BackColor = mAlertColor
                        'iFlagLoiCMT = 1
                    End If
                End If
            End If
            
            
            'check CT 12 < CT11
            
            .GetText .ColLetterToNumber("AI"), .Row, varTemp
            If Trim(varTemp) = "1" Then
                .Col = .ColLetterToNumber("I")
                .BackColor = mErrorColor
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsgThuNhapGiamThue
                iFlagCheckCT12CT11 = 1
            Else
                .Col = .ColLetterToNumber("I")
                .BackColor = mNonErrorColor
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""
            End If
                
            ' check CT 13 < CT11
            .GetText .ColLetterToNumber("BG"), .Row, varTemp
            If Trim(varTemp) = "1" Then
                .Col = .ColLetterToNumber("J")
                .BackColor = mErrorColor
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsgGiamThueTheoHD
            Else
                .Col = .ColLetterToNumber("J")
                .BackColor = mNonErrorColor
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""
            End If
            
            ' kiem tra tong 12 + 13 <=11
            
            
            'check CT 19 < CT11
            .GetText .ColLetterToNumber("AK"), .Row, varTemp
            If Val(varTemp) = 1 Then
                .Col = .ColLetterToNumber("P")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsgSoThueKhauTru
                '.BackColor = mErrorColor
                .BackColor = mAlertColor
                'iFlagCheckCT17CT11 = 1
            Else
                .Col = .ColLetterToNumber("P")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            
            'check CT14 phai trong den hang tram nghin
            
            .GetText .ColLetterToNumber("BA"), .Row, varTemp
            If Val(varTemp) = 1 Then
                .Col = .ColLetterToNumber("K")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsgTongGiamTruGiaCanh
                .BackColor = mErrorColor
                iFlagLamTronTongGiamTruGiaCanh = 1
            Else
                .Col = .ColLetterToNumber("K")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            
            'check CT 12+13 <= CT11
            .GetText .ColLetterToNumber("BJ"), .Row, varTemp
            If Val(varTemp) = 1 Then
                .Col = .ColLetterToNumber("I")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsg1213
                .BackColor = mErrorColor
                '.BackColor = mAlertColor
                
                
                .Col = .ColLetterToNumber("J")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsg1213
                .BackColor = mErrorColor
                '.BackColor = mAlertColor
                'iFlagCheckCT17CT11 = 1
            Else
                .Col = .ColLetterToNumber("I")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""
                .BackColor = mNonErrorColor
                
                
                .Col = .ColLetterToNumber("J")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            
            ' check CT 17
            .GetText .ColLetterToNumber("BH"), .Row, varTemp
            If Trim(varTemp) = "1" Then
                .Col = .ColLetterToNumber("N")
                .BackColor = mErrorColor
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsgHuuTriKT
            Else
                .Col = .ColLetterToNumber("N")
                .BackColor = mNonErrorColor
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""
            End If
            
'            ' check chi tieu 16 phai nho hon chi tieu 11
'            .GetText .ColLetterToNumber("AG"), .Row, varTemp
'            If Val(varTemp) = 1 Then
'                .Col = .ColLetterToNumber("M")
'                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                .CellNote = vErrMsgBaoHiem
'                .BackColor = mErrorColor
'            Else
'                .Col = .ColLetterToNumber("M")
'                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'            End If

            
            
            tam = .Row
            
            If (CStr(chkUyQuyen) = "1" Or UCase(CStr(chkUyQuyen)) = "X") And Trim(vMST) <> vbNullString Then
                'result = kiemtra5AB(varMST)
                
                .Sheet = 3 'Bang ke 5A
                .Row = 22
                Do
                    ' Ma la null
                        'dhdang check trung MST
                        .GetText .ColLetterToNumber("D"), .Row, varTemp
                        If Trim(varTemp) = Trim(vMST) Then
                            .Col = .ColLetterToNumber("AA")
                            .Text = "1"
                            flagTrungMSTAB = 1
                        End If
        
                    .Row = .Row + 1
                    .Col = .ColLetterToNumber("B")
                Loop Until .Text = "aa"
                
                .Sheet = 2
                .Row = tam
                'neu mst trung voi 05B thi danh dau tren ca 05A
                If flagTrungMSTAB = 1 Then
                    .Col = .ColLetterToNumber("E")
                    .BackColor = mErrorColor
                    .CellNote = vErrMsgMSTAB
                    'set lai bien check mst trung ab
                    flagTrungMSTAB = 0
                End If
                
            End If
        
        End If
        
        ' Kiem tra cac ca nhan duoc MT b1
        If TAX_Utilities_v1.Year = "2012" Then
            .GetText .ColLetterToNumber("BD"), .Row, varTemp
             If varTemp = "1" Then
                soNguoiDuocMTB1 = soNguoiDuocMTB1 + 1
            End If
        End If
        
        countRow = countRow + 1
        .Row = countRow
        .Col = .ColLetterToNumber("B")
        
    Loop Until .Text = "aa"
    
    
    ' Kiem tra neu ca nhan b1 nam 2012 thi chuyen sang pl 27/MT
    Dim tongSoThuePN As Variant
    If TAX_Utilities_v1.Year = "2012" Then
        If soNguoiDuocMTB1 > 0 Then
            rowTemp = 22
            .Row = rowTemp
            i = 0
            ReDim arrData5A(soNguoiDuocMTB1, 8) As String
            Do
                .GetText .ColLetterToNumber("BB"), .Row, varTemp
                If varTemp = "1" Then
                     .GetText .ColLetterToNumber("D"), .Row, vHoVaTen
                    .GetText .ColLetterToNumber("E"), .Row, vMST
                    .GetText .ColLetterToNumber("F"), .Row, vCMTND
                    .GetText .ColLetterToNumber("O"), .Row, vTNTinhThue
                    .GetText .ColLetterToNumber("P"), .Row, vTNCNDaKT
                    .GetText .ColLetterToNumber("H"), .Row, vTongSo
                    .GetText .ColLetterToNumber("I"), .Row, vTNCNLamCanCuMG
                    .GetText .ColLetterToNumber("BE"), .Row, tongSoThuePN
                    
                    
                                        
                    arrData5A(i, 0) = vHoVaTen
                    arrData5A(i, 1) = vMST
                    arrData5A(i, 2) = vCMTND
                    arrData5A(i, 3) = vTNTinhThue
                    arrData5A(i, 4) = vTNCNDaKT
                    arrData5A(i, 5) = vTongSo
                    arrData5A(i, 6) = vTNCNLamCanCuMG
                    arrData5A(i, 7) = tongSoThuePN
                    i = i + 1
                End If
                                
                rowTemp = rowTemp + 1
                .Row = rowTemp
                .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
            
            ' Insert du lieu sang bang ke mien giam thue
            .Sheet = 5
            .Row = 22
            .Col = .ColLetterToNumber("B")
            Do
                countRowDel = countRowDel + 1
                .Row = countRowDel + 22
            Loop Until .Text = "aa"
            
            If countRowDel > 1 Then
                .DeleteRows 23, countRowDel - 1
                .MaxRows = .MaxRows - countRowDel + 1
            End If
            
            .Row = 22
            totalRow = UBound(arrData5A) - 1
            If totalRow >= 0 Then
                For i = 0 To totalRow
                    .Row = i + 22
                    ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                    .Col = .ColLetterToNumber("C")
                    .Text = "0"
                    For j = 0 To 7
                        ' Tu cot Ho va ten cho den cot thue da khau tru (D -> L). Dung cac thong tin theo bang ke 05
                        .Col = 3 + j
                        'sua loi tong hop sai dau "."
                        If j <= 4 Then
                            If .CellType = CellTypeNumber Then
                                .Text = Round(Val(arrData5A(i, j)), 0)
                            Else
                                .Text = arrData5A(i, j)
                            End If
                         Else
                            If j = 5 Then
                                .Col = .ColLetterToNumber("N")
                                If .CellType = CellTypeNumber Then
                                    .Text = Round(Val(arrData5A(i, j)), 0)
                                Else
                                    .Text = arrData5A(i, j)
                                End If
                            ElseIf j = 6 Then
                                .Col = .ColLetterToNumber("O")
                                If .CellType = CellTypeNumber Then
                                    .Text = Round(Val(arrData5A(i, j)), 0)
                                Else
                                    .Text = arrData5A(i, j)
                                End If
                            ElseIf j = 7 Then
                                .Col = .ColLetterToNumber("J")
                                If .CellType = CellTypeNumber Then
                                    .Text = Round(Val(arrData5A(i, j)), 0)
                                Else
                                    .Text = arrData5A(i, j)
                                End If
                            End If
                         End If
                    Next
                    
                    .Col = .ColLetterToNumber("B")
                    .Text = i + 1
                    
                    ' Trong truong hop la dong cuoi cung trong mang thi ko Insert them row nao.
                    If i <> totalRow Then
                        .ReDraw = False
                        .MaxRows = .MaxRows + 1
                        .InsertRows .Row + 1, 1
                        .CopyRowRange 22, 22, .Row + 1  ' Copy nguyen ban cac format cua dong bat dau la dong 22
                    End If
                Next
            End If
            ' set tong so nguoi duoc mien thue b1
            .Row = totalRow + 24
            .Col = .ColLetterToNumber("C")
            .Text = soNguoiDuocMTB1
        End If
        
        
    End If
    ' ket thuc kiem tra
    
    
    
          
    .Sheet = .SheetCount
    
    .SetText .ColLetterToNumber("B"), 21, IIf(iFlagTrongHoTen = 1, "0", "1")
    .SetText .ColLetterToNumber("B"), 27, IIf(iFlagLoiMST = 1, "0", "1")
    .SetText .ColLetterToNumber("B"), 22, IIf(iFlagLoiMST1 = 1, "0", "1")
    .SetText .ColLetterToNumber("B"), 23, IIf(iFlagLoiMST2 = 1, "0", "1")
    .SetText .ColLetterToNumber("B"), 24, IIf(iFlagMSTTrung = 1, "0", "1")
    .SetText .ColLetterToNumber("B"), 25, IIf(iFlagLoiCMT = 1, "0", "1")
    .SetText .ColLetterToNumber("B"), 26, IIf(iFlagLoiCMT1 = 1, "0", "1")
    .SetText .ColLetterToNumber("B"), 30, IIf(iFlagCheckCT12CT11 = 1, "0", "1")
    .SetText .ColLetterToNumber("B"), 31, IIf(iFlagCheckCT17CT11 = 1, "0", "1")
    .SetText .ColLetterToNumber("B"), 51, IIf(iFlagQTCQCT = 1, "0", "1")
    .SetText .ColLetterToNumber("B"), 57, IIf(iFlagLamTronTongGiamTruGiaCanh = 1, "0", "1")
    
    .EventEnabled(EventAllEvents) = True
    End With
    
End Sub


Public Sub fillDataPL27()
' Tinh so nguoi dc MTb1
    Dim soNguoiDuocMTB1 As Long
    Dim rowTemp As Integer
    Dim totalRow As Integer
    Dim countRowDel As Integer
    Dim countRow As Integer
    Dim varTemp As Variant
    Dim i As Integer, j As Integer
    Dim vHoVaTen As Variant, vMST As Variant, vCMTND As Variant, vTNTinhThue As Variant
    Dim vTNCNDaKT As Variant, vTongSo As Variant, vTNCNLamCanCuMG As Variant

    soNguoiDuocMTB1 = 0
    countRowDel = 0
     With fps
        .EventEnabled(EventAllEvents) = False
        
        countRow = 22
        .Row = countRow
        .Sheet = 2
        Do
            ' Kiem tra cac ca nhan duoc MT b1
            If TAX_Utilities_v1.Year = "2012" Then
                .GetText .ColLetterToNumber("BD"), .Row, varTemp
                 If varTemp = "1" Then
                    soNguoiDuocMTB1 = soNguoiDuocMTB1 + 1
                End If
            End If
            ' ket thuc kiem tra
                    
            countRow = countRow + 1
            .Row = countRow
            .Col = .ColLetterToNumber("B")
            
        Loop Until .Text = "aa"
        
        
        ' Kiem tra neu ca nhan b1 nam 2012 thi chuyen sang pl 27/MT
       Dim tongSoThuePN As Variant
        If TAX_Utilities_v1.Year = "2012" Then
            If soNguoiDuocMTB1 > 0 Then
                rowTemp = 22
                .Row = rowTemp
                i = 0
                ReDim arrData5A(soNguoiDuocMTB1, 8) As String
                Do
                    .GetText .ColLetterToNumber("BD"), .Row, varTemp
                    If varTemp = "1" Then
                         .GetText .ColLetterToNumber("D"), .Row, vHoVaTen
                        .GetText .ColLetterToNumber("E"), .Row, vMST
                        .GetText .ColLetterToNumber("F"), .Row, vCMTND
                        .GetText .ColLetterToNumber("O"), .Row, vTNTinhThue
                        .GetText .ColLetterToNumber("P"), .Row, vTNCNDaKT
                        .GetText .ColLetterToNumber("H"), .Row, vTongSo
                        .GetText .ColLetterToNumber("I"), .Row, vTNCNLamCanCuMG
                        .GetText .ColLetterToNumber("BE"), .Row, tongSoThuePN
                        
                        
                                            
                        arrData5A(i, 0) = vHoVaTen
                        arrData5A(i, 1) = vMST
                        arrData5A(i, 2) = vCMTND
                        arrData5A(i, 3) = Round(vTNTinhThue, 0)
                        arrData5A(i, 4) = Round(vTNCNDaKT, 0)
                        arrData5A(i, 5) = Round(vTongSo, 0)
                        arrData5A(i, 6) = Round(vTNCNLamCanCuMG, 0)
                        arrData5A(i, 7) = Round(tongSoThuePN, 0)
                        i = i + 1
                    End If
                                    
                    rowTemp = rowTemp + 1
                    .Row = rowTemp
                    .Col = .ColLetterToNumber("B")
                Loop Until .Text = "aa"
                
                ' Insert du lieu sang bang ke mien giam thue
                .Sheet = 5
                .Row = 22
                .Col = .ColLetterToNumber("B")
                Do
                    countRowDel = countRowDel + 1
                    .Row = countRowDel + 22
                Loop Until .Text = "aa"
                
                If countRowDel > 1 Then
                    .DeleteRows 23, countRowDel - 1
                    .MaxRows = .MaxRows - countRowDel + 1
                End If
                
                .Row = 22
                totalRow = UBound(arrData5A) - 1
                If totalRow >= 0 Then
                    For i = 0 To totalRow
                        .Row = i + 22
                        ' Set gia tri ban dau cua hop checkbox la 0, tuc la ko chon de in
                        .Col = .ColLetterToNumber("C")
                        .Text = "0"
                        For j = 0 To 7
                            ' Tu cot Ho va ten cho den cot thue da khau tru (D -> L). Dung cac thong tin theo bang ke 05
                            .Col = 3 + j
                            'sua loi tong hop sai dau "."
                            If j <= 4 Then
                                If .CellType = CellTypeNumber Then
                                    .Text = Round(Val(arrData5A(i, j)), 0)
                                Else
                                    .Text = arrData5A(i, j)
                                End If
                             Else
                                If j = 5 Then
                                    .Col = .ColLetterToNumber("N")
                                    If .CellType = CellTypeNumber Then
                                        .Text = Round(Val(arrData5A(i, j)), 0)
                                    Else
                                        .Text = arrData5A(i, j)
                                    End If
                                ElseIf j = 6 Then
                                    .Col = .ColLetterToNumber("O")
                                    If .CellType = CellTypeNumber Then
                                        .Text = Round(Val(arrData5A(i, j)), 0)
                                    Else
                                        .Text = arrData5A(i, j)
                                    End If
                                ElseIf j = 7 Then
                                    .Col = .ColLetterToNumber("J")
                                    If .CellType = CellTypeNumber Then
                                        .Text = Round(Val(arrData5A(i, j)), 0)
                                    Else
                                        .Text = arrData5A(i, j)
                                    End If
                                End If
                             End If
                        Next
                        
                        .Col = .ColLetterToNumber("B")
                        .Text = i + 1
                        
                        ' Trong truong hop la dong cuoi cung trong mang thi ko Insert them row nao.
                        If i <> totalRow Then
                            .ReDraw = False
                            .MaxRows = .MaxRows + 1
                            .InsertRows .Row + 1, 1
                            .CopyRowRange 22, 22, .Row + 1  ' Copy nguyen ban cac format cua dong bat dau la dong 22
                        End If
                    Next
                End If
                ' set tong so nguoi duoc mien thue b1
                .Row = totalRow + 24
                .Col = .ColLetterToNumber("C")
                .Text = soNguoiDuocMTB1
            End If
         End If
    ' ket thuc kiem tra
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

'Private Sub lamTron(lRow As Integer, lSheet As Integer)
'    Dim varThuNhap, varSoTienKH, varSoTienBH, varSoTienGiamThue, varSoThueKT, varThuNhapTT As Variant
'
'    With fps
'        .EventEnabled(EventAllEvents) = False
'        .AutoCalc = False
'        .Sheet = lSheet
'         .Row = lRow
'         If .Sheet = 2 Then
'
'            ' Clear ct
'            If .Row <= .MaxRows - 8 Then
'                .Col = .ColLetterToNumber("M")
'                .Formula = ""
'                .Col = .ColLetterToNumber("Q")
'                .Formula = ""
'            End If
'            ' End clear
'
'            ' Lam tron so cot thu nhap
'            .GetText .ColLetterToNumber("H"), .Row, varThuNhap
'            .Col = .ColLetterToNumber("H")
'            .Text = Round(Val(varThuNhap), 0)
'
'            ' Lam tron cot tu thien, khuyen hoc
'            .GetText .ColLetterToNumber("K"), .Row, varSoTienKH
'            .Col = .ColLetterToNumber("K")
'            .Text = Round(Val(varSoTienKH), 0)
'
'            'Lam tron cot bao hiem bat buoc
'            .GetText .ColLetterToNumber("L"), .Row, varSoTienBH
'            .Col = .ColLetterToNumber("L")
'            .Text = Round(Val(varSoTienBH), 0)
'
'            'Lam tron cot can cu giam thue
'            .GetText .ColLetterToNumber("I"), .Row, varSoTienGiamThue
'            .Col = .ColLetterToNumber("I")
'            .Text = Round(Val(varSoTienGiamThue), 0)
'
'            ' Lam tron cot so tien khau tru
'            .GetText .ColLetterToNumber("N"), .Row, varSoThueKT
'            .Col = .ColLetterToNumber("N")
'            .Text = Round(Val(varSoThueKT), 0)
'
'            ' Lam tron cot thu nhap tinh thue
'            .GetText .ColLetterToNumber("M"), .Row, varThuNhapTT
'            .Col = .ColLetterToNumber("M")
'            .Text = Round(Val(varThuNhapTT), 0)
'             End If
'             .AutoCalc = True
'            .EventEnabled(EventAllEvents) = True
'    End With
'
'End Sub


Private Sub CheckDynamicSheet2()
    Dim vErrMsgHoVaTen, vErrMsgMST, vErrMsgMST1, vErrMsgMST2, vErrMsgCMT, vErrMsgCMT1, vErrMsgMSTCMT As Variant
    Dim vErrMsgThuNhapGiamThue, vErrMsgSoThueDuocGiam, vErrMsgSoThueKhauTru As Variant
    
    Dim vErrMsgGiamThueTheoHD As Variant
    Dim vErrMsg1213 As Variant, vErrMsgCt1512 As Variant
    
    Dim vLastRow As Variant
    Dim varTemp As Variant
    Dim i, j As Integer
    Dim iFocusFlag As Integer
    Dim strFocusSheetName, strFocusCol, strFocusRow As String
    
    Dim vErrMsgMSTAB As Variant
    Dim flagTrungMSTAB As Integer, isheet As Integer
    
    '24/02/2012 dntai them bien sua lai phan check sheet2
    Dim iFlagHovaTen As Integer, iFlagLoiMST1 As Integer, iFlagMSTTrung As Integer, iFlagLoiMST As Integer
    Dim iFlagMSTCMND As Integer, iFlagCT12CT11 As Integer, iFlagCT13CT11 As Integer, iFlagLoiCMT1 As Integer, iFlagLoiCMT As Integer
    Dim CountRow1 As Long, rowDes As Long
    Dim strCheck As String
    
    flagTrungMSTAB = 0
    
    With fps
        .Sheet = .SheetCount
                
        .GetText .ColLetterToNumber("E"), 37, vErrMsgHoVaTen
        .GetText .ColLetterToNumber("E"), 38, vErrMsgMST
        .GetText .ColLetterToNumber("E"), 39, vErrMsgMST1
        .GetText .ColLetterToNumber("E"), 40, vErrMsgMST2
        .GetText .ColLetterToNumber("E"), 41, vErrMsgCMT
        .GetText .ColLetterToNumber("E"), 42, vErrMsgCMT1
        .GetText .ColLetterToNumber("E"), 43, vErrMsgMSTCMT
                
        .GetText .ColLetterToNumber("E"), 44, vErrMsgThuNhapGiamThue
        .GetText .ColLetterToNumber("E"), 45, vErrMsgSoThueKhauTru
        
        .GetText .ColLetterToNumber("E"), 47, vErrMsgMSTAB
        
        .GetText .ColLetterToNumber("E"), 69, vErrMsgGiamThueTheoHD
        .GetText .ColLetterToNumber("E"), 82, vErrMsg1213
        .GetText .ColLetterToNumber("E"), 70, vErrMsgCt1512
        
        
        
        .Sheet = 3 'Check tinh xac thuc cua cac thong tin tren bang ke 05A/TNCN
        isheet = .Sheet
        
         'dntai 24/02/2012 khai bao bien
         Dim vHoVaTen, vMST, vCMND, vChkCaNhanCC, vTongSo, vTNCNLamCCMG, vThueTNCNDaKT, vThueTNCNPhaiKT, vMSTCQCT As Variant
         Dim vGiamTruTheoHD As Variant
         Dim vThueGiamKKT As Variant
         Dim countRow As Integer
         
         'get MST cua CQCT
         .GetText .ColLetterToNumber("D"), 8, vMSTCQCT
         .Col = .ColLetterToNumber("B")
         countRow = 22
         .Row = countRow
         Do
         
            .GetText .ColLetterToNumber("C"), .Row, vHoVaTen
            .GetText .ColLetterToNumber("D"), .Row, vMST
            .GetText .ColLetterToNumber("E"), .Row, vCMND
            .GetText .ColLetterToNumber("F"), .Row, vChkCaNhanCC
            .GetText .ColLetterToNumber("G"), .Row, vTongSo
            .GetText .ColLetterToNumber("H"), .Row, vTNCNLamCCMG
            .GetText .ColLetterToNumber("I"), .Row, vGiamTruTheoHD
            .GetText .ColLetterToNumber("J"), .Row, vThueTNCNDaKT
            
            .GetText .ColLetterToNumber("K"), .Row, vThueGiamKKT
            If Trim(vHoVaTen) = vbNullString And Trim(vMST) = vbNullString And Trim(vCMND) = vbNullString And (Trim(vChkCaNhanCC) = "" Or Trim(vChkCaNhanCC) = "0") _
               And Trim(vTongSo) = "0" And Trim(vTNCNLamCCMG) = "0" And Trim(vThueTNCNDaKT) = "0" And Trim(vThueGiamKKT) = "0" Then
                
                'reset color and cellnote on sheet
               .Col = .ColLetterToNumber("C")
               .BackColor = mNonErrorColor
               .CellNote = ""
               .Col = .ColLetterToNumber("D")
               .BackColor = mNonErrorColor
               .CellNote = ""
               .Col = .ColLetterToNumber("E")
               .BackColor = mNonErrorColor
               .CellNote = ""
               .Col = .ColLetterToNumber("H")
               .BackColor = mNonErrorColor
               .CellNote = ""
               .Col = .ColLetterToNumber("I")
               .BackColor = mNonErrorColor
               .CellNote = ""
               .Col = .ColLetterToNumber("K")
               .BackColor = mNonErrorColor
               .CellNote = ""
            Else
                
                
                'check ho va ten khong duoc trong
                If Trim(vHoVaTen) = vbNullString Then
                    .Col = .ColLetterToNumber("C")
                    .BackColor = mErrorColor
                    .CellNote = vErrMsgHoVaTen
                    iFlagHovaTen = 1
                Else
                    .Col = .ColLetterToNumber("C")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If
                
                'check MST
                'set lai cell MST = ko loi
                .Col = .ColLetterToNumber("D")
                .CellNote = ""
                .BackColor = mNonErrorColor
                
                'check mst trung giua 2 bang ke 5A va 5B
                .GetText .ColLetterToNumber("AA"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    flagTrungMSTAB = 1
                    .Col = .ColLetterToNumber("D")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMSTAB
                    .BackColor = mErrorColor

                End If
                
                If Len(Trim(vMST)) > 0 Then
                    'check cau truc mst
                    If Len(Trim(vMST)) = 10 Then
                         strCheck = CheckTaxCode(Mid(Trim(vMST), 1, 1), Mid(Trim(vMST), 2, 1), Mid(Trim(vMST), 3, 1), Mid(Trim(vMST), 4, 1), Mid(Trim(vMST), 5, 1), Mid(Trim(vMST), 6, 1), Mid(Trim(vMST), 7, 1), Mid(Trim(vMST), 8, 1), Mid(Trim(vMST), 9, 1), Mid(Trim(vMST), 10, 1), Mid(Trim(vMST), 11, 1), Mid(Trim(vMST), 12, 1), Mid(Trim(vMST), 13, 1))
                         If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                            .Col = .ColLetterToNumber("D")
                            .CellNote = vErrMsgMST
                            .BackColor = mErrorColor
                            iFlagLoiMST = 1
                        Else
                            'check mst trung mstcqct
                            If Trim(vMST) = Trim(vMSTCQCT) Then
                                .Col = .ColLetterToNumber("D")
                                .CellNote = vErrMsgMST1
                                .BackColor = mErrorColor
                                iFlagLoiMST1 = 1
                            Else
                                 'check MST trung nhau
                                 
                                 CountRow1 = .SearchCol(.ColLetterToNumber("D"), -1, -1, vMST, SearchFlagsNone)
                                 rowDes = .SearchCol(.ColLetterToNumber("D"), CountRow1, -1, vMST, SearchFlagsNone)
                                If rowDes > -1 Then
                                    .Col = .ColLetterToNumber("D")
                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                    .CellNote = vErrMsgMST2       'static
                                    .BackColor = mErrorColor
                                    iFlagMSTTrung = 1
                                End If
                            End If
                         End If
                    Else
                            .Col = .ColLetterToNumber("D")
                            .CellNote = vErrMsgMST
                            .BackColor = mErrorColor
                            iFlagLoiMST = 1
                    End If
                End If
                
                ' check CMT
                'set cell CMT thanh k loi
                .Col = .ColLetterToNumber("E")
                .CellNote = ""
                .BackColor = mNonErrorColor
                
                'check khong duoc de trong dong thoi MST va CMTDN
                If Trim(vMST) = vbNullString And Trim(vCMND) = vbNullString Then
                    .Col = .ColLetterToNumber("D")
                    .CellNote = vErrMsgMSTCMT
                    .BackColor = mErrorColor
                    .Col = .ColLetterToNumber("E")
                    .CellNote = vErrMsgMSTCMT
                    .BackColor = mErrorColor
                    iFlagMSTCMND = 1
                End If
                
                'check cmt khong duoc nhap dau cach
                If Trim(vCMND) <> vbNullString Then
                    If InStr(1, vCMND, " ", vbTextCompare) <> 0 Then
                        .Col = .ColLetterToNumber("E")
                        .CellNote = vErrMsgCMT1
                        .BackColor = mErrorColor
                        iFlagLoiCMT1 = 1
                    Else
                        'check CMTND trung nhau
                         CountRow1 = .SearchCol(.ColLetterToNumber("E"), -1, -1, vCMND, SearchFlagsNone)
                         rowDes = .SearchCol(.ColLetterToNumber("E"), CountRow1, -1, vCMND, SearchFlagsNone)
                        If rowDes > -1 Then
                            .Col = .ColLetterToNumber("E")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgCMT        'static
                            .BackColor = mAlertColor
                            'iFlagLoiCMT = 1
                        End If
                    End If
                 End If
                 
                ' THu nhap lam can cu giam thue <= thu nhap
                .GetText .ColLetterToNumber("V"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("H")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgThuNhapGiamThue
                    .BackColor = mErrorColor
                    iFlagCT12CT11 = 1
                Else
                    .Col = .ColLetterToNumber("H")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                ' theo hiep dinh
                .GetText .ColLetterToNumber("AO"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("I")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgGiamThueTheoHD
                    .BackColor = mErrorColor
                Else
                    .Col = .ColLetterToNumber("I")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                'Thue da khau tru >= thu nhap
                .GetText .ColLetterToNumber("X"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("J")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgSoThueKhauTru
'                    .BackColor = mErrorColor
'                    iFlagCT13CT11 = 1
                     .BackColor = mErrorColor
                Else
                    .Col = .ColLetterToNumber("J")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                ' kiem tra tong chi tieu 12,13 <=11
                .GetText .ColLetterToNumber("AQ"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("H")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsg1213
                     .BackColor = mErrorColor
                     
                     .Col = .ColLetterToNumber("I")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsg1213
                     .BackColor = mErrorColor
                Else
                    .Col = .ColLetterToNumber("H")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                    
                    .Col = .ColLetterToNumber("I")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                 'check chi tieu 15
                .GetText .ColLetterToNumber("AP"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("K")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgCt1512
                    .BackColor = mAlertColor
                Else
                    .Col = .ColLetterToNumber("K")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            End If
            
            countRow = countRow + 1
            .Row = countRow
            .Col = .ColLetterToNumber("B")

         Loop Until .Text = "aa"
                                
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 37, IIf(iFlagHovaTen = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 38, IIf(iFlagLoiMST = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 39, IIf(iFlagLoiMST1 = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 40, IIf(iFlagMSTTrung = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 47, IIf(flagTrungMSTAB = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 43, IIf(iFlagMSTCMND = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 42, IIf(iFlagLoiCMT1 = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 41, IIf(iFlagLoiCMT = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 44, IIf(iFlagCT12CT11 = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 45, IIf(iFlagCT13CT11 = 1, "0", "1")
        
        .Sheet = isheet
    End With
End Sub


Private Sub CheckDynamicSheet3()
    Dim vHoVaTenNNT, vMSTNNT, vHoVaTenNPT, vNgaySinhNPT, vMSTNPT, vQuocTich, vSoCMNDNPT, vQuanHeNPT, vSo, vQuyenSo, vQuocGia As Variant
    Dim vTinh, vHuyen, vXa, vTuThang, vDenThang As Variant
    Dim vMaQuocGia As Variant, vMaTinh As Variant, vMaHuyen As Variant, vMaXa As Variant
    
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
    
    Dim vErrHuyen, vErrXa As Variant
    
    Dim vErrMsgMSTTrungNPT As Variant
    Dim CountRow1 As Long, rowDes As Long
    
    Dim tam As Long
    Dim flagTrungMSTAB As Integer
    Dim vErrThuoc05_1_2 As Variant
    Dim varTemp As Variant
    Dim countRowCheck As Long
    
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
        .GetText .ColLetterToNumber("E"), 86, vErrNgaySinhNPT1
        .GetText .ColLetterToNumber("E"), 87, vErrHuyen
        .GetText .ColLetterToNumber("E"), 88, vErrXa
        .GetText .ColLetterToNumber("E"), 89, vErrMsgMSTTrungNPT
        .GetText .ColLetterToNumber("E"), 90, vErrThuoc05_1_2
        
        
        .Sheet = 4
        .Row = 22
        Do
            .GetText .ColLetterToNumber("D"), .Row, vHoVaTenNNT
            .GetText .ColLetterToNumber("E"), .Row, vMSTNNT
            .GetText .ColLetterToNumber("F"), .Row, vHoVaTenNPT
            .GetText .ColLetterToNumber("G"), .Row, vNgaySinhNPT
            .GetText .ColLetterToNumber("G"), .Row, vNgaySinhNPT_Goc
            .GetText .ColLetterToNumber("H"), .Row, vMSTNPT
            .GetText .ColLetterToNumber("J"), .Row, vQuocTich
            .GetText .ColLetterToNumber("K"), .Row, vSoCMNDNPT
            .GetText .ColLetterToNumber("M"), .Row, vQuanHeNPT
            .GetText .ColLetterToNumber("N"), .Row, vSo
            .GetText .ColLetterToNumber("O"), .Row, vQuyenSo
            .GetText .ColLetterToNumber("Q"), .Row, vQuocGia
            .GetText .ColLetterToNumber("S"), .Row, vTinh
            .GetText .ColLetterToNumber("U"), .Row, vHuyen
            .GetText .ColLetterToNumber("W"), .Row, vXa
            .GetText .ColLetterToNumber("X"), .Row, vTuThang
            .GetText .ColLetterToNumber("Y"), .Row, vDenThang
            
            .GetText .ColLetterToNumber("P"), .Row, vMaQuocGia
            .GetText .ColLetterToNumber("R"), .Row, vMaTinh
            .GetText .ColLetterToNumber("T"), .Row, vMaHuyen
            .GetText .ColLetterToNumber("V"), .Row, vMaXa
            
            ' reset du lieu loi
            strColReset = "D~E~F~G~H~J~K~M~N~O~Q~S~U~W~X~Y"
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
'                'reset color and cellnote on sheet
'                strColReset = "D~E~F~G~H~J~K~M~N~O~Q~S~U~W~X~Y"
'                arrColReset = Split(strColReset, "~")
'                For i = 0 To UBound(arrColReset)
'                    .Col = .ColLetterToNumber(arrColReset(0))
'                    .BackColor = mNonErrorColor
'                    .CellNote = ""
'                Next
            Else
                ' check ho ten
                If Trim(vHoVaTenNNT) = vbNullString Then
                    .Col = .ColLetterToNumber("D")
                    .BackColor = mErrorColor
                    .CellNote = vErrHoVaTenNNT
                    iFlagHoVaTenNNT = 1
                Else
                    .Col = .ColLetterToNumber("D")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If
                
                ' check MST
                'check cau truc mst
                vMSTNNT = Replace$(vMSTNNT, "-", "")
                If Len(Trim(vMSTNNT)) = 10 Then
                     strCheck = CheckTaxCode(Mid(Trim(vMSTNNT), 1, 1), Mid(Trim(vMSTNNT), 2, 1), Mid(Trim(vMSTNNT), 3, 1), Mid(Trim(vMSTNNT), 4, 1), Mid(Trim(vMSTNNT), 5, 1), Mid(Trim(vMSTNNT), 6, 1), Mid(Trim(vMSTNNT), 7, 1), Mid(Trim(vMSTNNT), 8, 1), Mid(Trim(vMSTNNT), 9, 1), Mid(Trim(vMSTNNT), 10, 1), Mid(Trim(vMSTNNT), 11, 1), Mid(Trim(vMSTNNT), 12, 1), Mid(Trim(vMSTNNT), 13, 1))
                     If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMSTNNT) Then
                        .Col = .ColLetterToNumber("E")
                        .CellNote = vErrMSTNNT
                        .BackColor = mErrorColor
                        iFlagMSTNNT = 1
                    Else
                        'check MST trung MST NPT
                        rowDes = .SearchCol(.ColLetterToNumber("H"), -1, -1, vMSTNNT, SearchFlagsNone)
                        If rowDes > -1 Then
                            .Col = .ColLetterToNumber("E")
                            .CellNote = vErrMsgMSTTrungNPT       'static
                            .BackColor = mAlertColor
                        Else
                             .Col = .ColLetterToNumber("E")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                     End If
                Else
                        .Col = .ColLetterToNumber("E")
                        .CellNote = vErrMSTNNT
                        .BackColor = mErrorColor
                        iFlagMSTNNT = 1
                End If
                
                ' check ho ten NPT
                If Trim(vHoVaTenNPT) = vbNullString Then
                    .Col = .ColLetterToNumber("F")
                    .BackColor = mErrorColor
                    .CellNote = vErrHoVaTenNPT
                    iFlagHoVaTenNPT = 1
                Else
                    .Col = .ColLetterToNumber("F")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If
                
                ' kiem tra ngay sinh khong duoc lon hon ngay hien tai
                If Trim(vNgaySinhNPT) <> "" Then
                    If Format_ddmmyyyy1(CStr(vNgaySinhNPT_Goc)) <> "" Then
                        'If DateDiff("D", ToDate(CStr(vNgaySinhNPT), "dd/mm/yyyy"), Date) > 0 Then
                        vNgaySinhNPT = DateSerial(CInt(Mid$(vNgaySinhNPT, 7, 4)), CInt(Mid$(vNgaySinhNPT, 4, 2)), CInt(Mid$(vNgaySinhNPT, 1, 2)))
                        If vNgaySinhNPT > Date Then
                            .Col = .ColLetterToNumber("G")
                            .BackColor = mErrorColor
                            .CellNote = vErrNgaySinhNPT
                            iFlagNgaySinhNPT = 1
                        Else
                            If Trim(vTuThang) <> "" Then
                                If DateDiff("M", vNgaySinhNPT, Format(vTuThang, "mm/yyyy")) < 0 Then
                                    .Col = .ColLetterToNumber("G")
                                    .BackColor = mErrorColor
                                    .CellNote = vErrNgaySinhNPT1
                                    iFlagNgaySinhNPT = 1
                                Else
                                    .Col = .ColLetterToNumber("G")
                                    .BackColor = mNonErrorColor
                                    .CellNote = ""
                                End If
                            Else
                                .Col = .ColLetterToNumber("G")
                                .BackColor = mNonErrorColor
                                .CellNote = ""
                            End If
                        End If
                    Else
                        .Col = .ColLetterToNumber("G")
                        .BackColor = mErrorColor
                        .CellNote = vErrDuLieuKhongHL
                        iFlagNgaySinhNPT = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("G")
                    .BackColor = mErrorColor
                    .CellNote = vErrEmpty
                    iFlagNgaySinhNPT = 1
                End If
                
                ' Kiem tra MST NPT
                If Len(vMSTNPT) > 0 Then
                    vMSTNPT = Replace$(vMSTNPT, "-", "")
                    If Len(Trim(vMSTNPT)) = 10 Then
                         strCheck = CheckTaxCode(Mid(Trim(vMSTNPT), 1, 1), Mid(Trim(vMSTNPT), 2, 1), Mid(Trim(vMSTNPT), 3, 1), Mid(Trim(vMSTNPT), 4, 1), Mid(Trim(vMSTNPT), 5, 1), Mid(Trim(vMSTNPT), 6, 1), Mid(Trim(vMSTNPT), 7, 1), Mid(Trim(vMSTNPT), 8, 1), Mid(Trim(vMSTNPT), 9, 1), Mid(Trim(vMSTNPT), 10, 1), Mid(Trim(vMSTNPT), 11, 1), Mid(Trim(vMSTNPT), 12, 1), Mid(Trim(vMSTNPT), 13, 1))
                         If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMSTNPT) Then
                            .Col = .ColLetterToNumber("H")
                            .CellNote = vErrMSTNPT
                            .BackColor = mErrorColor
                            iFlagMSTNPT = 1
                        Else
                            'check MST trung MST NPT
                            rowDes = .SearchCol(.ColLetterToNumber("E"), -1, -1, vMSTNPT, SearchFlagsNone)
                            If rowDes > -1 Then
                                .Col = .ColLetterToNumber("H")
                                .CellNote = vErrMsgMSTTrungNPT       'static
                                .BackColor = mAlertColor
                             Else
                                .Col = .ColLetterToNumber("H")
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            End If
                         End If
                    Else
                            .Col = .ColLetterToNumber("H")
                            .CellNote = vErrMSTNPT
                            .BackColor = mErrorColor
                            iFlagMSTNPT = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("H")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                ' check CMND
                If Trim(vSoCMNDNPT) <> vbNullString Then
                    If InStr(1, vSoCMNDNPT, " ", vbTextCompare) <> 0 Then
                        .Col = .ColLetterToNumber("K")
                        .CellNote = vErrSoCMNDNPT
                        .BackColor = mErrorColor
                        iFlagSoCMNDNPT = 1
                    Else
                        .Col = .ColLetterToNumber("K")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                Else
                     ' kiem tra neu tuoi >15 bat buoc nhap CMT
                    If vNgaySinhNPT <> "" Then
                        If Format_ddmmyyyy1(CStr(vNgaySinhNPT_Goc)) <> "" Then
                            If Year(Date) - Year(Format(vNgaySinhNPT, "dd/mm/yyyy")) > 15 And Trim(vSoCMNDNPT) = "" Then
                                 .Col = .ColLetterToNumber("K")
                                .CellNote = vErrEmptyCMNDNPT
                                .BackColor = mErrorColor
                                iFlagEmpty = 1
                            End If
                         Else
                            .Col = .ColLetterToNumber("G")
                            .CellNote = vErrDuLieuKhongHL
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                         End If
                    End If
                End If
                 
                 ' Kiem tra tu thang bat buoc nhap
                 If Trim(vTuThang) <> "" Then
                    .Col = .ColLetterToNumber("X")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                 Else
                    .Col = .ColLetterToNumber("X")
                    .CellNote = vErrEmpty
                    .BackColor = mErrorColor
                    iFlagEmpty = 1
                 End If
                 
                 ' kiem tra den thang bat buoc nhap
'                 If Trim(vDenThang) <> "" Then
'                    .Col = .ColLetterToNumber("Y")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                 Else
'                    .Col = .ColLetterToNumber("Y")
'                    .CellNote = vErrEmpty
'                    .BackColor = mErrorColor
'                    iFlagEmpty = 1
'                 End If
                 
                ' Kiem tra tu thang, den thang
                If Trim(vTuThang) <> "" And Trim(vDenThang) <> "" Then
                    If DateDiff("M", Format(vTuThang, "mm/yyyy"), Format(vDenThang, "mm/yyyy")) < 0 Then
                            .Col = .ColLetterToNumber("X")
                            .CellNote = vErrTuThang
                            .BackColor = mErrorColor
                            
                            .Col = .ColLetterToNumber("Y")
                            .CellNote = vErrTuThang
                            .BackColor = mErrorColor
                            iFlagTuThang = 1
                    Else
                            .Col = .ColLetterToNumber("X")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                            
'                            .Col = .ColLetterToNumber("Y")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
                            ' kiem tra den thang voi den thang ky QT
'                            If vDenThang <> "" Then
'                                If DateDiff("M", Format(vDenThang, "mm/yyyy"), Format(TAX_Utilities_v1.LastDay, "mm/yyyy")) < 0 Then
'                                     .Col = .ColLetterToNumber("Y")
'                                    .CellNote = vErrDenThang
'                                    .BackColor = mErrorColor
'                                    iFlagDenThang = 1
'                                Else
                            .Col = .ColLetterToNumber("Y")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
'                                End If
'                            End If
                    End If
                 End If
                
                
                ' kiem tra bat buoc nhap cot 14
                If Trim(vQuanHeNPT) = "" Then
                    .Col = .ColLetterToNumber("M")
                    .CellNote = vErrEmpty
                    .BackColor = mErrorColor
                    iFlagEmpty = 1
                End If
                
                ' kiem tra khong co cot 11 va 13 bat buoc nhap cot 18,19,20
                If Trim(vMSTNPT) = "" And Trim(vSoCMNDNPT) = "" Then
                    ' bat buoc nhap thong tin quoc gia
                    If Trim(vQuocGia) = "" Then
                        .Col = .ColLetterToNumber("Q")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    ' bat buoc nhap thong tin tinh
                    If Trim(vTinh) = "" Then
                        .Col = .ColLetterToNumber("S")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    ' bat buoc nhap thong tin huyen
                    If Trim(vHuyen) = "" Then
                        .Col = .ColLetterToNumber("U")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    ' bat buoc nhap thong tin xa
                    If Trim(vXa) = "" Then
                        .Col = .ColLetterToNumber("W")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    ' kiem tra neu chon quoc gia Viet Nam
                    If Trim(vMaQuocGia) = "01" Then
                        If Trim$(vMaTinh) = "999" Then
                            .Col = .ColLetterToNumber("S")
                            .CellNote = vErrDuLieuKhongHL
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                        End If
                        
                        If Trim$(vMaHuyen) = "99999" Then
                            .Col = .ColLetterToNumber("U")
                            .CellNote = vErrDuLieuKhongHL
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                        End If
                        
                        If Trim$(vMaXa) = "9999999" Then
                            .Col = .ColLetterToNumber("W")
                            .CellNote = vErrDuLieuKhongHL
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                        End If
                        
                        ' bat buoc nhap thong tin tinh
                        If Trim(vTinh) = "" Then
                            .Col = .ColLetterToNumber("S")
                            .CellNote = vErrEmpty
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                        End If
                        
                        ' bat buoc nhap thong tin huyen
                        If Trim(vHuyen) = "" Then
                            .Col = .ColLetterToNumber("U")
                            .CellNote = vErrEmpty
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                        ElseIf Trim(vHuyen) <> "" And Trim(vTinh) <> "" And Trim(vMaTinh) <> Left(Trim$(vMaHuyen), 3) Then
                            .Col = .ColLetterToNumber("U")
                            .CellNote = vErrHuyen
                            .BackColor = mAlertColor
                        End If
                        
                        ' bat buoc nhap thong tin xa
                        If Trim(vXa) = "" Then
                            .Col = .ColLetterToNumber("W")
                            .CellNote = vErrEmpty
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                        ElseIf Trim(vHuyen) <> "" And Trim(vXa) <> "" And Trim(vMaHuyen) <> Left(Trim$(vMaXa), 5) Then
                            .Col = .ColLetterToNumber("W")
                            .CellNote = vErrXa
                            .BackColor = mAlertColor
                        End If
                    End If
                End If
                
                ' Kiem tra truong hop co nhap thong tin tinh, huyen , xa thi bat buoc nhap thong tin quoc gia
                If Trim$(vTinh) <> "" Or Trim$(vHuyen) <> "" Or Trim$(vXa) <> "" Then
                    ' bat buoc nhap thong tin quoc gia
                    If Trim(vQuocGia) = "" Then
                        .Col = .ColLetterToNumber("Q")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    ' bat buoc nhap thong tin tinh
                    If Trim(vTinh) = "" Then
                        .Col = .ColLetterToNumber("S")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    ' bat buoc nhap thong tin huyen
                    If Trim(vHuyen) = "" Then
                        .Col = .ColLetterToNumber("U")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                    ' bat buoc nhap thong tin xa
                    If Trim(vXa) = "" Then
                        .Col = .ColLetterToNumber("W")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    End If
                    
                End If
                
                
                ' bat buoc nhap thong so neu <= 15 t
                If Trim(vMSTNPT) = "" And Trim(vSoCMNDNPT) = "" Then
                    If Trim(vSo) = "" Then
                        If Format_ddmmyyyy1(CStr(vNgaySinhNPT_Goc)) <> "" Then
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
                
                ' bat buoc nhap thong quyen so neu <= 15 t
                If Trim(vMSTNPT) = "" And Trim(vSoCMNDNPT) = "" Then
                    If Trim(vQuyenSo) = "" Then
                        If Format_ddmmyyyy1(CStr(vNgaySinhNPT_Goc)) <> "" Then
                            If Year(Date) - Year(Format(vNgaySinhNPT, "dd/mm/yyyy")) <= 15 Then
                                .Col = .ColLetterToNumber("O")
                                .CellNote = vErrEmpty
                                .BackColor = mErrorColor
                                iFlagEmpty = 1
                            End If
                        End If
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
                ' kiem tra MST thuoc 05A, 05B
                tam = .Row
                .Sheet = 3 'Bang ke 5B
                .Row = 22
                flagTrungMSTAB = 0
                Do
                    ' Ma la null
                        'dhdang check trung MST
                        .GetText .ColLetterToNumber("D"), .Row, varTemp
                        If Trim(varTemp) = Trim(vMSTNNT) Then
                            flagTrungMSTAB = 1
                        End If
        
                    .Row = .Row + 1
                    .Col = .ColLetterToNumber("B")
                Loop Until .Text = "aa"
                
                If flagTrungMSTAB = 0 Then
                    .Sheet = 2 'Bang ke 5A
                    .Row = 22
                    Do
                        ' Ma la null
                            'dhdang check trung MST
                            .GetText .ColLetterToNumber("E"), .Row, varTemp
                            If Trim(varTemp) = Trim(vMSTNNT) Then
                                flagTrungMSTAB = 1
                            End If
                        .Row = .Row + 1
                        .Col = .ColLetterToNumber("B")
                    Loop Until .Text = "aa"
                End If
                
                .Sheet = 4
                .Row = tam
                'neu mst khong thuoc 05B, 05A thi canh bao do
                If flagTrungMSTAB = 0 Then
                    .Col = .ColLetterToNumber("E")
                    .BackColor = mErrorColor
                    .CellNote = vErrThuoc05_1_2
                    countRowCheck = countRowCheck + 1
                End If

            End If
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
        
        If countRowCheck > 0 Then
            flagTrungMSTAB = 0
        Else
            flagTrungMSTAB = 1
        End If
        
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
        .SetText .ColLetterToNumber("B"), 90, IIf(flagTrungMSTAB = 1, "1", "0")
    End With
End Sub



Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
            
    With fps
        'Because of unclearly-rised events, this function is here for turning it off
        'at the end of this sub, this event will be turned on
        .EventEnabled(EventAllEvents) = False
        .Sheet = .ActiveSheet
        mCurrentSheet = .ActiveSheet
        If .ActiveSheet = 2 Then
                mCurrentSheet = .Sheet
            ' chi tieu 21
            If Col = .ColLetterToNumber("D") And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 36
                .Formula = "Q36"
            End If
            
            ' chi tieu 24
            If Col = .ColLetterToNumber("P") And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 39
                .Formula = "Q39"
            End If
            
            ' chi tieu 26
            If Col = .ColLetterToNumber("J") And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 41
                .Formula = "Q41"
            End If
            
            ' chi tieu 28
            If Col = .ColLetterToNumber("H") And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 43
                .Formula = "Q43"
            End If
            
            ' chi tieu 36
            If Col = .ColLetterToNumber("P") And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 51
                .Formula = "Q51"
            End If
            
            
            ' chi tieu 32
            If (Col = .ColLetterToNumber("P") Or Col = .ColLetterToNumber("H")) And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 47
                .Formula = "Q47"
            End If
            
            
            ' chi tieu 40
            If (Col = .ColLetterToNumber("Q") Or Col = .ColLetterToNumber("I") Or Col = .ColLetterToNumber("J") Or Col = .ColLetterToNumber("H") Or Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("L") Or Col = .ColLetterToNumber("M") Or Col = .ColLetterToNumber("N")) _
            And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 55
                .Formula = "Q55"
            End If
            
            
            ' chi tieu 44
            If (Col = .ColLetterToNumber("D") Or Col = .ColLetterToNumber("G") Or Col = .ColLetterToNumber("P")) And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 62
                .Formula = "Q62"
            End If
        End If
        
        
        If .ActiveSheet = 3 Then
                mCurrentSheet = .Sheet
            ' chi tieu 21
            If Col = .ColLetterToNumber("C") And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 36
                .Formula = "Q36"
            End If
            
            ' chi tieu 24
            If (Col = .ColLetterToNumber("J") Or Col = .ColLetterToNumber("F")) And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 39
                .Formula = "Q39"
            End If
            
            
            ' chi tieu 25
            If (Col = .ColLetterToNumber("J") Or Col = .ColLetterToNumber("F")) And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 40
                .Formula = "Q40"
            End If
            
            ' chi tieu 26
            If Col = .ColLetterToNumber("I") And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 41
                .Formula = "Q41"
            End If
            
           
            
            ' chi tieu 29
            If (Col = .ColLetterToNumber("C") Or Col = .ColLetterToNumber("F") Or Col = .ColLetterToNumber("G")) And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 44
                .Formula = "Q44"
            End If
            
            
            ' chi tieu 30
            If (Col = .ColLetterToNumber("C") Or Col = .ColLetterToNumber("F") Or Col = .ColLetterToNumber("G")) And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 45
                .Formula = "Q45"
            End If
            
           
            
            ' chi tieu 33
            If (Col = .ColLetterToNumber("C") Or Col = .ColLetterToNumber("F") Or Col = .ColLetterToNumber("G") Or Col = .ColLetterToNumber("J")) And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 48
                .Formula = "Q48"
            End If
            
            ' chi tieu 34
            If (Col = .ColLetterToNumber("C") Or Col = .ColLetterToNumber("F") Or Col = .ColLetterToNumber("G") Or Col = .ColLetterToNumber("J")) And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 49
                .Formula = "Q49"
            End If
            
            ' chi tieu 37
            If (Col = .ColLetterToNumber("C") Or Col = .ColLetterToNumber("F") Or Col = .ColLetterToNumber("J")) And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 52
                .Formula = "Q52"
            End If
            ' chi tieu 38
            If (Col = .ColLetterToNumber("C") Or Col = .ColLetterToNumber("F") Or Col = .ColLetterToNumber("J")) And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 53
                .Formula = "Q53"
            End If
            
            ' chi tieu 42
            If (Col = .ColLetterToNumber("C") Or Col = .ColLetterToNumber("F") Or Col = .ColLetterToNumber("K")) And Row >= 22 Then
                .Sheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 57
                .Formula = "Q57"
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub




Sub setFormulaImport()
            
    With fps
        'Because of unclearly-rised events, this function is here for turning it off
        'at the end of this sub, this event will be turned on
        .EventEnabled(EventAllEvents) = False
        .Sheet = .ActiveSheet
        mCurrentSheet = .ActiveSheet
        If .ActiveSheet = 2 Then
                mCurrentSheet = .Sheet
            ' chi tieu 21
            .Sheet = 1
            .Col = .ColLetterToNumber("I")
            .Row = 36
            .Formula = "Q36"
            
            ' chi tieu 24
            .Col = .ColLetterToNumber("I")
            .Row = 39
            .Formula = "Q39"
            
            ' chi tieu 26
            .Col = .ColLetterToNumber("I")
            .Row = 41
            .Formula = "Q41"
            
            ' chi tieu 28
            .Col = .ColLetterToNumber("I")
            .Row = 43
            .Formula = "Q43"
            
            ' chi tieu 36
            .Col = .ColLetterToNumber("I")
            .Row = 51
            .Formula = "Q51"
            
            ' chi tieu 32
            .Col = .ColLetterToNumber("I")
            .Row = 47
            .Formula = "Q47"
            
            ' chi tieu 40
            .Col = .ColLetterToNumber("I")
            .Row = 55
            .Formula = "Q55"
            
            
            ' chi tieu 44
            .Col = .ColLetterToNumber("I")
            .Row = 62
            .Formula = "Q62"
        End If
        
        
        If .ActiveSheet = 3 Then
                mCurrentSheet = .Sheet
            ' chi tieu 21
            .Sheet = 1
            .Col = .ColLetterToNumber("I")
            .Row = 36
            .Formula = "Q36"
            
            ' chi tieu 24
            .Col = .ColLetterToNumber("I")
            .Row = 39
            .Formula = "Q39"
            
            
            ' chi tieu 25
            .Col = .ColLetterToNumber("I")
            .Row = 40
            .Formula = "Q40"
            
            ' chi tieu 26
            .Col = .ColLetterToNumber("I")
            .Row = 41
            .Formula = "Q41"
            
            ' chi tieu 29
            .Col = .ColLetterToNumber("I")
            .Row = 44
            .Formula = "Q44"
            
            ' chi tieu 30
            .Col = .ColLetterToNumber("I")
            .Row = 45
            .Formula = "Q45"
            
            ' chi tieu 33
            .Col = .ColLetterToNumber("I")
            .Row = 48
            .Formula = "Q48"

            
            ' chi tieu 34
            .Col = .ColLetterToNumber("I")
            .Row = 49
            .Formula = "Q49"
                
            ' chi tieu 37
            .Col = .ColLetterToNumber("I")
            .Row = 52
            .Formula = "Q52"

            ' chi tieu 38
            .Col = .ColLetterToNumber("I")
            .Row = 53
            .Formula = "Q53"

            
            ' chi tieu 42
            .Col = .ColLetterToNumber("I")
            .Row = 57
            .Formula = "Q57"
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
