VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_09TNCN"
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
Public flgLoadToKhai As Boolean

' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

Private arrTuThang() As String
Private arrDenThang() As String
Private soThang As Variant
Private soThang1 As Long, soThang2 As Long


Private Const phuLuc09_1_TNCN = 2
Private Const phuLuc09_2_TNCN = 3
Private Const phuLuc09_3_TNCN = 4
'Private Const phuLuc09_4_TNCN = 5
Private Const phuLuc09_MT_TNCN = 5

  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    Dim tuThang, denThang As Variant
    Dim varTuThang, varDenThang As Variant
    
    With fps
        .Sheet = 1 'To khai 09/TNCN
        ' Set ky cho thang bat dau tinh thue
        
        SetDateFormat fps, 1, 4, .ColLetterToNumber("N"), MMYYYY
        SetDateFormat fps, 1, 4, .ColLetterToNumber("P"), MMYYYY
            
        .Row = 4
        .Col = .ColLetterToNumber("N")
        .Text = Format(TAX_Utilities_v1.FirstDay, "mm/yyyy")
        .TypeHAlign = TypeHAlignLeft

        .Row = 4
        .Col = .ColLetterToNumber("P")
        .Text = Format(TAX_Utilities_v1.LastDay, "mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        ' Tinh tong so thang theo dinh dang tu ky den ky
        .GetText .ColLetterToNumber("N"), 4, tuThang
        .GetText .ColLetterToNumber("P"), 4, denThang
        
        'set so thang quyet toan
        .Row = 4
        .Col = .ColLetterToNumber("R")
        .Text = DateDiff("M", Format_mmyyyy(CStr(tuThang)), Format_mmyyyy(CStr(denThang))) + 1
        
        .GetText .ColLetterToNumber("R"), 4, soThang

        ' set cong thuc cho chi tieu 28
        arrTuThang = Split(tuThang, "/")
        arrDenThang = Split(denThang, "/")
        
        ' neu den thang < 7/2013
        If Val(arrDenThang(1)) < 2013 Or (Val(arrDenThang(1)) = 2013 And Val(arrDenThang(0)) < 7) Then
            .Row = 48
            .Col = .ColLetterToNumber("K")
            .value = 4000000 * Val(soThang)
        ElseIf (Val(arrTuThang(1)) > 2013) Or ((Val(arrTuThang(1)) = 2013) And (Val(arrTuThang(0)) >= 7)) Then
            .Row = 48
            .Col = .ColLetterToNumber("K")
            .value = 9000000 * Val(soThang)
        Else
            soThang1 = DateDiff("M", Format_mmyyyy(CStr(tuThang)), Format_mmyyyy("06/2013")) + 1
            soThang2 = DateDiff("M", Format_mmyyyy("07/2013"), Format_mmyyyy(CStr(denThang))) + 1
            .Row = 48
            .Col = .ColLetterToNumber("K")
            .value = 4000000 * soThang1 + 9000000 * soThang2
        End If
                
        ' Set format va gia tri cho ngay ky
        SetDateFormat fps, 1, 71, .ColLetterToNumber("O"), DDMMYYYY
                
        .Row = 71
        .Col = .ColLetterToNumber("O")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
                       
        ' Dinh dang ngay thang cho cot ngay sinh trong phu luc 09-3/TNCN
        SetDateFormat fps, phuLuc09_3_TNCN, 25, .ColLetterToNumber("E"), DDMMYYYY
        SetDateFormat fps, phuLuc09_3_TNCN, 33, .ColLetterToNumber("E"), DDMMYYYY
              
        SetDateFormat fps, phuLuc09_3_TNCN, 25, .ColLetterToNumber("Q"), MMYYYY
        SetDateFormat fps, phuLuc09_3_TNCN, 25, .ColLetterToNumber("S"), MMYYYY
              
        SetDateFormat fps, phuLuc09_3_TNCN, 33, .ColLetterToNumber("T"), MMYYYY
        SetDateFormat fps, phuLuc09_3_TNCN, 33, .ColLetterToNumber("U"), MMYYYY
              
        .Sheet = phuLuc09_3_TNCN
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
        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListMap
            arrValue = Split(GetAttribute(xmlNode, "value"), "###")
            strQuanHeNNT = strQuanHeNNT + arrValue(1) + Chr$(9)
            strQuanHeNNTID = strQuanHeNNTID + arrValue(0) + Chr$(9)
        Next
        
        .Row = 25
        .Col = .ColLetterToNumber("O")
        .TypeComboBoxList = strQuanHeNNT
        
        .Col = .ColLetterToNumber("N")
        .TypeComboBoxList = strQuanHeNNTID
        
        .Row = 33
        .Col = .ColLetterToNumber("S")
        .TypeComboBoxList = strQuanHeNNT
        
        .Col = .ColLetterToNumber("R")
        .TypeComboBoxList = strQuanHeNNTID
        
        
        ' lay du lieu danh muc dia ban tinh
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_Dia_Ban_Tinh.xml"))
        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListMap
            temValue = GetAttribute(xmlNode, "Value")
            arrValue = Split(temValue, "###")
            strTenTinh = strTenTinh + arrValue(2) + Chr$(9)
            strMaTinh = strMaTinh + arrValue(1) + Chr$(9)
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
            strTenHuyen = strTenHuyen + arrValue(3) + Chr$(9)
            strMaHuyen = strMaHuyen + arrValue(2) + Chr$(9)
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
            strTenXa = strTenXa + arrValue(4) + Chr$(9)
            strMaXa = strMaXa + arrValue(3) + Chr$(9)
        Next
        
        
        .Row = 33
        ' set danh sach tinh
        .Col = .ColLetterToNumber("K")
        .TypeComboBoxList = strTenTinh
        .Col = .ColLetterToNumber("J")
        .TypeComboBoxList = strMaTinh
        ' set danh sach huyen
        .Col = .ColLetterToNumber("M")
        .TypeComboBoxList = strTenHuyen
        .Col = .ColLetterToNumber("L")
        .TypeComboBoxList = strMaHuyen
        ' set danh sach xa
        .Col = .ColLetterToNumber("O")
        .TypeComboBoxList = strTenXa
        .Col = .ColLetterToNumber("N")
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
        
        .Row = 25
        ' set danh sach quoc tich
        .Col = .ColLetterToNumber("K")
        .TypeComboBoxList = strQuocTich
        .typeComboboxCurSel = 0
        .Col = .ColLetterToNumber("J")
        .TypeComboBoxList = strMaQuocTich
        .typeComboboxCurSel = 0
        
        .Row = 33
        ' set danh sach quoc gia
        .Col = .ColLetterToNumber("I")
        .TypeComboBoxList = strQuocTich
        .typeComboboxCurSel = 0
        .Col = .ColLetterToNumber("H")
        .TypeComboBoxList = strMaQuocTich
        .typeComboboxCurSel = 0
        
        ' set danh sach quoc tich
        .Col = .ColLetterToNumber("Q")
        .TypeComboBoxList = strQuocTich
        .typeComboboxCurSel = 0
        .Col = .ColLetterToNumber("P")
        .TypeComboBoxList = strMaQuocTich
        .typeComboboxCurSel = 0
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim tuThang, denThang As Variant
    'bien lay gia tri chi tieu 29,30 lay luy tien cua chi tieu 31
    Dim lT31, oldValue As Variant
    Dim xmlNode   As MSXML.IXMLDOMNode
    Dim iCol      As Long, iRow As Long
    Dim varCT15   As Variant, varCT31 As Variant
    Dim soThangKG As Integer
    Dim viTriMSTDL As String
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        
        If Strloaitk = "TKCT" And Trim(StrSolanBosung) = "" Then
            .Col = .ColLetterToNumber("E")
            .Row = 76
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("K")
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKCT" And Val(StrSolanBosung) > 0 Then
            .Col = .ColLetterToNumber("E")
            .Row = 76
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("K")
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
            
        
        'set so thang tinh thue neu la nam 2011 thi co tinh so thang mien giam
        .GetText .ColLetterToNumber("N"), 4, tuThang
        .GetText .ColLetterToNumber("P"), 4, denThang

        'tinh so thang ke khai
        'ke khai nam 2011 se settext vao cell Q, <>2011 set vao cell P
        'neu ke khai nam 2009 thi so thang quyet toan = 6
        soThangKG = DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThang, "mm/yyyy")) + 1
        If TAX_Utilities_v1.Year = "2011" Or TAX_Utilities_v1.Year = "2012" Then
            'tinh so thang Ke khai
            soThangKG = getSoThangQTMienGiam(CStr(tuThang), CStr(denThang))
            .SetText .ColLetterToNumber("S"), 4, soThangKG
            If isMienGiamPL26(CStr(tuThang), CStr(denThang)) Then
                .SetText .ColLetterToNumber("V"), 53, 1
            Else
                .SetText .ColLetterToNumber("V"), 53, 0
            End If
        ElseIf TAX_Utilities_v1.Year = "2009" Then
            .SetText .ColLetterToNumber("S"), 4, "6"
        Else
            .SetText .ColLetterToNumber("S"), 4, "1"
        End If
        .Col = .ColLetterToNumber("R")
        .Row = 4
        .Text = DateDiff("M", Format(CStr(tuThang), "mm/yyyy"), Format(CStr(denThang), "mm/yyyy")) + 1
        
        .GetText .ColLetterToNumber("R"), 4, soThang
        
        arrTuThang = Split(tuThang, "/")
        arrDenThang = Split(denThang, "/")
        
        ' neu den thang < 7/2013
        If Val(arrDenThang(1)) < 2013 Or (Val(arrDenThang(1)) = 2013 And Val(arrDenThang(0)) < 7) Then
            .Row = 48
            .Col = .ColLetterToNumber("K")
            .value = 4000000 * Val(soThang)
        ElseIf (Val(arrTuThang(1)) > 2013) Or ((Val(arrTuThang(1)) = 2013) And (Val(arrTuThang(0)) >= 7)) Then
            .Row = 48
            .Col = .ColLetterToNumber("K")
            .value = 9000000 * Val(soThang)
        Else
            soThang1 = DateDiff("M", Format_mmyyyy(CStr(tuThang)), Format_mmyyyy("06/2013")) + 1
            soThang2 = DateDiff("M", Format_mmyyyy("07/2013"), Format_mmyyyy(CStr(denThang))) + 1
            .Row = 48
            .Col = .ColLetterToNumber("K")
            .value = 4000000 * soThang1 + 9000000 * soThang2
        End If
        
        'check CT 15 < 0 thi khong cho nhap CT 16 tren PL 09-2
        .Sheet = phuLuc09_2_TNCN
        .GetText .ColLetterToNumber("I"), 26, varCT15

        If Val(varCT15) < 0 Then
            .Col = .ColLetterToNumber("I")
            .Row = 27
            .Lock = True
        Else
            .Col = .ColLetterToNumber("I")
            .Row = 27
            .Lock = False
        End If
        'check thong tin ma so thue dai ly
        'set vi tri theo thu tu MSTDL~TENNVDLT~ChungChiHN
        viTriMSTDL = "F_24~F_69~F_71"
        updateMSTDL fps, viTriMSTDL
        
        'set active vao cell tu ngay
'        .Sheet = 1
'        .SetActiveCell .ColLetterToNumber("N"), 4

        .Sheet = phuLuc09_3_TNCN
        .Row = 25
        Do
            .Col = .ColLetterToNumber("B")
            tinhGiamTruNPT_Load .Row
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
        
        .Row = .Row + 7
        Do
            .Col = .ColLetterToNumber("B")
            tinhGiamTruNPT_Load .Row
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
        
        ' chekc sheet 09-4
'        .Sheet = 5
'        Dim vCt6, vCt7 As Variant
'        .GetText .ColLetterToNumber("I"), 16, vCt6
'        .GetText .ColLetterToNumber("I"), 17, vCt7
'        If Val(vCt6) - Val(vCt7) < 0 Then
'            .Col = .ColLetterToNumber("I")
'            .Row = 19
'            .Lock = True
'        Else
'            .Col = .ColLetterToNumber("I")
'            .Row = 19
'            .Lock = False
'        End If
        
        
        .EventEnabled(EventAllEvents) = True
    End With

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

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    Dim intIndexCombo As Integer
    Dim varTemp       As Variant, varTemp2 As Variant
    Dim tuThang, denThang As Variant
    ' get so thang quyet toan
    Dim soThangKG As Integer
    Dim CT22      As Variant, ct26 As Variant, varCT31 As Variant, varCT33 As Variant
    Dim varCT15   As Variant
    Dim varCT06   As Variant, varCT07   As Variant
    Dim mSheet    As Long
    
    Dim chiTieu22 As Variant
    
    With fps
    
        .EventEnabled(EventAllEvents) = False
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("K") And Row = 54 Then
                .Formula = ""
            End If

            If Col = .ColLetterToNumber("K") And (Row = 48 Or Row = 49 Or Row = 50 Or Row = 51 Or Row = 52) Then
                .Col = .ColLetterToNumber("K")
                .Row = 54
                .Formula = "ROUND(T59,0)"
            End If
        End If

        If .ActiveSheet = 2 Then
            If .Col = .ColLetterToNumber("J") And (Row = 17 Or Row = 20) Then
                .Sheet = 1
                .Col = .ColLetterToNumber("K")
                .Row = 54
                .Formula = "ROUND(T59,0)"
            End If
        End If

        'check CT15 <0 thi khong cho nhap CT16 tren PL 09B
        If .ActiveSheet = 3 Then
            .Sheet = 3
            fps.EventEnabled(EventAllEvents) = False
            If Col = .ColLetterToNumber("I") And Row = 26 Then
                .GetText .ColLetterToNumber("I"), 26, varCT15
                If Val(varCT15) < 0 Then
                    .Col = .ColLetterToNumber("I")
                    .Row = 27
                    .Text = "0"
                    .Lock = True
                Else
                    .Col = .ColLetterToNumber("I")
                    .Row = 27
                    .Lock = False
                End If

            End If
            fps.EventEnabled(EventAllEvents) = True
        End If
        
                'check CT08 <0 thi set chi tieu 09 =0
'        If .ActiveSheet = 5 Then
'            .Sheet = 5
'            fps.EventEnabled(EventAllEvents) = False
'            If Col = .ColLetterToNumber("I") And (Row = 16 Or Row = 17) Then
'                .GetText .ColLetterToNumber("I"), 16, varCT06
'                .GetText .ColLetterToNumber("I"), 17, varCT07
'                If Val(varCT06) - Val(varCT07) < 0 Then
'                    .Col = .ColLetterToNumber("I")
'                    .Row = 19
'                    .Lock = True
'                    .Text = "0"
'                Else
'                    .Col = .ColLetterToNumber("I")
'                    .Row = 19
'                    .Lock = False
'                End If
'            End If
'            fps.EventEnabled(EventAllEvents) = True
'        End If

        
'            If .ActiveSheet = 4 Then
'                If Col = .ColLetterToNumber("L") And .Text <> "" Then
'                    intIndexCombo = .typeComboboxCurSel
'                    .Row = Row
'                    .Col = .ColLetterToNumber("R")
'                    .typeComboboxCurSel = intIndexCombo
'                    UpdateCell fps, .Col, .Row, .value
'                End If
'
'                If Col = .ColLetterToNumber("H") And .Text <> "" Then
'                    intIndexCombo = .typeComboboxCurSel
'                    .Row = Row
'                    .Col = .ColLetterToNumber("Q")
'                    .typeComboboxCurSel = intIndexCombo
'                    UpdateCell fps, .Col, .Row, .value
'                End If
'            End If
            
        .EventEnabled(EventAllEvents) = True
        End With
        tinhGiamTruNPT Col, Row
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
            ' chon tinh thanh
            If Col = .ColLetterToNumber("K") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("J")
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
                    .Col = .ColLetterToNumber("M")
                    .TypeComboBoxList = tenHuyen
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("L")
                    .TypeComboBoxList = maHuyen
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                   
                    .Row = Row
                    .Col = .ColLetterToNumber("O")
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("N")
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                                   
                
                End If
                
                
            End If
            ' chon quan he
            If Col = .ColLetterToNumber("O") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("N")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
            End If
            
        
            If Col = .ColLetterToNumber("I") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("H")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
                If .Text = "01" Then
                    .Col = .ColLetterToNumber("K")
                    .Row = Row
                    .Lock = False
                    
                    .Col = .ColLetterToNumber("M")
                    .Row = Row
                    .Lock = False
                    
                    .Col = .ColLetterToNumber("O")
                    .Row = Row
                    .Lock = False
                    
                    .Col = .ColLetterToNumber("J")
                    .Row = Row
                    .Text = ""
                    
                    .Col = .ColLetterToNumber("L")
                    .Row = Row
                    .Text = ""
                    
                    .Col = .ColLetterToNumber("N")
                    .Row = Row
                    .Text = ""
                Else
                    .Col = .ColLetterToNumber("K")
                    .Row = Row
                    .Text = "Khác"
                    .Lock = True
                    
                    .Col = .ColLetterToNumber("M")
                    .Row = Row
                    .Text = "Khác"
                    .Lock = True
                    
                    .Col = .ColLetterToNumber("O")
                    .Row = Row
                    .Text = "Khác"
                    .Lock = True
                    
                    
                    .Col = .ColLetterToNumber("J")
                    .Row = Row
                    .Text = "999"
                    
                    .Col = .ColLetterToNumber("L")
                    .Row = Row
                    .Text = "99999"
                    
                    .Col = .ColLetterToNumber("N")
                    .Row = Row
                    .Text = "9999999"
                End If
            End If
            
            
            If Col = .ColLetterToNumber("K") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("J")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
            End If
            
            ' chon huyen
            If Col = .ColLetterToNumber("M") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("L")
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
                    .Col = .ColLetterToNumber("O")
                    .TypeComboBoxList = tenXa
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("N")
                    .TypeComboBoxList = maXa
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                
                End If
            End If
            
            If Col = .ColLetterToNumber("O") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("N")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
            End If
            
            If Col = .ColLetterToNumber("Q") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("P")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
            End If
            
            If Col = .ColLetterToNumber("S") And .Text <> "" Then
                intIndexCombo = .typeComboboxCurSel
                 .Row = Row
                 .Col = .ColLetterToNumber("R")
                 .typeComboboxCurSel = intIndexCombo
                UpdateCell fps, .Col, .Row, .value
            End If
        End If
    End With
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim iCol As Long, iRow As Long
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tCol As Long, tRow As Long
Dim i As Integer
    With fps
        
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 4 Then
                    fps.EventEnabled(EventAllEvents) = False
                    
                    .Col = .ColLetterToNumber("B")
                    i = 1
                    .Row = 25
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = .Row + 1
                         i = i + 1
                    Loop Until .Text = "aa"
                    
                    
                    .Row = .Row + 7
                    i = 1
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = .Row + 1
                         i = i + 1
                    Loop Until .Text = "aa"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    
                    .SetActiveCell iCol, iRow
            End If
            
        End If
  
      fps.EventEnabled(EventAllEvents) = True
    End With
End Sub

 Private Sub fps_LeaveCell(ByVal Col As Long, _
                          ByVal Row As Long, _
                          ByVal NewCol As Long, _
                          ByVal NewRow As Long, _
                          Cancel As Boolean)
                          
    Dim varTemp As Variant
    Dim tuThang, denThang As Variant
    ' get so thang quyet toan
    Dim soThangQT As Integer, soThangKG As Integer
    Dim CT22      As Variant, ct26 As Variant
    
    Dim firstRowGroup2 As Long
    
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet

        If .ActiveSheet = 1 Then
            'check date
            If Col = .ColLetterToNumber("O") And Row = 71 Then
                .GetText Col, Row, varTemp

                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    Else
'                        .SetFocus
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
                    
        If .ActiveSheet = 4 Then
            If Col = .ColLetterToNumber("E") And Row >= 24 Then
                       .GetText Col, Row, varTemp
                       If varTemp <> "" And varTemp <> "../../...." Then
                           If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                               '.SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                               If CDate(Format_ddmmyyyy(CStr(varTemp))) < CDate(Format_ddmmyyyy(Format(Date, "dd/mm/yyyy"))) Then
'                                     .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                                Else
'                                     DisplayMessage "0175", msOKOnly, miCriticalError
'                                     .SetText Col, Row, ""
'                                     .SetActiveCell Col, Row
'                                End If
                                .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                           Else
                               .SetActiveCell Col, Row
                           End If
                       Else
                        .SetText Col, Row, ""
                       End If
                .Col = Col
                .Row = Row
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            ' Truong tu thang den thang group 1
            .Row = 25
            Do
                .Col = .ColLetterToNumber("B")
                .Row = .Row + 1
            Loop Until .Text = "aa"
            firstRowGroup2 = .Row + 7
            ' group 1
            If Col = .ColLetterToNumber("Q") And Row >= 25 And Row < firstRowGroup2 Then
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
            
            If Col = .ColLetterToNumber("S") And Row >= 25 And Row < firstRowGroup2 Then
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
            
            ' group 2
            If Col = .ColLetterToNumber("T") And Row >= firstRowGroup2 Then
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
            
            If Col = .ColLetterToNumber("U") And Row >= firstRowGroup2 Then
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
    tinhGiamTruNPT Col, Row
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
'        SetData
        
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
    'chuoi vitriheader tren to khai, chuoi vi tri tren sheet header de check lai
    Dim viTriHeaderTk As String, viTriSheetHeader As String, mangMessage As String, xmlNode As MSXML.IXMLDOMNode, strTenNVDDL, strCCHN As String
    Dim MSTDTNT As Variant, tSheet As Integer
    Dim ct12   As Variant, cT40 As Variant, mErrCT12 As Variant, iFlagErrCT12 As Integer
    
    With fps
        CheckErrorMST
        'check thong tin header
        tSheet = .Sheet
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 10, MSTDTNT
        MSTDTNT = Trim(MSTDTNT)
        'check thong tin header khong dc trong
        'lay vi tri header tren tk theo thu tu MSTDL, TENDLT, DIACHIDLT,QUAN/HUYEN,TP,SO HDDLT,NGAYDH,Ten NV DLT, CC Hanh Nghe
        viTriHeaderTk = "F_24~F_22~F_26~F_28~O_28~F_34~O_34~F_69~F_71"
        viTriSheetHeader = "B_68~B_69~B_70~B_71~B_72~B_73~B_74~B_75~B_76"
        mangMessage = "E_68~E_69~E_70~E_71~E_72~E_73~E_74~E_75~E_76"
        'check header
        checkErrorHeader fps, viTriHeaderTk, viTriSheetHeader, mangMessage, MSTDTNT
        
        .Sheet = tSheet
        ' Neu ton tai file data phu luc 09C thi kiem tra tren phu luc 09C
        If IIf(TAX_Utilities_v1.NodeValidity.childNodes(3).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            checkSheet4
        Else
            clearCheckSheet4
        End If
        
    End With
End Sub

Private Sub checkSheet4()
    Dim strCheck As String, iCurrentSheet As Integer
    Dim strFocusSheetName As String, strFocusCol As String, strFocusRow As String
    Dim vMST As Variant, vErrMsgMST As Variant
    Dim vCMT As Variant, vErrMsgCMT As Variant
    Dim vHoVaTen As Variant, vErrMsgHoVaTen As Variant
    Dim vNgaySinh As Variant, vErrMsgNgaySinh As Variant
    Dim vNgaySinh_Goc As Variant
    Dim vQuanHeNNT As Variant, vErrMsgQuanHeNNT As Variant
    Dim vLastRow As Variant
    Dim vTuThang, vDenThang As Variant
    Dim vCMND As Variant, iFlagCMDN_HC As Integer
    Dim vErrEmpty As Variant, iFlagEmty As Integer
    
    Dim iFocusFlag As Integer
    Dim mErrNgaySinhPhuThuoc As Variant, iFlagNgaySinhPhuThuoc As Integer
    Dim mErrMSTCT5Trung As Variant, iFlagMSTCT5Trung As Integer, rowMSTTrung As Integer
    Dim iFlagEmpty As Integer, iFlagMSTTrung As Integer
    Dim vErrTuThang As Variant, vErrMsgMST2 As Variant, vErrMsgMSTCMT As Variant
    Dim iFlagTuThang As Integer, iFlagTuThangTK As Integer, iFlagCMTTrung As Integer
    Dim CountRow1 As Long, rowDes As Long
    Dim iFlagvMst As Integer, iFlagvHoVaTen As Integer, iFlagvNgaySinh As Integer, iFlagLoiMST As Integer, iFlagDenThang As Integer
    Dim vErrDenThang As Variant, vErrTuThangTK As Variant, vErrMsgCMT2 As Variant
    Dim j As Integer, i As Integer
    Dim blCheck_S1, blCheck_S2, blCheck_S3, blCheck_S4, blCheck_S5
     Dim vErrDuLieuKhongHL As Variant
     Dim strColReset As String
    Dim arrColReset() As String
    Dim vErrEmptyCMNDNPT As Variant
    Dim mErrNgaySinhPhuThuoc1 As Variant
     
    With fps
        iCurrentSheet = .Sheet
        
        .Sheet = 1
        .GetText .ColLetterToNumber("N"), 4, vTuThang
        .GetText .ColLetterToNumber("P"), 4, vDenThang
        
        .Sheet = .SheetCount
        
        ' Lay thong bao sai ve ngay thanh toan trong sheet header
        .GetText .ColLetterToNumber("E"), 17, vErrMsgHoVaTen
        ' Lay thong bao sai ve ngay thanh toan trong sheet header
        .GetText .ColLetterToNumber("E"), 18, vErrMsgNgaySinh
        ' Lay thong bao ve ko chon quan he voi NNT trong sheet header
        .GetText .ColLetterToNumber("E"), 19, vErrMsgQuanHeNNT
        ' Lay thong bao ve ko nhap ma so thue trong sheet header
        .GetText .ColLetterToNumber("E"), 20, vErrMsgCMT
        ' Lay thong bao ve sai ma so thue trong sheet header
        .GetText .ColLetterToNumber("E"), 21, vErrMsgMST
        'lay thong bao ve nhap sai cmnd/ho chieu khi nguoi dung nhap dau cach
        .GetText .ColLetterToNumber("E"), 44, vErrMsgCMT
        'lay thong bao ve truong bat buoc phai nhap
        .GetText .ColLetterToNumber("E"), 45, vErrEmpty
        'lay thong bao ve ngay sinh cua nguoi phu thuoc phai truoc ky tinh thue
        .GetText .ColLetterToNumber("E"), 46, mErrNgaySinhPhuThuoc
        
        .GetText .ColLetterToNumber("E"), 47, mErrMSTCT5Trung
        
        .GetText .ColLetterToNumber("E"), 55, vErrTuThang
        .GetText .ColLetterToNumber("E"), 56, vErrMsgMST2
        .GetText .ColLetterToNumber("E"), 57, vErrMsgMSTCMT
        .GetText .ColLetterToNumber("E"), 59, vErrDenThang
        .GetText .ColLetterToNumber("E"), 60, vErrTuThangTK
        .GetText .ColLetterToNumber("E"), 61, vErrMsgCMT2
        .GetText .ColLetterToNumber("E"), 62, vErrDuLieuKhongHL
        .GetText .ColLetterToNumber("E"), 63, vErrEmptyCMNDNPT
        
        .GetText .ColLetterToNumber("E"), 64, mErrNgaySinhPhuThuoc1
        
        .Sheet = 4 'Check tinh xac thuc cua cac thong tin tren bang ke 09-3/TNCN
        
        blCheck_S1 = True
        blCheck_S2 = True
        blCheck_S3 = True
        blCheck_S4 = True
        blCheck_S5 = True
        iFlagNgaySinhPhuThuoc = 0
        iFlagMSTCT5Trung = 0
                
        ' Lay ma so thue tren phu luc 09C/TNCN chi tieu D11 (Ma so thue cua vo hoac chong)
        .GetText .ColLetterToNumber("E"), 12, vMST
        'lay ho va ten nguoi phu thuoc
        .GetText .ColLetterToNumber("D"), 25, vHoVaTen
        ' Check ma so thue tren phu luc 09C/TNCN chi tieu D11 (Ma so thue cua vo hoac chong)
        
        If Trim(vMST) <> vbNullString Then
            vMST = Left(vMST, 10)
           strCheck = ""
                   
           If Len(vMST) = 10 Or Len(vMST) = 13 Then
                strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
           End If
           If Mid(strCheck, 2, 1) <> "0" Then
                .Col = .ColLetterToNumber("E")
                .Row = 12
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsgMST   'static
                .BackColor = mErrorColor
                iFlagvMst = 1
                blCheck_S5 = False
                'capture the first error cell
                If iFocusFlag = 0 Then
                    strFocusSheetName = "'" & .SheetName & "'"
                    strFocusCol = str(.Col)
                    strFocusRow = str(.Row)
                    iFocusFlag = 1
                End If
            Else
                If Trim(vHoVaTen) <> vbNullString Then
                    rowMSTTrung = .SearchCol(.ColLetterToNumber("F"), 21, -1, vMST, SearchFlagsNone)
                    If rowMSTTrung <> -1 Then
                        .Col = .ColLetterToNumber("E")
                        .Row = 12
                        .CellNote = mErrMSTCT5Trung
                        .BackColor = mErrorColor
                        iFlagMSTCT5Trung = 1
                    Else
                        .Col = .ColLetterToNumber("E")
                        .Row = 12
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                Else
                 .Col = .ColLetterToNumber("E")
                 .Row = 12
                 .CellNote = ""
                 .BackColor = mNonErrorColor
                End If
            End If
        Else
            .Col = .ColLetterToNumber("E")
            .Row = 12
            .CellNote = ""
            .BackColor = mNonErrorColor
        End If
        
        'dntai 25/08/2011
        'Lay gia so cmnd , so ho chieu nguoi dung nhap vao
        .GetText .ColLetterToNumber("E"), 14, vCMND
        If InStr(1, CStr(vCMND), " ") <> 0 Then
            .Col = .ColLetterToNumber("E")
            .Row = 14
            .BackColor = mErrorColor
            .CellNote = vErrMsgCMT
            iFlagCMDN_HC = 1
        Else
            .Col = .ColLetterToNumber("E")
            .Row = 14
            .BackColor = mNonErrorColor
            .CellNote = ""
        End If
        'end
        
        For j = 25 To .MaxRows
            .Row = j
            
            .GetText .ColLetterToNumber("D"), .Row, vHoVaTen
            .GetText .ColLetterToNumber("E"), .Row, vNgaySinh
            .GetText .ColLetterToNumber("E"), .Row, vNgaySinh_Goc
            .GetText .ColLetterToNumber("F"), .Row, vMST
            .GetText .ColLetterToNumber("M"), .Row, vCMT
            .GetText .ColLetterToNumber("O"), .Row, vQuanHeNNT
            
            .GetText .ColLetterToNumber("Q"), .Row, vTuThang
            .GetText .ColLetterToNumber("S"), .Row, vDenThang
            
            .GetText .ColLetterToNumber("B"), .Row, vLastRow
            'reset du lieu loi
            .Col = .ColLetterToNumber("F")
            .CellNote = ""
            .BackColor = mNonErrorColor
            
            
            ' reset du lieu loi
            strColReset = "D~E~F~G~K~M~N~O~Q~S"
            arrColReset = Split(strColReset, "~")
            For i = 0 To UBound(arrColReset)
                .Col = .ColLetterToNumber(arrColReset(i))
                .BackColor = mNonErrorColor
                .CellNote = ""
            Next

            
            If UCase(vLastRow) = "AA" Then Exit For ' Check cho den het cac dong co du lieu thi thoi
            
            If Trim(vHoVaTen) = vbNullString And Trim(vNgaySinh) = vbNullString And Trim(vMST) = vbNullString And Trim(vCMT) = vbNullString And vTuThang = vbNullString And vDenThang = vbNullString Then
            Else
                'Check Ho va ten tren phu luc 09-3/TNCN
                If Trim(vHoVaTen) = vbNullString Then
                    .Col = .ColLetterToNumber("D")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgHoVaTen    'static
                    .BackColor = mErrorColor
                    iFlagvHoVaTen = 1
                    blCheck_S1 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("D")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                'Check Ngay sinh tren phu luc 09-3/TNCN
                If Trim(vNgaySinh) = vbNullString Then
                    .Col = .ColLetterToNumber("E")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgNgaySinh     'static
                    .BackColor = mErrorColor
                    iFlagvNgaySinh = 1
                    blCheck_S2 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                    'check ngay sinh phai sau ky tinh thue
                    'If Val(Right(CStr(Trim(vNgaySinh)), 4)) > Val(TAX_Utilities_v1.Year) Then
                    If Format_ddmmyyyy1(CStr(vNgaySinh_Goc)) <> "" Then
                            vNgaySinh = DateSerial(CInt(Mid$(vNgaySinh, 7, 4)), CInt(Mid$(vNgaySinh, 4, 2)), CInt(Mid$(vNgaySinh, 1, 2)))
                            If vNgaySinh > Date Then
                                .Col = .ColLetterToNumber("E")
                                .BackColor = mErrorColor
                                .CellNote = mErrNgaySinhPhuThuoc
                                iFlagNgaySinhPhuThuoc = 1
                            Else
                                 If Trim(vTuThang) <> "" Then
                                    If DateDiff("M", vNgaySinh, Format(vTuThang, "mm/yyyy")) < 0 Then
                                        .Col = .ColLetterToNumber("E")
                                        .BackColor = mErrorColor
                                        .CellNote = mErrNgaySinhPhuThuoc1
                                        iFlagNgaySinhPhuThuoc = 1
                                    Else
                                        .Col = .ColLetterToNumber("E")
                                        .BackColor = mNonErrorColor
                                        .CellNote = ""
                                    End If
                                Else
                                    .Col = .ColLetterToNumber("E")
                                    .BackColor = mNonErrorColor
                                    .CellNote = ""
                                End If
                            End If
                       Else
                            .Col = .ColLetterToNumber("E")
                            .BackColor = mErrorColor
                            .CellNote = vErrDuLieuKhongHL
                            iFlagNgaySinhPhuThuoc = 1
                       End If
                End If
                
                
                If Trim(vMST) = vbNullString And Trim(vCMT) = vbNullString Then
                    .Col = .ColLetterToNumber("F")
                    .CellNote = vErrMsgMSTCMT
                    .BackColor = mErrorColor
                    .Col = .ColLetterToNumber("M")
                    .CellNote = vErrMsgMSTCMT
                    .BackColor = mErrorColor
                    iFlagLoiMST = 1
                Else
                    .Col = .ColLetterToNumber("F")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                    .Col = .ColLetterToNumber("M")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                
                'Check ma so thue tren phu luc 09-3/TNCN
                If Trim(vMST) <> vbNullString Then
                   strCheck = ""
                   If Len(vMST) = 14 And InStr(1, vMST, "-") = 11 Then
                    vMST = Replace(vMST, "-", "")
                   End If
                   
                   If Len(vMST) = 10 Or Len(vMST) = 13 And InStr(1, vMST, "-") = 0 Then
                        strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                   End If
                   If Mid(strCheck, 2, 1) <> "0" Then
                        .Col = .ColLetterToNumber("F")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgMST   'static
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
                    Else
                        CountRow1 = .SearchCol(.ColLetterToNumber("F"), -1, -1, vMST, SearchFlagsNone)
                         rowDes = .SearchCol(.ColLetterToNumber("F"), CountRow1, -1, vMST, SearchFlagsNone)
                        If rowDes > -1 Then
                            .Col = .ColLetterToNumber("F")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgMST2       'static
                            .BackColor = mErrorColor
                            iFlagMSTTrung = 1
                        Else
                              .Col = .ColLetterToNumber("F")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                         .Col = .ColLetterToNumber("M")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                Else
                End If
                
                'check SoCMND khong dc co dau cach
                If Trim(vCMT) <> vbNullString Then
                    If InStr(1, CStr(vCMT), " ") <> 0 Then
                        .Col = .ColLetterToNumber("M")
                        .BackColor = mErrorColor
                        .CellNote = vErrMsgCMT
                        iFlagCMDN_HC = 1
                    Else
'                        .Col = .ColLetterToNumber("M")
'                        .BackColor = mNonErrorColor
'                        .CellNote = ""
                         CountRow1 = .SearchCol(.ColLetterToNumber("M"), -1, -1, vCMT, SearchFlagsNone)
                         rowDes = .SearchCol(.ColLetterToNumber("M"), CountRow1, -1, vCMT, SearchFlagsNone)
                        If rowDes > -1 Then
                            .Col = .ColLetterToNumber("M")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgCMT2
                            .BackColor = mErrorColor
                            iFlagCMTTrung = 1
                        Else
                            
                              .Col = .ColLetterToNumber("M")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    End If
                 Else
                     ' kiem tra neu tuoi >15 bat buoc nhap CMT
                    If vNgaySinh <> "" Then
                        If Format_ddmmyyyy1(CStr(vNgaySinh_Goc)) <> "" Then
                            If Year(Date) - Year(Format(vNgaySinh, "dd/mm/yyyy")) > 15 And Trim(vCMT) = "" Then
                                 .Col = .ColLetterToNumber("M")
                                .CellNote = vErrEmptyCMNDNPT
                                .BackColor = mErrorColor
                                iFlagEmpty = 1
                            End If
                        Else
                            .Col = .ColLetterToNumber("M")
                            .CellNote = vErrDuLieuKhongHL
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                        End If
                    End If
                End If
                ' Check quan he voi NNT
                'Check Ho va ten tren phu luc 09-3/TNCN
                If Trim(vQuanHeNNT) = vbNullString Then
                    .Col = .ColLetterToNumber("O")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgQuanHeNNT     'static
                    .BackColor = mErrorColor
                    blCheck_S3 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("O")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                ' Kiem tra tu thang bat buoc nhap
                 If Trim(vTuThang) <> "" Then
                    .Col = .ColLetterToNumber("Q")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                 Else
                    .Col = .ColLetterToNumber("Q")
                    .CellNote = vErrEmpty
                    .BackColor = mErrorColor
                    iFlagEmpty = 1
                 End If
                 
                 ' kiem tra den thang bat buoc nhap
                 If Trim(vDenThang) <> "" Then
                    .Col = .ColLetterToNumber("S")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                 Else
                    .Col = .ColLetterToNumber("S")
                    .CellNote = vErrEmpty
                    .BackColor = mErrorColor
                    iFlagEmpty = 1
                 End If
                 
                 ' Kiem tra tu thang, den thang
                If Trim(vTuThang) <> "" And Trim(vDenThang) <> "" Then
                    If DateDiff("M", Format(vTuThang, "mm/yyyy"), Format(vDenThang, "mm/yyyy")) < 0 Then
                            .Col = .ColLetterToNumber("Q")
                            .CellNote = vErrTuThang
                            .BackColor = mErrorColor
                            
                            .Col = .ColLetterToNumber("S")
                            .CellNote = vErrTuThang
                            .BackColor = mErrorColor
                            iFlagTuThang = 1
                    Else
                            ' kiem tra tu thang thuoc ky ke khai
                             If vTuThang <> "" Then
                                If DateDiff("M", Format(vTuThang, "mm/yyyy"), Format(TAX_Utilities_v1.FirstDay, "mm/yyyy")) > 0 Or DateDiff("M", Format(vTuThang, "mm/yyyy"), Format(TAX_Utilities_v1.LastDay, "mm/yyyy")) < 0 Then
                                     .Col = .ColLetterToNumber("Q")
                                    .CellNote = vErrTuThangTK
                                    .BackColor = mErrorColor
                                    iFlagTuThangTK = 1
                                Else
                                    .Col = .ColLetterToNumber("Q")
                                    .CellNote = ""
                                    .BackColor = mNonErrorColor
                                End If
                            End If
'                            .Col = .ColLetterToNumber("Q")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
                            
'                            .Col = .ColLetterToNumber("S")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
                             ' kiem tra den thang voi den thang ky QT
                            If vDenThang <> "" Then
                                If DateDiff("M", Format(vDenThang, "mm/yyyy"), Format(TAX_Utilities_v1.LastDay, "mm/yyyy")) < 0 Then
                                     .Col = .ColLetterToNumber("S")
                                    .CellNote = vErrDenThang
                                    .BackColor = mErrorColor
                                    iFlagDenThang = 1
                                Else
                                    .Col = .ColLetterToNumber("S")
                                    .CellNote = ""
                                    .BackColor = mNonErrorColor
                                End If
                            End If
                    End If
                 End If
            End If
        'ket thuc vong lap
        Next
        
        ' check group 2
        Dim vSo As Variant
        Dim vQuyenSo As Variant
        Dim vQuocGia, vTinh, vHuyen, vXa, vQuocTich As Variant
        Dim vMaQuocGia, vMaTinh, vMaHuyen, vMaXa As Variant
       
        
        .Row = .Row + 7
        For j = .Row To .MaxRows
            .Row = j
            
            .GetText .ColLetterToNumber("D"), .Row, vHoVaTen
            .GetText .ColLetterToNumber("E"), .Row, vNgaySinh
            .GetText .ColLetterToNumber("E"), .Row, vNgaySinh_Goc
            .GetText .ColLetterToNumber("F"), .Row, vSo
            .GetText .ColLetterToNumber("G"), .Row, vQuyenSo
            .GetText .ColLetterToNumber("I"), .Row, vQuocGia
            .GetText .ColLetterToNumber("K"), .Row, vTinh
            .GetText .ColLetterToNumber("M"), .Row, vHuyen
            .GetText .ColLetterToNumber("O"), .Row, vXa
            .GetText .ColLetterToNumber("Q"), .Row, vQuocTich
            .GetText .ColLetterToNumber("S"), .Row, vQuanHeNNT
            .GetText .ColLetterToNumber("T"), .Row, vTuThang
            .GetText .ColLetterToNumber("U"), .Row, vDenThang
            .GetText .ColLetterToNumber("B"), .Row, vLastRow
            
            .GetText .ColLetterToNumber("H"), .Row, vMaQuocGia
            .GetText .ColLetterToNumber("J"), .Row, vMaTinh
            .GetText .ColLetterToNumber("L"), .Row, vMaHuyen
            .GetText .ColLetterToNumber("N"), .Row, vMaXa
            
            'reset du lieu loi
            .Col = .ColLetterToNumber("F")
            .CellNote = ""
            .BackColor = mNonErrorColor


            ' reset du lieu loi
            strColReset = "D~E~F~G~I~K~M~N~O~Q~S~T~U"
            arrColReset = Split(strColReset, "~")
            For i = 0 To UBound(arrColReset)
                .Col = .ColLetterToNumber(arrColReset(i))
                .BackColor = mNonErrorColor
                .CellNote = ""
            Next
            
            
            If UCase(vLastRow) = "AA" Then Exit For ' Check cho den het cac dong co du lieu thi thoi
            
            If Trim(vHoVaTen) = vbNullString And Trim(vNgaySinh) = vbNullString And Trim(vSo) = vbNullString And Trim(vQuyenSo) = vbNullString _
            And Trim(vTinh) = vbNullString And Trim(vHuyen) = vbNullString And Trim(vXa) = vbNullString And Trim(vQuanHeNNT) = vbNullString And vTuThang = vbNullString And vDenThang = vbNullString Then
            Else
                'Check Ho va ten tren phu luc 09-3/TNCN
                If Trim(vHoVaTen) = vbNullString Then
                    .Col = .ColLetterToNumber("D")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgHoVaTen    'static
                    .BackColor = mErrorColor
                    iFlagvHoVaTen = 1
                    blCheck_S1 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("D")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                'Check Ngay sinh tren phu luc 09-3/TNCN
                If Trim(vNgaySinh) = vbNullString Then
                    .Col = .ColLetterToNumber("E")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgNgaySinh     'static
                    .BackColor = mErrorColor
                    iFlagvNgaySinh = 1
                    blCheck_S2 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                    'check ngay sinh phai sau ky tinh thue
                    'If Val(Right(CStr(Trim(vNgaySinh)), 4)) > Val(TAX_Utilities_v1.Year) Then
                    If Format_ddmmyyyy1(CStr(vNgaySinh_Goc)) <> "" Then
                        vNgaySinh = DateSerial(CInt(Mid$(vNgaySinh, 7, 4)), CInt(Mid$(vNgaySinh, 4, 2)), CInt(Mid$(vNgaySinh, 1, 2)))
                        If vNgaySinh > Date Then
                            .Col = .ColLetterToNumber("E")
                            .BackColor = mErrorColor
                            .CellNote = mErrNgaySinhPhuThuoc
                            iFlagNgaySinhPhuThuoc = 1
                        Else
                            If Trim(vTuThang) <> "" Then
                                If DateDiff("M", vNgaySinh, Format(vTuThang, "mm/yyyy")) < 0 Then
                                    .Col = .ColLetterToNumber("E")
                                    .BackColor = mErrorColor
                                    .CellNote = mErrNgaySinhPhuThuoc1
                                    iFlagNgaySinhPhuThuoc = 1
                                Else
                                    .Col = .ColLetterToNumber("E")
                                    .BackColor = mNonErrorColor
                                    .CellNote = ""
                                End If
                            Else
                                .Col = .ColLetterToNumber("E")
                                .BackColor = mNonErrorColor
                                .CellNote = ""
                            End If
                        End If
                    Else
                        .Col = .ColLetterToNumber("E")
                        .BackColor = mErrorColor
                        .CellNote = vErrDuLieuKhongHL
                        iFlagNgaySinhPhuThuoc = 1
                    End If
                End If
                
                
                
                ' Check quan he voi NNT
                'Check Ho va ten tren phu luc 09-3/TNCN
                If Trim(vQuanHeNNT) = vbNullString Then
                    .Col = .ColLetterToNumber("S")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgQuanHeNNT     'static
                    .BackColor = mErrorColor
                    blCheck_S3 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("S")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                
                ' Kiem tra tu thang bat buoc nhap
                 If Trim(vTuThang) <> "" Then
                    .Col = .ColLetterToNumber("T")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                 Else
                    .Col = .ColLetterToNumber("T")
                    .CellNote = vErrEmpty
                    .BackColor = mErrorColor
                    iFlagEmpty = 1
                 End If
                 
                 ' kiem tra den thang bat buoc nhap
                 If Trim(vDenThang) <> "" Then
                    .Col = .ColLetterToNumber("U")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                 Else
                    .Col = .ColLetterToNumber("U")
                    .CellNote = vErrEmpty
                    .BackColor = mErrorColor
                    iFlagEmpty = 1
                 End If
                 
                 ' Kiem tra tu thang, den thang
                If Trim(vTuThang) <> "" And Trim(vDenThang) <> "" Then
                    If DateDiff("M", Format(vTuThang, "mm/yyyy"), Format(vDenThang, "mm/yyyy")) < 0 Then
                            .Col = .ColLetterToNumber("T")
                            .CellNote = vErrTuThang
                            .BackColor = mErrorColor
                            
                            .Col = .ColLetterToNumber("U")
                            .CellNote = vErrTuThang
                            .BackColor = mErrorColor
                            iFlagTuThang = 1
                    Else
'                            .Col = .ColLetterToNumber("T")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
'
'                            .Col = .ColLetterToNumber("U")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
                                ' kiem tra tu thang thuoc ky ke khai
                                If vTuThang <> "" Then
                                   If DateDiff("M", Format(vTuThang, "mm/yyyy"), Format(TAX_Utilities_v1.FirstDay, "mm/yyyy")) > 0 Or DateDiff("M", Format(vTuThang, "mm/yyyy"), Format(TAX_Utilities_v1.LastDay, "mm/yyyy")) < 0 Then
                                        .Col = .ColLetterToNumber("T")
                                       .CellNote = vErrTuThangTK
                                       .BackColor = mErrorColor
                                       iFlagTuThangTK = 1
                                   Else
                                       .Col = .ColLetterToNumber("T")
                                       .CellNote = ""
                                       .BackColor = mNonErrorColor
                                   End If
                               End If
                               ' kiem tra den thang voi den thang ky QT
                               If vDenThang <> "" Then
                                   If DateDiff("M", Format(vDenThang, "mm/yyyy"), Format(TAX_Utilities_v1.LastDay, "mm/yyyy")) < 0 Then
                                        .Col = .ColLetterToNumber("U")
                                       .CellNote = vErrDenThang
                                       .BackColor = mErrorColor
                                       iFlagDenThang = 1
                                   Else
                                       .Col = .ColLetterToNumber("U")
                                       .CellNote = ""
                                       .BackColor = mNonErrorColor
                                   End If
                               End If
                    End If
                 End If
                 
                 
                 ' kiem tra neu chon quoc gia Viet Nam
                If Trim(vMaQuocGia) = "01" Then
                    ' bat buoc nhap thong tin tinh
                    If Trim(vTinh) = "" Then
                        .Col = .ColLetterToNumber("K")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    Else
                        If Trim$(vMaTinh) = "999" Then
                            .Col = .ColLetterToNumber("K")
                            .CellNote = vErrDuLieuKhongHL
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                        End If
                    End If
                    
                    ' bat buoc nhap thong tin huyen
                    If Trim(vHuyen) = "" Then
                        .Col = .ColLetterToNumber("M")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    Else
                        If Trim$(vMaHuyen) = "99999" Then
                            .Col = .ColLetterToNumber("M")
                            .CellNote = vErrDuLieuKhongHL
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                        End If
                    End If
                    
                    ' bat buoc nhap thong tin xa
                    If Trim(vXa) = "" Then
                        .Col = .ColLetterToNumber("O")
                        .CellNote = vErrEmpty
                        .BackColor = mErrorColor
                        iFlagEmpty = 1
                    Else
                        If Trim$(vMaXa) = "9999999" Then
                            .Col = .ColLetterToNumber("O")
                            .CellNote = vErrDuLieuKhongHL
                            .BackColor = mErrorColor
                            iFlagEmpty = 1
                        End If
                    End If
                End If
            End If
        Next
        

        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 17, IIf(blCheck_S1 = True, "1", "0")
        .SetText .ColLetterToNumber("B"), 18, IIf(blCheck_S2 = True, "1", "0")
        .SetText .ColLetterToNumber("B"), 19, IIf(blCheck_S3 = True, "1", "0")
        .SetText .ColLetterToNumber("B"), 20, IIf(blCheck_S3 = True, "1", "0")
        .SetText .ColLetterToNumber("B"), 21, IIf(blCheck_S3 = True, "1", "0")
        .SetText .ColLetterToNumber("B"), 45, IIf(iFlagEmty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 22, IIf(blCheck_S4 = True, "1", "0")
        .SetText .ColLetterToNumber("B"), 23, IIf(blCheck_S5 = True, "1", "0")
        .SetText .ColLetterToNumber("B"), 44, IIf(iFlagCMDN_HC = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 46, IIf(iFlagNgaySinhPhuThuoc = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 47, IIf(iFlagMSTCT5Trung = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 51, IIf(iFlagEmpty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 55, IIf(iFlagTuThang = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 56, IIf(iFlagMSTTrung = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 59, IIf(iFlagDenThang = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 60, IIf(iFlagTuThangTK = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 61, IIf(iFlagCMTTrung = 1, "0", "1")
        
        
        .Sheet = iCurrentSheet
    End With
End Sub

Private Sub clearCheckSheet4()
    With fps
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 17, "1"
        .SetText .ColLetterToNumber("B"), 18, "1"
        .SetText .ColLetterToNumber("B"), 19, "1"
        .SetText .ColLetterToNumber("B"), 20, "1"
        .SetText .ColLetterToNumber("B"), 21, "1"
        .SetText .ColLetterToNumber("B"), 22, "1"
        .SetText .ColLetterToNumber("B"), 27, "1"
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
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode, xmlCellNodeChild As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "09-1"
                blnNullValue = True

                For Each xmlCellNode In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in E column and I column
'                    If (lCol = fps.ColLetterToNumber("I") And lRow = 16) Or (lCol = fps.ColLetterToNumber("I") And lRow = 18) Or (lCol = fps.ColLetterToNumber("I") And lRow = 20) Or (lCol = fps.ColLetterToNumber("I") And lRow = 23) Or (lCol = fps.ColLetterToNumber("I") And lRow = 26) Then
                        strValue = GetAttribute(xmlCellNode, "Value")
                        If Not IsNullValue(strValue) Then
                            blnNullValue = False
                            Exit For
                        End If
'                    End If
                Next

                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If

             Case "09-2"
                blnNullValue = True

                For Each xmlCellNode In TAX_Utilities_v1.Data(2).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in E column and I column
'                    If (lCol = fps.ColLetterToNumber("I") And lRow = 11) Or (lCol = fps.ColLetterToNumber("I") And lRow = 17) _
'                            Or (lCol = fps.ColLetterToNumber("I") And lRow = 26) Or (lCol = fps.ColLetterToNumber("I") And lRow = 27) _
'                                    Or (lCol = fps.ColLetterToNumber("I") And lRow = 28) Or (lCol = fps.ColLetterToNumber("I") And lRow = 29) Then
                        strValue = GetAttribute(xmlCellNode, "Value")
                        If Not IsNullValue(strValue) Then
                            blnNullValue = False
                            Exit For
                        End If
'                    End If
                Next

                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If

            Case "09-3"
                blnNullValue = True

                For Each xmlCellNode In TAX_Utilities_v1.Data(3).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
                    If lCol = fps.ColLetterToNumber("D") Or _
                       lCol = fps.ColLetterToNumber("E") Then
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
'            Case "09-4"
'                blnNullValue = True
'
'                For Each xmlCellNode In TAX_Utilities_v1.Data(4).getElementsByTagName("Cell")
'                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
'                    'Check value in E column and I column
''                    If (lCol = fps.ColLetterToNumber("I") And lRow = 11) Or (lCol = fps.ColLetterToNumber("I") And lRow = 17) _
''                            Or (lCol = fps.ColLetterToNumber("I") And lRow = 26) Or (lCol = fps.ColLetterToNumber("I") And lRow = 27) _
''                                    Or (lCol = fps.ColLetterToNumber("I") And lRow = 28) Or (lCol = fps.ColLetterToNumber("I") And lRow = 29) Then
'                        strValue = GetAttribute(xmlCellNode, "Value")
'                        If Not IsNullValue(strValue) Then
'                            blnNullValue = False
'                            Exit For
'                        End If
''                    End If
'                Next
'
'                If blnNullValue Then
'                    SetAttribute xmlNodeValid, "Active", "0"
'                End If
            Case "26MT"
                blnNullValue = True

                For Each xmlCellNode In TAX_Utilities_v1.Data(4).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
                    If lCol = fps.ColLetterToNumber("I") Then
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
    Set xmlCellNodeChild = Nothing
End Sub
Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean
    Dim isheet As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
            ' set gia tri CT8 tren 09C
'            isheet = mCurrentSheet
'            .Sheet = 1
'            .SetText .ColLetterToNumber("D"), 20, ""
'            UpdateCell fps, .ColLetterToNumber("D"), 20, ""
'            .SetText .ColLetterToNumber("M"), 20, ""
'            UpdateCell fps, .ColLetterToNumber("M"), 20, ""
'            mCurrentSheet = isheet
        .EventEnabled(EventAllEvents) = True
    End With
End Function

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()
Dim isheet As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
            ' set gia tri CT8 tren 09C
            isheet = mCurrentSheet
'            .Sheet = 4
'            .SetText .ColLetterToNumber("B"), 24, 1
            If .ActiveSheet = 4 Then
                .Row = 25
                .Col = fps.ColLetterToNumber("B")
                .Formula = "B24 + 1"
                
                .Row = 33
                .Col = fps.ColLetterToNumber("B")
                .Formula = "B32 + 1"
            End If
            mCurrentSheet = isheet
        .EventEnabled(EventAllEvents) = True
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
        .Sheet = 1
        mCurrentSheet = .Sheet
        
        If (GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "Active") <> "0") Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'               If lCol = .ColLetterToNumber("P") And lRow = 52 Then
'
'               Else
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
'               End If
             Next
        End If
        If (GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(1), "Active") <> "0") Then
        .Sheet = 2
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
        If (GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(2), "Active") <> "0") Then
        .Sheet = 3
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(2).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
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
        
        If (GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(4), "Active") <> "0") Then
            .Sheet = 5
            mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(4).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
        
        
'         If (GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(5), "Active") <> "0") Then
'        .Sheet = 6
'        mCurrentSheet = .Sheet
'            For Each xmlCellNodeData In TAX_Utilities_v1.Data(5).getElementsByTagName("Cell")
'               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'               .GetText lCol, lRow, varTemp
'               UpdateCell fps, lCol, lRow, varTemp
'             Next
'        End If
      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

Public Sub FinishImport()

End Sub


'Private Function getSoThangQTMienGiam(ByVal tuThang As String, ByVal denThang As String) As Integer
'    Dim soThang As Integer
'    Dim tuThangTmp As Variant, denThangTmp As Variant
'    Dim hieuTuThang As Integer, hieuDenThang As Integer
'    'set Thang mien giam thue tu thang 08/2011 den thang 12/2011
'    tuThangTmp = "08/2011"
'    denThangTmp = "12/2011"
'    'set so Thang quyet toan Thue
'    soThang = 1
'    hieuTuThang = DateDiff("M", Format(tuThangTmp, "mm/yyyy"), Format(tuThang, "mm/yyyy"))
'    hieuDenThang = DateDiff("M", Format(denThangTmp, "mm/yyyy"), Format(denThang, "mm/yyyy"))
'    If hieuTuThang < 0 And hieuDenThang <= 0 And DateDiff("M", Format(tuThangTmp, "mm/yyyy"), Format(denThang, "mm/yyyy")) >= 0 Then ' tinh so thang quyet toan trong TH : tuthang < 08/2011 va den thang < 12/2011
'        denThang = tuThangTmp
'        soThang = DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThang, "mm/yyyy"))
'    ElseIf hieuTuThang >= 0 And DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThangTmp, "mm/yyyy")) >= 0 And hieuDenThang > 0 Then ' tinh so thang quyet toan trong TH : tuThang > 08/2011 va denThang > 12/2011
'        tuThang = denThangTmp
'        soThang = DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThang, "mm/yyyy"))
'    ElseIf hieuTuThang < 0 And hieuDenThang > 0 Then  ' tinh so thang quyet toan trong TH : tu Thang < 08 / 2011 va denThang > 12/2011
'        soThang = DateDiff("M", Format(tuThang, "mm/yyyy"), Format(tuThangTmp, "mm/yyyy"))
'        soThang = soThang + DateDiff("M", Format(denThangTmp, "mm/yyyy"), Format(denThang, "mm/yyyy"))
'    ElseIf hieuTuThang >= 0 And hieuDenThang <= 0 Then   ' tinh so thang quyet toan trong TH : thThang > 08/2011 va denThang < 12/2011
'        soThang = 0
'    Else   ' tinh so thang trong truong hop thang ke khai khong nam trong khoang thang mien giam
'        soThang = DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThang, "mm/yyyy")) + 1
'    End If
'
'    ' tinh so thang mien giam cho 2012
'
'
'
'    getSoThangQTMienGiam = soThang
'End Function


Private Function getSoThangQTMienGiam(ByVal tuThang As String, ByVal denThang As String) As Integer
    Dim soThang As Integer
    'set so Thang quyet toan Thue
    soThang = 0

    ' tinh so thang mien giam cho 2012
    Dim arrDanhsach() As String
    Dim thangMg As String
    Dim countMonth As Integer
    Dim tempDate As Date
    Dim strDataFileName As String
    Dim i As Integer
    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode

    countMonth = DateDiff("M", Format(tuThang, "mm/yyyy"), Format(denThang, "mm/yyyy")) + 1
    If countMonth > 1 Then
        ReDim arrDanhsach(countMonth - 1)
        arrDanhsach(0) = tuThang
        For i = 1 To countMonth - 1
            tempDate = DateAdd("M", i, Format(tuThang, "mm/yyyy"))
            arrDanhsach(i) = Format(tempDate, "mm/yyyy")
        Next
    Else
        ReDim arrDanhsach(1)
        arrDanhsach(0) = tuThang
    End If

    strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_Thang_MT.xml")
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListCell
            If GetAttribute(xmlNode, "Value") <> "" Then
                thangMg = Trim$(GetAttribute(xmlNode, "Value"))
                For i = 0 To countMonth - 1
                    If thangMg = arrDanhsach(i) Then
                        soThang = soThang + 1
                        Exit For
                    End If
                Next
            End If
        Next
    End If


    getSoThangQTMienGiam = soThang
End Function

Private Sub fps_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
    Dim varTemp1 As Variant
    Dim varTemp2 As Variant
    With fps
    
'        If OldSheet = 3 Then
'
'        End If
'        If NewSheet = 1 And OldSheet = 3 Then
'            .Sheet = 3
'            .GetText .ColLetterToNumber("I"), 33, varTemp1
'            .Sheet = 1
'            .GetText .ColLetterToNumber("R"), 56, varTemp2
'            If (varTemp2 <> varTemp1) Then
'                .Col = .ColLetterToNumber("I")
'                .Row = 52
'                .Formula = "ROUND(R57,0)"
'                .SetText .ColLetterToNumber("R"), 56, varTemp1
'            End If
'        End If
'        If NewSheet = 4 Then
'            .Sheet = 1
'            .GetText .ColLetterToNumber("I"), 48, varTemp1
'            .SetText .ColLetterToNumber("Q"), 48, varTemp1
'            .Sheet = NewSheet
'        End If
    End With
End Sub


' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 20
    denRowDL = 25
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "10_12_2~10_12_3~18_20_4~13_15_5"
    ' end
End Sub


' Ham kiem tra thang Mien giam thuoc 7/2012 - 12/2012
Private Function isMienGiamPL26(ByVal tuThang As String, ByVal denThang As String) As Boolean
    Dim arrThang() As String
    Dim result As Boolean
    arrThang = Split(tuThang, "/")
    result = False
    ' kiem tra tu thang thuoc 7-12/2012
    If Val(arrThang(1)) = 2012 Then
        If Val(arrThang(0)) >= 7 And Val(arrThang(0)) <= 12 Then
            result = True
        End If
    End If
    
    ' kiem tra den thang thuoc 7-12/2012
    arrThang = Split(denThang, "/")
     If Val(arrThang(1)) = 2012 Then
        If Val(arrThang(0)) >= 7 And Val(arrThang(0)) <= 12 Then
            result = True
        End If
    End If
    
    isMienGiamPL26 = result
End Function



Private Sub tinhGiamTruNPT(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
    Dim tuThang As Variant
    Dim denThang As Variant
    Dim arrTuThang() As String
    Dim arrDenThang() As String
    
    Dim group As Variant
    Dim colTuThangBK As String, colDenThangBK As String
    Dim colSoThangBK As String, colSoTienBK As String
    
    With fps
        If .ActiveSheet = phuLuc09_3_TNCN Then
            .Sheet = phuLuc09_3_TNCN
            
            ' kiem tra group
            .Col = .ColLetterToNumber("AE")
            .GetText .Col, .Row, group
            
            colSoThangBK = "Y"
            colSoTienBK = "Z"
            If group = "1" Then
                colTuThangBK = "Q"
                colDenThangBK = "S"
            ElseIf group = "2" Then
                colTuThangBK = "T"
                colDenThangBK = "U"
            End If
            
            If (Col = .ColLetterToNumber(colTuThangBK) Or Col = .ColLetterToNumber(colDenThangBK)) And Row >= 24 Then
                .Col = .ColLetterToNumber(colTuThangBK)
                .Row = Row
                .GetText .Col, .Row, tuThang
                .Col = .ColLetterToNumber(colDenThangBK)
                .GetText .Col, .Row, denThang
                If tuThang <> vbNullString And denThang <> vbNullString And tuThang <> "../...." And denThang <> "../...." And Format_mmyyyy(CStr(denThang)) <> "" And Format_mmyyyy(CStr(tuThang)) <> "" Then
                    .Col = .ColLetterToNumber(colSoThangBK)
                    .Row = Row
                    .Text = DateDiff("M", Format(Format_mmyyyy(CStr(tuThang)), "mm/yyyy"), Format(Format_mmyyyy(CStr(denThang)), "mm/yyyy")) + 1
                    
                    .GetText .ColLetterToNumber(colSoThangBK), Row, soThang
                    
                    arrTuThang = Split(tuThang, "/")
                    arrDenThang = Split(denThang, "/")
                    
                    ' neu den thang < 7/2013
                    If Val(arrDenThang(1)) < 2013 Or (Val(arrDenThang(1)) = 2013 And Val(arrDenThang(0)) < 7) Then
                        .Row = Row
                        .Col = .ColLetterToNumber(colSoTienBK)
                        .value = 1600000 * Val(soThang)
                    ElseIf (Val(arrTuThang(1)) > 2013) Or ((Val(arrTuThang(1)) = 2013) And (Val(arrTuThang(0)) >= 7)) Then
                        .Row = Row
                        .Col = .ColLetterToNumber(colSoTienBK)
                        .value = 3600000 * Val(soThang)
                    Else
                        soThang1 = DateDiff("M", Format_mmyyyy(CStr(tuThang)), Format_mmyyyy("06/2013")) + 1
                        soThang2 = DateDiff("M", Format_mmyyyy("07/2013"), Format_mmyyyy(CStr(denThang))) + 1
                        .Row = Row
                        .Col = .ColLetterToNumber(colSoTienBK)
                        .value = 1600000 * soThang1 + 3600000 * soThang2
                    End If
                    tuThang = ""
                    denThang = ""
                Else
                    .Row = Row
                    .Col = .ColLetterToNumber(colSoTienBK)
                    .value = 0
                End If
            End If
        End If
    End With
End Sub





Private Sub tinhGiamTruNPT_Load(ByVal Row As Long)
    Dim tuThang As Variant
    Dim denThang As Variant
    Dim arrTuThang() As String
    Dim arrDenThang() As String
    Dim group As Variant
    Dim colTuThangBK As String, colDenThangBK As String
    Dim colSoThangBK As String, colSoTienBK As String
    
    With fps
            .Sheet = phuLuc09_3_TNCN
            
            ' kiem tra group
            .Col = .ColLetterToNumber("AE")
            .GetText .Col, .Row, group
            
            colSoThangBK = "Y"
            colSoTienBK = "Z"
            If group = "1" Then
                colTuThangBK = "Q"
                colDenThangBK = "S"
            ElseIf group = "2" Then
                colTuThangBK = "T"
                colDenThangBK = "U"
            End If
            
            .Col = .ColLetterToNumber(colTuThangBK)
            .Row = Row
            .GetText .Col, .Row, tuThang
            .Col = .ColLetterToNumber(colDenThangBK)
            .GetText .Col, .Row, denThang
            
            
            If tuThang <> vbNullString And denThang <> vbNullString And tuThang <> "../...." And denThang <> "../...." And Format_mmyyyy(CStr(denThang)) <> "" And Format_mmyyyy(CStr(tuThang)) <> "" Then
                .Col = .ColLetterToNumber(colSoThangBK)
                .Row = Row
                .Text = DateDiff("M", Format(CStr(tuThang), "mm/yyyy"), Format(CStr(denThang), "mm/yyyy")) + 1
                
                .GetText .ColLetterToNumber(colSoThangBK), Row, soThang
                
                arrTuThang = Split(tuThang, "/")
                arrDenThang = Split(denThang, "/")
                
                ' neu den thang < 7/2013
                If Val(arrDenThang(1)) < 2013 Or (Val(arrDenThang(1)) = 2013 And Val(arrDenThang(0)) < 7) Then
                    .Row = Row
                    .Col = .ColLetterToNumber(colSoTienBK)
                    .value = 1600000 * Val(soThang)
                ElseIf (Val(arrTuThang(1)) > 2013) Or ((Val(arrTuThang(1)) = 2013) And (Val(arrTuThang(0)) >= 7)) Then
                    .Row = Row
                    .Col = .ColLetterToNumber(colSoTienBK)
                    .value = 3600000 * Val(soThang)
                Else
                    soThang1 = DateDiff("M", Format_mmyyyy(CStr(tuThang)), Format_mmyyyy("06/2013")) + 1
                    soThang2 = DateDiff("M", Format_mmyyyy("07/2013"), Format_mmyyyy(CStr(denThang))) + 1
                    .Row = Row
                    .Col = .ColLetterToNumber(colSoTienBK)
                    .value = 1600000 * soThang1 + 3600000 * soThang2
                End If
                tuThang = ""
                denThang = ""
            Else
                .Row = Row
                .Col = .ColLetterToNumber(colSoTienBK)
                .value = 0
            End If
    End With
End Sub


' ham set thong tin sau khi tai XML
Public Sub finishImportXML()
    Dim varTemp As Variant
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = phuLuc09_3_TNCN
        .Row = 25
        Do
            .Col = .ColLetterToNumber("B")
            tinhGiamTruNPT_Load .Row
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
        
        .Row = .Row + 7
        Do
            .Col = .ColLetterToNumber("B")
            tinhGiamTruNPT_Load .Row
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
'        .Sheet = 1
'        mCurrentSheet = .Sheet
'        .Row = 47
'        .Col = .ColLetterToNumber("K")
'        .Formula = "'09-3'!AA23"
'        UpdateCell fps, .Col, .Row, .Text
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
