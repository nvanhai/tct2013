VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01BTNDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Thu Nhap Doanh Nghiep Tam tinh" interface sheets
'this Class is belong to TAX_Business_v1 project which will be compline to DLL

Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Private vFormular As String
Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
'Dim xNodeData0 As MSXML.IXMLDOMNode
Dim xNode As MSXML.IXMLDOMNode
Dim arrData() As String
Public loaiKHBS As String
Public FlagThaydoi As String
Public Strloaitk As String
Public StrSolanBosung As String
Public strLoaiNNKD As String
Private Const Row_MSTDN = 12
Private Const Col_MSTDN = "E"
Private Const rowStart = 35

' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

Public hanNopTKTNDN As String

 
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
        Dim strComboCQT As String
        Dim strComboCQT_Hide As String
        Dim strDataFileName1 As String
        Dim arrDanhsach() As String
        Dim strComboHien As String
        Dim xmlDomDataCQT As New MSXML.DOMDocument, xmlDomCurrentData1 As New MSXML.DOMDocument
        Dim xmlNodeListItemsCQT As MSXML.IXMLDOMNodeList
        Dim xmlNodeCQT As MSXML.IXMLDOMNode
        Dim fldListCQT() As String
        
        
    With fps
        SetDateFormat fps, 1, 69, .ColLetterToNumber("J"), DDMMYYYY
        SetDateFormat fps, 1, 60, .ColLetterToNumber("H"), DDMMYYYY
        .Sheet = 1
        .Row = 69
        .Col = .ColLetterToNumber("J")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
          'dhdang sua check gia han nop theo thong tu
         ' q Quy ke khai
        ' y nam ke khai
        ' dNgayTaiChinh ngay tai chinh lay tren man hinh HTKK
        ' dThangTaiChinh thang tai chinh tren phan thong tin chung HTKK
        ' dType: 0 tra ve quy, 1 tra ve nam
        Dim chkValue As Variant
        
        Dim Quynamtc As Integer
        Dim Namtc As Integer
        Dim QuyKK As Integer
        Dim NamKK As Integer
        Dim iNgayTaiChinh As Integer
        Dim iThangTaiChinh As Integer
        Dim strNgayTaiChinh As String
        Dim Thongtu As String
        
        Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
        Dim strDataFileName As String
        Dim xmlNodeListItems As MSXML.IXMLDOMNodeList
        Dim strCboLyDo As String
        Dim strCboLyDoAn As String
        Dim xmlNode As MSXML.IXMLDOMNode
        Dim i As Integer
        Dim fldList() As String
                

        QuyKK = TAX_Utilities_v1.ThreeMonths
        NamKK = TAX_Utilities_v1.Year
        
        
        'Lay ngay bat dau nam tai chinh
        If GetAttribute(TAX_Utilities_v1.NodeMenu, "FinanceYear") = "1" Then
            strNgayTaiChinh = GetNgayBatDauNamTaiChinh
                iNgayTaiChinh = GetNgayTaiChinh(strNgayTaiChinh)
                iThangTaiChinh = GetThangTaiChinh(strNgayTaiChinh)
        Else
            strNgayTaiChinh = "01/01"
            iNgayTaiChinh = 1
            iThangTaiChinh = 1
        End If
        
        Quynamtc = GetQuyNamTaiChinh(QuyKK, NamKK, iNgayTaiChinh, iThangTaiChinh, 0)
        Namtc = GetQuyNamTaiChinh(QuyKK, NamKK, iNgayTaiChinh, iThangTaiChinh, 1)
'        If (Namtc = 2010 And Quynamtc = 4) Then
'            Thongtu = GetAttribute(GetMessageCellById("0187"), "Msg")
'            .SetText .ColLetterToNumber("C"), 48, Thongtu
'        ElseIf Namtc = 2011 And (Quynamtc = 1 Or Quynamtc = 2) Then
'            Thongtu = GetAttribute(GetMessageCellById("0260"), "Msg")
'            .SetText .ColLetterToNumber("C"), 48, Thongtu
'        ElseIf Namtc = 2011 And (Quynamtc = 3 Or Quynamtc = 4) Then
'            Thongtu = GetAttribute(GetMessageCellById("0243"), "Msg")
'            .SetText .ColLetterToNumber("C"), 48, Thongtu
'        ElseIf Namtc = 2010 And Quynamtc <> 4 Then
'            Thongtu = GetAttribute(GetMessageCellById("0186"), "Msg")
'            .SetText .ColLetterToNumber("C"), 48, Thongtu
'        ElseIf Namtc = 2009 Then
'            Thongtu = GetAttribute(GetMessageCellById("0185"), "Msg")
'            .SetText .ColLetterToNumber("C"), 48, Thongtu
''        ElseIf Namtc = 2013 And (Quynamtc = 1 Or Quynamtc = 2 Or Quynamtc = 3) Then
''            Thongtu = GetAttribute(GetMessageCellById("0263"), "Msg")
''            .SetText .ColLetterToNumber("C"), 48, Thongtu
'        Else
''            .SetText .ColLetterToNumber("C"), 48, ""
''            .Row = 48
''            .Col = .ColLetterToNumber("B")
''            .Text = ""
''            .Lock = True
'            Thongtu = GetAttribute(GetMessageCellById("0263"), "Msg")
'            .SetText .ColLetterToNumber("C"), 48, Thongtu
            
            ' Ky dc gia han moi set gia tri ly do gia han
            strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_DM_Lydo_GH_TT16.xml")
            If xmlDomData.Load(strDataFileName) Then
                 Set xmlNodeListItems = xmlDomData.getElementsByTagName("Item")
                 For Each xmlNode In xmlNodeListItems
                     fldList = Split(GetAttribute(xmlNode, "Value"), "###")
                     strCboLyDo = strCboLyDo + fldList(1) + Chr$(9)
                     strCboLyDoAn = strCboLyDoAn + fldList(0) + Chr$(9)
                 Next
             End If
                  'set ly do gia han
            .Row = 58
            .Col = .ColLetterToNumber("H")
            .TypeComboBoxList = strCboLyDo
            ' set Ma ly do gia han
            .Col = .ColLetterToNumber("P")
            .TypeComboBoxList = strCboLyDoAn
        'End If
        '-------------------------------------------------------------------------------
        'SetDateFormat fps, 2, 9, .ColLetterToNumber("D"), DDMMYYYY
        
        SetDateFormat fps, .SheetCount - 1, 31, .ColLetterToNumber("BE"), DDMMYYYY
        SetDateFormat fps, .SheetCount - 1, 46, .ColLetterToNumber("BG"), DDMMYYYY
        .Sheet = .SheetCount - 1 'KHBS
        .Row = 46
        .Col = .ColLetterToNumber("BG")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        ' reset cong thuc
'        If TAX_Utilities_v1.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0" Then
'            resetFomula
'        End If
                
    End With
     ' set thong tin CQT quan ly
    setCQTQuanLy
        ' set CQT ra quyet dinh hoan
    setCQTQuanLyHoanThue fps
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim chkValue As Variant
    Dim i As Long
    Dim intCtrl As Integer
    Dim strId As Variant
    Dim xmlNode As MSXML.IXMLDOMNode
    With fps
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        'set to khai bo xung hay chinh thuc
            If Strloaitk = "TKCT" Then
                .Col = .ColLetterToNumber("E")
                .Row = 14
                .Text = 1
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("G")
                .Row = 14
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
             ElseIf Strloaitk = "TKBS" Then
                .Col = .ColLetterToNumber("E")
                .Row = 14
                .Text = 0
                  UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("G")
                .Row = 14
                .Text = Val(StrSolanBosung)
                UpdateCell fps, .Col, .Row, .Text
            End If
        'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
            'ma NNKD
        .Col = .ColLetterToNumber("L")
        .Row = 14
        .Text = strLoaiNNKD
        UpdateCell fps, .Col, .Row, .Text
         'end
        ' lock 2 chi tieu 27, 28
        .Col = .ColLetterToNumber("K")
        .Row = 48
        .Lock = True
        .Row = 49
        .Lock = True
        'end
        'unlock chi tieu 30
        .Col = .ColLetterToNumber("K")
        .Row = 51
        .Lock = False
        'end
'        .SetText .ColLetterToNumber("E"), 7, 0
        .GetText .ColLetterToNumber("H"), 47, chkValue
        .Col = .ColLetterToNumber("K")
        .Row = 50
        vFormular = .Formula
        If chkValue = "0" Then
            .Col = .ColLetterToNumber("K")
            .Row = 47
            If .CellType = 1 Or .CellType = 5 Then
                FormatTextPercent fps, 1, .ColLetterToNumber("K"), 47, True
                '.SetText .ColLetterToNumber("K"), .Row, "0"
            End If
            .Col = .ColLetterToNumber("K")
            .Row = 50
            .Lock = True
            .Formula = "ROUND(K41*K42/100*K47/100,0)"
        ElseIf chkValue = "1" Or chkValue = "x" Then
            'FormatText fps, 1, .ColLetterToNumber("J"), 42
            .Col = .ColLetterToNumber("K")
            .Row = 50
            .Lock = False
        Else
            FormatTextPercent fps, 1, .ColLetterToNumber("K"), 47, False
            .Col = .ColLetterToNumber("K")
            .Row = 50
            .Lock = True
            .Formula = "ROUND(K41*K42/100*K47/100,0)"
        End If
        
        
        If Not TAX_Utilities_v1.DataKHBS Is Nothing Then
            Set xmlNode = TAX_Utilities_v1.DataKHBS.nodeFromID(GetCellID(fps, fps.ColLetterToNumber("J"), 47))
            vThuephainop = GetAttribute(xmlNode, "Value")
            Set xmlNode = Nothing
          
          .Sheet = .SheetCount - 1 'KHBS
            Dim cCol As Long, cRow As Long
          '  .ActiveSheet = .Sheet
            Set xmlNode = TAX_Utilities_v1.Data(.SheetCount - 2).childNodes(2).childNodes(2).lastChild.lastChild
            mCurrentSheet = .SheetCount - 1
            fps.Sheet = mCurrentSheet
            If GetAttribute(xmlNode, "Value") <> "" Then
                If GetAttribute(xmlNode, "Value") < 0 Then
                    ParserCellID fps, GetAttribute(xmlNode, "CellID"), cCol, cRow
                    .Col = .ColLetterToNumber("BE")
                    .Row = cRow + 3
                    .value = 0
                    UpdateCell fps, .Col, .Row, .value
                End If
            End If
            i = 1
            .Col = .ColLetterToNumber("B")
            .Row = 9
            Do
              .GetText .ColLetterToNumber("BC"), .Row, strId
              If strId <> "" Then
                .Col = .ColLetterToNumber("BF")
                .CellTag = .Row & "~" & Trim(Mid(strId, 100, 20))
              End If
              .Col = .ColLetterToNumber("B")
              .Row = .Row + 1
            Loop Until .Text = "bb"
            Set xmlNode = TAX_Utilities_v1.Data(.SheetCount - 2).childNodes(2).firstChild
            SetAttribute xmlNode, "DateKHBS", TAX_Utilities_v1.DateKHBS
            SetAttribute xmlNode, "loaiKHBS", loaiKHBS
            If loaiKHBS = "frmKHBS_BS" Then
                .SetText .ColLetterToNumber("B"), 5, GetAttribute(GetMessageCellById("0115"), "Msg") & _
                GetAttribute(GetMessageCellById("0117"), "Msg") & Mid(TAX_Utilities_v1.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 5, 4)
            Else
                .SetText .ColLetterToNumber("B"), 5, GetAttribute(GetMessageCellById("0116"), "Msg") & _
                GetAttribute(GetMessageCellById("0117"), "Msg") & Mid(TAX_Utilities_v1.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 5, 4)
            End If
                .EventEnabled(EventAllEvents) = False
                .Sheet = 1
                mCurrentSheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 24
                .Formula = "TRIM(Header!T5)"
                UpdateCell fps, .Col, .Row, .Text
                .Sheet = 1 'To khai GTGT
                .Row = 26
                .Col = .ColLetterToNumber("I")
                .Text = Format(Date, "dd/mm/yyyy")
                .TypeHAlign = TypeHAlignLeft
                .EventEnabled(EventAllEvents) = True
            ReDim arrData(TAX_Utilities_v1.DataKHBS.getElementsByTagName("Cell").length)
            Dim j As Integer
            For Each xNode In TAX_Utilities_v1.DataKHBS.getElementsByTagName("Cell")
                arrData(j) = GetAttribute(xNode, "CellID") & ":" & GetAttribute(xNode, "Value")
                j = j + 1
            Next
            Set xNode = Nothing
         End If
         
         
         ' Neu check gia han moi cho phep chon GT
         .Sheet = 1
        .Col = .ColLetterToNumber("F")
        .Row = 56
        If .Text = "1" Or UCase$(.Text) = "X" Then
            '
            .Col = .ColLetterToNumber("H")
            .Row = 56
            .Lock = False
            '
            .Col = .ColLetterToNumber("H")
            .Row = 58
            .Lock = False
            '
            .Col = .ColLetterToNumber("H")
            .Row = 60
            .Lock = False
            '
            .Col = .ColLetterToNumber("H")
            .Row = 62
            .Lock = False
        Else
            .Col = .ColLetterToNumber("H")
            .Row = 56
            .Lock = True
            '
            .Col = .ColLetterToNumber("H")
            .Row = 58
            .Lock = True
            '
            .Col = .ColLetterToNumber("H")
            .Row = 60
            .Lock = True
            '
            .Col = .ColLetterToNumber("H")
            .Row = 62
            .Lock = True
        End If
        ' End set gia han
         
         'set lai stt cho cac cot dong sheet 2
         .Sheet = 2
         If .SheetVisible = True Then
                    i = 1
                   .Col = .ColLetterToNumber("C")
                    .Row = 35
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("C")
                         .Row = i + 35
                         i = i + 1
                    Loop Until .Text = "aa"
            End If
        
'        If TAX_Utilities_v1.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0" Then
'            setFomula
'        End If
         'end
         
         ' set lock CQT phan bo
        .Sheet = 3
        Dim vChon As Variant
        If .SheetVisible = True Then
        .Row = 42
         Do
                .GetText .ColLetterToNumber("G"), .Row, vChon
                If vChon = 1 Then
                    .Col = .ColLetterToNumber("P")
                    .Lock = True
                End If
                
                .Row = .Row + 1
                .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
         End If
    End With
End Sub

Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim iCol As Long, iRow As Long, tCol As Long, tRow As Long
    Dim i As Long
    Dim lResult As String
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
    
    Dim blTemp As Boolean
    
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .ActiveSheet
        Select Case mCurrentSheet
        Case 1
            'Thuc hien khi chon Checkbox
             ' neu check thue suat 20 % thi set chi tieu 25 = 20%
'            If (Col = fps.ColLetterToNumber("E") And Row = 6) Then
'                fps.Sheet = fps.ActiveSheet
'                fps.Col = Col
'                fps.Row = Row
'                If (fps.value = 1 Or UCase$(fps.value) = "X") Then
'                    fps.Row = 41
'                    fps.Col = fps.ColLetterToNumber("J")
'                    fps.Text = 20
'                    fps.Lock = True
'                    UpdateCell fps, fps.Col, fps.Row, fps.Text
'                Else
'                    fps.Row = 41
'                    fps.Col = fps.ColLetterToNumber("J")
'                    fps.Text = 0
'                    fps.Lock = False
'                    UpdateCell fps, fps.Col, fps.Row, fps.Text
'                End If
'            End If
            
            'xu ly voi checkbox doanh nghiep co hach toan phu thuoc
            If (Col = .ColLetterToNumber("F") And Row = 7) Or (Col = .ColLetterToNumber("F") And Row = 6) Then
                .Sheet = .ActiveSheet
                .Row = Row
                .Col = Col
                UpdateCell fps, .Col, .Row, IIf(.value = 1, "x", vbNullString)
            End If
            'end
            ' neu check thue suat 20 % thi set chi tieu 25 = 20%
'            If (Col = .ColLetterToNumber("E") And Row = 6) Then
'                .Sheet = .ActiveSheet
'                .Col = Col
'                .Row = Row
'                If (.value = 1 Or UCase$(.value) = "X") Then
'                    .Row = 41
'                    .Col = .ColLetterToNumber("J")
'                    .Text = 20
'                    UpdateCell fps, .Col, .Row, .Text
'                Else
'                    .Row = 41
'                    .Col = .ColLetterToNumber("J")
'                    .Text = 0
'                    UpdateCell fps, .Col, .Row, .Text
'                End If
'            End If
            
            'xu ly voi checkbox gia han
            If Col = .ColLetterToNumber("F") And Row = 56 Then
                .Sheet = .ActiveSheet
                .Row = Row
                .Col = Col
                UpdateCell fps, .Col, .Row, IIf(.value = 1, "x", vbNullString)
            End If
            'end
            If Col = .ColLetterToNumber("H") And Row = 47 Then
                .Sheet = .ActiveSheet
                .Row = Row
                .Col = Col
                If .value = 1 Then 'Khi danh dau Checkbox
                    'Hien thong bao co muon xoa du lieu cu
                    lResult = DisplayMessage("0111", msYesNo, miQuestion, , mrNo)
                    If lResult = 3 Then 'Khong muon xoa thi thoat ra
                        .value = 0
                        UpdateCell fps, .Col, .Row, IIf(.value = 1, "x", vbNullString)
                        
                        .EventEnabled(EventAllEvents) = True
                        Exit Sub
                    Else 'Khi chon xoa
                        'Xoa du lieu chi tieu 15, 18 va cho phep nhap cot 18, format lai cot 15
                        
                        Set xmlCellNode = TAX_Utilities_v1.Data(0).nodeFromID(GetCellID(fps, .ColLetterToNumber("K"), 47))
                        Set xmlCellsNode = xmlCellNode.parentNode
                        UpdateCell fps, .ColLetterToNumber("H"), .Row, "x"
                        Do
                            For i = 1 To xmlCellsNode.childNodes.length
                                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
                                Select Case iCol
                                    Case .ColLetterToNumber("K")
                                        'Xoa trang chi tieu 26
                                        .Row = 47
                                        .Col = iCol
                                        .Text = ""
                                        UpdateCell fps, iCol, 47, .Text
                                        FormatText fps, 1, .ColLetterToNumber("K"), 47
                                        
                                        'Xoa trang chi tieu 29
                                        .Row = 50
                                        .Col = iCol
                                        .Formula = ""
                                        .Text = "0"
                                        .value = "0"
                                        .Lock = False ' Bo enable cho phep nhap
                                        UpdateCell fps, iCol, 50, .value
                                End Select
                            Next
                            If xmlCellsNode.nextSibling Is Nothing Then Exit Do
                            Set xmlCellsNode = xmlCellsNode.nextSibling
                        Loop Until xmlCellsNode Is Nothing
                    End If
                ElseIf .value = 0 Then 'Huy danh dau Checkbox
                        
                    Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("K"), 47))
                    Set xmlCellsNode = xmlCellNode.parentNode
                    UpdateCell fps, .ColLetterToNumber("H"), .Row, ""
                    Do
                        For i = 1 To xmlCellsNode.childNodes.length
                            ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
                            Select Case iCol
                                Case .ColLetterToNumber("K")
                                    ' Dinh dang lai cho chi tieu 15 (Cho phep nhap nhieu muc thue)
                                    .Row = 47
                                    .Col = iCol
                                    .value = 0
                                    UpdateCell fps, iCol, 47, .value
                                    FormatTextPercent fps, 1, .ColLetterToNumber("K"), 47, False
                                
                                    ' Thiet lap lai cong thuc cho chi tieu 18
                                    .Row = 50
                                    .Col = iCol
                                    .Formula = "ROUND(K41*K42/100*K47/100,0)"
'                                    .Formula = vFormular
                                    .Lock = True
                                    UpdateCell fps, iCol, 50, .value
                            End Select
                        Next
                        If xmlCellsNode.nextSibling Is Nothing Then Exit Do
                        Set xmlCellsNode = xmlCellsNode.nextSibling
                    Loop Until xmlCellsNode Is Nothing
                End If
            End If
        TAX_Utilities_v1.AdjustData(.Sheet) = True
        End Select
        .EventEnabled(EventAllEvents) = True
    End With
                    
    fps.EventEnabled(EventAllEvents) = False
    ' set gia lock cell ly do gia han
    If fps.ActiveSheet = 1 Then
            ' Neu check gia han moi cho phep chon GT
            If Col = fps.ColLetterToNumber("F") And Row = 56 Then
                fps.Col = fps.ColLetterToNumber("F")
                fps.Row = 56
                If fps.Text = "1" Or UCase$(fps.Text) = "X" Then
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 58
                    fps.Lock = False
                    '
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 60
                    fps.Lock = False
                    '
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 62
                    fps.Lock = False
                    '
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 64
                    fps.Lock = False
                    fps.Formula = "K54-H62"
                Else
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 58
                    fps.Lock = True
                    fps.SetText fps.ColLetterToNumber("H"), 58, ""
                    fps.SetText fps.ColLetterToNumber("P"), 58, "00"
                    UpdateCell fps, fps.ColLetterToNumber("P"), 58, "00"
                    UpdateCell fps, fps.ColLetterToNumber("H"), 58, ""
                    '
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 60
                    fps.Lock = True
                    fps.SetText fps.ColLetterToNumber("H"), 60, ""
                    UpdateCell fps, fps.ColLetterToNumber("H"), 60, ""
                    '
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 62
                    fps.Lock = True
                    fps.SetText fps.ColLetterToNumber("H"), 62, "0"
                    UpdateCell fps, fps.ColLetterToNumber("H"), 62, "0"
                    '
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 64
                    fps.Lock = True
                    fps.Formula = ""
                    fps.SetText fps.ColLetterToNumber("H"), 64, "0"
                    UpdateCell fps, fps.ColLetterToNumber("H"), 64, "0"
                End If
            End If
    End If
    
    With fps

            If Col = .ColLetterToNumber("G") And Row >= 42 And .ActiveSheet = 3 Then
                   fps.EventEnabled(EventAllEvents) = False
                   .Sheet = 3
                   .Row = Row
                   .Col = .ColLetterToNumber("G")
                   If (.value = "1") Then
                       .Row = Row
                       ' chi tieu STT
                       .Col = .ColLetterToNumber("P")
                       .Lock = True
                       
                       .Text = ""
                       UpdateCell fps, .Col, .Row, .Text
                       
                       .Col = .ColLetterToNumber("AH")
                       .Text = ""
                       UpdateCell fps, .Col, .Row, .Text
                   Else
                       .Row = Row
                       ' chi tieu STT
                       .Col = .ColLetterToNumber("P")
                       .Lock = False
                      
                   End If
                   fps.EventEnabled(EventAllEvents) = True
               End If
    End With
    
    fps.EventEnabled(EventAllEvents) = True
    
End Sub

Public Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
Dim varTemp As Variant
Dim strId As Variant
Dim iCol As Long, iRow As Long
Dim cCol As Long, cRow As Long
Dim cValue As Variant, temp As Variant
Dim cValueG As Variant
Dim cTagF As String
Dim cValueF As Variant
Dim vO34 As Variant, vP34 As Variant, vQ34 As Variant, vR34 As Variant, vTLPB As Variant
Dim i As Integer
Dim xmlNodeH15 As MSXML.IXMLDOMNode
With fps
    If .ActiveSheet = .SheetCount - 1 Then
         fps.EventEnabled(EventAllEvents) = False
        GetCellSpan fps, Col, Row
        If blGiaTriCu And (Col = .ColLetterToNumber("BD") Or Col = .ColLetterToNumber("BE")) Then
                 DisplayMessage "0110", msOKOnly, miWarning
                .Row = RowGiaTriCu
                .Col = .ColLetterToNumber("BD")
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("BE")
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("BF")
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("BG")
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("BH")
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                blGiaTriCu = False
                fps.SetFocus
                fps.EventEnabled(EventAllEvents) = True
                Exit Sub
         End If
         
         'get gia tri da ke khai
        .GetText .ColLetterToNumber("BC"), Row, strId
          If strId <> "" Then
                Dim arr() As String
                 For i = 0 To UBound(arrData)
                     arr() = Split(arrData(i), ":")
                     If arr(0) = Trim(Mid(strId, 100, 20)) Then
                         cValue = arr(1)
                         Exit For
                     End If
                 Next
               .Col = .ColLetterToNumber("BF")
               .Row = Row
               cTagF = .CellTag
               If cTagF <> "" Then
                         For i = 0 To UBound(arrData)
                               arr() = Split(arrData(i), ":")
                               If arr(0) = Right(cTagF, Len(cTagF) - InStr(cTagF, "~")) Then
                                   cValueF = arr(1)
                                   Exit For
                               End If
                           Next
                End If
               .CellTag = .Row & "~" & Trim(Mid(strId, 100, 20))
               .value = cValue
               UpdateCell fps, .Col, .Row, .value
               .Col = .ColLetterToNumber("BG")
               .Row = Row
               UpdateCell fps, .Col, .Row, .value
               .Col = .ColLetterToNumber("BH")
               .Row = Row
               UpdateCell fps, .Col, .Row, .value
                
               If loaiKHBS = "frmKHBS_TT" Then
                    .GetText .ColLetterToNumber("BG"), Row, cValueG
               Else
                    .GetText .ColLetterToNumber("BH"), Row, cValueG
               End If
              
              If cValueG <> "" Then
                      ParserCellID fps, Trim(Mid(strId, 100, 20)), cCol, cRow
                      .Sheet = 1
                      .SetText cCol, cRow, cValueG
                      If loaiKHBS = "frmKHBS_TT" Then
                        UpdateKHBSCell fps, cCol, cRow, cValueG
                      Else
                        UpdateKHBSCell fps, cCol, cRow, cValueG + cValueF
                      End If
              End If
              
' Set lai gia tri cu cua sheet 01
               .Sheet = .SheetCount - 1
               .Col = .ColLetterToNumber("BF")
               .Row = Row
               
               If cTagF <> "" And cTagF <> .CellTag Then
                  If Left(cTagF, InStr(cTagF, "~") - 1) = .Row Then
                   ' .GetText .ColLetterToNumber("F"), Row, temp
                    ParserCellID fps, Right(cTagF, Len(cTagF) - InStr(cTagF, "~")), iCol, iRow
                    UpdateKHBSCell fps, iCol, iRow, cValueF
                    .Sheet = 1
                    If loaiKHBS = "frmKHBS_TT" Then
                        .SetText iCol, iRow, cValueF
                    Else
                        .SetText iCol, iRow, 0
                    End If
                    .Sheet = .SheetCount - 1
                    If loaiKHBS = "frmKHBS_TT" Then
                        .GetText .ColLetterToNumber("BG"), Row, varTemp
                    Else
                        .GetText .ColLetterToNumber("BH"), Row, varTemp
                    End If
                    If varTemp <> "" Then
                        ParserCellID fps, Trim(Mid(strId, 100, 20)), iCol, iRow
                        If loaiKHBS = "frmKHBS_TT" Then
                            UpdateKHBSCell fps, cCol, cRow, varTemp
                        Else
                            UpdateKHBSCell fps, cCol, cRow, varTemp + cValueF
                        End If
                       ' UpdateKHBSCell fps, iCol, iRow, varTemp
                        .Sheet = 1
                        .SetText iCol, iRow, varTemp
                        .Sheet = .SheetCount - 1
                    End If
                   End If
               End If
'Xoa dong
               If mblKeyF6 = True Then
                ParserCellID fps, Right(cTagF, Len(cTagF) - InStr(cTagF, "~")), iCol, iRow
                UpdateKHBSCell fps, iCol, iRow, cValueF
                .Sheet = 1
                .SetText iCol, iRow, 0
                .Sheet = .SheetCount - 1
                .Col = .ColLetterToNumber("BF")
                .Row = Row
                .value = 0
                .Col = .ColLetterToNumber("BG")
                .Row = Row
                .value = 0
                .Col = .ColLetterToNumber("BH")
                .Row = Row
                .value = 0
               End If
              
            .Sheet = 1
            .GetText .ColLetterToNumber("K"), 54, varTemp
            .Sheet = .SheetCount - 1
            Set xmlNodeH15 = TAX_Utilities_v1.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
                 ParserCellID fps, GetAttribute(xmlNodeH15, "CellID"), cCol, cRow
                .Col = cCol
                .Row = cRow
                If loaiKHBS = "frmKHBS_TT" Then
                    .value = CDbl(varTemp) - CDbl(vThuephainop)
                     UpdateCell fps, .Col, .Row, .value
                Else
                    .value = CDbl(varTemp)
                     UpdateCell fps, .Col, .Row, .value
                End If
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
             Set xmlNodeH15 = Nothing
              
          End If
            fps.EventEnabled(EventAllEvents) = True
       End If
     mblKhoiDauToKhai = False
   End With
   setChangeComb Col, Row
End Sub


Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, iCol As Long, iRow As Long
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tCol As Long, tRow As Long
Dim strId As Variant, Col7 As Variant, intIndexCombo As Integer, cRow As Integer
Dim NguoiKyCol As Long, NguoiKyRow As Long
Dim vO34 As Variant, vP34 As Variant, vQ34 As Variant, vR34 As Variant
    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 2 Then
                   fps.EventEnabled(EventAllEvents) = True
'                   .GetText .ColLetterToNumber("O"), 34, vO34
'                   .GetText .ColLetterToNumber("P"), 34, vP34
'                   .GetText .ColLetterToNumber("Q"), 34, vQ34
'                   .GetText .ColLetterToNumber("R"), 34, vR34
                    
                    i = 1
                   .Col = .ColLetterToNumber("C")
                    .Row = 35
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("C")
                         .Row = i + 35
                         i = i + 1
                    Loop Until .Text = "aa"
                   
                    .Row = iRow
                    .Col = iCol
                     fps.EventEnabled(EventAllEvents) = False
                     .SetActiveCell iCol, iRow
                   
            End If
            
            ' set lock CQT phan bo
            If .ActiveSheet = 3 Then
                .Sheet = 3
                Dim vChon As Variant
                If .SheetVisible = True Then
                .Row = 42
                 Do
                        .GetText .ColLetterToNumber("G"), .Row, vChon
                        If vChon = 1 Then
                            .Col = .ColLetterToNumber("P")
                            .Lock = True
                        Else
                            .Col = .ColLetterToNumber("P")
                            .Lock = False
                        End If
                        
                        .Row = .Row + 1
                        .Col = .ColLetterToNumber("B")
                    Loop Until .Text = "aa"
                 End If
            End If
            
            If .ActiveSheet = .SheetCount - 1 Then
                 fps_Change iCol, iRow
                 fps.EventEnabled(EventAllEvents) = False
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 9
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 9
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 3
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
                 fps.EventEnabled(EventAllEvents) = True
            End If
         End If
        .EventEnabled(EventAllEvents) = True
    End With
    
End Sub

Private Sub fps_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim i As Integer, iCol As Long, iRow As Long
     With fps
       If KeyCode = vbKeyF6 Then
            iCol = .ActiveCol
            iRow = .ActiveRow
            mblKeyF6 = True
            fps_Change iCol, iRow
            mblKeyF6 = False
        End If
     End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim varCT11 As Variant
    Dim varCT12 As Variant
    Dim tCol As Long, tRow As Long
    Dim cCol As Long, cRow As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    With fps
    .EventEnabled(EventAllEvents) = False
    mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            ' Neu chi tieu [16] nho hon khong thi thiet lap chi tieu 19,20=0
            'dhdang sua cho ca truong hop kk bo sung
            .GetText .ColLetterToNumber("K"), 48, varTemp
            If loaiKHBS <> "frmKHBS_BS" Then
                If varTemp <= 0 Then
                    .Row = 51
                    .Col = .ColLetterToNumber("K")
                    .Text = "0"
                    .Lock = True
                    UpdateCell fps, .Col, .Row, .Text
                    .Row = 54
                    .Col = .ColLetterToNumber("K")
                    .value = 0
                    UpdateCell fps, .Col, .Row, .Text
                Else
                    .Row = 51
                    .Col = .ColLetterToNumber("K")
                    .Lock = False
                End If
            End If
            'check ngay nhap
    
           If Col = .ColLetterToNumber("J") And Row = 69 Then
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
           
           
           If Col = .ColLetterToNumber("H") And Row = 60 Then
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
           
            ' Update so chenh lech cho KHBS
'            .GetText .ColLetterToNumber("J"), 21, varTemp
'            If TAX_Utilities_v1.NodeValidity.childNodes(.SheetCount - 2).Attributes.getNamedItem("Active").nodeValue <> "0" Then
'                .Sheet = .SheetCount - 1
'                Set xmlNode = TAX_Utilities_v1.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
'                ParserCellID fps, GetAttribute(xmlNode, "CellID"), cCol, cRow
'                .Col = cCol
'                .Row = cRow
'                .Value = CDbl(varTemp) - CDbl(vThuephainop)
'                mCurrentSheet = .SheetCount - 1
'                UpdateCell fps, .Col, .Row, .Value
'                .Col = .ColLetterToNumber("BE")
'                .Row = .Row + 3
'                UpdateCell fps, .Col, .Row, .Value
'                .Sheet = 1
'                 mCurrentSheet = 1
'            End If
           
           
        End If
'        If .ActiveSheet = 2 Then
'           If Col = .ColLetterToNumber("D") Then
'               .GetText Col, Row, varTemp
'               If varTemp <> "" And varTemp <> "../../...." Then
'                   If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
'                       .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                   Else
'                        .SetFocus
'                        .SetActiveCell Col, Row
'                   End If
'               Else
'                    .SetText Col, Row, ""
'               End If
'               .Col = Col
'               .Row = Row
'               UpdateCell fps, Col, Row, .Text
'           End If
'        End If
        
        'KHBS
'        If .ActiveSheet = .SheetCount - 1 Then
'         Dim count1 As Integer
'            .Row = 9
'            .Col = .ColLetterToNumber("B")
'            Do While .value <> "bb"
'                count1 = count1 + 1
'                .Row = .Row + 1
'            Loop
'                If Col = .ColLetterToNumber("BG") And Row = 23 + count1 - 5 Then
'                       .GetText Col, Row, varTemp
'                       If varTemp <> "" And varTemp <> "../../...." Then
'                           If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
'                               .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                               .TypeHAlign = TypeHAlignLeft
'                           Else
'                                .SetFocus
'                                .SetActiveCell Col, Row
'                           End If
'                       Else
'                        .SetText Col, Row, ""
'                       End If
'                      .Col = Col
'                      .Row = Row
'                      UpdateCell fps, .Col, .Row, .Text
'                   End If
'                .Sheet = .SheetCount - 1
'        End If
    .EventEnabled(EventAllEvents) = True
    End With
End Sub

'Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
'    Dim xmlNodeH15 As MSXML.IXMLDOMNode
'    Dim cCol As Long, cRow As Long
'    Dim varTemp As Variant
'With fps
'    .EventEnabled(EventAllEvents) = False
'    .Sheet = .ActiveSheet
'    Select Case .Sheet
'        Case .SheetCount - 1
'         mCurrentSheet = .Sheet
'            .Sheet = 1
'            .GetText .ColLetterToNumber("J"), 21, varTemp
'            .Sheet = .SheetCount - 1
'            Set xmlNodeH15 = TAX_Utilities_v1.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
'                 ParserCellID fps, GetAttribute(xmlNodeH15, "CellID"), cCol, cRow
'                .Col = cCol
'                .Row = cRow
'                If loaiKHBS = "frmKHBS_TT" Then
'                    .Value = CDbl(varTemp) - CDbl(vThuephainop)
'                     UpdateCell fps, .Col, .Row, .Value
'                Else
'                    .Value = CDbl(varTemp)
'                     UpdateCell fps, .Col, .Row, .Value
'                End If
'                If .Value < 0 Then
'                    .Col = .ColLetterToNumber("BG")
'                    .Row = cRow + 3
'                    .Value = 0
'                    UpdateCell fps, .Col, .Row, 0
'                Else
'                    .Col = .ColLetterToNumber("BG")
'                    .Row = cRow + 3
'                    UpdateCell fps, .Col, .Row, .Value
'                End If
'
'          End Select
'         .Sheet = mCurrentSheet
'         .EventEnabled(EventAllEvents) = True
'    End With
'
'End Sub

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim strId As Variant
Dim iCol As Long, iRow As Long
Dim intIndexCombo As Integer, i As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim itemp As Variant, temp As Variant
Dim cValue As Variant

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
   If .ActiveSheet = .SheetCount - 1 Then
               'lay ve dong combobox
'            Set xmlCellNode = TAX_Utilities_v1.Data(.SheetCount - 2).nodeFromID(GetCellID(fps, .ColLetterToNumber("BI"), 5))
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
     ElseIf .ActiveSheet = 1 Then
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("H") And Row = 58 Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("P")
            .typeComboboxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
        End If
    End If
    
End With
    
    TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
End Sub


'Calling this function whenever we don't need this class anymore.
'No parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
    Dim iCol As Long, iRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        
        CheckDynamicError 'Set Exception Error on cells of interface
            
        .Sheet = 1
        .Col = .ColLetterToNumber("H")
        .Row = 47
        If .value <> "chk" Then
            If .value = 0 Then
                FormatTextPercent fps, 1, .ColLetterToNumber("K"), 47, False
            End If
        End If

        UpdateSheets
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    Dim iFlagTongPhanBo As Variant, iFlagCheckDate As Integer, iFlagSoUuDai As Integer, iFlagCheckMaSTDN As Integer, iFlagCheckMaSTDN10 As Integer, iFlagCheckSoThue As Integer
    Dim strCheck As String
    Dim blCheck_S2 As Boolean
    Dim vFlag As Variant, vErrMsg As Variant
    Dim varTemp As Variant, vMST As Variant, checkboxThue As Variant, vCT26 As Variant, arrCT26() As String
    Dim errMST As Variant, errTongTLPB As Variant, errEmty As Variant, errDate As Variant, errSoThueUuDai As Variant, errMSTDN As Variant, errMst10 As Variant, errTongSoThue As Variant
    Dim i As Integer
    
    
    ' kiem tra nam dc gh
    Dim Quynamtc As Integer
    Dim Namtc As Integer
    Dim QuyKK As Integer
    Dim NamKK As Integer
    Dim iNgayTaiChinh As Integer
    Dim iThangTaiChinh As Integer
    Dim strNgayTaiChinh As String
        
    Dim isCheckGh As Boolean
    '
    Dim arrDateGH() As String
    Dim arrDateHanNop() As String
    Dim vErrorCheckHN As Variant
    ' PL 01
    Dim iFlagMSTTrung As Integer, mErrMSTTrung As Variant
    
    ' check gh
    Dim vCheckGH As Variant, vLyDoGH As Variant, vErrorCheckGH As Variant, vThoiHanGH As Variant, vErrorCheckThoiHanGH As Variant, vSoThueGH As Variant, vErrorCheckSoThueGH As Variant
    Dim iflagCheckGH As Integer, iflagCheckThoiHanGH As Integer, iflagCheckSoThueGH As Integer
    
    With fps
    '.Visible = False
        '.ReDraw = False
        .Sheet = .SheetCount
        
        .GetText .ColLetterToNumber("E"), 20, vErrMsg
        .GetText .ColLetterToNumber("E"), 21, errTongTLPB
        .GetText .ColLetterToNumber("E"), 31, errMST
        .GetText .ColLetterToNumber("E"), 22, errEmty
        .GetText .ColLetterToNumber("E"), 24, errDate
        .GetText .ColLetterToNumber("E"), 25, errSoThueUuDai
        .GetText .ColLetterToNumber("B"), 21, iFlagTongPhanBo
        .GetText .ColLetterToNumber("E"), 26, errMSTDN
        .GetText .ColLetterToNumber("E"), 27, errMst10
        
        .GetText .ColLetterToNumber("E"), 29, mErrMSTTrung
        
        .GetText .ColLetterToNumber("E"), 30, vErrorCheckGH
        .GetText .ColLetterToNumber("E"), 32, errTongSoThue
        
        .GetText .ColLetterToNumber("E"), 34, vErrorCheckThoiHanGH
        .GetText .ColLetterToNumber("E"), 35, vErrorCheckSoThueGH
        
        .GetText .ColLetterToNumber("E"), 38, vErrorCheckHN
        
        vFlag = 0
            
        .Sheet = 1
        iFlagSoUuDai = 0
        iFlagCheckDate = 0
        iFlagCheckMaSTDN = 0
        iFlagCheckMaSTDN10 = 0
        iFlagCheckSoThue = 0
        mCurrentSheet = .Sheet
        .GetText .ColLetterToNumber("J"), 69, varTemp
        .Col = .ColLetterToNumber("J")
        .Row = 69
        If varTemp <> "" And varTemp <> "../../...." Then
            If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                .SetText .Col, .Row, Format_ddmmyyyy(CStr(varTemp))
                .BackColor = mNonErrorColor
                .CellNote = ""
            Else
    '                .SetFocus
                .BackColor = mErrorColor
                .CellNote = errDate
                .SetActiveCell .Col, .Row
                iFlagCheckDate = 1
            End If
        Else
            .SetText .Col, .Row, ""
            .BackColor = mNonErrorColor
            .CellNote = ""
        End If
        UpdateCell fps, .Col, .Row, varTemp
        'check dinh dang so trong chi tieu 26
        .GetText .ColLetterToNumber("H"), 47, checkboxThue
        If (Val(checkboxThue) = 1) Then
     
            .GetText .ColLetterToNumber("K"), 47, vCT26
            arrCT26 = Split(CStr(vCT26), ";")
            For i = 0 To UBound(arrCT26) - 1
                If IsNumeric(arrCT26(i)) = False Then
                    iFlagSoUuDai = 1
                End If
            Next
            If iFlagSoUuDai = 1 Then
                .Col = .ColLetterToNumber("K")
                .Row = 47
                .CellNote = errSoThueUuDai
                .BackColor = mErrorColor
            Else
                .Col = .ColLetterToNumber("K")
                .Row = 47
                .BackColor = mNonErrorColor
                .CellNote = ""
            End If
        ElseIf (Val(checkboxThue) = 0) Then
                .Col = .ColLetterToNumber("K")
                .Row = 47
                .BackColor = mNonErrorColor
                .CellNote = ""
        End If
    
        ' check ly do gia han
         ' cac ky khong check ly do gia han
        QuyKK = TAX_Utilities_v1.ThreeMonths
        NamKK = TAX_Utilities_v1.Year
        
        
        'Lay ngay bat dau nam tai chinh
        If GetAttribute(TAX_Utilities_v1.NodeMenu, "FinanceYear") = "1" Then
            strNgayTaiChinh = GetNgayBatDauNamTaiChinh
                iNgayTaiChinh = GetNgayTaiChinh(strNgayTaiChinh)
                iThangTaiChinh = GetThangTaiChinh(strNgayTaiChinh)
        Else
            strNgayTaiChinh = "01/01"
            iNgayTaiChinh = 1
            iThangTaiChinh = 1
        End If
        
        Quynamtc = GetQuyNamTaiChinh(QuyKK, NamKK, iNgayTaiChinh, iThangTaiChinh, 0)
        Namtc = GetQuyNamTaiChinh(QuyKK, NamKK, iNgayTaiChinh, iThangTaiChinh, 1)
'        If (Namtc = 2010 And Quynamtc = 4) Then
'            isCheckGh = False
'        ElseIf Namtc = 2011 And (Quynamtc = 3 Or Quynamtc = 4) Then
'            isCheckGh = False
'        ElseIf Namtc = 2011 And (Quynamtc = 1 Or Quynamtc = 2) Then
'            isCheckGh = False
'        ElseIf Namtc = 2010 And Quynamtc <> 4 Then
'            isCheckGh = False
'        ElseIf Namtc = 2009 Then
'            isCheckGh = False
'        Else
'            isCheckGh = True
'        End If
         ' end
        isCheckGh = True
        
        .GetText .ColLetterToNumber("F"), 56, vCheckGH
        .GetText .ColLetterToNumber("H"), 58, vLyDoGH
        
        .GetText .ColLetterToNumber("H"), 60, vThoiHanGH
        .GetText .ColLetterToNumber("H"), 62, vSoThueGH
        
         If (Trim$(vCheckGH) = "1" Or Trim$(UCase$(vCheckGH)) = "X") And isCheckGh Then
            ' kiem tra ly do gia han
             If Trim$(vLyDoGH) = "" Then
                 .Col = .ColLetterToNumber("H")
                 .Row = 58
                 .CellNote = vErrorCheckGH  'static
                 .BackColor = mErrorColor
                 iflagCheckGH = 1
             Else
                 .Col = .ColLetterToNumber("H")
                 .Row = 58
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckGH = 0
             End If
             ' Kiem tra thoi han gia han
             If Trim$(vThoiHanGH) = "" Or Trim$(vThoiHanGH) = "../../...." Then
                 .Col = .ColLetterToNumber("H")
                 .Row = 60
                 .CellNote = vErrorCheckThoiHanGH  'static
                 .BackColor = mErrorColor
                 iflagCheckThoiHanGH = 1
             Else
                    ' kiem tra han nop
                    arrDateGH = Split(vThoiHanGH, "/")
                    arrDateHanNop = Split(hanNopTKTNDN, "/")
                    Dim dThoiHanGH As Date
                    Dim dHanNop As Date

                    dThoiHanGH = DateSerial(CInt(arrDateGH(2)), CInt(arrDateGH(1)), CInt(arrDateGH(0)))
                    dHanNop = DateSerial(CInt(arrDateHanNop(2)), CInt(arrDateHanNop(1)), CInt(arrDateHanNop(0)))
                    If DateDiff("D", dThoiHanGH, dHanNop) > 0 Then
                        .Col = .ColLetterToNumber("H")
                        .Row = 60
                        .CellNote = vErrorCheckHN  'static
                        .BackColor = mAlertColor
                    Else
                        .Col = .ColLetterToNumber("H")
                        .Row = 60
                        .CellNote = ""  'static
                        .BackColor = mNonErrorColor
                    End If
             End If
             ' Kiem tra so thue
             If Val(vSoThueGH) <= 0 Then
                 .Col = .ColLetterToNumber("H")
                 .Row = 62
                 .CellNote = vErrorCheckSoThueGH
                 .BackColor = mErrorColor
                 iflagCheckSoThueGH = 1
             Else
                 .Col = .ColLetterToNumber("H")
                 .Row = 62
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckSoThueGH = 0
             End If
         Else
                ' kiem tra ly do gia han
                 .Col = .ColLetterToNumber("H")
                 .Row = 58
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckGH = 0
                 ' Thoi han gia han
                 .Col = .ColLetterToNumber("H")
                 .Row = 60
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckThoiHanGH = 0
                 ' So thue gia han
                 .Col = .ColLetterToNumber("H")
                 .Row = 62
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckSoThueGH = 0
         End If
    
    '    check tren sheet 2
        .Sheet = 2
        If .SheetVisible = True Then

            mCurrentSheet = .Sheet
            Dim iFlagEmpty As Integer, iFlagCheckMaST As Integer
            Dim vTenDn As Variant, vCQTCapCuc As Variant, vCqtChiCuc As Variant, vTyLe As Variant, vSoThue As Variant
            Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer
            Dim iWhiteRow As Integer
            Dim iFlagTyLe  As Integer
            Dim vTongTyLe As Double
            Dim vTongSoThue As Double
            Dim vCt5 As Variant
            Dim vColor As Variant
            iFlagEmpty = 0
            iFlagTyLe = 0
            iFlagCheckMaST = 0
            iWhiteRow = 1
            
            ' lay so thue cua to khai
            .GetText .ColLetterToNumber("F"), 7, vCt5
            
            'check not empty ten doanh nghiep
            .Row = 35
            Do
                .GetText .ColLetterToNumber("D"), .Row, vTenDn
                .GetText .ColLetterToNumber("F"), .Row, vMST
                .GetText .ColLetterToNumber("G"), .Row, vCQTCapCuc
                .GetText .ColLetterToNumber("H"), .Row, vCqtChiCuc
                .GetText .ColLetterToNumber("M"), .Row, vTyLe
                .GetText .ColLetterToNumber("P"), .Row, vSoThue
                
                If Trim$(vTenDn) <> "" Or Trim$(vMST) <> "" Or Trim$(vCQTCapCuc) <> "" Or Trim$(vCqtChiCuc) <> "" Or Trim$(vTyLe) <> "0" Or Trim$(vSoThue) <> "0" Then iWhiteRow = 0

                If Trim$(vMST) <> vbNullString Then
                       vMST = Replace(vMST, " ", "")
                       vMST = Replace(vMST, "-", "")
                        'vMST = Left(vMST, 10)

                       strCheck = ""
                       If Len(vMST) = 10 Or Len(vMST) = 13 Then
                            strCheck = CheckTaxCode(Mid$(vMST, 1, 1), Mid$(vMST, 2, 1), Mid$(vMST, 3, 1), Mid$(vMST, 4, 1), Mid$(vMST, 5, 1), Mid$(vMST, 6, 1), Mid$(vMST, 7, 1), Mid$(vMST, 8, 1), Mid$(vMST, 9, 1), Mid$(vMST, 10, 1), Mid$(vMST, 11, 1), Mid$(vMST, 12, 1), Mid$(vMST, 13, 1))
                       End If
                       
                       If Mid$(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                            .Col = .ColLetterToNumber("F")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = errMST   'static
                            .BackColor = mErrorColor
                            iFlagCheckMaST = 1
                            
                            'capture the first error cell
                            If iFocusFlag = 0 Then
                                strFocusSheetName = "'" & .SheetName & "'"
                                strFocusCol = str$(.Col)
                                strFocusRow = str$(.Row)
                                iFocusFlag = 1
                            End If
                        Else
                            .Col = .ColLetterToNumber("F")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                Else
                    .Col = .ColLetterToNumber("F")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If

                    'check MST trung nhau
'                     CountRow = .SearchCol(.ColLetterToNumber("F"), -1, -1, vMSTCheckTrung, SearchFlagsNone)
'                     rowDes = .SearchCol(.ColLetterToNumber("F"), CountRow, -1, vMSTCheckTrung, SearchFlagsNone)
'                    If rowDes > -1 Then
'                        .Col = .ColLetterToNumber("F")
'                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                        .CellNote = mErrMSTTrung      'static
'                        .BackColor = mErrorColor
'                        iFlagMSTTrung = 1
'                    End If


                'check du lieu bat buoc phai nhap
                If (vTenDn = "" And iWhiteRow = 0) Then
                    .Col = .ColLetterToNumber("D")
                    .BackColor = mErrorColor
                    .CellNote = errEmty
                    iFlagEmpty = 1
                Else
                    .Col = .ColLetterToNumber("D")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If
                
                If (vCQTCapCuc = "" And iWhiteRow = 0) Then
                    .Col = .ColLetterToNumber("G")
                    .BackColor = mErrorColor
                    .CellNote = errEmty
                    iFlagEmpty = 1
                Else
                    .Col = .ColLetterToNumber("G")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If
                
                If (vCqtChiCuc = "" And iWhiteRow = 0) Then
                    .Col = .ColLetterToNumber("H")
                    .BackColor = mErrorColor
                    .CellNote = errEmty
                    iFlagEmpty = 1
                Else
                    .Col = .ColLetterToNumber("H")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If

                ' kiem tra tong ty le
                vTongTyLe = vTongTyLe + CDbl(vTyLe)
                vTongSoThue = vTongSoThue + CDbl(vSoThue)

                .Col = .ColLetterToNumber("C")
                .Row = .Row + 1
            Loop Until .Text = "aa"
            
            If Round(vTongTyLe, 2) <> 100 And iWhiteRow = 0 Then
                    .Col = .ColLetterToNumber("M")
                    .Row = .Row + 1
                    .BackColor = mErrorColor
                    .CellNote = errTongTLPB
                    iFlagTyLe = 1
            Else
                    .Row = .Row + 1
                    .Col = .ColLetterToNumber("C")
                     vColor = .BackColor
                    .Col = .ColLetterToNumber("M")
                    .BackColor = vColor
                    .CellNote = ""
            End If
            ' Kiem tra so thue
            If (CDbl(vCt5) - vTongSoThue > 5 Or vTongSoThue - CDbl(vCt5) > 5) Then
                    .Col = .ColLetterToNumber("P")
                    .BackColor = mErrorColor
                    .CellNote = errTongSoThue
                    iFlagCheckSoThue = 1
            Else
                    .Col = .ColLetterToNumber("C")
                     vColor = .BackColor
                    .Col = .ColLetterToNumber("P")
                    .BackColor = vColor
                    .CellNote = ""
            End If
            
            
        End If
'        'end
        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        .Sheet = .SheetCount
            If blCheck_S2 Then
                .SetText .ColLetterToNumber("B"), 20, IIf(vFlag = 1, "0", "1")
                .SetText .ColLetterToNumber("B"), 21, IIf(iFlagTyLe = 1, "0", "1")
                .SetText .ColLetterToNumber("B"), 27, IIf(iFlagCheckMaSTDN10 = 1, "0", "1")
                .SetText .ColLetterToNumber("B"), 32, IIf(iFlagCheckSoThue = 1, "0", "1")
            Else
                .SetText .ColLetterToNumber("B"), 20, "1"
                .SetText .ColLetterToNumber("B"), 21, "1"
                .SetText .ColLetterToNumber("B"), 27, "1"
                .SetText .ColLetterToNumber("B"), 32, "1"
            End If
            
            .SetText .ColLetterToNumber("B"), 31, IIf(iFlagCheckMaST = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 22, IIf(iFlagEmpty = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 24, IIf(iFlagCheckDate = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 25, IIf(iFlagSoUuDai = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 26, IIf(iFlagCheckMaSTDN = 1, "0", "1")
            
            .SetText .ColLetterToNumber("B"), 29, IIf(iFlagMSTTrung = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 30, IIf(iflagCheckGH = 1, "0", "1")
            
            .SetText .ColLetterToNumber("B"), 34, IIf(iflagCheckThoiHanGH = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 35, IIf(iflagCheckSoThueGH = 1, "0", "1")
        
            CheckErrorMST
        
        ' check phu luc phan bo
        .Sheet = .SheetCount
        
        Dim vChon As Variant
        Dim vTenNhaMay, vCQTCapDiaPhuong, vTyLePhanBo, vSoThuePhaiNop As Variant
        Dim vMessRequire As Variant
        
        Dim flagSoThue As Integer
        
        .GetText .ColLetterToNumber("E"), 39, vMessRequire
        .GetText .ColLetterToNumber("E"), 41, errMST
        
        .Sheet = 3
        
        If .SheetVisible = True Then
            mCurrentSheet = .Sheet
            'check phu luc 01_2
            .Row = 42
            iWhiteRow = 1
            
            Do
                .GetText .ColLetterToNumber("C"), .Row, vTenNhaMay
                .GetText .ColLetterToNumber("G"), .Row, vChon
                .GetText .ColLetterToNumber("H"), .Row, vMST
                .GetText .ColLetterToNumber("P"), .Row, vCQTCapDiaPhuong
                .GetText .ColLetterToNumber("R"), .Row, vTyLePhanBo
                .GetText .ColLetterToNumber("U"), .Row, vSoThuePhaiNop
                
                
                If Trim$(vTenNhaMay) <> "" Or Trim$(vMST) <> "" Or Trim$(vCQTCapDiaPhuong) <> "" Or Trim$(vTyLePhanBo) <> "0" Or Trim$(vSoThuePhaiNop) <> "0" Then
                    iWhiteRow = 0
                End If
            
                .Col = .ColLetterToNumber("C")
                If (Trim(.Text) = vbNullString And Trim(vChon) = "1" And iWhiteRow = 0) Then
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vMessRequire      'static
                    .BackColor = mErrorColor
    
                    iFlagEmpty = 1
                Else
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                End If
    
                
                .Col = .ColLetterToNumber("P")
                If (Trim(.Text) = vbNullString And Trim(vChon) = "0" And iWhiteRow = 0) Then
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vMessRequire      'static
                    .BackColor = mErrorColor
    
                    iFlagEmpty = 1
                Else
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                End If
                
                ' bat buoc nhap cot so thue
                .Col = .ColLetterToNumber("U")
                If Val(vSoThuePhaiNop) = 0 And iWhiteRow = 0 Then
                    .CellNote = vMessRequire      'static
                    .BackColor = mErrorColor
    
                    flagSoThue = 1
                Else
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If
                
                'check mst
                If Trim$(vMST) <> vbNullString Then
                    vMST = Replace(vMST, " ", "")
                    vMST = Replace(vMST, "-", "")
                    'vMST = Left(vMST, 10)

                    strCheck = ""

                    If Len(vMST) = 10 Or Len(vMST) = 13 Then
                        strCheck = CheckTaxCode(Mid$(vMST, 1, 1), Mid$(vMST, 2, 1), Mid$(vMST, 3, 1), Mid$(vMST, 4, 1), Mid$(vMST, 5, 1), Mid$(vMST, 6, 1), Mid$(vMST, 7, 1), Mid$(vMST, 8, 1), Mid$(vMST, 9, 1), Mid$(vMST, 10, 1), Mid$(vMST, 11, 1), Mid$(vMST, 12, 1), Mid$(vMST, 13, 1))
                    End If
                       
                    If Mid$(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                        .Col = .ColLetterToNumber("H")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errMST   'static
                        .BackColor = mErrorColor

                        iFlagCheckMaST = 1
                            
                        'capture the first error cell
'                        If iFocusFlag = 0 Then
'                            strFocusSheetName = "'" & .SheetName & "'"
'                            strFocusCol = str$(.Col)
'                            strFocusRow = str$(.Row)
'
'                            iFocusFlag = 1
'                        End If

                    Else
                        .Col = .ColLetterToNumber("H")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If

                Else
                    If (Trim$(vMST) = vbNullString And iWhiteRow = 0) Then
                        .Col = .ColLetterToNumber("H")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        
                        'iFlagEmty = 1
                    End If
                End If
                
                .Row = .Row + 1
                .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
        End If
        
        'set flag
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 39, IIf(iFlagEmpty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 41, IIf(iFlagCheckMaST = 1, "0", "1")
        
        .SetText .ColLetterToNumber("B"), 42, IIf(flagSoThue = 1, "0", "1")
        
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
             Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(3).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(3).getElementsByTagName("Cell").length - 1)
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

End Function

Public Sub SetData()
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
        ' clear checkbox
'        .Row = 6
'        .Col = .ColLetterToNumber("E")
'        .value = 0
'        UpdateCell fps, .Col, .Row, .Text

      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing        '
End Sub

Public Sub Prepare_KHBS()
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    Dim strPath As String
    Dim strComboName As String
    Dim strComboMCT As String
    Dim strCombo As String
    
    mblKhoiDauToKhai = True
    fps.Sheet = fps.SheetCount - 1
    fps.ActiveSheet = fps.Sheet
    
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT.xml"
    
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
                strCombo = strCombo + CPab(GetAttribute(xmlNodeItem, "Caption"), 100) + CPab(GetAttribute(xmlNodeItem, "CellID"), 20) + GetAttribute(xmlNodeItem, "MCT") + Chr$(9)
                strComboName = strComboName + GetAttribute(xmlNodeItem, "Caption") + Chr$(9)
                strComboMCT = strComboMCT + GetAttribute(xmlNodeItem, "MCT") + Chr$(9)
    Next
     With fps
        
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        .Col = .ColLetterToNumber("BC")
        .Row = 9
        .TypeComboBoxList = strCombo
        .Row = 13
        .TypeComboBoxList = strCombo
        .Col = .ColLetterToNumber("BD")
        .Row = 9
        .TypeComboBoxList = strComboName
        .Row = 13
        .TypeComboBoxList = strComboName
        .Col = .ColLetterToNumber("BE")
        .Row = 9
        .TypeComboBoxList = strComboMCT
        .Row = 13
        .TypeComboBoxList = strComboMCT
' Cho phep nhap am cac chi tieu 11, 12
        .Sheet = 1
        .Col = .ColLetterToNumber("J")
        .Row = 11
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Row = 12
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
    End With
    
End Sub


Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

Public Sub UpdateChangeKHBS()
    Dim varTemp As Variant
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    With fps
        .Sheet = 1
        For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
            ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
            .GetText lCol, lRow, varTemp
            ' Trong truong hop check co nhieu thue suat thi set lai gia tri varTemp = "x" khong thi varTemp = ""
            If (lCol = .ColLetterToNumber("E") And lRow = 7) And (Trim(varTemp) = "1" Or Trim(varTemp) = "x") Then
                varTemp = "x"
            ElseIf (lCol = .ColLetterToNumber("E") And lRow = 7) And (Trim(varTemp) = "0" Or Trim(varTemp) = "") Then
                varTemp = ""
            End If
            If (lCol = .ColLetterToNumber("E") And lRow = 6) And (Trim(varTemp) = "1" Or Trim(varTemp) = "x") Then
                varTemp = "x"
            ElseIf (lCol = .ColLetterToNumber("E") And lRow = 6) And (Trim(varTemp) = "0" Or Trim(varTemp) = "") Then
                varTemp = ""
            End If
            If (lCol = .ColLetterToNumber("G") And lRow = 42) And (Trim(varTemp) = "1" Or Trim(varTemp) = "x") Then
                varTemp = "x"
            ElseIf (lCol = .ColLetterToNumber("G") And lRow = 42) And (Trim(varTemp) = "0" Or Trim(varTemp) = "") Then
                varTemp = ""
            End If
            
            If (lCol = .ColLetterToNumber("B") And lRow = 48) And (Trim(varTemp) = "1" Or Trim(varTemp) = "x") Then
                varTemp = "x"
            ElseIf (lCol = .ColLetterToNumber("B") And lRow = 48) And (Trim(varTemp) = "0" Or Trim(varTemp) = "") Then
                varTemp = ""
            End If
            
            UpdateLastKHBSCell fps, lCol, lRow, varTemp
        Next
        .Sheet = mCurrentSheet
    End With
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

' update gia tri tren man hinh xuong DOM
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
        For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
           ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               If (lRow = 7 And lCol = .ColLetterToNumber("F")) Or (lRow = 6 And lCol = .ColLetterToNumber("F")) Or (lRow = 56 And lCol = .ColLetterToNumber("F")) Or (lRow = 47 And lCol = .ColLetterToNumber("H")) Then
                    .GetText lCol, lRow, varTemp
                    If varTemp = "1" Or UCase(CStr(varTemp)) = "X" Then
                        UpdateCell fps, lCol, lRow, "x"
                    Else
                        UpdateCell fps, lCol, lRow, ""
                    End If
               Else
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
               End If
        Next
        
        
        blCheckS2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
       If blCheckS2 = True Then
           .Sheet = 2
           mCurrentSheet = .Sheet
               For Each xmlCellNodeData In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
                  ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
    '              If (lRow < 52) Or (lRow > 58) Then
                       .GetText lCol, lRow, varTemp
                       UpdateCell fps, lCol, lRow, varTemp
     '              End If
               Next
        End If
        
        
       .Sheet = 3
        mCurrentSheet = .Sheet
        
        blCheckS2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheckS2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(2).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                  If (lRow >= 42 And lCol = .ColLetterToNumber("G")) Or (lRow >= 42 And lCol = .ColLetterToNumber("AI")) Then
                        .GetText lCol, lRow, varTemp
                        If varTemp = "1" Or UCase(CStr(varTemp)) = "X" Then
                            UpdateCell fps, lCol, lRow, "x"
                        Else
                            UpdateCell fps, lCol, lRow, ""
                        End If
                  Else
                        .GetText lCol, lRow, varTemp
                        UpdateCell fps, lCol, lRow, varTemp
                  End If
            Next
        End If
        
       ' Update lai sheet KHBS
        .Sheet = .SheetCount - 1
        mCurrentSheet = .Sheet
        blCheckS2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheckS2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(.Sheet - 1).getElementsByTagName("Cell")
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
Public Function getValueTK(strFileLoadData As String) As String

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
            Set xmlNode = .nodeFromID("K_49")
            result = "CT28_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("K_50")
            result = result & "~" & "CT29_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("K_51")
            result = result & "~" & "CT30_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01BTNDN", "getValueTK", Err.number, Err.Description
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
                .GetText .ColLetterToNumber("K"), 49, varTemp
                strCurrentValue = "CT28_" & Trim(varTemp)
                .GetText .ColLetterToNumber("K"), 50, varTemp
                strCurrentValue = strCurrentValue & "~CT29_" & Trim(varTemp)
                .GetText .ColLetterToNumber("K"), 51, varTemp
                strCurrentValue = strCurrentValue & "~CT30_" & Trim(varTemp)
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

                If Left(strMaCT, 3) = "30_" Then
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
                        ' dieu chinh giam
                        strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
                        ' dieu chinh tang
                        strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                Else
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
                        ' dieu chinh giam
                        strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
                        ' dieu chinh tang
                        strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
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
    SaveErrorLog "01B_TNDN", "getMapCT", Err.number, Err.Description
End Function


' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("K"), 54, varTemp
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
            Set xmlNode = .nodeFromID("K_54")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01BTNDN", "getValueCTDC", Err.number, Err.Description
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
                Set xmlCellNode = TAX_Utilities_v1.Data(3).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(3).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(3).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(3).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v1.Data(3).getElementsByTagName("Cell").length > 31 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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


' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 24
    denRowDL = 34
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "14_16_2~11_12_3~20_30_4"
    ' end
End Sub

' ham reset formula
Private Sub resetFomula()
     With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 2
        .Row = 34
        .Col = .ColLetterToNumber("O")
        .Formula = ""
        .Col = .ColLetterToNumber("P")
        .Formula = ""
        .Col = .ColLetterToNumber("Q")
        .Formula = ""
        .Col = .ColLetterToNumber("R")
        .Formula = ""
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

' ham set formula
Private Sub setFomula()
     With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 2
        .Row = 34
        .Col = .ColLetterToNumber("O")
        .Formula = "IF(I49=""1"",'01'!J47,0)"
        .Col = .ColLetterToNumber("P")
        .Formula = "IF(I49=""2"",'01'!J47,0)"
        .Col = .ColLetterToNumber("Q")
        .Formula = "IF(I49=""3"",'01'!J47,0)"
        .Col = .ColLetterToNumber("R")
        .Formula = "IF(I49=""4"",'01'!J47,0)"
        .EventEnabled(EventAllEvents) = True
    End With
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
            .Sheet = 2 'set thong tin CQL
            '.EventEnabled(EventAllEvents) = False
            
            .Row = rowStart
            .Col = .ColLetterToNumber("G")
            .TypeComboBoxList = tenCucThue
            
            .Row = rowStart
            .Col = .ColLetterToNumber("S")
            .TypeComboBoxList = maCucThue
            
            .Row = rowStart
            .Col = .ColLetterToNumber("H")
            .TypeComboBoxList = tenChiCucThue
    
            .Row = rowStart
            .Col = .ColLetterToNumber("R")
            .TypeComboBoxList = maChiCucThue
            '.EventEnabled(EventAllEvents) = True
            
             .Sheet = 3
            
            .Row = 42
            .Col = .ColLetterToNumber("P")
            .TypeComboBoxList = tenCucThue
            
            .Row = 42
            .Col = .ColLetterToNumber("AH")
            .TypeComboBoxList = maCucThue
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
        If .ActiveSheet = 2 Then
            .Sheet = 2
            mCurrentSheet = 2
            
            .Col = Col
            .Row = Row
            
            If Col = .ColLetterToNumber("G") Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot G va gia tri Index cot S
                intIndexCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("S")
                .Row = Row
                .typeComboboxCurSel = intIndexCuc
                'Nhap lai Combo cot R
                .Col = .ColLetterToNumber("S")
                UpdateCell fps, .Col, .Row, .Text
                ' Lay Id cua cuc thue vua chon o tren
                fps.GetText .Col, .Row, varIdCucThue
                
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
                    .Col = .ColLetterToNumber("H")
                    .TypeComboBoxList = tenChiCucThue
                    
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("R")
                    .TypeComboBoxList = maChiCucThue
                    
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                   
                End If
            End If
                        
            If Col = .ColLetterToNumber("H") Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot L va gia tri Index cot R
                intIndexChiCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("R")
                .Row = Row
                .typeComboboxCurSel = intIndexChiCuc
                'Nhap lai Combo cot R
                .Col = .ColLetterToNumber("R")
                UpdateCell fps, .Col, .Row, .Text
            End If
        ElseIf .ActiveSheet = 3 Then
            If Col = .ColLetterToNumber("P") And .Text <> "" Then
                intIndexChiCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("AH")
                .typeComboboxCurSel = intIndexChiCuc
                'Update combo C
                UpdateCell fps, .Col, .Row, .Text
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
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
                     .GetText .ColLetterToNumber("S"), .Row, CQT_CAPCUC
                     .GetText .ColLetterToNumber("R"), .Row, CQT_CAPCHICUC
                     GetCQTByMaCQT CQT_CAPCUC, tCQT_CAPCUC
                     GetCQTByMaCQT CQT_CAPCHICUC, tCQT_CAPCHICUC
                     .Col = .ColLetterToNumber("G")
                     .Text = tCQT_CAPCUC
                     UpdateCell fps, .Col, .Row, .Text
                     
                     .Col = .ColLetterToNumber("H")
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
    Dim i As Integer
    Dim CQT As Variant
    Dim idx As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        .Col = .ColLetterToNumber("P")
        .Row = 58
        idx = .typeComboboxCurSel
        ' set cho cot ten gia han
        .Col = .ColLetterToNumber("H")
        .Row = 58
        .typeComboboxCurSel = idx
        
         .Sheet = 3
         If .SheetVisible = True Then
                .Row = 42
                .Col = .ColLetterToNumber("B")
                i = 1
                Do
                     .GetText .ColLetterToNumber("P"), .Row, CQT
                     If Trim$(CStr(CQT)) = "" Or CQT = vbNullString Then
                        .Text = i
                        .Col = .ColLetterToNumber("G")
                        .value = 1
                        .Col = .ColLetterToNumber("P")
                        .Lock = True
                        i = i + 1
                     End If
                     .Col = .ColLetterToNumber("B")
                    .Row = .Row + 1
                Loop Until .Text = "aa"
                
         End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
