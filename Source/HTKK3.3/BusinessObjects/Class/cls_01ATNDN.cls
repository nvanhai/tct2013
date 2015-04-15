VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01ATNDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Thu Nhap Doanh Nghiep Tam tinh" interface sheets
'this Class is belong to TAX_Business_v2 project which will be compline to DLL

Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
'Dim xNodeData0 As MSXML.IXMLDOMNode
Dim xNode As MSXML.IXMLDOMNode
Dim arrData() As String
Public loaiKHBS As String
Public Strloaitk As String
Public StrSolanBosung As String
Public FlagThaydoi As String
Public strLoaiNNKD As String
Private Const Row_MSTDN = 19
Private Const Col_MSTDN = "F"
Private Const rowStart = 35
' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

Public hanNopTKTNDN As String

Private vFormular As String



'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    With fps
        SetDateFormat fps, 1, 56, .ColLetterToNumber("J"), DDMMYYYY
        SetDateFormat fps, 1, 47, .ColLetterToNumber("H"), DDMMYYYY
        .Sheet = 1
'        .Row = 7
'        .Col = .ColLetterToNumber("E")
'        .TypeEditMultiLine = True
        
        .Row = 56
        .Col = .ColLetterToNumber("J")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        'dhdang sua check gia han nop theo thong tu
         ' q Quy ke khai
        ' y nam ke khai
        ' dNgayTaiChinh ngay tai chinh lay tren man hinh HTKK
        ' dThangTaiChinh thang tai chinh tren phan thong tin chung HTKK
        ' dType: 0 tra ve quy, 1 tra ve nam
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
        
        QuyKK = TAX_Utilities_v2.ThreeMonths
        NamKK = TAX_Utilities_v2.Year
        
        
        'Lay ngay bat dau nam tai chinh
'        If GetAttribute(TAX_Utilities_v2.NodeMenu, "FinanceYear") = "1" Then
'            strNgayTaiChinh = GetNgayBatDauNamTaiChinh
'                iNgayTaiChinh = GetNgayTaiChinh(strNgayTaiChinh)
'                iThangTaiChinh = GetThangTaiChinh(strNgayTaiChinh)
'        Else
'            strNgayTaiChinh = "01/01"
'            iNgayTaiChinh = 1
'            iThangTaiChinh = 1
'        End If
'
'        Quynamtc = GetQuyNamTaiChinh(QuyKK, NamKK, iNgayTaiChinh, iThangTaiChinh, 0)
'        Namtc = GetQuyNamTaiChinh(QuyKK, NamKK, iNgayTaiChinh, iThangTaiChinh, 1)
'        If (Namtc = 2010 And Quynamtc = 4) Then
'            Thongtu = GetAttribute(GetMessageCellById("0187"), "Msg")
'            .SetText .ColLetterToNumber("C"), 34, Thongtu
'        ElseIf Namtc = 2011 And (Quynamtc = 3 Or Quynamtc = 4) Then
'            Thongtu = GetAttribute(GetMessageCellById("0243"), "Msg")
'            .SetText .ColLetterToNumber("C"), 34, Thongtu
'        ElseIf Namtc = 2011 And (Quynamtc = 1 Or Quynamtc = 2) Then
'            Thongtu = GetAttribute(GetMessageCellById("0260"), "Msg")
'            .SetText .ColLetterToNumber("C"), 34, Thongtu
'        ElseIf Namtc = 2010 And Quynamtc <> 4 Then
'            Thongtu = GetAttribute(GetMessageCellById("0186"), "Msg")
'            .SetText .ColLetterToNumber("C"), 34, Thongtu
'        ElseIf Namtc = 2009 Then
'            Thongtu = GetAttribute(GetMessageCellById("0185"), "Msg")
'            .SetText .ColLetterToNumber("C"), 34, Thongtu
''        ElseIf Namtc = 2013 And (Quynamtc = 1 Or Quynamtc = 2 Or Quynamtc = 3) Then
''            Thongtu = GetAttribute(GetMessageCellById("0263"), "Msg")
''            .SetText .ColLetterToNumber("C"), 34, Thongtu
'        Else
''            .SetText .ColLetterToNumber("C"), 34, ""
''            .Row = 34
''            .Col = .ColLetterToNumber("B")
''            .Text = ""
''            .Lock = True
'            Thongtu = GetAttribute(GetMessageCellById("0263"), "Msg")
'            .SetText .ColLetterToNumber("C"), 34, Thongtu
            
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
            .Row = 45
            .Col = .ColLetterToNumber("H")
            .TypeComboBoxList = strCboLyDo
            ' set Ma ly do gia han
            .Col = .ColLetterToNumber("P")
            .TypeComboBoxList = strCboLyDoAn
'        End If
        '-------------------------------------------------------------------------------
        'SetDateFormat fps, 2, 9, .ColLetterToNumber("D"), DDMMYYYY
        
        SetDateFormat fps, .SheetCount - 1, 31, .ColLetterToNumber("BE"), DDMMYYYY
        SetDateFormat fps, .SheetCount - 1, 46, .ColLetterToNumber("BG"), DDMMYYYY
        .Sheet = .SheetCount - 1 'KHBS
        .Row = 46
        .Col = .ColLetterToNumber("BG")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
    End With
    setCQTQuanLy
    
        ' set CQT ra quyet dinh hoan
    setCQTQuanLyHoanThue fps
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim i As Long
    Dim intCtrl As Integer
    Dim strId As Variant, chkValue As Variant
    Dim xmlNode As MSXML.IXMLDOMNode
    With fps

      If Not TAX_Utilities_v2.DataKHBS Is Nothing Then
        Set xmlNode = TAX_Utilities_v2.DataKHBS.nodeFromID(GetCellID(fps, fps.ColLetterToNumber("J"), 19))
        vThuephainop = GetAttribute(xmlNode, "Value")
          Set xmlNode = Nothing
         .Sheet = .SheetCount - 1 'KHBS
            Dim cCol As Long, cRow As Long
          '  .ActiveSheet = .Sheet
            Set xmlNode = TAX_Utilities_v2.Data(.SheetCount - 2).childNodes(2).childNodes(2).lastChild.lastChild
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
            Set xmlNode = TAX_Utilities_v2.Data(.SheetCount - 2).childNodes(2).firstChild
            SetAttribute xmlNode, "DateKHBS", TAX_Utilities_v2.DateKHBS
            SetAttribute xmlNode, "loaiKHBS", loaiKHBS
            If loaiKHBS = "frmKHBS_BS" Then
                .SetText .ColLetterToNumber("B"), 5, GetAttribute(GetMessageCellById("0115"), "Msg") & _
                GetAttribute(GetMessageCellById("0117"), "Msg") & Mid(TAX_Utilities_v2.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_v2.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_v2.DateKHBS, 5, 4)
            Else
                .SetText .ColLetterToNumber("B"), 5, GetAttribute(GetMessageCellById("0116"), "Msg") & _
                GetAttribute(GetMessageCellById("0117"), "Msg") & Mid(TAX_Utilities_v2.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_v2.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_v2.DateKHBS, 5, 4)
            End If
                .EventEnabled(EventAllEvents) = False
                .Sheet = 1
                mCurrentSheet = 1
                .Col = .ColLetterToNumber("I")
                .Row = 22
                .Formula = "TRIM(Header!T5)"
                UpdateCell fps, .Col, .Row, .Text
                .Sheet = 1 'To khai GTGT
                .Row = 24
                .Col = .ColLetterToNumber("I")
                .Text = Format(Date, "dd/mm/yyyy")
                .TypeHAlign = TypeHAlignLeft
                .EventEnabled(EventAllEvents) = True
            ReDim arrData(TAX_Utilities_v2.DataKHBS.getElementsByTagName("Cell").length)
            Dim j As Integer
            For Each xNode In TAX_Utilities_v2.DataKHBS.getElementsByTagName("Cell")
                arrData(j) = GetAttribute(xNode, "CellID") & ":" & GetAttribute(xNode, "Value")
                j = j + 1
            Next
            Set xNode = Nothing
        End If
       .Sheet = 1
       mCurrentSheet = .Sheet
        'set to khai bo xung hay chinh thuc
            If Strloaitk = "TKCT" Then
                .Col = .ColLetterToNumber("E")
                .Row = 13
                .Text = 1
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("G")
                .Row = 13
                .Text = 0
                UpdateCell fps, .Col, .Row, .Text
             ElseIf Strloaitk = "TKBS" Then
                .Col = .ColLetterToNumber("E")
                .Row = 13
                .Text = 0
                  UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("G")
                .Row = 13
                .Text = Val(StrSolanBosung)
                UpdateCell fps, .Col, .Row, .Text
            End If
            'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
            'ma NNKD
            .Col = .ColLetterToNumber("L")
            .Row = 15
            .Text = strLoaiNNKD
            UpdateCell fps, .Col, .Row, .Text
         'end
         
         
         .GetText .ColLetterToNumber("F"), 34, chkValue
        .Col = .ColLetterToNumber("K")
        .Row = 37
        vFormular = .Formula
        If chkValue = "0" Then
            .Col = .ColLetterToNumber("K")
            .Row = 34
            If .CellType = 1 Or .CellType = 5 Then
                FormatTextPercent fps, 1, .ColLetterToNumber("K"), 34, True
                '.SetText .ColLetterToNumber("K"), .Row, "0"
            End If
            .Col = .ColLetterToNumber("K")
            .Row = 37
            .Lock = True
            .Formula = "ROUND(K31*20/100,0) + ROUND(K32*22/100,0) +ROUND(K33*K34/100,0) + ROUND(K35*K36/100,0)"
        ElseIf chkValue = "1" Or chkValue = "x" Then
            'FormatText fps, 1, .ColLetterToNumber("J"), 42
            .Col = .ColLetterToNumber("K")
            .Row = 37
            .Lock = False
        Else
            FormatTextPercent fps, 1, .ColLetterToNumber("K"), 34, False
            .Col = .ColLetterToNumber("K")
            .Row = 37
            .Lock = True
            .Formula = "ROUND(K31*20/100,0) + ROUND(K32*22/100,0) +ROUND(K33*K34/100,0) + ROUND(K35*K36/100,0)"
        End If
         
         
         ' Neu check gia han moi cho phep chon GT
        .Col = .ColLetterToNumber("F")
        .Row = 43
        If .Text = "1" Or UCase$(.Text) = "X" Then
            .Col = .ColLetterToNumber("H")
            .Row = 45
            .Lock = False
            
            .Col = .ColLetterToNumber("H")
            .Row = 47
            .Lock = False
            
            .Col = .ColLetterToNumber("H")
            .Row = 49
            .Lock = False
            
            .Col = .ColLetterToNumber("H")
            .Row = 51
            .Lock = False
        Else
            .Col = .ColLetterToNumber("H")
            .Row = 45
            .Lock = True
            
            .Col = .ColLetterToNumber("H")
            .Row = 47
            .Lock = True
            
            .Col = .ColLetterToNumber("H")
            .Row = 49
            .Lock = True
            
            .Col = .ColLetterToNumber("H")
            .Row = 51
            .Lock = True
        End If
        ' End set gia han
         
        'set lai stt cho cac cot dong sheet 2
         .Sheet = 2
         If .SheetVisible = True Then
                    i = 0
                    .Col = .ColLetterToNumber("C")
                    .Row = 35
                    Do
                        i = i + 1
                        .Text = str(i)
                        .Col = .ColLetterToNumber("C")
                        .Row = .Row + 1
                    Loop Until .Text = "aa"
            End If
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
    Dim lResult As Integer
    Dim blTemp As Boolean
    
    Dim i As Integer, iCol As Long, iRow As Long
    
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
    
    Dim ct26 As Variant, ct27 As Variant, ct28 As Variant, vCheck20 As Variant
    
    fps.EventEnabled(EventAllEvents) = False
' set gia lock cell ly do gia han
    If fps.ActiveSheet = 1 Then
'             ' neu check thue suat 20 % thi set chi tieu 30 = 20%
'            If (Col = fps.ColLetterToNumber("D") And Row = 11) Then
'                fps.Sheet = fps.ActiveSheet
'                fps.Col = Col
'                fps.Row = Row
'                If (fps.value = 1 Or UCase$(fps.value) = "X") Then
'                    fps.Row = 31
'                    fps.Col = fps.ColLetterToNumber("K")
'                    fps.Text = 20
'                    fps.Lock = True
'                    UpdateCell fps, fps.Col, fps.Row, fps.Text
'                    ' kiem tra rang buoc
'                    fps.Sheet = 1
'                    fps.GetText fps.ColLetterToNumber("K"), 27, CT26
'                    fps.GetText fps.ColLetterToNumber("K"), 28, CT27
'                    fps.GetText fps.ColLetterToNumber("K"), 29, CT28
'                    fps.GetText fps.ColLetterToNumber("D"), 11, vCheck20
'                    ' set chi tieu 28
'                    'sua bat loi theo chi tieu [26]<0 thi lock 28-32
'                    ' set lai gia tri 30,31,32
'                    If Val(CT26) < 0 Then
'                       '30
'                       fps.Row = 31
'                       fps.Col = fps.ColLetterToNumber("K")
'                       fps.value = 0
'                       fps.Lock = True
'                       UpdateCell fps, fps.Col, fps.Row, 0
'                    Else
'    '                    .Row = 29
'    '                   .Col = .ColLetterToNumber("K")
'    '                   .Lock = False
'                        If Val(CT26) - Val(CT27) - Val(CT28) >= 0 Then
'                            fps.Row = 31
'                            fps.Col = fps.ColLetterToNumber("K")
'                            ' set lai thue suat
'                            If vCheck20 = "1" Or UCase$(CStr(vCheck20)) = "X" Then
'                                fps.value = 20
'                                fps.Lock = True
'                                UpdateCell fps, fps.Col, fps.Row, fps.Text
'                            End If
'                        Else
'                            fps.Row = 31
'                            fps.Col = fps.ColLetterToNumber("K")
'                            ' set lai thue suat
'                            If vCheck20 = "1" Or UCase$(CStr(vCheck20)) = "X" Then
'                                fps.value = 0
'                                fps.Lock = True
'                                UpdateCell fps, fps.Col, fps.Row, fps.Text
'                            End If
'                        End If
'                    End If
'                Else
'                    fps.Row = 31
'                    fps.Col = fps.ColLetterToNumber("K")
'                    fps.Text = 0
'                    fps.Lock = False
'                    UpdateCell fps, fps.Col, fps.Row, fps.Text
'                End If
'            End If
             ' check nhieu muc thue uu dai
             If Col = fps.ColLetterToNumber("F") And Row = 34 Then
                fps.Sheet = fps.ActiveSheet
                fps.Row = Row
                fps.Col = Col
                If fps.value = 1 Then 'Khi danh dau Checkbox
                    'Hien thong bao co muon xoa du lieu cu
                    lResult = DisplayMessage("0294", msYesNo, miQuestion, , mrNo)
                    If lResult = 3 Then 'Khong muon xoa thi thoat ra
                        fps.value = 0
                        UpdateCell fps, fps.Col, fps.Row, IIf(fps.value = 1, "x", vbNullString)
                        
                        fps.EventEnabled(EventAllEvents) = True
                        Exit Sub
                    Else 'Khi chon xoa
                        'Xoa du lieu chi tieu 15, 18 va cho phep nhap cot 18, format lai cot 15
                        
                        Set xmlCellNode = TAX_Utilities_v2.Data(0).nodeFromID(GetCellID(fps, fps.ColLetterToNumber("K"), 34))
                        Set xmlCellsNode = xmlCellNode.parentNode
                        UpdateCell fps, fps.ColLetterToNumber("F"), fps.Row, "x"
                        Do
                            For i = 1 To xmlCellsNode.childNodes.length
                                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
                                Select Case iCol
                                    Case fps.ColLetterToNumber("K")
                                        'Xoa trang chi tieu 32a
                                        fps.Row = 34
                                        fps.Col = iCol
                                        fps.Text = ""
                                        UpdateCell fps, iCol, 34, fps.Text
                                        FormatText fps, 1, fps.ColLetterToNumber("K"), 34
                                        
                                        'Xoa trang chi tieu 34
                                        fps.Row = 37
                                        fps.Col = iCol
                                        fps.Formula = ""
                                        fps.Text = "0"
                                        fps.value = "0"
                                        fps.Lock = False ' Bo enable cho phep nhap
                                        UpdateCell fps, iCol, 37, fps.value
                                End Select
                            Next
                            If xmlCellsNode.nextSibling Is Nothing Then Exit Do
                            Set xmlCellsNode = xmlCellsNode.nextSibling
                        Loop Until xmlCellsNode Is Nothing
                    End If
                ElseIf fps.value = 0 Then 'Huy danh dau Checkbox
                        
                    Set xmlCellNode = TAX_Utilities_v2.Data(1 - 1).nodeFromID(GetCellID(fps, fps.ColLetterToNumber("K"), 34))
                    Set xmlCellsNode = xmlCellNode.parentNode
                    UpdateCell fps, fps.ColLetterToNumber("F"), fps.Row, ""
                    Do
                        For i = 1 To xmlCellsNode.childNodes.length
                            ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
                            Select Case iCol
                                Case fps.ColLetterToNumber("K")
                                    ' Dinh dang lai cho chi tieu 32a (Cho phep nhap nhieu muc thue)
                                    fps.Row = 34
                                    fps.Col = iCol
                                    fps.value = 0
                                    UpdateCell fps, iCol, 34, fps.value
                                    FormatTextPercent fps, 1, fps.ColLetterToNumber("K"), 34, False
                                
                                    ' Thiet lap lai cong thuc cho chi tieu 34
                                    fps.Row = 37
                                    fps.Col = iCol
                                    fps.Formula = "ROUND(K31*20/100,0) + ROUND(K32*22/100,0) +ROUND(K33*K34/100,0) + ROUND(K35*K36/100,0)"
'                                    .Formula = vFormular
                                    fps.Lock = True
                                    UpdateCell fps, iCol, 50, fps.value
                            End Select
                        Next
                        If xmlCellsNode.nextSibling Is Nothing Then Exit Do
                        Set xmlCellsNode = xmlCellsNode.nextSibling
                    Loop Until xmlCellsNode Is Nothing
                End If
            End If

            ' Neu check gia han moi cho phep chon GT
            If Col = fps.ColLetterToNumber("F") And Row = 43 Then
                fps.Col = fps.ColLetterToNumber("F")
                fps.Row = 43
                If fps.Text = "1" Or UCase$(fps.Text) = "X" Then
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 45
                    fps.Lock = False
                    
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 47
                    fps.Lock = False
                    
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 49
                    fps.Lock = False
                    
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 51
                    fps.Lock = False
                    fps.Formula = "K41-H49"
                Else
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 45
                    fps.Lock = True
                    
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 47
                    fps.Lock = True
                    
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 49
                    fps.Lock = True
                    
                    fps.Col = fps.ColLetterToNumber("H")
                    fps.Row = 51
                    fps.Lock = True
                    fps.Formula = ""
                    
                    fps.SetText fps.ColLetterToNumber("H"), 45, ""
                    fps.SetText fps.ColLetterToNumber("P"), 45, "00"
                    UpdateCell fps, fps.ColLetterToNumber("P"), 45, "00"
                    UpdateCell fps, fps.ColLetterToNumber("H"), 45, ""
                    
                    fps.SetText fps.ColLetterToNumber("H"), 47, ""
                    fps.SetText fps.ColLetterToNumber("H"), 49, ""
                    fps.SetText fps.ColLetterToNumber("H"), 51, ""
                    UpdateCell fps, fps.ColLetterToNumber("H"), 47, ""
                    UpdateCell fps, fps.ColLetterToNumber("H"), 49, ""
                    UpdateCell fps, fps.ColLetterToNumber("H"), 51, ""
                    
                End If
            End If
    End If
    
    With fps

            If Col = .ColLetterToNumber("G") And Row >= 42 And .ActiveSheet = 3 Then
                   fps.EventEnabled(EventAllEvents) = False
                   .Sheet = 23
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
Dim i As Integer
Dim xmlNodeH15 As MSXML.IXMLDOMNode

Dim ct26, ct27, ct28, ct29, ct30, ct31, ct32, ct32a, CT33, ct33a, ct34, CT35, ct35a, ct35b, CT36 As Variant
'Dim vCheck20 As Variant

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
                    .SetText iCol, iRow, 0
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
            .GetText .ColLetterToNumber("J"), 19, varTemp
            .Sheet = .SheetCount - 1
            Set xmlNodeH15 = TAX_Utilities_v2.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
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
       
       If .ActiveSheet = 1 Then
            fps.EventEnabled(EventAllEvents) = False
            If Col = .ColLetterToNumber("K") And (Row = 27 Or Row = 30) Then
                ' 12062011
                ' set lai cac rang buoc cong thuc
                .Sheet = 1
                .GetText .ColLetterToNumber("K"), 27, ct26
                .GetText .ColLetterToNumber("K"), 28, ct27
                .GetText .ColLetterToNumber("K"), 29, ct28
                .GetText .ColLetterToNumber("K"), 30, ct29
                .GetText .ColLetterToNumber("K"), 31, ct30
                .GetText .ColLetterToNumber("K"), 32, ct31
                .GetText .ColLetterToNumber("K"), 33, ct32
                .GetText .ColLetterToNumber("K"), 34, ct32a
                .GetText .ColLetterToNumber("K"), 35, CT33
                .GetText .ColLetterToNumber("K"), 36, ct33a
                .GetText .ColLetterToNumber("K"), 37, ct34
                .GetText .ColLetterToNumber("K"), 38, CT35
                .GetText .ColLetterToNumber("K"), 39, ct35a
                .GetText .ColLetterToNumber("K"), 40, ct35b
                .GetText .ColLetterToNumber("K"), 41, CT36
                
                '.GetText .ColLetterToNumber("D"), 11, vCheck20
                
                ' set chi tieu 28
                'sua bat loi theo chi tieu [26]<0 thi lock 28-32
                ' set lai gia tri 30,31,32
                If Val(ct26) <= 0 Then
                    '28
                   .Row = 29
                   .Col = .ColLetterToNumber("K")
                   .value = 0
                   .Lock = True
                   UpdateCell fps, .Col, .Row, 0
                   '30
                   .Row = 31
                   .Col = .ColLetterToNumber("K")
                   .value = 0
                   .Lock = True
                   UpdateCell fps, .Col, .Row, 0
                   '31
                   .Row = 32
                   .Col = .ColLetterToNumber("K")
                   .value = 0
                   .Lock = True
                   UpdateCell fps, .Col, .Row, 0
                   '32
                   .Row = 33
                   .Col = .ColLetterToNumber("K")
                   .value = 0
                   .Lock = True
                   UpdateCell fps, .Col, .Row, 0
                   ' 32a
                   .Row = 34
                   .Col = .ColLetterToNumber("K")
                   .value = 0
                   .Lock = True
                   UpdateCell fps, .Col, .Row, 0
                   '33
                   .Row = 35
                   .Col = .ColLetterToNumber("K")
                   .value = 0
                   .Lock = True
                   UpdateCell fps, .Col, .Row, 0
                    ' 33a
                   .Row = 36
                   .Col = .ColLetterToNumber("K")
                   .value = 0
                   .Lock = True
                   UpdateCell fps, .Col, .Row, 0
                   ' 35
                   .Row = 38
                   .Col = .ColLetterToNumber("K")
                   .value = 0
                   .Lock = True
                   UpdateCell fps, .Col, .Row, 0
                   '35a
                   .Row = 39
                   .Col = .ColLetterToNumber("K")
                   .value = 0
                   .Lock = True
                   UpdateCell fps, .Col, .Row, 0
                   '35b
                   .Row = 40
                   .Col = .ColLetterToNumber("K")
                   .value = 0
                   .Lock = True
                   UpdateCell fps, .Col, .Row, 0
                Else
                    If Val(ct26) - Val(ct27) < 0 Then
                        '28
                       .Row = 29
                       .Col = .ColLetterToNumber("K")
                       .value = 0
                       .Lock = True
                       UpdateCell fps, .Col, .Row, 0
                    Else
                        '28
                       .Row = 29
                       .Col = .ColLetterToNumber("K")
                       .Lock = False
                    End If
'                    .Row = 29
'                   .Col = .ColLetterToNumber("K")
'                   .Lock = False
                    If Val(ct26) - Val(ct27) - Val(ct28) < 0 Then
                        '30
                        .Row = 31
                        .Col = .ColLetterToNumber("K")
                        .value = 0
                        .Lock = True
                        '31
                        .Row = 32
                        .Col = .ColLetterToNumber("K")
                        .value = 0
                        .Lock = True
                        '32
                        .Row = 33
                        .Col = .ColLetterToNumber("K")
                        .value = 0
                        .Lock = True
                        
                        '35
                        .Row = 38
                        .Col = .ColLetterToNumber("K")
                        .value = 0
                        .Lock = True
                        '35a
                        .Row = 39
                        .Col = .ColLetterToNumber("K")
                        .value = 0
                        .Lock = True
                        '35b
                        .Row = 40
                        .Col = .ColLetterToNumber("K")
                        .value = 0
                        .Lock = True
                    Else
                        ' 30
                        .Row = 31
                        .Col = .ColLetterToNumber("K")
                        .Lock = False
                        ' set lai thue suat
'                        If vCheck20 = "1" Or UCase$(CStr(vCheck20)) = "X" Then
'                            .value = 20
'                            .Lock = True
'                        End If
                        ' 31
                        .Row = 32
                        .Col = .ColLetterToNumber("K")
                        .Lock = False
                        '32
                        .Row = 33
                        .Col = .ColLetterToNumber("K")
                        .Lock = False
                        
                        '35
                        .Row = 38
                        .Col = .ColLetterToNumber("K")
                        .Lock = False
                        '35a
                        .Row = 39
                        .Col = .ColLetterToNumber("K")
                        .Lock = False
                        '35b
                        .Row = 40
                        .Col = .ColLetterToNumber("K")
                        .Lock = False
                    End If
                    
                        '32a
                        .Row = 34
                        .Col = .ColLetterToNumber("K")
                        .Lock = False
                        '33
                        .Row = 35
                        .Col = .ColLetterToNumber("K")
                        .Lock = False
                        '33a
                        .Row = 36
                        .Col = .ColLetterToNumber("K")
                        .Lock = False
                End If
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
Dim strId As Variant, Col7 As Variant, intIndexCombo As Integer
Dim NguoiKyCol As Long, NguoiKyRow As Long
    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 2 Then
                   fps.EventEnabled(EventAllEvents) = True
                    i = 0
                    .Col = .ColLetterToNumber("C")
                    .Row = 35
                    Do
                         i = i + 1
                         .Text = str(i)
                         .Col = .ColLetterToNumber("C")
                         .Row = .Row + 1
                    Loop Until .Text = "aa"
                    .SetActiveCell iCol, iRow

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
    Dim tCol As Long, tRow As Long
    Dim cCol As Long, cRow As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    With fps
    mCurrentSheet = .Sheet
    .EventEnabled(EventAllEvents) = False
        If .ActiveSheet = 1 Then
'            .GetText .ColLetterToNumber("J"), 16, varTemp
'            If varTemp < 0 Then
'                .Row = 17
'                .Col = .ColLetterToNumber("J")
'                .Text = "0"
'                .Lock = True
'                UpdateCell fps, .Col, .Row, .Text
'
'                .Row = 18
'                .Col = .ColLetterToNumber("J")
'                .Text = "0"
'                .Lock = True
'
'                UpdateCell fps, .Col, .Row, .Text
'                .Row = 19
'                .Col = .ColLetterToNumber("J")
'                .value = 0
'                .Lock = True ' sua 12032011 theo yeu cau cua ptly
'                UpdateCell fps, .Col, .Row, .value
'            Else
'                .Row = 17
'                .Col = .ColLetterToNumber("J")
'                .Lock = False
'                UpdateCell fps, .Col, .Row, .value
'
'                .Row = 18
'                .Col = .ColLetterToNumber("J")
'                .Lock = False
'                UpdateCell fps, .Col, .Row, .value
'
'                .Row = 19
'                .Col = .ColLetterToNumber("J")
'                .Lock = False
'                UpdateCell fps, .Col, .Row, .value
'            End If
        End If
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("J") And Row = 56 Then
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
          
          
            If Col = .ColLetterToNumber("H") And Row = 47 Then
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
'          ' Update so chenh lech cho KHBS
'            .GetText .ColLetterToNumber("J"), 19, varTemp
'            If TAX_Utilities_v2.NodeValidity.childNodes(.SheetCount - 2).Attributes.getNamedItem("Active").nodeValue <> "0" Then
'                .Sheet = .SheetCount - 1
'                Set xmlNode = TAX_Utilities_v2.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
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
'                            .SetFocus
'                            .SetActiveCell Col, Row
'                           End If
'                       Else
'                        .SetText Col, Row, ""
'                       End If
'                      .Col = Col
'                      .Row = Row
'                      UpdateCell fps, .Col, .Row, .Text
'                   End If
'                .Sheet = .SheetCount - 1
'
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
'            mCurrentSheet = .Sheet
'
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
       
       
    ElseIf .ActiveSheet = 1 Then
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("H") And Row = 45 Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("P")
            .typeComboboxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
        End If
    ElseIf .ActiveSheet = 2 Then
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("G") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("S")
            .typeComboboxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        If Col = .ColLetterToNumber("H") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("R")
            .typeComboboxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
        End If
    ElseIf .ActiveSheet = 3 Then
        If Col = .ColLetterToNumber("P") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("AH")
            .typeComboboxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
        End If
    End If
   End With
   TAX_Utilities_v2.AdjustData(mCurrentSheet - 1) = True
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
        
        .Sheet = 1
        .Col = .ColLetterToNumber("F")
        .Row = 34
        If .value <> "chk" Then
            If .value = 0 Then
                FormatTextPercent fps, 1, .ColLetterToNumber("K"), 34, False
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
    Dim iFlagCheckMaST As Integer, iFlagCheckMaSTDN As Integer, iFlagCheckMaSTDN10 As Integer
    Dim strCheck As String
    Dim blCheck_S2 As Boolean
    Dim vFlag As Variant, vErrMsg As Variant, errTongTLPB As Variant, errMST As Variant, errMSTDN As Variant, errEmty As Variant, iFlagTongPhanBo As Variant, errMst10 As Variant
    Dim varTemp As Variant, vMST As Variant
       ' PL01-2
    Dim vTyLe As Variant
    Dim iWhiteRow As Integer
    Dim iFlagEmpty As Integer
    Dim vTenDn As Variant, vCQTCapCuc As Variant, vCqtChiCuc As Variant, vSoThue As Variant
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer
    Dim errTongSoThue As Variant
    Dim iFlagTyLe  As Integer
    Dim iFlagCheckSoThue As Integer
    Dim vTongTyLe As Double
    Dim vTongSoThue As Double
    Dim vCt5 As Variant
    Dim vColor As Variant
    
    ' kiem tra nam dc gh
    Dim Quynamtc As Integer
    Dim Namtc As Integer
    Dim QuyKK As Integer
    Dim NamKK As Integer
    Dim iNgayTaiChinh As Integer
    Dim iThangTaiChinh As Integer
    Dim strNgayTaiChinh As String
        
    Dim isCheckGh As Boolean
    ' end
    Dim arrDateGH() As String
    Dim arrDateHanNop() As String
    Dim vErrorCheckHN As Variant
    
    
    Dim vCT26 As Variant, arrCT26() As String
    Dim errSoThueUuDai As Variant, checkboxThue As Variant
    Dim iFlagSoUuDai As Integer
    Dim i As Integer
    ' PL 05
    Dim mErrMSTTrung As Variant
    ' check gh
    Dim vCheckGH As Variant, vLyDoGH As Variant, vErrorCheckGH As Variant, vThoiHanGH As Variant, vErrorCheckThoiHanGH As Variant, vSoThueGH As Variant, vErrorCheckSoThueGH As Variant
    Dim iflagCheckGH As Integer, iflagCheckThoiHanGH As Integer, iflagCheckSoThueGH As Integer
    
    With fps
        '.Visible = False
        '.ReDraw = False
        .Sheet = .SheetCount
        
        .GetText .ColLetterToNumber("E"), 19, vErrMsg
        .GetText .ColLetterToNumber("E"), 21, errTongTLPB
        .GetText .ColLetterToNumber("E"), 23, errMST
        .GetText .ColLetterToNumber("E"), 22, errEmty
        .GetText .ColLetterToNumber("E"), 25, errMSTDN
        .GetText .ColLetterToNumber("E"), 27, errMst10
        
        
        .GetText .ColLetterToNumber("E"), 32, mErrMSTTrung
        .GetText .ColLetterToNumber("E"), 33, vErrorCheckGH
        .GetText .ColLetterToNumber("E"), 34, errTongSoThue
        
        .GetText .ColLetterToNumber("E"), 39, vErrorCheckThoiHanGH
        .GetText .ColLetterToNumber("E"), 42, vErrorCheckSoThueGH
        
        .GetText .ColLetterToNumber("E"), 25, errSoThueUuDai
        .GetText .ColLetterToNumber("E"), 44, vErrorCheckHN
        
        
        vFlag = 0
        iFlagCheckMaST = 0
        iFlagCheckMaSTDN = 0
        iFlagCheckMaSTDN10 = 0
        iFlagTongPhanBo = 0
        .Sheet = 1
        mCurrentSheet = .Sheet
        .GetText .ColLetterToNumber("J"), 56, varTemp
        .Col = .ColLetterToNumber("J")
        .Row = 56
        If varTemp <> "" And varTemp <> "../../...." Then
            If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                .SetText .Col, .Row, Format_ddmmyyyy(CStr(varTemp))
            Else
                .Col = .ColLetterToNumber("J")
                .Row = 56
                .SetFocus
                .SetActiveCell .Col, .Row
            End If
        Else
            .SetText .Col, .Row, ""
        End If
           UpdateCell fps, .Col, .Row, varTemp
       
         
         
         ' check ly do gia han
         ' cac ky khong check ly do gia han
        QuyKK = TAX_Utilities_v2.ThreeMonths
        NamKK = TAX_Utilities_v2.Year
        
        
        'Lay ngay bat dau nam tai chinh
        If GetAttribute(TAX_Utilities_v2.NodeMenu, "FinanceYear") = "1" Then
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
         isCheckGh = True
         ' end
        .GetText .ColLetterToNumber("F"), 43, vCheckGH
        .GetText .ColLetterToNumber("H"), 45, vLyDoGH
        .GetText .ColLetterToNumber("H"), 47, vThoiHanGH
        .GetText .ColLetterToNumber("H"), 49, vSoThueGH
        
         If (Trim$(vCheckGH) = "1" Or Trim$(UCase$(vCheckGH)) = "X") And isCheckGh Then
             ' Kiem tra ly do gia han
             If Trim$(vLyDoGH) = "" Then
                 .Col = .ColLetterToNumber("H")
                 .Row = 45
                 .CellNote = vErrorCheckGH  'static
                 .BackColor = mErrorColor
                 iflagCheckGH = 1
             Else
                 .Col = .ColLetterToNumber("H")
                 .Row = 45
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckGH = 0
             End If
             ' Kiem tra thoi han gia han
             If Trim$(vThoiHanGH) = "" Or Trim$(vThoiHanGH) = "../../...." Then
                 .Col = .ColLetterToNumber("H")
                 .Row = 47
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
                    .Row = 47
                    .CellNote = vErrorCheckHN  'static
                    .BackColor = mAlertColor
                Else
                    .Col = .ColLetterToNumber("H")
                    .Row = 47
                    .CellNote = ""  'static
                    .BackColor = mNonErrorColor
                 End If
             End If
             ' Kiem tra so thue
             If Val(vSoThueGH) <= 0 Then
                 .Col = .ColLetterToNumber("H")
                 .Row = 49
                 .CellNote = vErrorCheckSoThueGH
                 .BackColor = mErrorColor
                 iflagCheckSoThueGH = 1
             Else
                 .Col = .ColLetterToNumber("H")
                 .Row = 49
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckSoThueGH = 0
             End If
             
         Else
                ' Ly do gia han
                 .Col = .ColLetterToNumber("H")
                 .Row = 45
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckGH = 0
                 ' Thoi han gia han
                 .Col = .ColLetterToNumber("H")
                 .Row = 47
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckThoiHanGH = 0
                 ' So thue gia han
                 .Col = .ColLetterToNumber("H")
                 .Row = 49
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckSoThueGH = 0
         End If
       
       
        
        'check dinh dang so trong chi tieu 32a
       iFlagSoUuDai = 0
       
        .GetText .ColLetterToNumber("F"), 34, checkboxThue
        If (Val(checkboxThue) = 1) Then
     
            .GetText .ColLetterToNumber("K"), 34, vCT26
            arrCT26 = Split(CStr(vCT26), ";")
            For i = 0 To UBound(arrCT26) - 1
                If IsNumeric(arrCT26(i)) = False Then
                    iFlagSoUuDai = 1
                End If
            Next
            If iFlagSoUuDai = 1 Then
                .Col = .ColLetterToNumber("K")
                .Row = 34
                .CellNote = errSoThueUuDai
                .BackColor = mErrorColor
            Else
                .Col = .ColLetterToNumber("K")
                .Row = 34
                .BackColor = mNonErrorColor
                .CellNote = ""
            End If
        ElseIf (Val(checkboxThue) = 0) Then
                .Col = .ColLetterToNumber("K")
                .Row = 34
                .BackColor = mNonErrorColor
                .CellNote = ""
        End If
       
       
        .Sheet = 2
        If .SheetVisible = True Then

            mCurrentSheet = .Sheet
            iFlagEmpty = 0
            iFlagTyLe = 0
            iFlagCheckMaST = 0
            iFlagCheckSoThue = 0
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
        'end
        blCheck_S2 = IIf(TAX_Utilities_v2.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        .Sheet = .SheetCount
        If blCheck_S2 Then
            .SetText .ColLetterToNumber("B"), 21, IIf(iFlagTyLe = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 35, IIf(iFlagCheckMaSTDN10 = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 34, IIf(iFlagCheckSoThue = 1, "0", "1")
        Else
            .SetText .ColLetterToNumber("B"), 21, "1"
            .SetText .ColLetterToNumber("B"), 34, "1"
            .SetText .ColLetterToNumber("B"), 35, "1"
        End If
        
        .SetText .ColLetterToNumber("B"), 22, IIf(iFlagEmpty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 23, IIf(iFlagCheckMaST = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 25, IIf(iFlagCheckMaSTDN = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 33, IIf(iflagCheckGH = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 39, IIf(iflagCheckThoiHanGH = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 42, IIf(iflagCheckSoThueGH = 1, "0", "1")
        
        .SetText .ColLetterToNumber("B"), 43, IIf(iFlagSoUuDai = 1, "0", "1")
            
        CheckErrorMST
        
        '.ReDraw = True
        '.Visible = True
        
        ' Kiem tra phu luc phan bo
        .Sheet = .SheetCount
        
        Dim vChon As Variant
        Dim vTenNhaMay, vCQTCapDiaPhuong, vTyLePhanBo, vSoThuePhaiNop As Variant
        Dim vMessRequire As Variant
        
        Dim flagSoThue As Integer
        
        .GetText .ColLetterToNumber("E"), 45, vMessRequire
        .GetText .ColLetterToNumber("E"), 47, errMST
        
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
        .SetText .ColLetterToNumber("B"), 45, IIf(iFlagEmpty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 47, IIf(iFlagCheckMaST = 1, "0", "1")
        
        .SetText .ColLetterToNumber("B"), 48, IIf(flagSoThue = 1, "0", "1")
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
                Set xmlCellNode = TAX_Utilities_v2.Data(3).getElementsByTagName("Cell") _
                    (TAX_Utilities_v2.Data(3).getElementsByTagName("Cell").length - 1)
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
'        .Row = 11
'        .Col = .ColLetterToNumber("D")
'        .value = 0
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
    
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT.xml"
    
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
        
' Cho phep nhap am cac chi tieu 10, 11, 13, 14, 15
        .Sheet = 1
        .Col = .ColLetterToNumber("J")
        .Row = 10
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Row = 11
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Row = 13
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Row = 14
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Row = 15
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
        For Each xmlCellNodeData In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
            ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
            If (lRow <> 20) And (lCol <> .ColLetterToNumber("B")) Then
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
            End If
        Next
        .Sheet = mCurrentSheet
    End With
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub


Private Sub UpdateSheets()
    Dim varTemp As Variant
    Dim ssSheet As Integer
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    
    Dim blCheck_S2 As Boolean
    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet
        For Each xmlCellNodeData In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
           ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               If (lRow = 12 And lCol = .ColLetterToNumber("D")) Or (lRow = 11 And lCol = .ColLetterToNumber("D")) Or (lRow = 43 And lCol = .ColLetterToNumber("F")) Or (lRow = 34 And lCol = .ColLetterToNumber("F")) Then
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
        .Sheet = 2
        mCurrentSheet = .Sheet
        blCheck_S2 = IIf(TAX_Utilities_v2.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
        
        .Sheet = 3
        mCurrentSheet = .Sheet
        
        blCheck_S2 = IIf(TAX_Utilities_v2.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(2).getElementsByTagName("Cell")
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
        blCheck_S2 = IIf(TAX_Utilities_v2.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
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
            Set xmlNode = .nodeFromID("K_37")
            result = "CT34_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("K_38")
            result = result & "~" & "CT35_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01ATNDN", "getValueTK", Err.number, Err.Description
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
                .GetText .ColLetterToNumber("K"), 37, varTemp
                strCurrentValue = "CT34_" & Trim(varTemp)
                .GetText .ColLetterToNumber("K"), 38, varTemp
                strCurrentValue = strCurrentValue & "~CT35_" & Trim(varTemp)
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

                If Left(strMaCT, 3) = "35_" Then
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
    SaveErrorLog "cls_01ATNDN", "getMapCT", Err.number, Err.Description
End Function



' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("K"), 41, varTemp
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
            Set xmlNode = .nodeFromID("K_41")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01ATNDN", "getValueCTDC", Err.number, Err.Description
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
                Set xmlCellNode = TAX_Utilities_v2.Data(3).getElementsByTagName("Cell") _
                    (TAX_Utilities_v2.Data(3).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v2.Data(3).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v2.Data(3).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v2.Data(3).getElementsByTagName("Cell").length > 31 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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
    tuRowDL = 23
    denRowDL = 33
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "14_16_2~11_12_3~20_30_4"
    ' end
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
    With fps
        .EventEnabled(EventAllEvents) = False
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
