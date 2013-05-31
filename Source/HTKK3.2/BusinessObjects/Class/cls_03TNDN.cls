VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_03TNDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Quyet toan nam" interface sheets
'this Class is belong to TAX_Business project which will be compline to DLL

Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'Private Const S1TxtB24Row = 32
'Private Const S1TxtB24Col = "G"
'Private Const S1TxtB27Row = 35
'Private Const S1TxtB27Col = "G"
'Private Const S1TxtB25Row = 33
'Private Const S1TxtB25Col = "G"
'Private Const S1TxtB28Row = 36
'Private Const S1TxtB28Col = "G"
'Private Const S1TxtB30Row = 38
'Private Const S1TxtB30Col = "G"
'Private Const S1TxtB31Row = 39
'Private Const S1TxtB31Col = "G"
'Private Const S1TxtC2Row = 42
'Private Const S1TxtC2Col = "G"
Private Const S1TxtC8Row = 47
Private Const S1TxtC8Col = "O"
Private Const S1TxtC9Row = 48
Private Const S1TxtC9Col = "O"
'Private Const S1TxtC6Row = 46
'Private Const S1TxtC6Col = "G"
' Cac bien de sinh tu dong cac nam duoc chuyen lo, 6 nam lien nhau (PL 03-2/TNDN)
'Private Const S2TxtF5Row = 5
'Private Const S2TxtF5Col = "F"
'Private Const S2TxtI5Row = 5
'Private Const S2TxtI5Col = "I"
'Private Const S2TxtL5Row = 5
'Private Const S2TxtL5Col = "L"
'Private Const S2TxtO5Row = 5
'Private Const S2TxtO5Col = "O"
'Private Const S2TxtR5Row = 5
'Private Const S2TxtR5Col = "R"
Private Const S2TxtB26Row = 26
Private Const S2TxtB26Col = "B"
Private Const S2TxtB27Row = 27
Private Const S2TxtB27Col = "B"
Private Const S2TxtB28Row = 28
Private Const S2TxtB28Col = "B"
Private Const S2TxtB29Row = 29
Private Const S2TxtB29Col = "B"
Private Const S2TxtB30Row = 30
Private Const S2TxtB30Col = "B"


Private Const S5TxtU39Row = 39
Private Const S5TxtU39Col = "U"

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


' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'no parameter
Public Sub Prepare1()
    Dim arrResul() As String
    Dim blCheck_S05 As Boolean
    
    Dim strDataFileName As String
    Dim strComboHien As String
    Dim strCombo As String
    Dim strCombo1 As String
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim arrDanhsach() As String
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim xmlNodeListItems As MSXML.IXMLDOMNodeList
    Dim strCboLyDo As String
    Dim strCboLyDoAn As String
    Dim i As Integer
    Dim fldList() As String

    Dim arrRowCom() As String
    Dim strTemp As String
    
    
With fps
    ' TK 03
    SetDateFormat fps, 1, 67, .ColLetterToNumber("N"), DDMMYYYY
    .Sheet = 1 ' Mac dinh ngay ky la ngay he thong (To khai quyet toan TNDN 03/TNDN)
    .Row = 67
    .Col = .ColLetterToNumber("N")
    .Text = Format(Date, "dd/mm/yyyy")
    .TypeHAlign = TypeHAlignLeft
    If TAX_Utilities_New.Year = 2009 Then
        .SetText .ColLetterToNumber("C"), 62, GetAttribute(GetMessageCellById("0185"), "Msg")
    ElseIf TAX_Utilities_New.Year = 2010 Then
        .SetText .ColLetterToNumber("C"), 62, GetAttribute(GetMessageCellById("0187"), "Msg")
    ElseIf TAX_Utilities_New.Year = 2011 Then
        .SetText .ColLetterToNumber("C"), 62, GetAttribute(GetMessageCellById("0243"), "Msg")
'    ElseIf TAX_Utilities_New.Year = 2013 Then
'        .SetText .ColLetterToNumber("C"), 62, GetAttribute(GetMessageCellById("0263"), "Msg")
'    ElseIf Val(TAX_Utilities_New.Year) = 2012 Or Val(TAX_Utilities_New.Year) > 2013 Then
    Else
         .SetText .ColLetterToNumber("C"), 62, GetAttribute(GetMessageCellById("0263"), "Msg")
'        .SetText .ColLetterToNumber("C"), 62, ""
'        .Row = 62
'        .Col = .ColLetterToNumber("B")
'        .Text = ""
'        .Lock = True
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
            .Row = 63
            .Col = .ColLetterToNumber("D")
            .TypeComboBoxList = strCboLyDo
            ' set Ma ly do gia han
            .Col = .ColLetterToNumber("P")
            .TypeComboBoxList = strCboLyDoAn
    End If
    
    
    ' PL 03-3A
    FormatTextNumber fps, 7, .ColLetterToNumber("N"), 38
    SetDateFormat fps, 7, 38, .ColLetterToNumber("U"), yyyy
    FormatTextNumber fps, 7, .ColLetterToNumber("J"), 39
    SetDateFormat fps, 7, 39, .ColLetterToNumber("Q"), yyyy
    FormatTextNumber fps, 7, .ColLetterToNumber("N"), 40
    SetDateFormat fps, 7, 40, .ColLetterToNumber("U"), yyyy
    ' PL 03-3B
    FormatTextNumber fps, 8, .ColLetterToNumber("N"), 44
    SetDateFormat fps, 8, 44, .ColLetterToNumber("U"), yyyy
    FormatTextNumber fps, 8, .ColLetterToNumber("J"), 45
    SetDateFormat fps, 8, 45, .ColLetterToNumber("Q"), yyyy
    FormatTextNumber fps, 8, .ColLetterToNumber("N"), 46
    SetDateFormat fps, 8, 46, .ColLetterToNumber("U"), yyyy
    SetDateFormat fps, 8, 41, .ColLetterToNumber("S"), DDMMYYYY
    ' PL 03-3C
    SetDateFormat fps, 9, 27, .ColLetterToNumber("V"), DDMMYYYY
    SetDateFormat fps, 9, 31, .ColLetterToNumber("Y"), DDMMYYYY

    
     SetDateFormat fps, .SheetCount - 1, 23, .ColLetterToNumber("BG"), DDMMYYYY
     
     .Sheet = 10
     .Col = .ColLetterToNumber("AK")
     .Row = 26
     .Formula = "25"
     
    .Sheet = .SheetCount - 1 'KHBS
    .Row = 23
    .Col = .ColLetterToNumber("BG")
    .Text = Format(Date, "dd/mm/yyyy")
    .TypeHAlign = TypeHAlignLeft
    ' Lay gia tri cua cac quy khac de set vao phu luc 05/TNDN
    arrResul = GetValue32_31TK_Quy
    blCheck_S05 = IIf(TAX_Utilities_New.NodeValidity.childNodes(11).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    If blCheck_S05 = True Then
        .Sheet = 12
        mCurrentSheet = .Sheet
        .Row = 34
        ' Quy 1
        .Col = .ColLetterToNumber("O")
        .Text = arrResul(0)
        ' Quy 2
        .Col = .ColLetterToNumber("P")
        .Text = arrResul(1)
        ' Quy 3
        .Col = .ColLetterToNumber("Q")
        .Text = arrResul(2)
        ' Quy 4
        .Col = .ColLetterToNumber("R")
        .Text = arrResul(3)
        
    End If
    
    ' Sheet GCN/QLT
       strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_PPDG_TNDN.xml")
        ' Lay danh muc loai hoa don
        ' 15/11/2010
        If xmlDomData.Load(strDataFileName) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    strCombo = strCombo + arrDanhsach(0) + Chr$(9)
                    strComboHien = strComboHien + arrDanhsach(1) + Chr$(9)
                End If
            Next
        End If
        .Sheet = 14
        strTemp = "37~41~46~50~54~59~63~68~72"
        arrRowCom = Split(strTemp, "~")
        For i = 0 To UBound(arrRowCom)
            .Row = CLng(arrRowCom(i))
            ' PP dinh gia
            .Col = .ColLetterToNumber("Q")
            .TypeComboBoxList = strComboHien
            
            .Col = .ColLetterToNumber("AB")
            .TypeComboBoxList = strComboHien
            
            .Col = .ColLetterToNumber("AJ")
            .TypeComboBoxList = strCombo


            .Col = .ColLetterToNumber("AK")
            .TypeComboBoxList = strCombo

        Next
        SetDateFormat fps, 13, 26, .ColLetterToNumber("C"), yyyy
        
        
    .ActiveSheet = 1
End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'no parameter
Public Sub Prepare2()
    Dim intCtrl As Integer
    Dim strId As Variant
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim i As Long, DataText As Variant
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    'PL 03-3A
    Dim Row_PL_3A_U38 As Long, Col_PL_3A_U38 As Long, Row_PL_3A_Q39 As Long, Col_PL_3A_Q39 As Long, Row_PL_3A_U40 As Long, Col_PL_3A_U40 As Long
    'PL 03-3B
    Dim Row_PL_3B_U44 As Long, Col_PL_3B_U44 As Long, Row_PL_3B_Q45 As Long, Col_PL_3B_Q45 As Long, Row_PL_3B_U46 As Long, Col_PL_3B_U46 As Long, Row_PL_3B_S41 As Long, Col_PL_3B_S41 As Long
    'PL 03-3C
    Dim Row_PL_3C_V27 As Long, Col_PL_3C_V27 As Long, Row_PL_3C_Y31 As Long, Col_PL_3C_Y31 As Long
    '
    Dim lSheet1 As Long, lSheet03_1A As Long, lSheet03_1B As Long, lSheet03_1C As Long, lSheet03_2A As Long
    Dim lSheet03_3A As Long, lSheet03_3B As Long, lSheet03_3C As Long
    
    
    ' kiem tra gia tri chi tieu C4
    Dim chiTieuC4 As Variant

    
    lSheet03_3A = 7
    lSheet03_3B = 8
    lSheet03_3C = 9
    
    
    Dim arrCellTag() As String
    
    Set arrErrorCells = New Scripting.Dictionary
    
    With fps
        .Sheet = 1 'To khai TNDN (SheetName: 03)
        .ActiveSheet = .Sheet
            'Khoi tao trang thai ban dau cho cac chi tieu B24, B25, B30
'            If loaiKHBS <> "frmKHBS_BS" Then
'                CellChange .ColLetterToNumber(S1TxtB24Col), S1TxtB24Row
'                CellChange .ColLetterToNumber(S1TxtB25Col), S1TxtB25Row
'                CellChange .ColLetterToNumber(S1TxtB30Col), S1TxtB30Row
'            End If
        .Sheet = 5 'Phu luc 03_2A TNDN (SheetName: 03_2)
            .ActiveSheet = .Sheet
            'Khoi tao cac truong nam cho phu luc 03_2
            SetData
        .Sheet = 6 'Phu luc 03_2B TNDN (SheetName: 03_2)
            .ActiveSheet = .Sheet
            'Khoi tao cac truong nam cho phu luc 03_2
            SetData
        .Sheet = lSheet03_3A 'Phu luc 03_3A TNDN (SheetName: 03_3A)
            mCurrentSheet = .Sheet
            Set xmlCellNode = TAX_Utilities_New.Data(lSheet03_3A - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("U"), 38))
            Set xmlCellsNode = xmlCellNode.parentNode

            Do
                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(lSheet03_3A, 38, .ColLetterToNumber("U")) - 1), "CellID"), Col_PL_3A_U38, Row_PL_3A_U38
                ValidateDateError fps, lSheet03_3A, Row_PL_3A_U38, Col_PL_3A_U38, yyyy

                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(lSheet03_3A, 39, .ColLetterToNumber("Q")) - 1), "CellID"), Col_PL_3A_Q39, Row_PL_3A_Q39
                ValidateDateError fps, lSheet03_3A, Row_PL_3A_Q39, Col_PL_3A_Q39, yyyy

                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(lSheet03_3A, 40, .ColLetterToNumber("U")) - 1), "CellID"), Col_PL_3A_U40, Row_PL_3A_U40
                ValidateDateError fps, lSheet03_3A, Row_PL_3A_U40, Col_PL_3A_U40, yyyy

                Set xmlCellsNode = xmlCellsNode.nextSibling
            Loop Until xmlCellsNode Is Nothing

        .Sheet = lSheet03_3B 'Phu luc 03_3B TNDN (SheetName: 03_3B)
            mCurrentSheet = .Sheet

            Set xmlCellNode = TAX_Utilities_New.Data(lSheet03_3B - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("U"), 44))
            Set xmlCellsNode = xmlCellNode.parentNode

            Do
                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(lSheet03_3B, 44, .ColLetterToNumber("U")) - 1), "CellID"), Col_PL_3B_U44, Row_PL_3B_U44
                ValidateDateError fps, lSheet03_3B, Row_PL_3B_U44, Col_PL_3B_U44, yyyy

                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(lSheet03_3B, 45, .ColLetterToNumber("Q")) - 1), "CellID"), Col_PL_3B_Q45, Row_PL_3B_Q45
                ValidateDateError fps, lSheet03_3B, Row_PL_3B_Q45, Col_PL_3B_Q45, yyyy

                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(lSheet03_3B, 46, .ColLetterToNumber("U")) - 1), "CellID"), Col_PL_3B_U46, Row_PL_3B_U46
                ValidateDateError fps, lSheet03_3B, Row_PL_3B_U46, Col_PL_3B_U46, yyyy

                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(lSheet03_3B, 41, .ColLetterToNumber("S")) - 1), "CellID"), Col_PL_3B_S41, Row_PL_3B_S41
                ValidateDateError fps, lSheet03_3B, Row_PL_3B_S41, Col_PL_3B_S41, DDMMYYYY

                Set xmlCellsNode = xmlCellsNode.nextSibling
            Loop Until xmlCellsNode Is Nothing

        .Sheet = lSheet03_3C 'Phu luc 03_3C TNDN (SheetName: 03_3C)
            mCurrentSheet = .Sheet

            Set xmlCellNode = TAX_Utilities_New.Data(lSheet03_3C - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("V"), 27))
            Set xmlCellsNode = xmlCellNode.parentNode

            Do
                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(lSheet03_3C, 27, .ColLetterToNumber("V")) - 1), "CellID"), Col_PL_3C_V27, Row_PL_3C_V27
                ValidateDateError fps, lSheet03_3C, Row_PL_3C_V27, Col_PL_3C_V27, DDMMYYYY

                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(lSheet03_3C, 31, .ColLetterToNumber("Y")) - 1), "CellID"), Col_PL_3C_Y31, Row_PL_3C_Y31
                ValidateDateError fps, lSheet03_3C, Row_PL_3C_Y31, Col_PL_3C_Y31, DDMMYYYY

                Set xmlCellsNode = xmlCellsNode.nextSibling
            Loop Until xmlCellsNode Is Nothing

         .Sheet = 1
        mCurrentSheet = .Sheet
        'set to khai bo xung hay chinh thuc
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("E")
            .Row = 14
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 14
            .Text = 0
            UpdateCell fps, .Col, .Row, .Text
         ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("E")
            .Row = 14
            .Text = 0
              UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 14
            .Text = Val(StrSolanBosung)
            UpdateCell fps, .Col, .Row, .Text
            ' set unlock cac cell trong truong hop to khai KHBS
            unLockCell (False)
        End If
       'ma NNKD
        .Col = .ColLetterToNumber("J")
        .Row = 14
        .Text = strLoaiNNKD
        UpdateCell fps, .Col, .Row, .Text

        ' Lock cac cell neu co tong hop tu phu luc
        unLockCellPL (Strloaitk)


        ' Kiem tra neu C4 <0 thi lock C5
        .Sheet = 1
        .GetText .ColLetterToNumber("O"), 43, chiTieuC4
        If Val(chiTieuC4) < 0 Then
            .Col = .ColLetterToNumber("O")
            .Row = 44
            .Lock = True
        Else
            .Col = .ColLetterToNumber("O")
            .Row = 44
            .Lock = False
        End If
        ' End check
        
        ' Lay ve row check gia han
        Dim count As Integer
        count = 0
        .Sheet = 1
        .Row = 60
        .Col = .ColLetterToNumber("B")
        Do
            count = count + 1
            .Row = .Row + 1
        Loop Until .Text = "aa"
        ' Neu check gia han moi cho phep chon GT
        .Col = .ColLetterToNumber("B")
        .Row = 61 + count
        If .Text = "1" Or UCase$(.Text) = "X" Then
            .Col = .ColLetterToNumber("D")
            .Row = 62 + count
            .Lock = False
        Else
            .Col = .ColLetterToNumber("D")
            .Row = 62 + count
            .Lock = True
        End If
        ' End set gia han
        
    .ActiveSheet = 1
    .Sheet = 1
    End With
    
End Sub

'inner class event: Change
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

Dim chiTieuC4 As Variant

Dim intIndexCombo As Integer
Dim strMa As Variant

Dim lSheet03_7 As Long

lSheet03_7 = 14

With fps
    If .ActiveSheet = lSheet03_7 Then
        .Col = Col
        .Row = Row
        If Col = .ColLetterToNumber("Q") Then 'Khi chon Combo
            intIndexCombo = .TypeComboBoxCurSel
            .Col = .ColLetterToNumber("AJ")
            .TypeComboBoxCurSel = intIndexCombo
            .GetText .Col, .Row, strMa
            'Nhap lai Combo cot AJ: Combo AJ
            
            .Col = .ColLetterToNumber("AH")
            .SetText .Col, .Row, Left$(CStr(strMa), 3)
            UpdateCell fps, .Col, .Row, Left$(CStr(strMa), 3)
        End If
        
        .Col = Col
        .Row = Row
        If Col = .ColLetterToNumber("AB") Then 'Khi chon Combo
            intIndexCombo = .TypeComboBoxCurSel
            .Col = .ColLetterToNumber("AK")
            .TypeComboBoxCurSel = intIndexCombo
            .GetText .Col, .Row, strMa
            'Nhap lai Combo cot AK: Combo AK
            
            .Col = .ColLetterToNumber("AI")
            .SetText .Col, .Row, Left$(CStr(strMa), 3)
            UpdateCell fps, .Col, .Row, Left$(CStr(strMa), 3)
        End If
    
    End If
    
    ' Kiem tra neu C4 <0 thi lock C5
    .Sheet = 1
    .GetText .ColLetterToNumber("O"), 43, chiTieuC4
    If Val(chiTieuC4) < 0 Then
        .Col = .ColLetterToNumber("O")
        .Row = 44
        .Lock = True
    Else
        .Col = .ColLetterToNumber("O")
        .Row = 44
        .Lock = False
    End If
'    If .ActiveSheet = .SheetCount - 1 Then
'         fps.EventEnabled(EventAllEvents) = False
'        GetCellSpan fps, Col, Row
'        If blGiaTriCu And (Col = .ColLetterToNumber("BD") Or Col = .ColLetterToNumber("BE")) Then
'                 DisplayMessage "0110", msOKOnly, miWarning
'                .Row = RowGiaTriCu
'                .Col = .ColLetterToNumber("BD")
'                .Text = ""
'                UpdateCell fps, .Col, .Row, .Text
'                .Col = .ColLetterToNumber("BE")
'                .Text = ""
'                UpdateCell fps, .Col, .Row, .Text
'                .Col = .ColLetterToNumber("BF")
'                .Text = ""
'                UpdateCell fps, .Col, .Row, .Text
'                .Col = .ColLetterToNumber("BG")
'                .Text = ""
'                UpdateCell fps, .Col, .Row, .Text
'                .Col = .ColLetterToNumber("BH")
'                .Text = ""
'                UpdateCell fps, .Col, .Row, .Text
'                blGiaTriCu = False
'                fps.SetFocus
'                fps.EventEnabled(EventAllEvents) = True
'                Exit Sub
'         End If
'
'       'get gia tri da ke khai
'        .GetText .ColLetterToNumber("BC"), Row, strId
'          If strId <> "" Then
'                Dim arr() As String
'                 For i = 0 To UBound(arrData)
'                     arr() = Split(arrData(i), ":")
'                     If arr(0) = Trim(Mid(strId, 100, 20)) Then
'                         cValue = arr(1)
'                         Exit For
'                     End If
'                 Next
'               .Col = .ColLetterToNumber("BF")
'               .Row = Row
'               cTagF = .CellTag
'               If cTagF <> "" Then
'                         For i = 0 To UBound(arrData)
'                               arr() = Split(arrData(i), ":")
'                               If arr(0) = Right(cTagF, Len(cTagF) - InStr(cTagF, "~")) Then
'                                   cValueF = arr(1)
'                                   Exit For
'                               End If
'                           Next
'                End If
'               .CellTag = .Row & "~" & Trim(Mid(strId, 100, 20))
'               .value = cValue
'               UpdateCell fps, .Col, .Row, .value
'               .Col = .ColLetterToNumber("BG")
'               .Row = Row
'               UpdateCell fps, .Col, .Row, .value
'               .Col = .ColLetterToNumber("BH")
'               .Row = Row
'               UpdateCell fps, .Col, .Row, .value
'
'               If loaiKHBS = "frmKHBS_TT" Then
'                    .GetText .ColLetterToNumber("BG"), Row, cValueG
'               Else
'                    .GetText .ColLetterToNumber("BH"), Row, cValueG
'               End If
'
'              If cValueG <> "" Then
'                      ParserCellID fps, Trim(Mid(strId, 100, 20)), cCol, cRow
'                      .Sheet = 1
'                      .SetText cCol, cRow, cValueG
'                      If loaiKHBS = "frmKHBS_TT" Then
'                        UpdateKHBSCell fps, cCol, cRow, cValueG
'                      Else
'                        UpdateKHBSCell fps, cCol, cRow, cValueG + cValueF
'                      End If
'              End If
'
'' Set lai gia tri cu cua sheet 01
'               .Sheet = .SheetCount - 1
'               .Col = .ColLetterToNumber("BF")
'               .Row = Row
'
'               If cTagF <> "" And cTagF <> .CellTag Then
'                  If Left(cTagF, InStr(cTagF, "~") - 1) = .Row Then
'                   ' .GetText .ColLetterToNumber("F"), Row, temp
'                    ParserCellID fps, Right(cTagF, Len(cTagF) - InStr(cTagF, "~")), iCol, iRow
'                    UpdateKHBSCell fps, iCol, iRow, cValueF
'                    .Sheet = 1
'                    If loaiKHBS = "frmKHBS_TT" Then
'                        .SetText iCol, iRow, cValueF
'                    Else
'                        .SetText iCol, iRow, 0
'                    End If
'                    .Sheet = .SheetCount - 1
'                    If loaiKHBS = "frmKHBS_TT" Then
'                        .GetText .ColLetterToNumber("BG"), Row, varTemp
'                    Else
'                        .GetText .ColLetterToNumber("BH"), Row, varTemp
'                    End If
'                    If varTemp <> "" Then
'                        ParserCellID fps, Trim(Mid(strId, 100, 20)), iCol, iRow
'                        If loaiKHBS = "frmKHBS_TT" Then
'                            UpdateKHBSCell fps, cCol, cRow, varTemp
'                        Else
'                            UpdateKHBSCell fps, cCol, cRow, varTemp + cValueF
'                        End If
'                       ' UpdateKHBSCell fps, iCol, iRow, varTemp
'                        .Sheet = 1
'                        .SetText iCol, iRow, varTemp
'                        .Sheet = .SheetCount - 1
'                    End If
'                   End If
'               End If
''Xoa dong
'               If mblKeyF6 = True Then
'                ParserCellID fps, Right(cTagF, Len(cTagF) - InStr(cTagF, "~")), iCol, iRow
'                UpdateKHBSCell fps, iCol, iRow, cValueF
'                .Sheet = 1
'                .SetText iCol, iRow, 0
'                .Sheet = .SheetCount - 1
'                .Col = .ColLetterToNumber("BF")
'                .Row = Row
'                .value = 0
'                .Col = .ColLetterToNumber("BG")
'                .Row = Row
'                .value = 0
'                .Col = .ColLetterToNumber("BH")
'                .Row = Row
'                .value = 0
'               End If
'
'            .Sheet = 1
'            .GetText .ColLetterToNumber("G"), 50, varTemp
'            .Sheet = .SheetCount - 1
'            Set xmlNodeH15 = TAX_Utilities_New.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
'                 ParserCellID fps, GetAttribute(xmlNodeH15, "CellID"), cCol, cRow
'                .Col = cCol
'                .Row = cRow
'                If loaiKHBS = "frmKHBS_TT" Then
'                    .value = CDbl(varTemp) - CDbl(vThuephainop)
'                     UpdateCell fps, .Col, .Row, .value
'                Else
'                    .value = CDbl(varTemp)
'                     UpdateCell fps, .Col, .Row, .value
'                End If
'                If .value < 0 Then
'                    .Col = .ColLetterToNumber("BE")
'                    .Row = cRow + 3
'                    .value = 0
'                    UpdateCell fps, .Col, .Row, 0
'                Else
'                    .Col = .ColLetterToNumber("BE")
'                    .Row = cRow + 3
'                    UpdateCell fps, .Col, .Row, .value
'                End If
'             Set xmlNodeH15 = Nothing
'
'          End If
'            fps.EventEnabled(EventAllEvents) = True
'       End If
     mblKhoiDauToKhai = False
   End With

    fps.Col = Col
    fps.Row = Row
    If arrErrorCells.Exists(mCurrentSheet & "_" & GetCellID(fps, Col, Row)) Then
        fps.CellNote = ""
        fps.BackColor = arrErrorCells.Item(mCurrentSheet & "_" & GetCellID(fps, Col, Row))
        arrErrorCells.Remove mCurrentSheet & "_" & GetCellID(fps, Col, Row)
    End If
    CellChange Col, Row
End Sub

'inner class event: ButtonClicked
Public Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
     TAX_Utilities_New.AdjustData(fps.ActiveSheet - 1) = True
     'CellChange Col, Row
     
    Dim blTemp As Boolean
    Dim lResult As Integer
    Dim count As Integer
    With fps
        fps.EventEnabled(EventAllEvents) = False
        ' set gia lock cell ly do gia han
        If fps.ActiveSheet = 1 Then
                ' Lay ve row check gia han
                count = 0
                .Sheet = 1
                .Row = 60
                .Col = .ColLetterToNumber("B")
                Do
                    count = count + 1
                    .Row = .Row + 1
                Loop Until .Text = "aa"
        
                ' Neu check gia han moi cho phep chon GT
                If Col = fps.ColLetterToNumber("B") And Row = 61 + count Then
                    fps.Col = fps.ColLetterToNumber("B")
                    fps.Row = 61 + count
                    If fps.Text = "1" Or UCase$(fps.Text) = "X" Then
                        fps.Col = fps.ColLetterToNumber("D")
                        fps.Row = 62 + count
                        fps.Lock = False
                    Else
                        fps.Col = fps.ColLetterToNumber("D")
                        fps.Row = 62 + count
                        fps.Lock = True
                        fps.SetText fps.ColLetterToNumber("D"), fps.Row = 62 + count, ""
                        fps.SetText fps.ColLetterToNumber("P"), fps.Row = 62 + count, "00"
                        UpdateCell fps, fps.ColLetterToNumber("P"), fps.Row = 62 + count, "00"
                        UpdateCell fps, fps.ColLetterToNumber("D"), fps.Row = 62 + count, ""
                    End If
                End If
        End If
        
        fps.EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim strId As Variant
Dim iCol As Long, iRow As Long
Dim intIndexCombo As Integer, i As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim itemp As Variant, temp As Variant
Dim cValue As Variant
With fps

'   If .ActiveSheet = .SheetCount - 1 Then
'        .Col = Col
'        .Row = Row
'         If Col = .ColLetterToNumber("BD") And .Text <> "" Then
'            intIndexCombo = .TypeComboBoxCurSel
'            .Col = .ColLetterToNumber("BC")
'            .TypeComboBoxCurSel = intIndexCombo
'            'Update combo C
'            UpdateCell fps, .Col, .Row, .Text
'            'Chon lai combo cua colF
'            .Col = .ColLetterToNumber("BE")
'            .GetText .ColLetterToNumber("BC"), Row, strId
'            .Text = Trim(Right(strId, 10))
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'         If Col = .ColLetterToNumber("BE") And .Text <> "" Then
'            intIndexCombo = .TypeComboBoxCurSel
'            .Col = .ColLetterToNumber("BC")
'            .TypeComboBoxCurSel = intIndexCombo
'            'Update combo C
'            UpdateCell fps, .Col, .Row, .Text
'            'Chon lai combo cua colF
'            .Col = .ColLetterToNumber("BD")
'            .GetText .ColLetterToNumber("BC"), Row, strId
'            .Text = Trim(Left(strId, 100))
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'
'
'        Set xmlCellNode = TAX_Utilities_New.Data(.SheetCount - 2).nodeFromID(GetCellID(fps, Col, Row))
'        Set xmlCellNode = xmlCellNode.parentNode.parentNode.firstChild.firstChild
'        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'        blGiaTriCu = False
'        If Col = fps.ColLetterToNumber("BD") Or Col = fps.ColLetterToNumber("BE") Then
'            fps.Col = .ColLetterToNumber("BC")
'            fps.Row = Row
'            temp = fps.Text
'            Set xmlCellsNode = xmlCellNode.parentNode
'            i = 1
'            .Col = .ColLetterToNumber("B")
'            ParserCellID fps, "BD_9", iCol, iRow
'            .Row = iRow
'            Do
'                 .Col = .ColLetterToNumber("BC")
'                 itemp = .Text
'                 If .Row <> Row Then
'                    If itemp = temp Then 'Neu co thi dat co blGiaTriCu = True
'                        RowGiaTriCu = Row
'                        ColGiaTriCu = .ColLetterToNumber("BD")
'                        .Row = Row
'                        .Col = .ColLetterToNumber("BD")
'                        blGiaTriCu = True
'                        Exit Sub
'                    End If
'                 End If
'                 .Col = .ColLetterToNumber("B")
'                 .Row = i + iRow
'                 i = i + 1
'            Loop Until .Text = "bb"
'       End If
'
'    End If
    
        If .ActiveSheet = 1 Then
            Dim count As Integer
            count = 0
            .Sheet = 1
            .Row = 60
            .Col = .ColLetterToNumber("B")
            Do
                count = count + 1
                .Row = .Row + 1
            Loop Until .Text = "aa"
        
            .Col = Col
            .Row = Row
             If Col = .ColLetterToNumber("D") And Row = 62 + count Then
                intIndexCombo = .TypeComboBoxCurSel
                .Col = .ColLetterToNumber("P")
                .TypeComboBoxCurSel = intIndexCombo
                'Update combo C
                UpdateCell fps, .Col, .Row, .Text
            End If
        End If
End With
    
    CellChange Col, Row
    TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
End Sub


Private Sub fps_EditChange(ByVal Col As Long, ByVal Row As Long)
   'MsgBox "Col: " & Col & " Row: " & Row & " Sheet: " & fps.ActiveSheet
End Sub

Private Sub fps_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer, iCol As Long, iRow As Long
    If KeyCode = vbKeyClear Or KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        With fps
'            If .ActiveCol = .ColLetterToNumber(S1TxtB24Col) And .ActiveRow = S1TxtB24Row Then
'                XoaDauAm .ActiveCol, .ActiveRow
'            End If
'            If .ActiveCol = .ColLetterToNumber(S1TxtB25Col) And .ActiveRow = S1TxtB25Row Then
'                XoaDauAm .ActiveCol, .ActiveRow
'            End If
            If .ActiveCol = .ColLetterToNumber(S1TxtC8Col) And .ActiveRow = S1TxtC8Row Then
                XoaDauAm .ActiveCol, .ActiveRow
            End If
        End With
    End If
 ' for KHBS only
    If KeyCode = vbKeyF6 Then
     With fps
            iCol = .ActiveCol
            iRow = .ActiveRow
            mblKeyF6 = True
            fps_Change iCol, iRow
            mblKeyF6 = False
     End With
    End If
    
    
End Sub

'inner class event: Leavecell
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
            If Col = .ColLetterToNumber("F") And Row = 61 Then
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
'            .GetText .ColLetterToNumber("G"), 50, varTemp
'            If TAX_Utilities_New.NodeValidity.childNodes(.SheetCount - 2).Attributes.getNamedItem("Active").nodeValue <> "0" Then
'                .Sheet = .SheetCount - 1
'                Set xmlNode = TAX_Utilities_New.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
'                ParserCellID fps, GetAttribute(xmlNode, "CellID"), cCol, cRow
'                .Col = cCol
'                .Row = cRow
'                .Value = CDbl(varTemp) - CDbl(vThuephainop)
'                mCurrentSheet = .SheetCount - 1
'                UpdateCell fps, .Col, .Row, .Value
'                .Col = .ColLetterToNumber("BG")
'                .Row = .Row + 3
'                UpdateCell fps, .Col, .Row, .Value
'                .Sheet = 1
'                 mCurrentSheet = 1
'            End If
        End If
        'KHBS
        If .ActiveSheet = .SheetCount - 1 Then
         Dim count1 As Integer
            .Row = 9
            .Col = .ColLetterToNumber("B")
            Do While .value <> "bb"
                count1 = count1 + 1
                .Row = .Row + 1
            Loop
                If Col = .ColLetterToNumber("BG") And Row = 20 + count1 - 5 Then
                       .GetText Col, Row, varTemp
                       If varTemp <> "" And varTemp <> "../../...." Then
                           If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                               .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                               .TypeHAlign = TypeHAlignLeft
                           Else
                                .SetFocus
                                .SetActiveCell Col, Row
                           End If
                       Else
                        .SetText Col, Row, ""
                       End If
                      .Col = Col
                      .Row = Row
                      UpdateCell fps, .Col, .Row, .Text
                   End If
                .Sheet = .SheetCount - 1
        End If
        
    .EventEnabled(EventAllEvents) = True
    End With

    CellChange Col, Row
End Sub
'When cell's data is changed or a control at a cell is handled, This sub is called
'parameter Col: the column of cell that rise event
'parameter Row: the row of cell that rise event
'parameter f: the flag to mark some special option
'**: when this sub is on executing, all events of fpspread will be disabled until the end of the sub
Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
    Dim cCol As Long, cRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim varTemp As Variant
    Dim pL_3A_U50, pL_3B_U60, pL_3A_U56, pL_3B_U65, pL_3C_X43 As Double
    Dim strValue As String
    Dim arrRowCols() As String
    Dim arrRowCol() As String
    Dim idx As Long
    
    Dim lSheet03_2A As Long, lSheet03_2B As Long, lSheet03_3A As Long, lSheet03_3B As Long, lSheet03_3C As Long
    Dim lSheet03_4 As Long, lSheet03_5 As Long, lSheet03_05 As Long
    
    lSheet03_2A = 5
    lSheet03_2B = 6
    lSheet03_3A = 7
    lSheet03_3B = 8
    lSheet03_3C = 9
    lSheet03_4 = 10
    lSheet03_5 = 11
    lSheet03_05 = 12
            
With fps
    'Because of unclearly-rised events, this function is here for turning it off
    'at the end of this sub, this event will be turned on
    .EventEnabled(EventAllEvents) = False
    .Sheet = .ActiveSheet
    mCurrentSheet = .ActiveSheet
    Select Case .Sheet
        Case 1 ' To khai TNDN
            mCurrentSheet = .Sheet
            GetCellSpan fps, Col, Row
        Case 2 ' Phu luc 03-1A
            mCurrentSheet = .Sheet
            If Row = 41 And Col = .ColLetterToNumber("E") Then
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 23
                .Formula = "P23"
            End If
        Case 3 ' Phu luc 03-1B
            mCurrentSheet = .Sheet
            If Row = 38 And Col = .ColLetterToNumber("E") Then
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 23
                .Formula = "P23"
            End If
            
        Case 4 ' Phu luc 03-1C
            mCurrentSheet = .Sheet
            If Row = 50 And Col = .ColLetterToNumber("E") Then
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 23
                .Formula = "P23"
            End If
            
        Case lSheet03_2A ' Phu luc 03-2A
            mCurrentSheet = .Sheet
            If Row = 32 And Col = .ColLetterToNumber("M") Then
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 42
                .Formula = "P42"
            End If
        Case lSheet03_2B ' Phu luc 03-2B
            mCurrentSheet = .Sheet
'            If Row = 32 And Col = .ColLetterToNumber("M") Then
'                .Sheet = 1
'                .Col = .ColLetterToNumber("O")
'                .Row = 42
'                .Formula = "P42"
'            End If
            
        Case lSheet03_3A ' Phu luc 03-3A
            mCurrentSheet = .Sheet
            If IsValidatedCell(lSheet03_3A, 38, .ColLetterToNumber("N"), Row, Col) Then
                 ValidateNumberFormat fps, lSheet03_3A, Row, Col
            End If
            If IsValidatedCell(lSheet03_3A, 39, .ColLetterToNumber("J"), Row, Col) Then
                 ValidateNumberFormat fps, lSheet03_3A, Row, Col
            End If
            If IsValidatedCell(lSheet03_3A, 40, .ColLetterToNumber("N"), Row, Col) Then
                 ValidateNumberFormat fps, lSheet03_3A, Row, Col
            End If
            If IsValidatedCell(lSheet03_3A, 38, .ColLetterToNumber("U"), Row, Col) Then
                 ValidateDate fps, lSheet03_3A, Row, Col, yyyy
            End If
            If IsValidatedCell(lSheet03_3A, 39, .ColLetterToNumber("Q"), Row, Col) Then
                 ValidateDate fps, lSheet03_3A, Row, Col, yyyy
            End If
            If IsValidatedCell(lSheet03_3A, 40, .ColLetterToNumber("U"), Row, Col) Then
                 ValidateDate fps, lSheet03_3A, Row, Col, yyyy
            End If
            ' TOng hop len chi tieu C8
            pL_3A_U50 = Sum_Cell(lSheet03_3A, 50, .ColLetterToNumber("U")) 'PL 03-3A
            pL_3B_U60 = Sum_Cell(lSheet03_3B, 60, .ColLetterToNumber("U")) 'PL 03-3B
            
            strValue = getArrRowCol(lSheet03_3A, 50, .ColLetterToNumber("U"))
            If InStr(1, strValue, "###") > 0 Then
                arrRowCols = Split(strValue, "###")
                If UBound(arrRowCols) > 0 Then
                    For idx = 0 To UBound(arrRowCols)
                        arrRowCol = Split(arrRowCols(idx), "_")
                        If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                            .Sheet = 1
                            .SetText .ColLetterToNumber("P"), 47, pL_3A_U50 + pL_3B_U60
                            .Col = .ColLetterToNumber("O")
                            .Row = 47
                            .Formula = "P47"
                            Exit For
                        End If
                    Next idx
                End If
            Else
                arrRowCol = Split(strValue, "_")
                If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                    .Sheet = 1
                    .SetText .ColLetterToNumber("P"), 47, pL_3A_U50 + pL_3B_U60
                    .Col = .ColLetterToNumber("O")
                    .Row = 47
                    .Formula = "P47"
                End If
            End If
            
            ' TOng hop len chi tieu C9
            pL_3A_U56 = Sum_Cell(lSheet03_3A, 56, .ColLetterToNumber("U")) 'PL 03-3A
            pL_3B_U65 = Sum_Cell(lSheet03_3B, 65, .ColLetterToNumber("U")) 'PL 03-3B
            pL_3C_X43 = Sum_Cell(lSheet03_3C, 43, .ColLetterToNumber("X")) 'PL 03-3C
            
            strValue = getArrRowCol(lSheet03_3A, 56, .ColLetterToNumber("U"))
            If InStr(1, strValue, "###") > 0 Then
                arrRowCols = Split(strValue, "###")
                If UBound(arrRowCols) > 0 Then
                    For idx = 0 To UBound(arrRowCols)
                        arrRowCol = Split(arrRowCols(idx), "_")
                        If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                            .Sheet = 1
                            .SetText .ColLetterToNumber("P"), 48, pL_3A_U56 + pL_3B_U65 + pL_3C_X43
                            .Col = .ColLetterToNumber("O")
                            .Row = 48
                            .Formula = "P48"
                            Exit For
                        End If
                    Next idx
                End If
             Else
                arrRowCol = Split(strValue, "_")
                If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                    .Sheet = 1
                    .SetText .ColLetterToNumber("P"), 48, pL_3A_U56 + pL_3B_U65 + pL_3C_X43
                    .Col = .ColLetterToNumber("O")
                    .Row = 48
                    .Formula = "P48"
                End If
            End If
            
        
        Case lSheet03_3B ' Phu luc 03-3B
            mCurrentSheet = .Sheet
            If IsValidatedCell(lSheet03_3B, 44, .ColLetterToNumber("N"), Row, Col) Then
                 ValidateNumberFormat fps, lSheet03_3B, Row, Col
            End If
            If IsValidatedCell(lSheet03_3B, 45, .ColLetterToNumber("J"), Row, Col) Then
                 ValidateNumberFormat fps, lSheet03_3B, Row, Col
            End If
            If IsValidatedCell(lSheet03_3B, 46, .ColLetterToNumber("N"), Row, Col) Then
                 ValidateNumberFormat fps, lSheet03_3B, Row, Col
            End If
            If IsValidatedCell(lSheet03_3B, 44, .ColLetterToNumber("U"), Row, Col) Then
                 ValidateDate fps, lSheet03_3B, Row, Col, yyyy
            End If
            If IsValidatedCell(lSheet03_3B, 45, .ColLetterToNumber("Q"), Row, Col) Then
                 ValidateDate fps, lSheet03_3B, Row, Col, yyyy
            End If
            If IsValidatedCell(lSheet03_3B, 46, .ColLetterToNumber("U"), Row, Col) Then
                 ValidateDate fps, lSheet03_3B, Row, Col, yyyy
            End If
            If IsValidatedCell(lSheet03_3B, 41, .ColLetterToNumber("S"), Row, Col) Then
                 ValidateDate fps, lSheet03_3B, Row, Col, DDMMYYYY
            End If
        
            ' Tong hop len chi tieu C8
            pL_3A_U50 = Sum_Cell(lSheet03_3A, 50, .ColLetterToNumber("U")) 'PL 03-3A
            pL_3B_U60 = Sum_Cell(lSheet03_3B, 60, .ColLetterToNumber("U")) 'PL 03-3B
            
            strValue = getArrRowCol(lSheet03_3B, 60, .ColLetterToNumber("U"))
            If InStr(1, strValue, "###") > 0 Then
                arrRowCols = Split(strValue, "###")
                If UBound(arrRowCols) > 0 Then
                    For idx = 0 To UBound(arrRowCols)
                        arrRowCol = Split(arrRowCols(idx), "_")
                        If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                            .Sheet = 1
                            .SetText .ColLetterToNumber("P"), 47, pL_3A_U50 + pL_3B_U60
                            .Col = .ColLetterToNumber("O")
                            .Row = 47
                            .Formula = "P47"
                            Exit For
                        End If
                    Next idx
                End If
            Else
                arrRowCol = Split(strValue, "_")
                If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                    .Sheet = 1
                    .SetText .ColLetterToNumber("P"), 47, pL_3A_U50 + pL_3B_U60
                    .Col = .ColLetterToNumber("O")
                    .Row = 47
                    .Formula = "P47"
                End If
            End If
            
            ' TOng hop len chi tieu C9
            pL_3A_U56 = Sum_Cell(lSheet03_3A, 56, .ColLetterToNumber("U")) 'PL 03-3A
            pL_3B_U65 = Sum_Cell(lSheet03_3B, 65, .ColLetterToNumber("U")) 'PL 03-3B
            pL_3C_X43 = Sum_Cell(lSheet03_3C, 43, .ColLetterToNumber("X")) 'PL 03-3C
            
            strValue = getArrRowCol(lSheet03_3B, 65, .ColLetterToNumber("U"))
            If InStr(1, strValue, "###") > 0 Then
                arrRowCols = Split(strValue, "###")
                If UBound(arrRowCols) > 0 Then
                    For idx = 0 To UBound(arrRowCols)
                        arrRowCol = Split(arrRowCols(idx), "_")
                        If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                            .Sheet = 1
                            .SetText .ColLetterToNumber("P"), 48, pL_3A_U56 + pL_3B_U65 + pL_3C_X43
                            .Col = .ColLetterToNumber("O")
                            .Row = 48
                            .Formula = "P48"
                            Exit For
                        End If
                    Next idx
                End If
             Else
                arrRowCol = Split(strValue, "_")
                If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                    .Sheet = 1
                    .SetText .ColLetterToNumber("P"), 48, pL_3A_U56 + pL_3B_U65 + pL_3C_X43
                    .Col = .ColLetterToNumber("O")
                    .Row = 48
                    .Formula = "P48"
                End If
            End If
        
        Case lSheet03_3C ' Phu luc 03-3C
            mCurrentSheet = .Sheet
            If IsValidatedCell(lSheet03_3C, 27, .ColLetterToNumber("V"), Row, Col) Then
                 ValidateDate fps, lSheet03_3C, Row, Col, DDMMYYYY
            End If
            If IsValidatedCell(lSheet03_3C, 31, .ColLetterToNumber("Y"), Row, Col) Then
                 ValidateDate fps, lSheet03_3C, Row, Col, DDMMYYYY
            End If
            
            ' TOng hop len chi tieu C9
            pL_3A_U56 = Sum_Cell(lSheet03_3A, 56, .ColLetterToNumber("U")) 'PL 03-3A
            pL_3B_U65 = Sum_Cell(lSheet03_3B, 65, .ColLetterToNumber("U")) 'PL 03-3B
            pL_3C_X43 = Sum_Cell(lSheet03_3C, 43, .ColLetterToNumber("X")) 'PL 03-3C
            
            strValue = getArrRowCol(lSheet03_3C, 43, .ColLetterToNumber("X"))
            If InStr(1, strValue, "###") > 0 Then
                arrRowCols = Split(strValue, "###")
                If UBound(arrRowCols) > 0 Then
                    For idx = 0 To UBound(arrRowCols)
                        arrRowCol = Split(arrRowCols(idx), "_")
                        If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                            .Sheet = 1
                            .SetText .ColLetterToNumber("P"), 48, pL_3A_U56 + pL_3B_U65 + pL_3C_X43
                            .Col = .ColLetterToNumber("O")
                            .Row = 48
                            .Formula = "P48"
                            Exit For
                        End If
                    Next idx
                End If
            Else
                arrRowCol = Split(strValue, "_")
                If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                    .Sheet = 1
                    .SetText .ColLetterToNumber("P"), 48, pL_3A_U56 + pL_3B_U65 + pL_3C_X43
                    .Col = .ColLetterToNumber("O")
                    .Row = 48
                    .Formula = "P48"
                End If
            End If
        Case lSheet03_4 ' Phu luc 03-4
            mCurrentSheet = .Sheet
            Set xmlCellNode = TAX_Utilities_New.Data(lSheet03_4 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("E"), 26))
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode.lastChild
            Set xmlCellNode = xmlCellNode.lastChild
            Set xmlCellNode = xmlCellNode.lastChild
            ' Lay ve col vaf row cua dong tong cong
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), cCol, cRow
            If Row = cRow And Col = cCol Then
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 49
                .Formula = "P49"
            End If
        Case lSheet03_5 ' Phu luc 03-5
            mCurrentSheet = .Sheet
             If Row = 38 And Col = .ColLetterToNumber("E") Then
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 51
                .Formula = "P51"
            End If
        Case lSheet03_05 ' Phu luc 03-05
            mCurrentSheet = .Sheet
        Case .SheetCount - 1 'KHBS
         mCurrentSheet = .Sheet
            
    End Select
    .EventEnabled(EventAllEvents) = True
End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    Dim i As Integer, iCol As Long, iRow As Long
    
    Dim col3 As Long, col4 As Long, col5 As Long
    Dim col6 As Long, Col7 As Long, col8 As Long
    Dim col9 As Long, Col10 As Long, Col11 As Long
    
    Dim ColLoi As Long, RowLoi As Long
    Dim ColLoi7 As Long, RowLoi7 As Long
    Dim ColLoi8 As Long, RowLoi8 As Long
    Dim ColLoi10 As Long, RowLoi10 As Long
    
    Dim vErrMsg1 As Variant
    
    Dim ValCol3 As Variant, ValCol4 As Variant, ValCol5 As Variant
    Dim ValCol6 As Variant, ValCol7 As Variant, ValCol8 As Variant
    Dim ValCol9 As Variant, ValCol10 As Variant, ValCol11 As Variant
    'PL 14
    Dim XacDinhLoiCol7 As Boolean, XacDinhLoiCol8 As Boolean, XacDinhLoiCol10 As Boolean
    
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim GiaTri As Double, TongSo As Double
    Dim lCtrl As Long, RowTongSo As Long, ColTongSo As Long
    
    Dim XacDinhLoi As Boolean
    Dim blCheck_S03_1A As Boolean, blCheck_S03_1B As Boolean, blCheck_S03_1C  As Boolean
    Dim blCheck_S03_3 As Boolean, blCheck_S03_3C As Boolean, blCheck_S03_4 As Boolean
    Dim blCheck_S03_3A As Boolean
    Dim strTenDN As Variant, strMaST As Variant, strValue As Variant, vMSTDN As Variant
    Dim errEmty As Variant, iFlagTongPhanBo As Variant, errTongTLPB As Variant, vMST As Variant
    Dim errMSTDN As Variant, errMst10 As Variant, iFlagCheckMaSTDN10 As Variant, errMST As Variant
    Dim blCheck_S11 As Boolean
    Dim iFlagEmpty As Integer, iFlagCheckMaSTDN As Integer, iFlagCheckMaST As Integer
    Dim strKyHieu As Variant, strCheckMST As String

    ' PL 05
    Dim CountRow As Long, rowDes As Long, iFlagMSTTrung As Integer, mErrMSTTrung As Variant
    Dim vMSTCheckTrung As Variant
    
    ' PL03-6
    Dim vSTT, vCTieu, vTyLe, vThueGTGT As Variant
    Dim blCheck_S12 As Boolean
    Dim iWhiteRow As Integer, iFlagCtieu As Integer, iFlagMST As Integer
    Dim errCTieu, vMSTtemp, alertThueTNDN As Variant
    Dim strCheck As String
    Dim vThuePL05 As Variant
    Dim tempRow As Long, tempRow1 As Long
    ' end
    ' PL 03-3A
    Dim vCtieu1, vCTieu5, vCTieuC4TK As Variant
    Dim vErrCtieuC4 As Variant
    ' lSheet
    Dim lSheet03_2A As Long, lSheet03_2B As Long, lSheet03_3A As Long, lSheet03_3B As Long, lSheet03_3C As Long
    Dim lSheet03_4 As Long, lSheet03_5 As Long, lSheet03_05 As Long
    Dim lSheet03_6 As Long, lSheet03_7 As Long
    
    ' check gh
    Dim vCheckGH, vLyDoGH, vErrorCheckGH As Variant
    Dim iflagCheckGH As Integer, count As Integer
    
    lSheet03_3A = 7
    lSheet03_3B = 8
    lSheet03_3C = 9
    lSheet03_4 = 10
    lSheet03_5 = 11
    lSheet03_05 = 12
    lSheet03_6 = 13
    lSheet03_7 = 14
    
    
    CheckErrorMST
    With fps
        .EventEnabled(EventAllEvents) = False
        ' Phu Luc 03-1A
        .Sheet = 2
        blCheck_S03_1A = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        
'        ' Neu Phu Luc 03-1A ma Visible thi khong can check va set lai cac gia tri cho B27 trong sheet Header = 1
'        If blCheck_S03_1A = False Then
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 27, "1"
'        Else ' Hien tai theo yeu cau luon luon ko kiem tra cac chi tieu nay tren to khai quyet toan
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 27, "1"
'        End If
        
        ' Phu Luc 03-1B
'        .Sheet = 3
'        blCheck_S03_1B = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'        ' Neu Phu Luc 03-1A ma Visible thi khong can check va set lai cac gia tri cho B28 trong sheet Header = 1
'        If blCheck_S03_1B = False Then
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 28, "1"
'        Else ' Hien tai theo yeu cau luon luon ko kiem tra cac chi tieu nay tren to khai quyet toan
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 28, "1"
'        End If
        
        ' Phu Luc 03-1C
'        .Sheet = 4
'        blCheck_S03_1C = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'        ' Neu Phu Luc 03-1A ma Visible thi khong can check va set lai cac gia tri cho B29 trong sheet Header = 1
'        If blCheck_S03_1C = False Then
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 29, "1"
'        Else ' Hien tai theo yeu cau luon luon ko kiem tra cac chi tieu nay tren to khai quyet toan
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 29, "1"
'        End If
                   
        ' Chuyen den Sheet Header de lay cac thong bao loi
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 25, vErrMsg1
        .GetText .ColLetterToNumber("E"), 27, errTongTLPB
        .GetText .ColLetterToNumber("E"), 30, errMST
        .GetText .ColLetterToNumber("E"), 29, errEmty
        .GetText .ColLetterToNumber("B"), 27, iFlagTongPhanBo
        .GetText .ColLetterToNumber("E"), 31, errMSTDN
        .GetText .ColLetterToNumber("E"), 28, errMst10
        
        .GetText .ColLetterToNumber("E"), 35, mErrMSTTrung
        .GetText .ColLetterToNumber("E"), 36, vErrCtieuC4
        
        .GetText .ColLetterToNumber("E"), 46, vErrorCheckGH
        
        ' check ly do gia han
        .Sheet = 1
         count = 0
        .Row = 60
        .Col = .ColLetterToNumber("B")
        Do
            count = count + 1
            .Row = .Row + 1
        Loop Until .Text = "aa"
        
        .GetText .ColLetterToNumber("B"), 61 + count, vCheckGH
        
        .GetText .ColLetterToNumber("D"), 62 + count, vLyDoGH
         If Trim(vCheckGH) = "1" Or Trim$(UCase(vCheckGH)) = "X" Then
             If Trim$(vLyDoGH) = "" Then
                 .Col = .ColLetterToNumber("D")
                 .Row = 62 + count
                 .CellNote = vErrorCheckGH  'static
                 .BackColor = mErrorColor
                 iflagCheckGH = 1
             Else
                 .Col = .ColLetterToNumber("D")
                 .Row = 62 + count
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckGH = 0
             End If
         Else
                 .Col = .ColLetterToNumber("D")
                 .Row = 62 + count
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckGH = 0
         End If
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 46, IIf(iflagCheckGH = 1, "0", "1")
        
        
        ' Kiem tra validate cua PL 03-3C
        .Sheet = lSheet03_3C
        blCheck_S03_3C = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S03_3C = True Then
            checkPL03_3C
        End If
        
        
        ' Chuyen ve Sheet 9 (Phu Luc 03-4) lay gia tri cac cot can validate du lieu
        mCurrentSheet = lSheet03_4
        .Sheet = lSheet03_4
        blCheck_S03_4 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S03_4 = True Then
            XacDinhLoi = False
            i = 1
            col3 = .ColLetterToNumber("O")
            col4 = .ColLetterToNumber("R")
            col5 = .ColLetterToNumber("U")
            col6 = .ColLetterToNumber("Z")
            Col7 = .ColLetterToNumber("AE")
            col8 = .ColLetterToNumber("AH")
            col9 = .ColLetterToNumber("AK")
            Col10 = .ColLetterToNumber("AN")
            Col11 = .ColLetterToNumber("AR")
            
            .Col = .ColLetterToNumber("C")
            .Row = 26
            
            Do
                 .GetText col3, .Row, ValCol3
                 .GetText col4, .Row, ValCol4
                 .GetText col5, .Row, ValCol5
                 .GetText col6, .Row, ValCol6
                 .GetText Col7, .Row, ValCol7
                 .GetText col8, .Row, ValCol8
                 .GetText col9, .Row, ValCol9
                 .GetText Col10, .Row, ValCol10
                 .GetText Col11, .Row, ValCol11
                 
                 ' Kiem tra Cot 11 (AR) & Cot 10 (AN) hoac Cot 11 (AR) & Cot 6 (Z)
                 If (ValCol11 > ValCol10) Or (ValCol11 > ValCol6) Then
                        .Col = Col11
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsg1
                        .BackColor = mErrorColor
                        If XacDinhLoi = False Then
                            ColLoi = Col11
                            RowLoi = .Row
                        End If
                        XacDinhLoi = True
                        
                 Else
                         .Col = Col11
                         .CellNote = ""
                         .BackColor = mNonErrorColor
                 End If
                 
                 .Row = i + 26
                 .Col = .ColLetterToNumber("C")
                 i = i + 1
            Loop Until .Text = "aa"
            If XacDinhLoi Then
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 25, IIf(XacDinhLoi, "0", "1")
                Dim strFormula As String
                strFormula = "03-4!" & Trim(.ColNumberToLetter(CLng(ColLoi))) & Trim(str(RowLoi))
                .Col = .ColLetterToNumber("C")
                .Row = 25
                .Formula = strFormula
            Else
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 25, "1"
            End If
         End If
        
        ' Check phu  luc 05
        .Sheet = lSheet03_05
        If .SheetVisible = True Then
            mCurrentSheet = .Sheet
            .GetText .ColLetterToNumber("F"), 38, vMSTDN
            iFlagEmpty = 0
            'check not empty ten doanh nghiep
            .Row = 39
            Do
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
                'iFlagEmpty = 0
            End If
            'lay du lieu nhap vao
            .GetText .ColLetterToNumber("E"), .Row, strKyHieu
            If strKyHieu = "T" And Val(iFlagTongPhanBo) = 0 Then
                .Col = .ColLetterToNumber("L")
                .BackColor = mErrorColor
                .CellNote = errTongTLPB
            ElseIf Val(iFlagTongPhanBo) = 1 Then
                .Col = .ColLetterToNumber("L")
                .BackColor = mNonErrorColor
                .CellNote = ""
            End If
            If strKyHieu = "M" And .Row <> 38 Then
                .Col = .ColLetterToNumber("F")
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
                        If Left(vMST, 10) <> Left(Trim(vMSTDN), 10) Then
                            .Col = .ColLetterToNumber("F")
                            .BackColor = mErrorColor
                            .CellNote = errMSTDN
                            iFlagCheckMaSTDN = 1
                        ElseIf Len(vMST) = 10 And strCheckMST = "000" Then
                            .Col = .ColLetterToNumber("F")
                            .BackColor = mErrorColor
                            .CellNote = errMst10
                            iFlagCheckMaSTDN10 = 1
                        ElseIf Len(vMST) = 10 And strCheckMST <> "000" Then
                            .Col = .ColLetterToNumber("F")
                            .BackColor = mErrorColor
                            .CellNote = errMST
                            iFlagCheckMaST = 1
                        ElseIf Len(vMST) = 13 And (strCheckMST <> "000" Or IsNumeric(vMST) = False) Then
                            .Col = .ColLetterToNumber("F")
                            .BackColor = mErrorColor
                            .CellNote = errMST
                            iFlagCheckMaST = 1
                        Else
                            .Col = .ColLetterToNumber("F")
                            .BackColor = mNonErrorColor
                            .CellNote = ""
                            'iFlagCheckMaST = 0
                        End If
                    ElseIf Len(vMST) <> 10 Or Len(vMST) <> 13 Then
                           .Col = .ColLetterToNumber("F")
                            .BackColor = mErrorColor
                            .CellNote = errMST
                            iFlagCheckMaST = 1
                    
                    End If
                    
                      'check MST trung nhau
                     CountRow = .SearchCol(.ColLetterToNumber("F"), -1, -1, vMSTCheckTrung, SearchFlagsNone)
                     rowDes = .SearchCol(.ColLetterToNumber("F"), CountRow, -1, vMSTCheckTrung, SearchFlagsNone)
                    If rowDes > -1 Then
                        .Col = .ColLetterToNumber("F")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = mErrMSTTrung      'static
                        .BackColor = mErrorColor
                        iFlagMSTTrung = 1
                    End If
                    
                Else
                         .Col = .ColLetterToNumber("F")
    '                    .BackColor = mNonErrorColor
    '                    .CellNote = ""
                        'iFlagCheckMaST = 0
                        .BackColor = mErrorColor
                        .CellNote = errEmty
                        iFlagEmpty = 1
                End If
            End If
            
            .Col = .ColLetterToNumber("C")
            .Row = .Row + 1
            Loop Until .Text = "aa"
        End If
        'end
        blCheck_S11 = IIf(TAX_Utilities_New.NodeValidity.childNodes(lSheet03_05 - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        .Sheet = .SheetCount
        If blCheck_S11 Then
            .SetText .ColLetterToNumber("B"), 28, IIf(iFlagCheckMaSTDN10 = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 35, IIf(iFlagMSTTrung = 1, "0", "1")
        Else
            .SetText .ColLetterToNumber("B"), 28, "1"
            .SetText .ColLetterToNumber("B"), 35, "1"
        End If
        
        .SetText .ColLetterToNumber("B"), 29, IIf(iFlagEmpty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 30, IIf(iFlagCheckMaST = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 31, IIf(iFlagCheckMaSTDN = 1, "0", "1")
        
        ' Kiem tra PL 03-6
'        .Sheet = .SheetCount
'        .GetText .ColLetterToNumber("E"), 32, errCTieu
'        .GetText .ColLetterToNumber("E"), 33, errMST
'        .GetText .ColLetterToNumber("E"), 34, alertThueTNDN
'
'        .Sheet = 12
'            blCheck_S12 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'            blCheck_S11 = IIf(TAX_Utilities_New.NodeValidity.childNodes(10).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'            If blCheck_S12 = True Then
'                mCurrentSheet = .Sheet
'                .Col = .ColLetterToNumber("B")
'                .Row = 41
'                Do
'                      iWhiteRow = 1
'                      ' Clear cell note
'                        .Col = .ColLetterToNumber("L")
'                        .CellNote = ""
'                        .BackColor = mNonErrorColor
'                        .Col = .ColLetterToNumber("C")
'                        .CellNote = ""
'                        .BackColor = mNonErrorColor
'                        .Col = .ColLetterToNumber("U")
'                        .CellNote = ""
'                        .BackColor = mNonErrorColor
'                      ' End
'
'                      .GetText .ColLetterToNumber("B"), .Row, vSTT
'                      .GetText .ColLetterToNumber("C"), .Row, vCTieu
'                      .GetText .ColLetterToNumber("L"), .Row, vMST
'                      .GetText .ColLetterToNumber("R"), .Row, vTyLe
'                      .GetText .ColLetterToNumber("U"), .Row, vThueGTGT
'                      If Trim(vSTT) <> "" Or Trim(vCTieu) <> "" Or Trim(vMST) <> "" Or Val(Trim(vTyLe)) <> 0 Or Trim(vThueGTGT) <> "0" Then
'                        iWhiteRow = 0
'                      End If
'                      ' Kiem tra ctieu khong duoc trong
'                      If Trim$(vCTieu) = "" And iWhiteRow = 0 Then
'                            .Col = .ColLetterToNumber("C")
'                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                            .CellNote = errCTieu
'                            .BackColor = mErrorColor
'                            iFlagCtieu = 1
'                     Else
'                            .Col = .ColLetterToNumber("C")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
'                      End If
'                      ' Kiem tra cau truc MST
'                      If Trim(vMST) <> vbNullString Then
'                        vMSTtemp = vMST
'                        vMST = Replace(vMST, " ", "")
'                        vMST = Replace(vMST, "-", "")
'                         'vMST = Left(vMST, 10)
'                        strCheck = ""
'                           If Len(vMST) = 10 Or Len(vMST) = 13 Then
'                                If Len(vMST) = Len(vMSTtemp) Or Mid(vMSTtemp, 11, 1) = "-" Then
'                                    strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
'                                End If
'                           End If
'                           If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
'                                .Col = .ColLetterToNumber("L")
'                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                .CellNote = errMST
'                                .BackColor = mErrorColor
'                                iFlagMST = 1
'                            End If
'                        Else
'                            .Col = .ColLetterToNumber("L")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
'                        End If
'
'                        ' Kiem tra MST voi PL 05/TNDN
'                        tempRow = .Row
'                        If blCheck_S11 = True Then
'                            If Trim(vMST) <> vbNullString Then
'                                .Sheet = 11
'                                .Row = 38
'                                .Col = .ColLetterToNumber("C")
'                                Do
'                                    .Sheet = 11
'                                    .GetText .ColLetterToNumber("F"), .Row, vMSTtemp
'                                    If Len(Trim$(vMSTtemp)) = 13 Then
'                                        vMSTtemp = Left(vMSTtemp, 10) & "-" & Right(vMSTtemp, 3)
'                                    End If
'                                    If Len(Trim$(vMST)) = 13 Then
'                                        vMST = Left(vMST, 10) & "-" & Right(vMST, 3)
'                                    End If
'                                    If Trim(vMST) = Trim(vMSTtemp) Then
'                                        .GetText .ColLetterToNumber("W"), .Row - 1, vThuePL05
'                                        If Val(vThueGTGT) <> Val(vThuePL05) Then
'                                            .Sheet = 12
'                                            tempRow1 = .Row
'                                            .Row = tempRow
'                                            .Col = .ColLetterToNumber("U")
'                                            .CellNote = alertThueTNDN
'                                            .BackColor = mAlertColor
'                                            .Sheet = 11
'                                            .Row = tempRow1
'                                            Exit Do
'                                        End If
'                                    End If
'                                    .Row = .Row + 1
'                                    .Col = .ColLetterToNumber("C")
'                                Loop Until .Text = "aa"
'
'                                .Sheet = 12
'                            End If
'                      End If
'                     .Col = .ColLetterToNumber("B")
'                     .Row = tempRow + 1
'                Loop Until .Text = "aa"
'                .Sheet = .SheetCount
'                .SetText .ColLetterToNumber("B"), 32, IIf(iFlagCtieu = 1, "0", "1")
'                .SetText .ColLetterToNumber("B"), 33, IIf(iFlagMST = 1, "0", "1")
'
'            End If
            
            .Sheet = lSheet03_3A
            blCheck_S03_3A = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
            If blCheck_S03_3A = True Then
                ' Lay gia tri chi tieu tren PL 03-3A
                .GetText .ColLetterToNumber("U"), 47, vCtieu1
                .GetText .ColLetterToNumber("U"), 52, vCTieu5
                ' Lay gia tri chi tieu C4
                .Sheet = 1
                .GetText .ColLetterToNumber("O"), 43, vCTieuC4TK
'                If Val(vCTieuC4TK) < Val(vCtieu1) Or Val(vCTieuC4TK) < Val(vCTieu5) Then
'                    .Col = .ColLetterToNumber("O")
'                    .Row = 43
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                Else
'                    .Col = .ColLetterToNumber("O")
'                    .Row = 43
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrCtieuC4
'                    .BackColor = mAlertColor
'                End If
            End If
            
            ' Check validate sheet 01/GCN-QLT
            Dim vCheckMST, vDiaChi, vTenLK As Variant
            Dim vCheck(13) As Variant
            Dim strCol As String
            Dim ArrCol() As String
            Dim isCheck As Boolean
            Dim iFlagCheck As Integer
            Dim errHinhThucLK As Variant
            
            .Sheet = .SheetCount
            .GetText .ColLetterToNumber("E"), 33, errMST
            .GetText .ColLetterToNumber("E"), 38, errHinhThucLK
            .Sheet = lSheet03_7
            blCheck_S12 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
            If blCheck_S12 = True Then
                .Row = 82
                .Col = .ColLetterToNumber("B")
                strCol = "U~V~W~X~Y~Z~AA~AB~AC~AD~AE~AF~AG"
                ArrCol = Split(strCol, "~")
                Do
                    iWhiteRow = 1
                    .GetText .ColLetterToNumber("AH"), .Row, vCheckMST
                    If Trim$(CStr(vCheckMST)) = "GB" Then
                       .GetText .ColLetterToNumber("D"), .Row, vDiaChi
                       .GetText .ColLetterToNumber("I"), .Row, vDiaChi
                       .GetText .ColLetterToNumber("O"), .Row, vMST
                        
                        
                        ' Clear cell note
                        .Col = .ColLetterToNumber("O")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        
                        isCheck = False
                        For i = 0 To UBound(ArrCol)
                            .GetText .ColLetterToNumber(ArrCol(i)), .Row, vCheck(i)
                            
                            .Col = .ColLetterToNumber(ArrCol(i))
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                            If UCase$(Trim(CStr(vCheck(i)))) = "X" Or Trim(CStr(vCheck(i))) = "1" Then
                                isCheck = True
                            End If
                        Next
                        
                        '
                        If Trim$(vTenLK) <> "" Or Trim$(vDiaChi) <> "" Or Trim$(vMST) <> "" Or isCheck Then
                            iWhiteRow = 0
                        End If
                      ' Kiem tra cau truc MST
                      If Trim(vMST) <> vbNullString Then
                            vMSTtemp = vMST
                            vMST = Replace(vMST, " ", "")
                            vMST = Replace(vMST, "-", "")
                             'vMST = Left(vMST, 10)
                            strCheck = ""
                           If Len(vMST) = 10 Or Len(vMST) = 13 Then
                                If Len(vMST) = Len(vMSTtemp) Or Mid(vMSTtemp, 11, 1) = "-" Then
                                    strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                                End If
                           End If
                           If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                                .Col = .ColLetterToNumber("O")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = errMST
                                .BackColor = mErrorColor
                                iFlagMST = 1
                            End If
                        Else
                            .Col = .ColLetterToNumber("O")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                        
                        ' Kiem tra rang buoc check
                        If isCheck = False And iWhiteRow = 0 Then
                            For i = 0 To UBound(ArrCol)
                                .Col = .ColLetterToNumber(ArrCol(i))
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = errHinhThucLK
                                .BackColor = mErrorColor
                            Next
                            iFlagCheck = 1
                        Else
                            For i = 0 To UBound(ArrCol)
                                .Col = .ColLetterToNumber(ArrCol(i))
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            Next
                        End If
                
                    End If
                    .Row = .Row + 1
                    .Col = .ColLetterToNumber("B")
                Loop Until .Text = "kk"
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 33, IIf(iFlagMST = 1, "0", "1")
                .SetText .ColLetterToNumber("B"), 38, IIf(iFlagCheck = 1, "0", "1")
            End If
        
        .EventEnabled(EventAllEvents) = True
    End With
    
End Sub

'Calling this function whenever we don't need this class anymore.
'It will do some special task with special Error/ special Data
'no parameter
Public Sub finish()
    Dim pL_3A_U50, pL_3B_U60 As Double
    Dim pL_3A_U56, pL_3B_U65, pL_3C_X43 As Double
    Dim blCheck_S3A, blCheck_S3B, blCheck_S3C As Boolean
    
    ' lSheet
    Dim lSheet03_2A As Long, lSheet03_2B As Long, lSheet03_3A As Long, lSheet03_3B As Long, lSheet03_3C As Long
    Dim lSheet03_4 As Long, lSheet03_5 As Long, lSheet03_05 As Long
    
    lSheet03_3A = 7
    lSheet03_3B = 8
    lSheet03_3C = 9
    lSheet03_4 = 10
    lSheet03_5 = 11
    lSheet03_05 = 12
    
    ResetWarning

    With fps
        blCheck_S3A = IIf(TAX_Utilities_New.NodeValidity.childNodes(lSheet03_3A - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        blCheck_S3B = IIf(TAX_Utilities_New.NodeValidity.childNodes(lSheet03_3B - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        blCheck_S3C = IIf(TAX_Utilities_New.NodeValidity.childNodes(lSheet03_3C - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        ' Lay gia tri cua chi tieu dong trong PL 03-3A, 03-3B, 03-3C
        pL_3A_U50 = Sum_Cell(lSheet03_3A, 50, .ColLetterToNumber("U")) 'PL 03-3A
        pL_3B_U60 = Sum_Cell(lSheet03_3B, 60, .ColLetterToNumber("U")) 'PL 03-3B
'        ' Lay gia tri chi tieu C8 tren to khai quyet toan
        mCurrentSheet = 1
        .Sheet = mCurrentSheet
        If blCheck_S3A = True Or blCheck_S3B = True Then
            Dim vC47 As Variant
            .GetText .ColLetterToNumber("O"), 47, vC47
            If (vC47 = pL_3A_U50 + pL_3B_U60) Then
                ' Update gia tri dung (khong loi) trong sheet header
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 23, "1"
                ' Chuyen chi tieu C8 ve trang thai dung
                mCurrentSheet = 1
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 47
                .CellNote = ""
                .BackColor = mNonErrorColor
            Else
                ' Update gia tri sai (co loi) trong sheet header
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 23, "0"
                ' Lay thong bao loi
                Dim vErrMsg1 As Variant
                .GetText .ColLetterToNumber("E"), 23, vErrMsg1
                ' Chuyen chi tieu ve trang thai loi
                mCurrentSheet = 1
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 47
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsg1 'static
                .BackColor = mErrorColor
            End If
        End If
'        ' Lay gia tri dong o cac Phu Luc tuong ung
        pL_3A_U56 = Sum_Cell(lSheet03_3A, 56, .ColLetterToNumber("U"))
        pL_3B_U65 = Sum_Cell(lSheet03_3B, 65, .ColLetterToNumber("U"))
        pL_3C_X43 = Sum_Cell(lSheet03_3C, 43, .ColLetterToNumber("X"))
'        ' Lay gia tri cua chi tieu C8 tren to khai
        mCurrentSheet = 1
        .Sheet = mCurrentSheet
        If blCheck_S3A = True Or blCheck_S3B = True Or blCheck_S3C = True Then
            Dim vC48 As Variant
            .GetText .ColLetterToNumber("O"), 48, vC48
            If (vC48 = pL_3A_U56 + pL_3B_U65 + pL_3C_X43) Then
                ' Update gia tri dung trong sheet header
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 24, "1"
                ' Chuyen trang thai dung cho C9
                mCurrentSheet = 1
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 48
                .CellNote = ""
                .BackColor = mNonErrorColor
            Else
                ' Update gia tri sai trong sheet header
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 24, "0"
                ' Lay thong bao loi
                Dim vErrMsg2 As Variant
                .GetText .ColLetterToNumber("E"), 24, vErrMsg2
                ' Chuyen C9 ve trang thai bao loi
                mCurrentSheet = 1
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 48
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsg2 'static
                .BackColor = mErrorColor
            End If
        End If
'        ' Phu Luc 03-5
        UpdateSheets
        CheckDynamicError
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
        .GetText .ColLetterToNumber("E"), 56, vError1
        .GetText .ColLetterToNumber("E"), 57, vError2
        .GetText .ColLetterToNumber("E"), 58, vError3
        
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
            .Row = 56
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 56
            .Text = "1"
        End If
        
        If iFlagTaxCode2 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 57
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 57
            .Text = "1"
        End If
        
        
        If iFlagTaxCode3 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 58
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 58
            .Text = "1"
        End If
        
        .Sheet = 1
        .Col = .ColLetterToNumber("E")
        .Row = 3
        .CellNote = ""
        '.BackColor = mFormColor
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

'******************************
'Description: Set active properties in validity node
'Author: ThanhDX
'Date: 24/11/2005

'******************************
Public Sub SetActiveSheet() 'ByRef xmlNodeValid As MSXML.IXMLDOMNode)
    
End Sub

'******************************
'Description: Reset all error satus in sheet contain dynamic error.
'Author: ThanhDX
'Date: 31/12/2005

'******************************
Public Sub ResetErrorCells()
Dim i As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim GiaTri As Double, TongSo As Double
Dim lCtrl As Long, RowTongSo As Long, ColTongSo As Long

'    With fps
'            ' GD2
'            i = 1
'            .Col = .ColLetterToNumber("C")
'            .Row = 26
'
'            mCurrentSheet = 9
'            .Sheet = 9
'            Do
'                .Col = .ColLetterToNumber("AR")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'                 .Row = i + 26
'                 .Col = .ColLetterToNumber("C")
'                 i = i + 1
'             Loop Until .Text = "aa"
'
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 26, "1"
'    End With
End Sub

Public Function ResetData() As Boolean
    Dim startRow As Long
    Dim totalCell1  As Long
    Dim countRowReset1 As Long
    Dim endRow As Long, endCol As Long
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlNodeReset As MSXML.IXMLDOMNode
    With fps
        Set xmlCellNode = TAX_Utilities_New.Data(0).nodeFromID(GetCellID(fps, .ColLetterToNumber("O"), 23))
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode.lastChild
        Set xmlCellNode = xmlCellNode.firstChild
        Set xmlCellNode = xmlCellNode.firstChild
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), endCol, endRow

        totalCell1 = TAX_Utilities_New.Data(fps.SheetCount - 2).getElementsByTagName("Cell").length
        countRowReset1 = (totalCell1 - 8) / 6
        
         For Each xmlNodeReset In TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            .Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            .Col = lCol
            .Row = lRow
            
            If (lRow < 23 Or lRow >= endRow) And mCurrentSheet = 1 Then
                '
            ElseIf (lRow >= 20 + countRowReset1) And mCurrentSheet = (fps.SheetCount - 1) Then
            Else
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
                TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
            End If
        Next
    End With
End Function
'Xu ly khi nut Nhap lai dc kich hoat
Public Sub SetData()
    Dim intCurrSheet As Integer
    Dim arrResul() As String
    Dim blCheck_S05 As Boolean
    
    
        ' lSheet
    Dim lSheet03_2A As Long, lSheet03_2B As Long, lSheet03_3A As Long, lSheet03_3B As Long, lSheet03_3C As Long
    Dim lSheet03_4 As Long, lSheet03_5 As Long, lSheet03_05 As Long
    
    lSheet03_2A = 5
    lSheet03_2B = 6
    
    lSheet03_3A = 7
    lSheet03_3B = 8
    lSheet03_3C = 9
    lSheet03_4 = 10
    lSheet03_5 = 11
    lSheet03_05 = 12
        intCurrSheet = fps.Sheet
        With fps
            If .ActiveSheet = 1 Then
                mCurrentSheet = 1
                .Sheet = mCurrentSheet
                ' set lai trang thai dung cho C8, C9 to khai
                .Col = .ColLetterToNumber("O")
                .Row = 47
                .CellNote = ""
                .BackColor = mNonErrorColor

                .Col = .ColLetterToNumber("O")
                .Row = 48
                .CellNote = ""
                .BackColor = mNonErrorColor
'                'Set lai trang thai cho cac chi tieu B24, B25, B30
'                CellChange .ColLetterToNumber(S1TxtB24Col), S1TxtB24Row
'                CellChange .ColLetterToNumber(S1TxtB25Col), S1TxtB25Row
'                CellChange .ColLetterToNumber(S1TxtB30Col), S1TxtB30Row
            End If
             'set du lieu cho 03-2A
            If fps.ActiveSheet = lSheet03_2A Then
                mCurrentSheet = lSheet03_2A
                fps.Sheet = mCurrentSheet
                'Ham tao cac truong nam co trong PL
                Gene_Year
            End If
    
             'set du lieu cho 03-2B
            If fps.ActiveSheet = lSheet03_2B Then
                mCurrentSheet = lSheet03_2B
                fps.Sheet = mCurrentSheet
                'Ham tao cac truong nam co trong PL
                Gene_Year
            End If
    
    
            If fps.ActiveSheet = lSheet03_05 Then
                .Sheet = lSheet03_05
                 mCurrentSheet = .Sheet
                .Row = 34
                .Col = .ColLetterToNumber("L")
                .Text = 100
                UpdateCell fps, .Col, .Row, .Text
                .Row = 39
                .Col = .ColLetterToNumber("C")
                .Text = 3
                UpdateCell fps, .Col, .Row, .Text
           
             End If
            fps.Sheet = intCurrSheet
            mCurrentSheet = intCurrSheet
        End With
End Sub
' Sinh cac truong du lieu nam co trong PL
Sub Gene_Year()
    With fps
        .Col = fps.ColLetterToNumber(S2TxtB26Col)
        .Row = S2TxtB26Row
        .Text = TAX_Utilities_New.Year - 5
        .Lock = True
        .TypeHAlign = TypeHAlignCenter
        UpdateCell fps, .Col, .Row, .Text

        .Col = fps.ColLetterToNumber(S2TxtB27Col)
        .Row = S2TxtB27Row
        .Text = TAX_Utilities_New.Year - 4
        .Lock = True
        .TypeHAlign = TypeHAlignCenter
        UpdateCell fps, .Col, .Row, .Text

        .Col = fps.ColLetterToNumber(S2TxtB28Col)
        .Row = S2TxtB28Row
        .Text = TAX_Utilities_New.Year - 3
        .Lock = True
        .TypeHAlign = TypeHAlignCenter
        UpdateCell fps, .Col, .Row, .Text

        .Col = fps.ColLetterToNumber(S2TxtB29Col)
        .Row = S2TxtB29Row
        .Text = TAX_Utilities_New.Year - 2
        .Lock = True
        .TypeHAlign = TypeHAlignCenter
        UpdateCell fps, .Col, .Row, .Text

        .Col = fps.ColLetterToNumber(S2TxtB30Col)
        .Row = S2TxtB30Row
        .Text = TAX_Utilities_New.Year - 1
        .Lock = True
        .TypeHAlign = TypeHAlignCenter
        UpdateCell fps, .Col, .Row, .Text
    End With
End Sub

'Description: IsValidatedCell function check whether a Cell with lCol, lRow
'             is validated cell which the validation of this cell and the Cell
'             with lAnchorRow, lAnchorCol are the same.
'Author:ThanhDX
'Date:09/02/2006
'Input:
'Return:
Private Function IsValidatedCell(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long, ByVal lRow As Long, ByVal lCol As Long) As Boolean
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long
    
    If lAnchorCol = lCol And lAnchorRow = lRow Then
        IsValidatedCell = True
        Exit Function
    End If
    
    Set xmlCellNode = TAX_Utilities_New.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
    Set xmlCellsNode = xmlCellNode.parentNode
    
    'Get Index of anchor cell
    For lCtrl = 1 To xmlCellsNode.childNodes.length
        If GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID") = GetAttribute(xmlCellNode, "CellID") Then _
            Exit For
    Next
    
    Do
        Set xmlCellsNode = xmlCellsNode.nextSibling
        If Not xmlCellsNode Is Nothing Then
            If GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID") = fps.ColNumberToLetter(lCol) & "_" & lRow Then
                IsValidatedCell = True
                Exit Function
            End If
        End If
    Loop Until xmlCellsNode Is Nothing
    
End Function
' Ham tinh tong 1 truong trong PL dong
Private Function Sum_Cell(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long) As Double
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long
    
    Set xmlCellNode = TAX_Utilities_New.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
    Set xmlCellsNode = xmlCellNode.parentNode
    ' Gan gia tri bang gia tri ban dau cho ham
    Sum_Cell = CDbl(GetAttribute(xmlCellNode, "Value"))
    ' Lay so thu tu lCtrl cua truong ban dau
    For lCtrl = 1 To xmlCellsNode.childNodes.length
        If GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID") = GetAttribute(xmlCellNode, "CellID") Then
            Exit For
        End If
    Next
    ' Tim cac truong giong truong ban dau de tinh tong
    Do
        Set xmlCellsNode = xmlCellsNode.nextSibling
        If Not xmlCellsNode Is Nothing Then
            'Tinh tong
            Sum_Cell = Sum_Cell + CDbl(GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "Value"))
        End If
    Loop Until xmlCellsNode Is Nothing
End Function

' Ham lay ve Col and row
Private Function getArrRowCol(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long) As String
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim strValue As String
    Dim lCtrl As Long
    Dim cCol As Long
    Dim cRow As Long
    Set xmlCellNode = TAX_Utilities_New.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
    Set xmlCellsNode = xmlCellNode.parentNode
    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), cCol, cRow
    
    ' Lay so thu tu lCtrl cua truong ban dau
    For lCtrl = 1 To xmlCellsNode.childNodes.length
        If GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID") = GetAttribute(xmlCellNode, "CellID") Then
            Exit For
        End If
    Next
    ' Tim cac truong giong truong ban dau
    strValue = cCol & "_" & cRow
    Do
        Set xmlCellsNode = xmlCellsNode.nextSibling
        If Not xmlCellsNode Is Nothing Then
                  ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID"), cCol, cRow
                  strValue = strValue & "###" & cCol & "_" & cRow
        End If
    Loop Until xmlCellsNode Is Nothing
    getArrRowCol = strValue
End Function

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Integer, iCol As Long, iRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim tCol As Long, tRow As Long
    
        ' lSheet
    Dim lSheet03_2A As Long, lSheet03_2B As Long, lSheet03_3A As Long, lSheet03_3B As Long, lSheet03_3C As Long
    Dim lSheet03_4 As Long, lSheet03_5 As Long, lSheet03_05 As Long, lSheet03_6 As Long, lSheet03_7 As Long
    
    lSheet03_2A = 5
    lSheet03_2B = 6
    
    lSheet03_3A = 7
    lSheet03_3B = 8
    lSheet03_3C = 9
    lSheet03_4 = 10
    lSheet03_5 = 11
    lSheet03_05 = 12
    
    lSheet03_7 = 14

    If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
        With fps
            .Sheet = .ActiveSheet
            mCurrentSheet = .ActiveSheet
            iCol = .ActiveCol
            iRow = .ActiveRow
            GetCellSpan fps, iCol, iRow
    
            Select Case .Sheet
                Case lSheet03_3A ' Phu Luc 03-3A
                    mCurrentSheet = .Sheet
                    Set xmlCellNode = TAX_Utilities_New.Data(.Sheet - 1).nodeFromID(GetCellID(fps, iCol, iRow))
                    Set xmlCellsNode = xmlCellNode.parentNode
                    ' Chi tieu 3.3
                    ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(.Sheet, 49, .ColLetterToNumber("U")) - 1), "CellID"), tCol, tRow
                    .Row = tRow
                    .Col = tCol
                    UpdateCell fps, tCol, tRow, .value
                    ' Chi tieu 3.4
                    ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(.Sheet, 50, .ColLetterToNumber("U")) - 1), "CellID"), tCol, tRow
                    .Row = tRow
                    .Col = tCol
                    UpdateCell fps, tCol, tRow, .value
                    ' Chi tieu 4.3
                    ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(.Sheet, 54, .ColLetterToNumber("U")) - 1), "CellID"), tCol, tRow
                    .Row = tRow
                    .Col = tCol
                    UpdateCell fps, tCol, tRow, .value
                    ' Chi tieu 4.5
                    ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(.Sheet, 56, .ColLetterToNumber("U")) - 1), "CellID"), tCol, tRow
                    .Row = tRow
                    .Col = tCol
                    UpdateCell fps, tCol, tRow, .value

                Case lSheet03_3B ' Phu Luc 03-3B
                    mCurrentSheet = .Sheet
                    Set xmlCellNode = TAX_Utilities_New.Data(.Sheet - 1).nodeFromID(GetCellID(fps, iCol, iRow))
                    Set xmlCellsNode = xmlCellNode.parentNode
                    ' Chi tieu 4.1
                    ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(.Sheet, 58, .ColLetterToNumber("U")) - 1), "CellID"), tCol, tRow
                    .Row = tRow
                    .Col = tCol
                    UpdateCell fps, tCol, tRow, .value
                    ' Chi tieu 4.3
                    ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(.Sheet, 60, .ColLetterToNumber("U")) - 1), "CellID"), tCol, tRow
                    .Row = tRow
                    .Col = tCol
                    UpdateCell fps, tCol, tRow, .value
                    ' Chi tieu 5.2
                    ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(.Sheet, 63, .ColLetterToNumber("U")) - 1), "CellID"), tCol, tRow
                    .Row = tRow
                    .Col = tCol
                    UpdateCell fps, tCol, tRow, .value
                     ' Chi tieu 5.4
                    ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(.Sheet, 65, .ColLetterToNumber("U")) - 1), "CellID"), tCol, tRow
                    .Row = tRow
                    .Col = tCol
                    UpdateCell fps, tCol, tRow, .value
                Case lSheet03_3C ' Phu Luc 03-3C
                    mCurrentSheet = .Sheet
                Case lSheet03_4 ' Phu Luc 03-4
                    mCurrentSheet = .Sheet
                    .Row = iRow
                    .Col = .ColLetterToNumber("AE")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AH")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AN")
                    UpdateCell fps, .Col, .Row, .value
                Case lSheet03_05 ' Phu Luc 03-05
                    mCurrentSheet = .Sheet
            End Select
        End With
        With fps
            .Sheet = .ActiveSheet
            mCurrentSheet = .ActiveSheet
            ' To khai
            If mCurrentSheet = 1 Then
                i = 1
                .Col = .ColLetterToNumber("B")
                .Row = 60
                Do
                     .Text = str(i)
                     ' UpdateCell fps, .Col, .Row, .Text
                     .Col = .ColLetterToNumber("B")
                     .Row = i + 60
                     i = i + 1
                 Loop Until .Text = "aa"
            End If
            
            ' PL 03-4
            If KeyCode = vbKeyF5 Then
                If mCurrentSheet = lSheet03_4 Then
                    .SetText .ColLetterToNumber("AK"), .ActiveRow + 1, "25"
                    UpdateCell fps, .ColLetterToNumber("AK"), .ActiveRow + 1, "25"
                End If
            End If
            
            If mCurrentSheet = lSheet03_4 Then
                fps.EventEnabled(EventAllEvents) = False
'                .Row = iRow + 1
'                .Col = .ColLetterToNumber("AK")
'                .Text = "25"
'                UpdateCell fps, .Col, .Row, .Text
                
                i = 1
                .Col = .ColLetterToNumber("C")
                .Row = 26
                Do
                     .Text = str(i)
                     ' UpdateCell fps, .Col, .Row, .Text
                     .Col = .ColLetterToNumber("AK")
'                     .Text = "25"
'                     UpdateCell fps, .Col, .Row, .Text
                     
                     .Col = .ColLetterToNumber("C")
                     .Row = i + 26
                     i = i + 1
                 Loop Until .Text = "aa"
                 fps.EventEnabled(EventAllEvents) = True
            End If
            
            ' PL 03-05
            If .ActiveSheet = lSheet03_05 Then
                   fps.EventEnabled(EventAllEvents) = True
                    i = 0
                    .Col = .ColLetterToNumber("C")
                    .Row = 39
                    If iRow >= 39 Then
                        Do
                             If (.Row = 39 + i * 3) Then
                                .Text = str(i + 3)
                                i = i + 1
                             End If
                             .Col = .ColLetterToNumber("C")
                             .Row = .Row + 1
 '                            i = i + 1
                        Loop Until .Text = "aa"
                        .SetActiveCell iCol, iRow

                        .Row = iRow
                        .Col = iCol
                         fps.EventEnabled(EventAllEvents) = False
                         .SetActiveCell iCol, iRow
                    End If
            End If
            
            ' PL 03-7
            If .ActiveSheet = lSheet03_7 Then
                   fps.EventEnabled(EventAllEvents) = True
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 37
                    Do
                     .Text = "1.1." & Trim(str(i))
                     ' UpdateCell fps, .Col, .Row, .Text
                     .Row = i + 37
                     i = i + 1
                    Loop Until .Text = "aa"
                    
                   i = 1
                   .Row = .Row + 3
                    Do
                         .Text = "1.2." & Trim(str(i))
                         .Col = .ColLetterToNumber("B")
                      '   UpdateCell fps, .Col, .Row, .Text
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"


                   i = 1
                   .Row = .Row + 4
                    Do
                         .Text = "2.1." & Trim(str(i))
                         .Col = .ColLetterToNumber("B")
                       '  UpdateCell fps, .Col, .Row, .Text
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "cc"

                   i = 1
                   .Row = .Row + 3
                    Do
                         .Text = "2.2." & Trim(str(i))
                         .Col = .ColLetterToNumber("B")
                        ' UpdateCell fps, .Col, .Row, .Text
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "dd"


                   i = 1
                   .Row = .Row + 3
                    Do
                         .Text = "2.3." & Trim(str(i))
                         .Col = .ColLetterToNumber("B")
                         'UpdateCell fps, .Col, .Row, .Text
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "ee"

                   i = 1
                   .Row = .Row + 4
                    Do
                         .Text = "2.4.1." & Trim(str(i))
                         .Col = .ColLetterToNumber("B")
                         'UpdateCell fps, .Col, .Row, .Text
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "ff"

                   i = 1
                   .Row = .Row + 3
                    Do
                         .Text = "2.4.2." & Trim(str(i))
                         .Col = .ColLetterToNumber("B")
                         'UpdateCell fps, .Col, .Row, .Text
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "gg"


                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = "2.5.1." & Trim(str(i))
                         .Col = .ColLetterToNumber("B")
                         'UpdateCell fps, .Col, .Row, .Text
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "hh"

                    i = 1
                   .Row = .Row + 3
                    Do
                         .Text = "2.5.2." & Trim(str(i))
                         .Col = .ColLetterToNumber("B")
                         'UpdateCell fps, .Col, .Row, .Text
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "ii"
                    
                    fps.EventEnabled(EventAllEvents) = False
                    i = 1
                   .Row = .Row + 9
                    Do
                         .Text = i
                         .Col = .ColLetterToNumber("B")
                         'UpdateCell fps, .Col, .Row, .Text
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "kk"
                    


                   fps.EventEnabled(EventAllEvents) = True
            End If
            
            .EventEnabled(EventAllEvents) = True
        End With
    End If
End Sub
Private Sub XoaDauAm(Col As Long, Row As Long)
    With fps
            Dim vGiatriam As Variant
            .GetText Col, Row, vGiatriam
            If vGiatriam = 0 Then
                .SetText Col, Row, 0
            End If
    End With
End Sub

Private Function GetValidatedCellIndex(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long) As Integer
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long
    
  
    
    Set xmlCellNode = TAX_Utilities_New.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
    Set xmlCellsNode = xmlCellNode.parentNode
    
    'Get Index of anchor cell
    For lCtrl = 1 To xmlCellsNode.childNodes.length
        If GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID") = GetAttribute(xmlCellNode, "CellID") Then _
            Exit For
    Next
    GetValidatedCellIndex = lCtrl
    
    
End Function

Private Sub XacDinhLoiDong(ByVal XacDinhLoi As Boolean, ByVal ColLoi As Long, ByVal RowLoi As Long, ByVal RowHeader As Long, ByVal TenPL As String)
    With fps
        If XacDinhLoi Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), RowHeader, IIf(XacDinhLoi, "0", "1")
            Dim fstrFormula As String
            fstrFormula = TenPL & "!" & Trim(.ColNumberToLetter(CLng(ColLoi))) & Trim(str(RowLoi))
            .Col = .ColLetterToNumber("C")
            .Row = RowHeader
            .Formula = fstrFormula
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), RowHeader, "1"
        End If
    End With
End Sub


Private Sub ResetWarning()
    Dim arrStrPositions As Variant, arrStrPosition() As String
    Dim lCurSheet As Long, i As Long
    
    lCurSheet = fps.Sheet
    
    fps.EventEnabled(EventAllEvents) = False
    
    arrStrPositions = arrErrorCells.Keys
    For i = 1 To arrErrorCells.count
        arrStrPosition = Split(arrStrPositions(i - 1), "_")
        fps.Sheet = CLng(arrStrPosition(0))
        fps.Col = fps.ColLetterToNumber(arrStrPosition(1))
        fps.Row = CLng(arrStrPosition(2))
        fps.CellNote = ""
        fps.BackColor = arrErrorCells.Item(arrStrPositions(i - 1))
    Next
    arrErrorCells.RemoveAll
    fps.Sheet = lCurSheet
    
    fps.EventEnabled(EventAllEvents) = True
    
End Sub

Public Sub Prepare_KHBS()
'    Dim xmlDocument As New MSXML.DOMDocument
'    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
'    Dim xmlNode As MSXML.IXMLDOMNode
'    Dim xmlNodeItem As MSXML.IXMLDOMNode
'    Dim strPath As String
'    Dim strComboName As String
'    Dim strComboMCT As String
'    Dim strCombo As String
'
'    mblKhoiDauToKhai = True
'    fps.Sheet = fps.SheetCount - 1
'    fps.ActiveSheet = fps.Sheet
'
'    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT.xml"
'
'    If xmlDocument.Load(strPath) Then
'        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
'        For Each xmlNode In xmlNodeListItem
'            If GetAttribute(xmlNode, "ID") = GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") Then
'                Exit For
'            End If
'        Next
'    End If
'
'
'    Set xmlNodeListItem = xmlNode.childNodes
'
'    For Each xmlNodeItem In xmlNodeListItem
'                strCombo = strCombo + CPab(GetAttribute(xmlNodeItem, "Caption"), 100) + CPab(GetAttribute(xmlNodeItem, "CellID"), 20) + GetAttribute(xmlNodeItem, "MCT") + Chr$(9)
'                strComboName = strComboName + GetAttribute(xmlNodeItem, "Caption") + Chr$(9)
'                strComboMCT = strComboMCT + GetAttribute(xmlNodeItem, "MCT") + Chr$(9)
'    Next
'     With fps
'
'        .Sheet = .ActiveSheet
'        mCurrentSheet = .Sheet
'        .Col = .ColLetterToNumber("BC")
'        .Row = 9
'        .TypeComboBoxList = strCombo
'        .Row = 13
'        .TypeComboBoxList = strCombo
'        .Col = .ColLetterToNumber("BD")
'        .Row = 9
'        .TypeComboBoxList = strComboName
'        .Row = 13
'        .TypeComboBoxList = strComboName
'        .Col = .ColLetterToNumber("BE")
'        .Row = 9
'        .TypeComboBoxList = strComboMCT
'        .Row = 13
'        .TypeComboBoxList = strComboMCT
'    End With
    
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
        For Each xmlCellNodeData In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
            ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
            .GetText lCol, lRow, varTemp
            .Col = lCol
            .Row = lRow
            If .CellType = CellTypeCheckBox Then
                UpdateLastKHBSCell fps, lCol, lRow, IIf(CStr(varTemp) = 1, "x", vbNullString)
            Else
                UpdateLastKHBSCell fps, lCol, lRow, varTemp
            End If
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
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    Dim i As Integer
    Dim blnSheet As Boolean
    
            ' lSheet
    Dim lSheet03_2A As Long, lSheet03_2B As Long, lSheet03_3A As Long, lSheet03_3B As Long, lSheet03_3C As Long
    Dim lSheet03_4 As Long, lSheet03_5 As Long, lSheet03_05 As Long, lSheet03_6 As Long, lSheet03_7 As Long
    
    lSheet03_2A = 5
    lSheet03_2B = 6
    
    lSheet03_3A = 7
    lSheet03_3B = 8
    lSheet03_3C = 9
    lSheet03_4 = 10
    lSheet03_5 = 11
    lSheet03_05 = 12
    
    lSheet03_7 = 14
    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               If (lRow <= 54) And (lRow >= 23) Then
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
                End If
            Next
            
       ' sheet 2
       For i = 2 To .SheetCount - 1
            .Sheet = i
            mCurrentSheet = .Sheet
            blnSheet = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
            If blnSheet = True Then
                For Each xmlCellNodeData In TAX_Utilities_New.Data(.Sheet - 1).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    If i = lSheet03_3A Then
                        If (lRow <= 56) And (lRow >= 47) Then
                            .GetText lCol, lRow, varTemp
                            UpdateCell fps, lCol, lRow, varTemp
                        End If
                    ElseIf i = lSheet03_3B Then
                        If (lRow <= 65) And (lRow >= 52) Then
                             .GetText lCol, lRow, varTemp
                             UpdateCell fps, lCol, lRow, varTemp
                        End If
                    ElseIf i = lSheet03_3C Then
                        If (lRow <= 44) And (lRow >= 40) Then
                            .GetText lCol, lRow, varTemp
                            UpdateCell fps, lCol, lRow, varTemp
                        End If
                    ElseIf i = lSheet03_7 Then
                        If lRow < TAX_Utilities_New.Data(.Sheet - 1).getElementsByTagName("cell").length - 13 Then
                            .GetText lCol, lRow, varTemp
                            UpdateCell fps, lCol, lRow, varTemp
                        End If
                    Else
                         .GetText lCol, lRow, varTemp
                         UpdateCell fps, lCol, lRow, varTemp
                    End If
                Next
                
            End If
       Next i
        
      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

' Lay ve danh sach cac chi tieu dieu chinh KHBS
Function getMapCT() As String
    On Error GoTo ErrHandle
    
    Dim tempValue As Variant
    Dim strPath As String
    Dim i As Integer
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT_TT28.xml"
    
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "ID") = GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") Then
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
    SaveErrorLog "03_TNDN", "getMapCT", Err.number, Err.Description
End Function


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
            Set xmlNode = .nodeFromID("O_46")
            result = "CTC7_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("O_47")
            result = result & "~" & "CTC8_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("O_48")
            result = result & "~" & "CTC9_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("O_49")
            result = result & "~" & "CTC10_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("O_51")
            result = result & "~" & "CTD_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_03TNDN", "getValueTK", Err.number, Err.Description
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
                .GetText .ColLetterToNumber("O"), 46, varTemp
                strCurrentValue = "CTC7_" & Trim(varTemp)
                .GetText .ColLetterToNumber("O"), 47, varTemp
                strCurrentValue = strCurrentValue & "~CTC8_" & Trim(varTemp)
                .GetText .ColLetterToNumber("O"), 48, varTemp
                strCurrentValue = strCurrentValue & "~CTC9_" & Trim(varTemp)
                .GetText .ColLetterToNumber("O"), 49, varTemp
                strCurrentValue = strCurrentValue & "~CTC10_" & Trim(varTemp)
                .GetText .ColLetterToNumber("O"), 51, varTemp
                strCurrentValue = strCurrentValue & "~CTD_" & Trim(varTemp)
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
                    
                    If Left(strMaCT, 3) = "C7_" Or Left(strMaCT, 3) = "C8_" Or Left(strMaCT, 2) = "D_" Then
                        If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
                            ' dieu chinh giam
                            strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                        End If
                        If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
                            ' dieu chinh tang
                            strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                        End If
                    Else
                        If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
                            ' dieu chinh giam
                            strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                        End If
                        If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
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

' Lay ve chi tieu D o to khai lan truoc
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
            Set xmlNode = .nodeFromID("O_52")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_03TNDN", "getValueCTDC", Err.number, Err.Description
End Function


' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("O"), 52, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function

' Ham check phu luc 03-3C/TNDN
Private Sub checkPL03_3C()
Dim startRow As Long
    Dim endRow As Long, endCol As Long
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellNodes As MSXML.IXMLDOMNode
    Dim countSection As Long
    Dim ct4, ct3 As Variant
    Dim idx As Long, i As Integer
    Dim XacDinhLoi As Boolean
    Dim errCT3_4 As Variant
    
            ' lSheet
    Dim lSheet03_2A As Long, lSheet03_2B As Long, lSheet03_3A As Long, lSheet03_3B As Long, lSheet03_3C As Long
    Dim lSheet03_4 As Long, lSheet03_5 As Long, lSheet03_05 As Long, lSheet03_6 As Long, lSheet03_7 As Long
    
    lSheet03_2A = 5
    lSheet03_2B = 6
    
    lSheet03_3A = 7
    lSheet03_3B = 8
    lSheet03_3C = 9
    lSheet03_4 = 10
    lSheet03_5 = 11
    lSheet03_05 = 12
    
    lSheet03_7 = 14
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 26, errCT3_4
        
        
        .Sheet = lSheet03_3C
        mCurrentSheet = lSheet03_3C
        
        Set xmlCellNode = TAX_Utilities_New.Data(lSheet03_3C - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("D"), 24))
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode
        countSection = xmlCellNode.childNodes.length
        Set xmlCellNode = xmlCellNode.parentNode.firstChild
        
        For idx = 0 To countSection - 1
            ' Lay gia tri ct3
            Set xmlCellNodes = xmlCellNode.childNodes(idx)
            Set xmlCellNodes = xmlCellNodes.firstChild
            For i = 0 To 12
                Set xmlCellNodes = xmlCellNodes.nextSibling
            Next i
            ParserCellID fps, GetAttribute(xmlCellNodes, "CellID"), endCol, endRow
            .Col = endCol
            .Row = endRow
            .GetText .Col, .Row, ct3
            ' Lay gia tri ct4
            Set xmlCellNodes = xmlCellNodes.nextSibling
            ParserCellID fps, GetAttribute(xmlCellNodes, "CellID"), endCol, endRow
            .Col = endCol
            .Row = endRow
            .GetText .Col, .Row, ct4
            If Val(ct4) > Val(ct3) Then
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = errCT3_4
                .BackColor = mErrorColor
                XacDinhLoi = True
            Else
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
        Next idx
        
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 26, IIf(XacDinhLoi, "0", "1")
        
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub


' Ham set unlock cho phep nhap cac chi tieu doi voi to khai B
Private Sub unLockCell(status As Boolean)
    With fps
        .Sheet = 1
        .Col = .ColLetterToNumber("O")
        ' chi tieu C3
        .Row = 42
        .Lock = status
        'chi tieu C8
        .Row = 47
        .Lock = status
        ' chi tieu C9
        .Row = 48
        .Lock = status
        ' chi tieu C10
        .Row = 49
        .Lock = status
    End With
End Sub

' Lock cac cell lay du lieu len tu phu luc
Public Sub unLockCellPL(strLoaiToKhai As String)
    Dim blCheck_S2A As Boolean, blCheck_S2B As Boolean, blCheck_S3A As Boolean, blCheck_S3B As Boolean
    Dim blCheck_S3C As Boolean, blCheck_S4 As Boolean, blCheck_S5 As Boolean
    
            ' lSheet
    Dim lSheet03_2A As Long, lSheet03_2B As Long, lSheet03_3A As Long, lSheet03_3B As Long, lSheet03_3C As Long
    Dim lSheet03_4 As Long, lSheet03_5 As Long, lSheet03_05 As Long, lSheet03_6 As Long, lSheet03_7 As Long
    
    lSheet03_2A = 5
    lSheet03_2B = 6
    
    lSheet03_3A = 7
    lSheet03_3B = 8
    lSheet03_3C = 9
    lSheet03_4 = 10
    lSheet03_5 = 11
    lSheet03_05 = 12
    
    lSheet03_7 = 14
    
    blCheck_S2A = IIf(TAX_Utilities_New.NodeValidity.childNodes(lSheet03_2A - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S2B = IIf(TAX_Utilities_New.NodeValidity.childNodes(lSheet03_2B - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S3A = IIf(TAX_Utilities_New.NodeValidity.childNodes(lSheet03_3A - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S3B = IIf(TAX_Utilities_New.NodeValidity.childNodes(lSheet03_3B - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S3C = IIf(TAX_Utilities_New.NodeValidity.childNodes(lSheet03_3C - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S4 = IIf(TAX_Utilities_New.NodeValidity.childNodes(lSheet03_4 - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S5 = IIf(TAX_Utilities_New.NodeValidity.childNodes(lSheet03_5 - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    If strLoaiToKhai = "TKCT" Then
         With fps
            .Sheet = 1
            ' lock chi tieu C3
            .Col = .ColLetterToNumber("O")
            'If blCheck_S2A Or blCheck_S2B Then
            If blCheck_S2A Then
                .Row = 42
                .Lock = True
            Else
                .Row = 42
                .Lock = False
            End If
            ' Lock chi tieu C8
            If blCheck_S3A Or blCheck_S3B Then
                .Row = 47
                .Lock = True
            Else
                .Row = 47
                .Lock = False
            End If
            ' Lock chi tieu c9
            If blCheck_S3A Or blCheck_S3B Or blCheck_S3C Then
                .Row = 48
                .Lock = True
            Else
                .Row = 48
                .Lock = False
            End If
            ' Chi tieu C10
            If blCheck_S4 Then
                .Row = 49
                .Lock = True
            Else
                .Row = 49
                .Lock = False
            End If
            ' Chi tieu D
            If blCheck_S5 Then
                .Row = 51
                .Lock = True
            Else
                .Row = 51
                .Lock = False
            End If
            
        End With
   End If
End Sub


' Lay ve cac gia tri cua to khai quy
Public Function GetValue32_31TK_Quy() As String()
     
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim fso As New FileSystemObject
    Dim strDataFileName1A As String, strDataFileName1B As String
    Dim str32_31TK_Quy As String
    Dim arrResult(3) As String
    Dim yearTK As Integer, i As Integer
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        yearTK = TAX_Utilities_New.Year
        For i = 1 To 4
            strDataFileName1A = TAX_Utilities_New.DataFolder & GetAttribute(GetValidityNode("11").childNodes(0), "DataFile") & "_0" & i & TAX_Utilities_New.Year & ".xml"
            strDataFileName1B = TAX_Utilities_New.DataFolder & GetAttribute(GetValidityNode("12").childNodes(0), "DataFile") & "_0" & i & TAX_Utilities_New.Year & ".xml"
            If fso.FileExists(strDataFileName1A) Then
                If .Load(strDataFileName1A) = True Then
                    Set xmlNode = .nodeFromID("K_33")  'chi tieu 32 tren to khai 01A_TNDN
                    arrResult(i - 1) = GetAttribute(xmlNode, "Value")
                    Set xmlNode = Nothing
                Else
                    arrResult(i - 1) = 0
                End If
            ElseIf fso.FileExists(strDataFileName1B) Then
                If .Load(strDataFileName1B) = True Then
                    Set xmlNode = .nodeFromID("J_47")  'chi tieu 31 tren to khai 01B_TNDN
                    arrResult(i - 1) = GetAttribute(xmlNode, "Value")
                    Set xmlNode = Nothing
                Else
                    arrResult(i - 1) = 0
                End If
            Else
                arrResult(i - 1) = 0
            End If
            
        Next i
        
        '***************************************
                
'        If .Load(strDataFileName) = True Then
'            Set xmlNode = .nodeFromID("L_33")  'L_33: col & row of 43 target
'            strValue = GetAttribute(xmlNode, "Value")
'            Set xmlNode = Nothing
'        ElseIf .parseError.reason <> vbNullString Then
'            If InStr(1, .parseError.errorCode, "2146697210") <> 0 Then
'                'file data of last month does not exist
'            Else
'                MsgBox .parseError.reason
'            End If
'            GetLast43Previous = False
'        End If
        
    End With
    GetValue32_31TK_Quy = arrResult
    Set xmlDomLastData = Nothing
    Set xmlNode = Nothing
    
    Exit Function
    
ErrorHandle:
    SaveErrorLog "cls001", "GetValue32_31TK_Quy", Err.number, Err.Description
End Function


' TOng hop gia tri len PL05
Public Sub tongHopPL05()
         ' Lay gia tri cua cac quy khac de set vao phu luc 05/TNDN
         Dim arrResul() As String
         Dim blCheck_S05 As Boolean
        arrResul = GetValue32_31TK_Quy
        blCheck_S05 = IIf(TAX_Utilities_New.NodeValidity.childNodes(11).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        With fps
            If blCheck_S05 = True Then
                fps.EventEnabled(EventAllEvents) = False
                .Sheet = 12
                mCurrentSheet = .Sheet
                .Row = 34
                ' Quy 1
                .Col = .ColLetterToNumber("O")
                .Text = arrResul(0)
                ' Quy 2
                .Col = .ColLetterToNumber("P")
                .Text = arrResul(1)
                ' Quy 3
                .Col = .ColLetterToNumber("Q")
                .Text = arrResul(2)
                ' Quy 4
                .Col = .ColLetterToNumber("R")
                .Text = arrResul(3)
                fps.EventEnabled(EventAllEvents) = True
            End If
        End With
End Sub

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
    For Each xmlNodeValid In TAX_Utilities_New.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
                Case "KHBS"
                    blnNullValue = True
                    'Check value of last node in Data
                    Set xmlCellNode = TAX_Utilities_New.Data(14).getElementsByTagName("Cell") _
                        (TAX_Utilities_New.Data(14).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_New.Data(14).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_New.Data(14).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_New.Data(14).getElementsByTagName("Cell").length > 20 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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
    tuRowDL = 16
    denRowDL = 21
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "17_18_2~17_18_3~17_18_4~16_18_5~16_18_6~19_21_7~20_22_8~19_21_9~14_16_10~15_17_11~14_16_12~20_30_15"
    ' end
End Sub

