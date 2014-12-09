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
'this Class is belong to TAX_Business_v1 project which will be compline to DLL

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
'Private Const S2TxtB26Row = 26
'Private Const S2TxtB26Col = "B"
'Private Const S2TxtB27Row = 27
'Private Const S2TxtB27Col = "B"
'Private Const S2TxtB28Row = 28
'Private Const S2TxtB28Col = "B"
'Private Const S2TxtB29Row = 29
'Private Const S2TxtB29Col = "B"
'Private Const S2TxtB30Row = 30
'Private Const S2TxtB30Col = "B"


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

Private mTempColor As Variant

Public hanNopTKTNDN As String

' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

' khai bao toa do cac cell va row
Private Const sheetTK03_TNDN = 1
Private Const sheetPl03_1A_TNDN = 2
Private Const sheetPl03_1B_TNDN = 3
Private Const sheetPl03_1C_TNDN = 4
Private Const sheetPl03_2A_TNDN = 5
Private Const sheetPl03_2B_TNDN = 6
Private Const sheetPl03_3A_TNDN = 7
Private Const sheetPl03_3B_TNDN = 8
Private Const sheetPl03_3C_TNDN = 9
Private Const sheetPl03_4_TNDN = 10
Private Const sheetPl03_5_TNDN = 11
Private Const sheetPl03_6_TNDN = 12
Private Const sheetPl03_7_TNDN = 13
Private Const sheetPl03_8_TNDN = 14
Private Const sheetPl03_9_TNDN = 15
Private Const sheetPl03_KHBS_TNDN = 16


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
    SetDateFormat fps, sheetTK03_TNDN, 87, .ColLetterToNumber("I"), DDMMYYYY
    SetDateFormat fps, sheetTK03_TNDN, 95, .ColLetterToNumber("J"), DDMMYYYY
    SetDateFormat fps, sheetTK03_TNDN, 95, .ColLetterToNumber("N"), DDMMYYYY
    SetDateFormat fps, sheetTK03_TNDN, 110, .ColLetterToNumber("N"), DDMMYYYY
    
    .Sheet = sheetTK03_TNDN ' Mac dinh ngay ky la ngay he thong (To khai quyet toan TNDN 03/TNDN)
    .Row = 110
    .Col = .ColLetterToNumber("N")
    .Text = Format(Date, "dd/mm/yyyy")
    .TypeHAlign = TypeHAlignLeft
    
    ' Set nganh nge co ty le doanh thu cao nhat
    
    strDataFileName = GetAbsolutePath("..\InterfaceIni\DM_nganh_TLDT_cao.xml")
    If xmlDomData.Load(strDataFileName) Then
         Set xmlNodeListItems = xmlDomData.getElementsByTagName("Item")
         For Each xmlNode In xmlNodeListItems
             fldList = Split(GetAttribute(xmlNode, "Value"), "###")
             'strCboLyDo = strCboLyDo + fldList(0) + "-" + fldList(1) + Chr$(9)
             strCboLyDo = strCboLyDo + fldList(1) + Chr$(9)
             strCboLyDoAn = strCboLyDoAn + fldList(0) + Chr$(9)
         Next
     End If
     'set nganh nghe co ty le doanh thu cao nhat
    .Row = 14
    .Col = .ColLetterToNumber("G")
    .TypeComboBoxList = strCboLyDo
    ' set Ma nganh nghe co ty le doanh thu cao nhat
    .Col = .ColLetterToNumber("P")
    .TypeComboBoxList = strCboLyDoAn
        
   ' Ky dc gia han moi set gia tri ly do gia han
   strCboLyDo = ""
   strCboLyDoAn = ""
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
    .Row = 85
    .Col = .ColLetterToNumber("I")
    .TypeComboBoxList = strCboLyDo
    ' set Ma ly do gia han
    .Col = .ColLetterToNumber("P")
    .TypeComboBoxList = strCboLyDoAn
   
    
    
    .Row = 46
    .Col = .ColLetterToNumber("O")
    mTempColor = .BackColor
    
    ' PL 03-3A
    FormatTextNumber fps, sheetPl03_3A_TNDN, .ColLetterToNumber("N"), 38
    SetDateFormat fps, sheetPl03_3A_TNDN, 38, .ColLetterToNumber("U"), yyyy
    FormatTextNumber fps, sheetPl03_3A_TNDN, .ColLetterToNumber("J"), 39
    SetDateFormat fps, sheetPl03_3A_TNDN, 39, .ColLetterToNumber("Q"), yyyy
    FormatTextNumber fps, sheetPl03_3A_TNDN, .ColLetterToNumber("N"), 40
    SetDateFormat fps, sheetPl03_3A_TNDN, 40, .ColLetterToNumber("U"), yyyy
    ' PL 03-3B
    FormatTextNumber fps, sheetPl03_3B_TNDN, .ColLetterToNumber("N"), 44
    SetDateFormat fps, sheetPl03_3B_TNDN, 44, .ColLetterToNumber("U"), yyyy
    FormatTextNumber fps, sheetPl03_3B_TNDN, .ColLetterToNumber("J"), 45
    SetDateFormat fps, sheetPl03_3B_TNDN, 45, .ColLetterToNumber("Q"), yyyy
    FormatTextNumber fps, sheetPl03_3B_TNDN, .ColLetterToNumber("N"), 46
    SetDateFormat fps, sheetPl03_3B_TNDN, 46, .ColLetterToNumber("U"), yyyy
    SetDateFormat fps, sheetPl03_3B_TNDN, 41, .ColLetterToNumber("S"), DDMMYYYY
    ' PL 03-3C
    SetDateFormat fps, sheetPl03_3C_TNDN, 27, .ColLetterToNumber("V"), DDMMYYYY
    SetDateFormat fps, sheetPl03_3C_TNDN, 31, .ColLetterToNumber("Y"), DDMMYYYY

    
     SetDateFormat fps, .SheetCount - 1, 23, .ColLetterToNumber("BG"), DDMMYYYY
     
    
     
     SetDateFormat fps, .SheetCount - 1, 31, .ColLetterToNumber("BE"), DDMMYYYY
     SetDateFormat fps, .SheetCount - 1, 46, .ColLetterToNumber("BG"), DDMMYYYY
    .Sheet = .SheetCount - 1 'KHBS
    .Row = 46
    .Col = .ColLetterToNumber("BG")
    .Text = Format(Date, "dd/mm/yyyy")
    .TypeHAlign = TypeHAlignLeft
    
    ' Lay gia tri cua cac quy khac de set vao phu luc 05/TNDN
    arrResul = GetValue32_31TK_Quy
    blCheck_S05 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_8_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    If blCheck_S05 = True Then
        .Sheet = sheetPl03_8_TNDN
        mCurrentSheet = .Sheet
        .Row = 41
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
    
    ' PL 03-7
       strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_PPDG_TNDN.xml")
        ' Lay danh muc phuong phap xac dinh gia
        ' 26/08/2014
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
        .Sheet = sheetPl03_7_TNDN
        strTemp = "51~55~60~64~68~73~77~81"
        arrRowCom = Split(strTemp, "~")
        For i = 0 To UBound(arrRowCom)
            .Row = CLng(arrRowCom(i))
            ' PP dinh gia
            .Col = .ColLetterToNumber("N")
            .TypeComboBoxList = strComboHien
            
            .Col = .ColLetterToNumber("AB")
            .TypeComboBoxList = strComboHien
            
            .Col = .ColLetterToNumber("AJ")
            .TypeComboBoxList = strCombo


            .Col = .ColLetterToNumber("AK")
            .TypeComboBoxList = strCombo

        Next
    .ActiveSheet = 1
End With

    ' set co quan thue quan ly
    setCQTQuanLy
    
    setCQTQuanLyHoanThue fps
    
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
    
    ' kiem tra gia tri chi tieu C4
    Dim chiTieuC4 As Variant

    Dim tuNgayPhatNC As String
    Dim ngayTC As String
    Dim thangTC As String
    
    Dim arrCellTag() As String
    
    Set arrErrorCells = New Scripting.Dictionary
    
    With fps
        ' set ngay nop cham tu ngay
        .Sheet = sheetTK03_TNDN
        mCurrentSheet = .Sheet
        If Len(CStr(GetNgayTaiChinh(GetNgayBatDauNamTaiChinh))) = 1 Then
            ngayTC = "0" & GetNgayTaiChinh(GetNgayBatDauNamTaiChinh)
        Else
            ngayTC = GetNgayTaiChinh(GetNgayBatDauNamTaiChinh)
        End If
        
        If Len(CStr(GetThangTaiChinh(GetNgayBatDauNamTaiChinh) + 1)) = 1 Then
            thangTC = "0" & GetThangTaiChinh(GetNgayBatDauNamTaiChinh) + 1
        Else
            thangTC = GetThangTaiChinh(GetNgayBatDauNamTaiChinh) + 1
        End If
        
        tuNgayPhatNC = ngayTC & "/" & thangTC & "/" & CInt(TAX_Utilities_v1.Year) + 1
        .Row = 95
        .Col = .ColLetterToNumber("J")
        .Text = tuNgayPhatNC
        .TypeHAlign = TypeHAlignLeft
        UpdateCell fps, .Col, .Row, .Text
        'Phu luc 03_2A TNDN
        .Sheet = sheetPl03_2A_TNDN
            .ActiveSheet = .Sheet
            mCurrentSheet = .Sheet
            'Khoi tao cac truong nam cho phu luc 03_2
            If .SheetVisible = True Then
                Gene_Year
            End If
        .Sheet = sheetPl03_2B_TNDN 'Phu luc 03_2B TNDN (SheetName: 03_2)
            .ActiveSheet = .Sheet
            mCurrentSheet = .Sheet
            'Khoi tao cac truong nam cho phu luc 03_2
            If .SheetVisible = True Then
                Gene_Year
            End If
         'Phu luc 03_3A TNDN (SheetName: 03_3A)
        .Sheet = sheetPl03_3A_TNDN
            mCurrentSheet = .Sheet
            Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl03_3A_TNDN - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("U"), 38))
            Set xmlCellsNode = xmlCellNode.parentNode

            Do
                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(sheetPl03_3A_TNDN, 38, .ColLetterToNumber("U")) - 1), "CellID"), Col_PL_3A_U38, Row_PL_3A_U38
                ValidateDateError fps, sheetPl03_3A_TNDN, Row_PL_3A_U38, Col_PL_3A_U38, yyyy

                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(sheetPl03_3A_TNDN, 39, .ColLetterToNumber("Q")) - 1), "CellID"), Col_PL_3A_Q39, Row_PL_3A_Q39
                ValidateDateError fps, sheetPl03_3A_TNDN, Row_PL_3A_Q39, Col_PL_3A_Q39, yyyy

                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(sheetPl03_3A_TNDN, 40, .ColLetterToNumber("U")) - 1), "CellID"), Col_PL_3A_U40, Row_PL_3A_U40
                ValidateDateError fps, sheetPl03_3A_TNDN, Row_PL_3A_U40, Col_PL_3A_U40, yyyy

                Set xmlCellsNode = xmlCellsNode.nextSibling
            Loop Until xmlCellsNode Is Nothing
        
        'Phu luc 03_3B TNDN (SheetName: 03_3B)
        .Sheet = sheetPl03_3B_TNDN
            mCurrentSheet = .Sheet

            Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl03_3B_TNDN - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("U"), 44))
            Set xmlCellsNode = xmlCellNode.parentNode

            Do
                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(sheetPl03_3B_TNDN, 44, .ColLetterToNumber("U")) - 1), "CellID"), Col_PL_3B_U44, Row_PL_3B_U44
                ValidateDateError fps, sheetPl03_3B_TNDN, Row_PL_3B_U44, Col_PL_3B_U44, yyyy

                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(sheetPl03_3B_TNDN, 45, .ColLetterToNumber("Q")) - 1), "CellID"), Col_PL_3B_Q45, Row_PL_3B_Q45
                ValidateDateError fps, sheetPl03_3B_TNDN, Row_PL_3B_Q45, Col_PL_3B_Q45, yyyy

                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(sheetPl03_3B_TNDN, 46, .ColLetterToNumber("U")) - 1), "CellID"), Col_PL_3B_U46, Row_PL_3B_U46
                ValidateDateError fps, sheetPl03_3B_TNDN, Row_PL_3B_U46, Col_PL_3B_U46, yyyy

                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(sheetPl03_3B_TNDN, 41, .ColLetterToNumber("S")) - 1), "CellID"), Col_PL_3B_S41, Row_PL_3B_S41
                ValidateDateError fps, sheetPl03_3B_TNDN, Row_PL_3B_S41, Col_PL_3B_S41, DDMMYYYY

                Set xmlCellsNode = xmlCellsNode.nextSibling
            Loop Until xmlCellsNode Is Nothing
        
        'Phu luc 03_3C TNDN (SheetName: 03_3C)
        .Sheet = sheetPl03_3C_TNDN
            mCurrentSheet = .Sheet

            Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl03_3C_TNDN - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("V"), 27))
            Set xmlCellsNode = xmlCellNode.parentNode

            Do
                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(sheetPl03_3C_TNDN, 27, .ColLetterToNumber("V")) - 1), "CellID"), Col_PL_3C_V27, Row_PL_3C_V27
                ValidateDateError fps, sheetPl03_3C_TNDN, Row_PL_3C_V27, Col_PL_3C_V27, DDMMYYYY

                ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(sheetPl03_3C_TNDN, 31, .ColLetterToNumber("Y")) - 1), "CellID"), Col_PL_3C_Y31, Row_PL_3C_Y31
                ValidateDateError fps, sheetPl03_3C_TNDN, Row_PL_3C_Y31, Col_PL_3C_Y31, DDMMYYYY

                Set xmlCellsNode = xmlCellsNode.nextSibling
            Loop Until xmlCellsNode Is Nothing

         .Sheet = sheetTK03_TNDN
         .EventEnabled(EventAllEvents) = False
         mCurrentSheet = .Sheet
        'set to khai bo sung hay chinh thuc
        
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("E")
            .Row = 22
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 22
            .Text = 0
            UpdateCell fps, .Col, .Row, .Text
         ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("E")
            .Row = 22
            .Text = 0
              UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 22
            .Text = Val(StrSolanBosung)
            UpdateCell fps, .Col, .Row, .Text
            ' set unlock cac cell trong truong hop to khai KHBS
            ' todo
            unLockCell (False)
        End If
       'ma NNKD
        .Col = .ColLetterToNumber("J")
        .Row = 22
        .Text = strLoaiNNKD
        UpdateCell fps, .Col, .Row, .Text

        ' Lock cac cell neu co tong hop tu phu luc
        ' todo
        unLockCellPL (Strloaitk)

        ' Kiem tra neu C4 <0 thi lock C5, C7, C8 va set =0
        .Sheet = sheetTK03_TNDN
        .GetText .ColLetterToNumber("O"), 53, chiTieuC4
        ' chi tieu C5, C7, C8
        If Val(chiTieuC4) < 0 Then
            .Col = .ColLetterToNumber("O")
            ' chi tieu C5
            .Row = 54
            .Lock = True
            .Text = 0
            ' chi tieu C7
            .Row = 56
            .Lock = True
            .Text = 0
            ' chi tieu C8
            .Row = 57
            .Lock = True
            .Text = 0
        Else
            .Col = .ColLetterToNumber("O")
            ' chi tieu C5
            .Row = 54
            .Lock = False
            ' chi tieu C7
            .Row = 56
            .Lock = False
            ' chi tieu C8
            .Row = 57
            .Lock = False
        End If
        ' End check
        ' Neu check gia han moi cho phep chon GT
        .Col = .ColLetterToNumber("F")
        .Row = 83
        If .Text = "1" Or UCase$(.Text) = "X" Then
            .Col = .ColLetterToNumber("I")
            .Row = 85
            .Lock = False
            
            .Col = .ColLetterToNumber("I")
            .Row = 87
            .Lock = False
            
            .Col = .ColLetterToNumber("I")
            .Row = 89
            .Lock = False
            
            .Col = .ColLetterToNumber("I")
            .Row = 91
            .Lock = False
        Else
            .Col = .ColLetterToNumber("I")
            .Row = 85
            .Lock = True
            
            .Col = .ColLetterToNumber("I")
            .Row = 87
            .Lock = True
            
            .Col = .ColLetterToNumber("I")
            .Row = 89
            .Lock = True
            
            .Col = .ColLetterToNumber("I")
            .Row = 91
            .Lock = True
        End If
        ' End set gia han
        
        ' Pl 03-5/TNDN
        .Sheet = sheetPl03_5_TNDN
        Dim ct01, ct03, ct04, ct05, ct06, ct07, ct08, ct10, ct11a As Variant
        If .SheetVisible = True Then
            .GetText .ColLetterToNumber("E"), 21, ct01
            .GetText .ColLetterToNumber("E"), 23, ct03
            .GetText .ColLetterToNumber("E"), 24, ct04
            .GetText .ColLetterToNumber("E"), 25, ct05
            .GetText .ColLetterToNumber("E"), 26, ct06
            .GetText .ColLetterToNumber("E"), 27, ct07
            .GetText .ColLetterToNumber("E"), 28, ct08
            
            .GetText .ColLetterToNumber("E"), 30, ct10
            .GetText .ColLetterToNumber("E"), 32, ct11a
            
            ' lock chi tieu 10
            If Val(ct01) - Val(ct03) - Val(ct04) - Val(ct05) - Val(ct06) - Val(ct07) - Val(ct08) < 0 Then
                .Row = 30
                .Col = .ColLetterToNumber("E")
                .Lock = True
                .Text = "0"
            Else
                .Row = 30
                .Col = .ColLetterToNumber("E")
                .Lock = False
            End If
            
            
            If Val(ct01) - Val(ct03) - Val(ct04) - Val(ct05) - Val(ct06) - Val(ct07) - Val(ct08) - Val(ct10) <= 0 Then
                ' chi tieu 11a
                .Row = 32
                .Col = .ColLetterToNumber("E")
                .Lock = False
'                ' chi tieu 12
'                .Row = 33
'                .Col = .ColLetterToNumber("E")
'                .Lock = True
'                .Text = "0"
            Else
                ' chi tieu 11a
                .Row = 32
                .Col = .ColLetterToNumber("E")
                .Lock = True
                .Text = "0"
'                ' chi tieu 12
'                .Row = 33
'                .Col = .ColLetterToNumber("E")
'                .Lock = False
            End If
               
            ' Kiem tra chi tieu 12
            If Val(ct01) - Val(ct03) - Val(ct04) - Val(ct05) - Val(ct06) - Val(ct07) - Val(ct08) - Val(ct10) < 0 Then
'                ' chi tieu 11a
'                .Row = 32
'                .Col = .ColLetterToNumber("E")
'                .Lock = False
                ' chi tieu 12
                .Row = 33
                .Col = .ColLetterToNumber("E")
                .Lock = True
                .Text = "0"
            Else
'                ' chi tieu 11a
'                .Row = 32
'                .Col = .ColLetterToNumber("E")
'                .Lock = True
'                .Text = "0"
                ' chi tieu 12
                .Row = 33
                .Col = .ColLetterToNumber("E")
                .Lock = False
            End If
            
            ' Kiem tra ct13
            If Val(ct01) - Val(ct03) - Val(ct04) - Val(ct05) - Val(ct06) - Val(ct07) - Val(ct08) - Val(ct10) < 0 And Val(ct11a) > 0 Then
                .Row = 34
                .Col = .ColLetterToNumber("E")
                .Formula = ""
                .Text = "0"
            Else
                .Row = 34
                .Col = .ColLetterToNumber("E")
                .Formula = "E31 -E33"
            End If
        
        
        End If
        
        
         ' set lock CQT phan bo
        .Sheet = sheetPl03_9_TNDN
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
         
         
     .EventEnabled(EventAllEvents) = True
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

Dim chiTieuC1, chiTieuC2, chiTieuC3a, chiTieuC3b, chiTieuC4 As Variant

Dim ct01, ct03, ct04, ct05, ct06, ct07, ct08, ct10, ct11a As Variant

Dim intIndexCombo As Integer
Dim strMa As Variant

Dim vTuNgay As Variant
Dim vDenNgay As Variant
Dim arrTemp() As String
Dim dTuNgay As Date
Dim dDenNgay As Date
Dim soNgay As Long

With fps
    If .ActiveSheet = sheetTK03_TNDN Then
        .Sheet = sheetTK03_TNDN
        mCurrentSheet = sheetTK03_TNDN
        soNgay = 0
        If Col = .ColLetterToNumber("N") And Row = 95 Then
            .GetText .ColLetterToNumber("J"), 95, vTuNgay
            .GetText .ColLetterToNumber("N"), 95, vDenNgay
            If Trim(vDenNgay) <> "" And Trim(vDenNgay) <> "../../...." Then
                arrTemp = Split(vTuNgay, "/")
                dTuNgay = DateSerial(Int(arrTemp(2)), Int(arrTemp(1)), Int(arrTemp(0)))
                arrTemp = Split(vDenNgay, "/")
                dDenNgay = DateSerial(Int(arrTemp(2)), Int(arrTemp(1)), Int(arrTemp(0)))
                soNgay = DateDiff("D", dTuNgay, dDenNgay) + 1
            End If
            If soNgay >= 0 Then
                .Col = .ColLetterToNumber("E")
                .Row = 95
                .Text = soNgay
                UpdateCell fps, .Col, .Row, .Text
            End If
        End If
    End If

    If .ActiveSheet = sheetPl03_5_TNDN Then
        .EventEnabled(EventAllEvents) = False
        .Sheet = sheetPl03_5_TNDN
        If Col = .ColLetterToNumber("E") And (Row >= 21 And Row <= 33) Then
            .GetText .ColLetterToNumber("E"), 21, ct01
            .GetText .ColLetterToNumber("E"), 23, ct03
            .GetText .ColLetterToNumber("E"), 24, ct04
            .GetText .ColLetterToNumber("E"), 25, ct05
            .GetText .ColLetterToNumber("E"), 26, ct06
            .GetText .ColLetterToNumber("E"), 27, ct07
            .GetText .ColLetterToNumber("E"), 28, ct08
            
            .GetText .ColLetterToNumber("E"), 30, ct10
            .GetText .ColLetterToNumber("E"), 32, ct11a
            
            ' lock chi tieu 10
            If Val(ct01) - Val(ct03) - Val(ct04) - Val(ct05) - Val(ct06) - Val(ct07) - Val(ct08) < 0 Then
                .Row = 30
                .Col = .ColLetterToNumber("E")
                .Lock = True
                .Text = "0"
            Else
                .Row = 30
                .Col = .ColLetterToNumber("E")
                .Lock = False
            End If
            
            
            If Val(ct01) - Val(ct03) - Val(ct04) - Val(ct05) - Val(ct06) - Val(ct07) - Val(ct08) - Val(ct10) <= 0 Then
                ' chi tieu 11a
                .Row = 32
                .Col = .ColLetterToNumber("E")
                .Lock = False
'                ' chi tieu 12
'                .Row = 33
'                .Col = .ColLetterToNumber("E")
'                .Lock = True
'                .Text = "0"
            Else
                ' chi tieu 11a
                .Row = 32
                .Col = .ColLetterToNumber("E")
                .Lock = True
                .Text = "0"
'                ' chi tieu 12
'                .Row = 33
'                .Col = .ColLetterToNumber("E")
'                .Lock = False
            End If
               
            ' Kiem tra chi tieu 12
            If Val(ct01) - Val(ct03) - Val(ct04) - Val(ct05) - Val(ct06) - Val(ct07) - Val(ct08) - Val(ct10) < 0 Then
'                ' chi tieu 11a
'                .Row = 32
'                .Col = .ColLetterToNumber("E")
'                .Lock = False
                ' chi tieu 12
                .Row = 33
                .Col = .ColLetterToNumber("E")
                .Lock = True
                .Text = "0"
            Else
'                ' chi tieu 11a
'                .Row = 32
'                .Col = .ColLetterToNumber("E")
'                .Lock = True
'                .Text = "0"
                ' chi tieu 12
                .Row = 33
                .Col = .ColLetterToNumber("E")
                .Lock = False
            End If
            
            ' Kiem tra ct13
            If Val(ct01) - Val(ct03) - Val(ct04) - Val(ct05) - Val(ct06) - Val(ct07) - Val(ct08) - Val(ct10) < 0 And Val(ct11a) > 0 Then
                .Row = 34
                .Col = .ColLetterToNumber("E")
                .Formula = ""
                .Text = "0"
            Else
                .Row = 34
                .Col = .ColLetterToNumber("E")
                .Formula = "E31 -E33"
            End If
               
        End If
        .EventEnabled(EventAllEvents) = True
    End If

    If .ActiveSheet = sheetPl03_7_TNDN Then
        .Col = Col
        .Row = Row
        If Col = .ColLetterToNumber("N") Then 'Khi chon Combo
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("AJ")
            .typeComboboxCurSel = intIndexCombo
            .GetText .Col, .Row, strMa
            'Nhap lai Combo cot AJ: Combo AJ
            
            .Col = .ColLetterToNumber("AH")
            .SetText .Col, .Row, Left$(CStr(strMa), 3)
            UpdateCell fps, .Col, .Row, Left$(CStr(strMa), 3)
        End If
        
        .Col = Col
        .Row = Row
        If Col = .ColLetterToNumber("AB") Then 'Khi chon Combo
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("AK")
            .typeComboboxCurSel = intIndexCombo
            .GetText .Col, .Row, strMa
            'Nhap lai Combo cot AK: Combo AK
            
            .Col = .ColLetterToNumber("AI")
            .SetText .Col, .Row, Left$(CStr(strMa), 3)
            UpdateCell fps, .Col, .Row, Left$(CStr(strMa), 3)
        End If
    
    End If
    
    ' Kiem tra neu C4 <0 thi lock C5 , C7, C8
    If .ActiveSheet = sheetTK03_TNDN Then
        .Sheet = sheetTK03_TNDN
        If Col = .ColLetterToNumber("O") And (Row = 45 Or Row = 49 Or Row = 51 Or Row = 52) Then
            .GetText .ColLetterToNumber("O"), 48, chiTieuC1
            .GetText .ColLetterToNumber("O"), 49, chiTieuC2
            .GetText .ColLetterToNumber("O"), 51, chiTieuC3a
            .GetText .ColLetterToNumber("O"), 52, chiTieuC3b
            If Val(chiTieuC1) - Val(chiTieuC2) - Val(chiTieuC3a) - Val(chiTieuC3b) < 0 Then
                .Col = .ColLetterToNumber("O")
                ' chi tieu C5
                .Row = 54
                .Lock = True
                .Text = 0
                ' chi tieu C7
                .Row = 56
                .Lock = True
                .Text = 0
                ' Chi tieu C8
                .Row = 57
                .Lock = True
                .Text = 0
            Else
                .Col = .ColLetterToNumber("O")
                ' chi tieu C5
                .Row = 54
                .Lock = False
                ' chi tieu C7
                .Row = 56
                .Lock = False
                ' Chi tieu C8
                .Row = 57
                .Lock = False
            End If
        End If
     End If
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
    setChangeComb Col, Row
End Sub

'inner class event: ButtonClicked
Public Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
     TAX_Utilities_v1.AdjustData(fps.ActiveSheet - 1) = True
     'CellChange Col, Row
     
    Dim blTemp As Boolean
    Dim lResult As Integer
    Dim count As Integer
    With fps
        fps.EventEnabled(EventAllEvents) = False
        If fps.ActiveSheet = 1 Then
        ' Neu check gia han moi cho phep chon GT
            If Col = fps.ColLetterToNumber("F") And Row = 83 Then
                fps.Col = fps.ColLetterToNumber("F")
                fps.Row = 83
                If fps.Text = "1" Or UCase$(fps.Text) = "X" Then
                    fps.Col = fps.ColLetterToNumber("I")
                    fps.Row = 85
                    fps.Lock = False
                    
                    fps.Col = fps.ColLetterToNumber("I")
                    fps.Row = 87
                    fps.Lock = False
                    
                    fps.Col = fps.ColLetterToNumber("I")
                    fps.Row = 89
                    fps.Lock = False
                    
                    fps.Col = fps.ColLetterToNumber("I")
                    fps.Row = 91
                    fps.Lock = False
                    fps.Formula = "O67-I89"
                Else
                    fps.Col = fps.ColLetterToNumber("I")
                    fps.Row = 85
                    fps.Lock = True
                    
                    fps.Col = fps.ColLetterToNumber("I")
                    fps.Row = 87
                    fps.Lock = True
                    
                    fps.Col = fps.ColLetterToNumber("I")
                    fps.Row = 89
                    fps.Lock = True
                    
                    fps.Col = fps.ColLetterToNumber("I")
                    fps.Row = 91
                    fps.Lock = True
                    fps.Formula = ""
                    
                    fps.SetText fps.ColLetterToNumber("I"), 85, ""
                    fps.SetText fps.ColLetterToNumber("P"), 85, "00"
                    UpdateCell fps, fps.ColLetterToNumber("P"), 85, "00"
                    UpdateCell fps, fps.ColLetterToNumber("I"), 85, ""
                    
                    fps.SetText fps.ColLetterToNumber("I"), 87, ""
                    fps.SetText fps.ColLetterToNumber("I"), 89, ""
                    fps.SetText fps.ColLetterToNumber("I"), 91, ""
                    UpdateCell fps, fps.ColLetterToNumber("I"), 87, ""
                    UpdateCell fps, fps.ColLetterToNumber("I"), 89, ""
                    UpdateCell fps, fps.ColLetterToNumber("I"), 91, ""
                End If
            End If
        End If
        
        
        If Col = .ColLetterToNumber("G") And Row >= 42 And .ActiveSheet = sheetPl03_9_TNDN Then
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
    If .ActiveSheet = 1 Then
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("I") And Row = 85 Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("P")
            .typeComboboxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
        End If
    ElseIf .ActiveSheet = sheetPl03_7_TNDN Then
            .Sheet = sheetPl03_7_TNDN
            mCurrentSheet = sheetPl03_7_TNDN
            
            .Col = Col
            .Row = Row
            If Col = .ColLetterToNumber("N") Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot AJ
                intIndexCombo = .typeComboboxCurSel
                .Col = .ColLetterToNumber("AJ")
                .Row = Row
                .typeComboboxCurSel = intIndexCombo
                'Nhap lai Combo cot AJ
                .Col = .ColLetterToNumber("AJ")
                UpdateCell fps, .Col, .Row, .Text
                
            End If
            
            If Col = .ColLetterToNumber("AB") Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot AK
                intIndexCombo = .typeComboboxCurSel
                .Col = .ColLetterToNumber("AK")
                .Row = Row
                .typeComboboxCurSel = intIndexCombo
                'Nhap lai Combo cot AK
                .Col = .ColLetterToNumber("AK")
                UpdateCell fps, .Col, .Row, .Text
                
            End If
    End If
End With
    
    CellChange Col, Row
    TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
End Sub


Private Sub fps_EditChange(ByVal Col As Long, ByVal Row As Long)
   'MsgBox "Col: " & Col & " Row: " & Row & " Sheet: " & fps.ActiveSheet
End Sub

Private Sub fps_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim i As Integer, iCol As Long, iRow As Long
'    If KeyCode = vbKeyClear Or KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
'        With fps
'            If .ActiveCol = .ColLetterToNumber(S1TxtC8Col) And .ActiveRow = S1TxtC8Row Then
'                XoaDauAm .ActiveCol, .ActiveRow
'            End If
'        End With
'    End If
' ' for KHBS only
'    If KeyCode = vbKeyF6 Then
'     With fps
'            iCol = .ActiveCol
'            iRow = .ActiveRow
'            mblKeyF6 = True
'            fps_Change iCol, iRow
'            mblKeyF6 = False
'     End With
'    End If
    
    
End Sub

'inner class event: Leavecell
Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim countRowDyn As Long
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("I") And Row = 87 Then
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
            
            
            If Col = .ColLetterToNumber("J") And Row = 95 Then
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
            
            
            If Col = .ColLetterToNumber("N") And Row = 95 Then
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
            
            'kiem tra truong ngay ky tren TK
            
            .Sheet = 1
            .Row = 103
            .Col = .ColLetterToNumber("B")
            countRowDyn = 0
            Do
                countRowDyn = countRowDyn + 1
                .Row = .Row + 1
            Loop Until .Text = "aa"
            
            If Col = .ColLetterToNumber("N") And Row = 110 + countRowDyn Then
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
    
    Dim chiTieuC4 As Variant
    
            
With fps
    'Because of unclearly-rised events, this function is here for turning it off
    'at the end of this sub, this event will be turned on
    .EventEnabled(EventAllEvents) = False
    .Sheet = .ActiveSheet
    mCurrentSheet = .ActiveSheet
    Select Case .Sheet
        Case sheetTK03_TNDN ' To khai TNDN
            mCurrentSheet = .Sheet
            GetCellSpan fps, Col, Row
        Case sheetPl03_1A_TNDN ' Phu luc 03-1A
            mCurrentSheet = .Sheet
            If Row = 41 And Col = .ColLetterToNumber("E") Then
                .Sheet = 1
                ' tong hop chi tieu 19 len chi tieu A1 cua to khai
                .Col = .ColLetterToNumber("O")
                .Row = 31
                .Formula = "P31"
            End If
        Case sheetPl03_1B_TNDN ' Phu luc 03-1B
            mCurrentSheet = .Sheet
            If Row = 38 And Col = .ColLetterToNumber("E") Then
                .Sheet = 1
                ' tong hop chi tieu 16 len chi tieu A1 cua to khai
                .Col = .ColLetterToNumber("O")
                .Row = 31
                .Formula = "P31"
            End If
            
        Case sheetPl03_1C_TNDN ' Phu luc 03-1C
            mCurrentSheet = .Sheet
            If Row = 50 And Col = .ColLetterToNumber("E") Then
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 31
                .Formula = "P31"
            End If
            
        Case sheetPl03_2A_TNDN ' Phu luc 03-2A
            mCurrentSheet = .Sheet
            If Row = 32 And Col = .ColLetterToNumber("M") Then
                .Sheet = 1
                ' tong hop lo cot (5) len chi tieu C3
'                .Col = .ColLetterToNumber("O")
'                .Row = 50
'                .Formula = "P50"
                
                ' tong hop lo cot (5) len chi tieu C3a
                .Col = .ColLetterToNumber("O")
                .Row = 51
                .Formula = "P51"
                
                
                
                .GetText .ColLetterToNumber("O"), 53, chiTieuC4
                If Val(chiTieuC4) < 0 Then
                    .Col = .ColLetterToNumber("O")
                    ' chi tieu C5
                    .Row = 54
                    .Lock = True
                    .Text = 0
                    ' chi tieu C7
                    .Row = 56
                    .Lock = True
                    .Text = 0
                    ' Chi tieu C8
                    .Row = 57
                    .Lock = True
                    .Text = 0
                Else
                    .Col = .ColLetterToNumber("O")
                    ' chi tieu C5
                    .Row = 54
                    .Lock = False
                    ' chi tieu C7
                    .Row = 56
                    .Lock = False
                    ' Chi tieu C8
                    .Row = 57
                    .Lock = False
                End If
            End If
        Case sheetPl03_3A_TNDN ' Phu luc 03-3A
            mCurrentSheet = .Sheet
            If IsValidatedCell(sheetPl03_3A_TNDN, 38, .ColLetterToNumber("N"), Row, Col) Then
                 ValidateNumberFormat fps, sheetPl03_3A_TNDN, Row, Col
            End If
            If IsValidatedCell(sheetPl03_3A_TNDN, 39, .ColLetterToNumber("J"), Row, Col) Then
                 ValidateNumberFormat fps, sheetPl03_3A_TNDN, Row, Col
            End If
            If IsValidatedCell(sheetPl03_3A_TNDN, 40, .ColLetterToNumber("N"), Row, Col) Then
                 ValidateNumberFormat fps, sheetPl03_3A_TNDN, Row, Col
            End If
            If IsValidatedCell(sheetPl03_3A_TNDN, 38, .ColLetterToNumber("U"), Row, Col) Then
                 ValidateDate fps, sheetPl03_3A_TNDN, Row, Col, yyyy
            End If
            If IsValidatedCell(sheetPl03_3A_TNDN, 39, .ColLetterToNumber("Q"), Row, Col) Then
                 ValidateDate fps, sheetPl03_3A_TNDN, Row, Col, yyyy
            End If
            If IsValidatedCell(sheetPl03_3A_TNDN, 40, .ColLetterToNumber("U"), Row, Col) Then
                 ValidateDate fps, sheetPl03_3A_TNDN, Row, Col, yyyy
            End If
            ' TOng hop len chi tieu C11
            pL_3A_U50 = Sum_Cell(sheetPl03_3A_TNDN, 50, .ColLetterToNumber("U")) 'PL 03-3A
            pL_3B_U60 = Sum_Cell(sheetPl03_3B_TNDN, 60, .ColLetterToNumber("U")) 'PL 03-3B
            
            strValue = getArrRowCol(sheetPl03_3A_TNDN, 50, .ColLetterToNumber("U"))
            If InStr(1, strValue, "###") > 0 Then
                arrRowCols = Split(strValue, "###")
                If UBound(arrRowCols) > 0 Then
                    For idx = 0 To UBound(arrRowCols)
                        arrRowCol = Split(arrRowCols(idx), "_")
                        If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                            ' chi tieu C11
                            .Sheet = 1
                            .SetText .ColLetterToNumber("P"), 61, pL_3A_U50 + pL_3B_U60
                            .Col = .ColLetterToNumber("O")
                            .Row = 61
                            .Formula = "P61"
                            Exit For
                        End If
                    Next idx
                End If
            Else
                arrRowCol = Split(strValue, "_")
                If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                    ' chi tieu C11
                    .Sheet = 1
                    .SetText .ColLetterToNumber("P"), 61, pL_3A_U50 + pL_3B_U60
                    .Col = .ColLetterToNumber("O")
                    .Row = 61
                    .Formula = "P61"
                End If
            End If
            
            ' TOng hop len chi tieu C12
            pL_3A_U56 = Sum_Cell(sheetPl03_3A_TNDN, 56, .ColLetterToNumber("U")) 'PL 03-3A
            pL_3B_U65 = Sum_Cell(sheetPl03_3B_TNDN, 65, .ColLetterToNumber("U")) 'PL 03-3B
            pL_3C_X43 = Sum_Cell(sheetPl03_3C_TNDN, 43, .ColLetterToNumber("X")) 'PL 03-3C
            
            strValue = getArrRowCol(sheetPl03_3A_TNDN, 56, .ColLetterToNumber("U"))
            If InStr(1, strValue, "###") > 0 Then
                arrRowCols = Split(strValue, "###")
                If UBound(arrRowCols) > 0 Then
                    For idx = 0 To UBound(arrRowCols)
                        arrRowCol = Split(arrRowCols(idx), "_")
                        If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                            ' chi tieu C12
                            .Sheet = 1
                            .SetText .ColLetterToNumber("P"), 62, pL_3A_U56 + pL_3B_U65 + pL_3C_X43
                            .Col = .ColLetterToNumber("O")
                            .Row = 62
                            .Formula = "P62"
                            Exit For
                        End If
                    Next idx
                End If
             Else
                arrRowCol = Split(strValue, "_")
                If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                    ' chi tieu C12
                    .Sheet = 1
                    .SetText .ColLetterToNumber("P"), 62, pL_3A_U56 + pL_3B_U65 + pL_3C_X43
                    .Col = .ColLetterToNumber("O")
                    .Row = 62
                    .Formula = "P62"
                End If
            End If
            
        
        Case sheetPl03_3B_TNDN ' Phu luc 03-3B
            mCurrentSheet = .Sheet
            If IsValidatedCell(sheetPl03_3B_TNDN, 44, .ColLetterToNumber("N"), Row, Col) Then
                 ValidateNumberFormat fps, sheetPl03_3B_TNDN, Row, Col
            End If
            If IsValidatedCell(sheetPl03_3B_TNDN, 45, .ColLetterToNumber("J"), Row, Col) Then
                 ValidateNumberFormat fps, sheetPl03_3B_TNDN, Row, Col
            End If
            If IsValidatedCell(sheetPl03_3B_TNDN, 46, .ColLetterToNumber("N"), Row, Col) Then
                 ValidateNumberFormat fps, sheetPl03_3B_TNDN, Row, Col
            End If
            If IsValidatedCell(sheetPl03_3B_TNDN, 44, .ColLetterToNumber("U"), Row, Col) Then
                 ValidateDate fps, sheetPl03_3B_TNDN, Row, Col, yyyy
            End If
            If IsValidatedCell(sheetPl03_3B_TNDN, 45, .ColLetterToNumber("Q"), Row, Col) Then
                 ValidateDate fps, sheetPl03_3B_TNDN, Row, Col, yyyy
            End If
            If IsValidatedCell(sheetPl03_3B_TNDN, 46, .ColLetterToNumber("U"), Row, Col) Then
                 ValidateDate fps, sheetPl03_3B_TNDN, Row, Col, yyyy
            End If
            If IsValidatedCell(sheetPl03_3B_TNDN, 41, .ColLetterToNumber("S"), Row, Col) Then
                 ValidateDate fps, sheetPl03_3B_TNDN, Row, Col, DDMMYYYY
            End If
        
            ' Tong hop len chi tieu C11
            pL_3A_U50 = Sum_Cell(sheetPl03_3A_TNDN, 50, .ColLetterToNumber("U")) 'PL 03-3A
            pL_3B_U60 = Sum_Cell(sheetPl03_3B_TNDN, 60, .ColLetterToNumber("U")) 'PL 03-3B
            
            strValue = getArrRowCol(sheetPl03_3B_TNDN, 60, .ColLetterToNumber("U"))
            If InStr(1, strValue, "###") > 0 Then
                arrRowCols = Split(strValue, "###")
                If UBound(arrRowCols) > 0 Then
                    For idx = 0 To UBound(arrRowCols)
                        arrRowCol = Split(arrRowCols(idx), "_")
                        If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                            ' chi tieu C11
                            .Sheet = 1
                            .SetText .ColLetterToNumber("P"), 61, pL_3A_U50 + pL_3B_U60
                            .Col = .ColLetterToNumber("O")
                            .Row = 61
                            .Formula = "P61"
                            Exit For
                        End If
                    Next idx
                End If
            Else
                arrRowCol = Split(strValue, "_")
                If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                    ' chi tieu C11
                    .Sheet = 1
                    .SetText .ColLetterToNumber("P"), 61, pL_3A_U50 + pL_3B_U60
                    .Col = .ColLetterToNumber("O")
                    .Row = 61
                    .Formula = "P61"
                End If
            End If
            
            ' TOng hop len chi tieu C12
            pL_3A_U56 = Sum_Cell(sheetPl03_3A_TNDN, 56, .ColLetterToNumber("U")) 'PL 03-3A
            pL_3B_U65 = Sum_Cell(sheetPl03_3B_TNDN, 65, .ColLetterToNumber("U")) 'PL 03-3B
            pL_3C_X43 = Sum_Cell(sheetPl03_3C_TNDN, 43, .ColLetterToNumber("X")) 'PL 03-3C
            
            strValue = getArrRowCol(sheetPl03_3B_TNDN, 65, .ColLetterToNumber("U"))
            If InStr(1, strValue, "###") > 0 Then
                arrRowCols = Split(strValue, "###")
                If UBound(arrRowCols) > 0 Then
                    For idx = 0 To UBound(arrRowCols)
                        arrRowCol = Split(arrRowCols(idx), "_")
                        If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                            ' chi tieu C12
                            .Sheet = 1
                            .SetText .ColLetterToNumber("P"), 62, pL_3A_U56 + pL_3B_U65 + pL_3C_X43
                            .Col = .ColLetterToNumber("O")
                            .Row = 62
                            .Formula = "P62"
                            Exit For
                        End If
                    Next idx
                End If
             Else
                arrRowCol = Split(strValue, "_")
                If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                    ' chi tieu C12
                    .Sheet = 1
                    .SetText .ColLetterToNumber("P"), 62, pL_3A_U56 + pL_3B_U65 + pL_3C_X43
                    .Col = .ColLetterToNumber("O")
                    .Row = 62
                    .Formula = "P62"
                End If
            End If
        
        Case sheetPl03_3C_TNDN ' Phu luc 03-3C
            mCurrentSheet = .Sheet
            If IsValidatedCell(sheetPl03_3C_TNDN, 27, .ColLetterToNumber("V"), Row, Col) Then
                 ValidateDate fps, sheetPl03_3C_TNDN, Row, Col, DDMMYYYY
            End If
            If IsValidatedCell(sheetPl03_3C_TNDN, 31, .ColLetterToNumber("Y"), Row, Col) Then
                 ValidateDate fps, sheetPl03_3C_TNDN, Row, Col, DDMMYYYY
            End If
            
            ' TOng hop len chi tieu C12
            pL_3A_U56 = Sum_Cell(sheetPl03_3A_TNDN, 56, .ColLetterToNumber("U")) 'PL 03-3A
            pL_3B_U65 = Sum_Cell(sheetPl03_3B_TNDN, 65, .ColLetterToNumber("U")) 'PL 03-3B
            pL_3C_X43 = Sum_Cell(sheetPl03_3C_TNDN, 43, .ColLetterToNumber("X")) 'PL 03-3C
            
            strValue = getArrRowCol(sheetPl03_3C_TNDN, 43, .ColLetterToNumber("X"))
            If InStr(1, strValue, "###") > 0 Then
                arrRowCols = Split(strValue, "###")
                If UBound(arrRowCols) > 0 Then
                    For idx = 0 To UBound(arrRowCols)
                        arrRowCol = Split(arrRowCols(idx), "_")
                        If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                            ' chi tieu C12
                            .Sheet = 1
                            .SetText .ColLetterToNumber("P"), 62, pL_3A_U56 + pL_3B_U65 + pL_3C_X43
                            .Col = .ColLetterToNumber("O")
                            .Row = 62
                            .Formula = "P62"
                            Exit For
                        End If
                    Next idx
                End If
            Else
                arrRowCol = Split(strValue, "_")
                If Row = arrRowCol(1) And Col = arrRowCol(0) Then
                    ' chi tieu C12
                    .Sheet = 1
                    .SetText .ColLetterToNumber("P"), 62, pL_3A_U56 + pL_3B_U65 + pL_3C_X43
                    .Col = .ColLetterToNumber("O")
                    .Row = 62
                    .Formula = "P62"
                End If
            End If
        Case sheetPl03_4_TNDN ' Phu luc 03-6
            mCurrentSheet = .Sheet
            Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl03_4_TNDN - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("E"), 26))
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode.lastChild
            Set xmlCellNode = xmlCellNode.lastChild
            Set xmlCellNode = xmlCellNode.lastChild
            ' Lay ve col vaf row cua dong tong cong
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), cCol, cRow
            If Row = cRow And Col = cCol Then
                ' chi tieu C15
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 65
                .Formula = "P65"
            End If
            
            If Row = cRow And Col = .ColLetterToNumber("Z") Then
                ' chi tieu C15
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 37
                .Formula = "P37"
            End If
            
        Case sheetPl03_6_TNDN ' Phu luc 03-6
            mCurrentSheet = .Sheet
            Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl03_6_TNDN - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("E"), 26))
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode.lastChild
            Set xmlCellNode = xmlCellNode.lastChild
            Set xmlCellNode = xmlCellNode.lastChild
            ' Lay ve col vaf row cua dong tong cong
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), cCol, cRow
            If Row = cRow And Col = cCol Then
                ' chi tieu C5
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 54
                .Formula = "P54"
            End If
        Case sheetPl03_5_TNDN ' Phu luc 03-5
            mCurrentSheet = .Sheet
             If Row = 36 And Col = .ColLetterToNumber("E") Then
                ' chi tieu D2
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 69
                .Formula = "P69"
            End If
            
'            If Row = 32 And Col = .ColLetterToNumber("E") Then
'                ' chi tieu C3b
'                .Sheet = 1
'                .Col = .ColLetterToNumber("O")
'                .Row = 52
'                .Formula = "P52"
'            End If
    End Select
    .EventEnabled(EventAllEvents) = True
End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    Dim i As Integer
    
    Dim col3 As Long, col4 As Long, col5 As Long
    Dim col6 As Long, Col7 As Long, col8 As Long
    Dim col9 As Long, Col10 As Long, Col11 As Long
    
    Dim ColLoi As Long, RowLoi As Long
    
    Dim vErrMsg1 As Variant
    
    Dim ValCol3 As Variant, ValCol4 As Variant, ValCol5 As Variant
    Dim ValCol6 As Variant, ValCol7 As Variant, ValCol8 As Variant
    Dim ValCol9 As Variant, ValCol10 As Variant, ValCol11 As Variant
    
    Dim XacDinhLoi As Boolean
    Dim blCheck_S03_1A As Boolean
    Dim blCheck_S03_3C As Boolean, blCheck_S03_4 As Boolean
    Dim errEmty As Variant, iFlagTongPhanBo As Variant, errTongTLPB As Variant, vMST As Variant
    Dim errMSTDN As Variant, errMst10 As Variant, errMST As Variant

    Dim iFlagEmpty As Integer, iFlagCheckMaST As Integer

    ' PL 05
    Dim mErrMSTTrung As Variant
    
    ' PL03-6
    Dim blCheck_S12 As Boolean
    Dim iWhiteRow As Integer, iFlagMST As Integer

    ' end
    ' PL 03-3A
    Dim vErrCtieuC4 As Variant
    
    ' check gh
    Dim vCheckGH As Variant, vLyDoGH As Variant, vErrorCheckGH As Variant, vThoiHanGH As Variant, vErrorCheckThoiHanGH As Variant, vSoThueGH As Variant, vErrorCheckSoThueGH As Variant
    Dim iflagCheckGH As Integer, iflagCheckThoiHanGH As Integer, iflagCheckSoThueGH As Integer
    Dim arrDateGH() As String
    Dim arrDateHanNop() As String
    Dim vErrorCheckHN As Variant
    Dim dThoiHanGH As Date
    Dim dHanNop As Date
    ' end
    Dim vDenNgayNC As Variant
    Dim vTuNgayNC As Variant
    Dim vErrDenNgayNC As Variant
    Dim vErrDenNgayNC1 As Variant
    Dim iflagDenNgayNC As Integer
    
    
    ' check TK
    Dim pL_3A_U50, pL_3B_U60 As Double
    Dim pL_3A_U56, pL_3B_U65, pL_3C_X43 As Double
    Dim blCheck_S3A, blCheck_S3B, blCheck_S3C As Boolean
    
    Dim vCheckLienKet, vCheckCSHachToan As Variant
    Dim ctA1 As Variant
    Dim errCheckCtA1 As Variant
    Dim iFlagCheckCtA1 As Integer
    Dim errCheckLienKet, errCheckCSHachToan As Variant
    Dim iFlagCheckLienKet As Integer, iFlagCheckCSHachToan As Integer
    
    'chek PL 03-7
    Dim vCheckMST, vDiaChi, vTenLK As Variant
    Dim vCheck(13) As Variant
    Dim strCol As String
    Dim arrCol() As String
    Dim isCheck As Boolean
    Dim iFlagCheck As Integer
    Dim errHinhThucLK As Variant
    
    Dim vCot5, vCot9 As Variant
    Dim errCot5, errCot9 As Variant
    Dim iFlagCot5 As Integer, iFlagCot9 As Integer
    Dim strRowGroup As String
    Dim strFinish As String
    Dim arrRowGroup() As String
    Dim arrFinish() As String
    
    ' PL 03-8
    Dim vTyLe As Variant
    Dim vTenDn As Variant, vCQTCapCuc As Variant, vCqtChiCuc As Variant, vSoThueQ1 As Variant, vSoThueQ2 As Variant, vSoThueQ3 As Variant, vSoThueQ4 As Variant, vSoThue As Variant
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer
    Dim errTongSoThue As Variant
    Dim iFlagTyLe  As Integer
    Dim iFlagCheckSoThue As Integer
    Dim vTongTyLe As Double
    Dim vTongSoThue As Double
    Dim vCt5 As Variant
    Dim vColor As Variant
    Dim strCheck As String
    
    Dim ctB14 As Variant
    

    iFlagCheckLienKet = 0
    iFlagCheckCSHachToan = 0
    
    CheckErrorMST
    With fps
        .EventEnabled(EventAllEvents) = False
        ' check validate TK
        .Sheet = .SheetCount
        ' xy ly to khai
        .GetText .ColLetterToNumber("E"), 47, errCheckCSHachToan
        .GetText .ColLetterToNumber("E"), 48, errCheckLienKet
        
        .GetText .ColLetterToNumber("E"), 46, vErrorCheckGH
        .GetText .ColLetterToNumber("E"), 53, vErrorCheckThoiHanGH
        .GetText .ColLetterToNumber("E"), 52, vErrorCheckHN
        .GetText .ColLetterToNumber("E"), 54, vErrorCheckSoThueGH
        
        .GetText .ColLetterToNumber("E"), 91, vErrDenNgayNC
        .GetText .ColLetterToNumber("E"), 92, vErrDenNgayNC1
        
        .GetText .ColLetterToNumber("E"), 97, errCheckCtA1
        
        ' end
        
        .Sheet = sheetTK03_TNDN
        .GetText .ColLetterToNumber("D"), 10, vCheckCSHachToan
        .GetText .ColLetterToNumber("D"), 12, vCheckLienKet
        
        .GetText .ColLetterToNumber("O"), 46, ctB14
        .GetText .ColLetterToNumber("O"), 31, ctA1
        
        
        If vCheckCSHachToan = "1" Or UCase$(vCheckCSHachToan) = "X" Then
            If TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_8_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0" Or Strloaitk = "TKBS" Then
            Else
                .Sheet = 1
                .ActiveSheet = 1
                If Not .Visible Then
                    .Visible = True
                    DisplayMessage "0337", msOKOnly, miWarning
                    .Visible = False
                Else
                    DisplayMessage "0337", msOKOnly, miWarning
                End If
                iFlagCheckCSHachToan = 1
            End If
        End If
        
        If vCheckLienKet = "1" Or UCase$(vCheckLienKet) = "X" Then
            If TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_7_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0" Or Strloaitk = "TKBS" Then
            Else
                .Sheet = 1
                .ActiveSheet = 1
                If Not .Visible Then
                    .Visible = True
                    DisplayMessage "0338", msOKOnly, miWarning
                    .Visible = False
                Else
                    DisplayMessage "0338", msOKOnly, miWarning
                End If
                iFlagCheckLienKet = 1
            End If
        End If
        
        If Val(ctA1) <> 0 Then
            If TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_1A_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0" _
            Or TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_1B_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0" _
            Or TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_1C_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0" Or Strloaitk = "TKBS" Then
            Else
                .Sheet = 1
                .ActiveSheet = 1
                If Not .Visible Then
                    .Visible = True
                    DisplayMessage "0347", msOKOnly, miWarning
                    .Visible = False
                Else
                    DisplayMessage "0347", msOKOnly, miWarning
                End If
                iFlagCheckCtA1 = 1
            End If
        End If
        
        
        blCheck_S3A = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_3A_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        blCheck_S3B = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_3B_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        blCheck_S3C = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_3C_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        ' Lay gia tri cua chi tieu dong trong PL 03-3A, 03-3B, 03-3C
        pL_3A_U50 = Sum_Cell(sheetPl03_3A_TNDN, 50, .ColLetterToNumber("U")) 'PL 03-3A
        pL_3B_U60 = Sum_Cell(sheetPl03_3B_TNDN, 60, .ColLetterToNumber("U")) 'PL 03-3B
'        ' Lay gia tri chi tieu C8 tren to khai quyet toan
        mCurrentSheet = 1
        .Sheet = mCurrentSheet
        If blCheck_S3A = True Or blCheck_S3B = True Then
            Dim vC59 As Variant
            .GetText .ColLetterToNumber("O"), 61, vC59
            If (vC59 = pL_3A_U50 + pL_3B_U60) Then
                ' Update gia tri dung (khong loi) trong sheet header
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 23, "1"
                ' Chuyen chi tieu C11 ve trang thai dung
                mCurrentSheet = 1
                ' chi tieu C11
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 61
                .CellNote = ""
                .BackColor = mNonErrorColor
            Else
                ' Update gia tri sai (co loi) trong sheet header
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 23, "0"
                ' Lay thong bao loi
                .GetText .ColLetterToNumber("E"), 23, vErrMsg1
                ' Chuyen chi tieu ve trang thai loi
                mCurrentSheet = 1
                .Sheet = 1
                ' chi tieu C11
                .Col = .ColLetterToNumber("O")
                .Row = 61
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsg1 'static
                .BackColor = mErrorColor
            End If
        End If
'        ' Lay gia tri dong o cac Phu Luc tuong ung
        pL_3A_U56 = Sum_Cell(sheetPl03_3A_TNDN, 56, .ColLetterToNumber("U"))
        pL_3B_U65 = Sum_Cell(sheetPl03_3B_TNDN, 65, .ColLetterToNumber("U"))
        pL_3C_X43 = Sum_Cell(sheetPl03_3C_TNDN, 43, .ColLetterToNumber("X"))
'        ' Lay gia tri cua chi tieu C8 tren to khai
        mCurrentSheet = 1
        .Sheet = mCurrentSheet
        If blCheck_S3A = True Or blCheck_S3B = True Or blCheck_S3C = True Then
            Dim vC60 As Variant
            .GetText .ColLetterToNumber("O"), 62, vC60
            If (vC60 = pL_3A_U56 + pL_3B_U65 + pL_3C_X43) Then
                ' Update gia tri dung trong sheet header
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 24, "1"
                ' Chuyen trang thai dung cho C12
                mCurrentSheet = 1
                ' chi tieu C12
                .Sheet = 1
                .Col = .ColLetterToNumber("O")
                .Row = 62
                .CellNote = ""
                .BackColor = mNonErrorColor
            Else
                ' Update gia tri sai trong sheet header
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 24, "0"
                ' Lay thong bao loi
                Dim vErrMsg2 As Variant
                .GetText .ColLetterToNumber("E"), 24, vErrMsg2
                ' Chuyen C12 ve trang thai bao loi
                mCurrentSheet = 1
                .Sheet = 1
                ' chi tieu C12
                .Col = .ColLetterToNumber("O")
                .Row = 62
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrMsg2 'static
                .BackColor = mErrorColor
            End If
        End If
        
        ' check gia han
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 83, vCheckGH
        .GetText .ColLetterToNumber("I"), 85, vLyDoGH
        .GetText .ColLetterToNumber("I"), 87, vThoiHanGH
        .GetText .ColLetterToNumber("I"), 89, vSoThueGH
        
         If (Trim$(vCheckGH) = "1" Or Trim$(UCase$(vCheckGH)) = "X") Then
             ' Kiem tra ly do gia han
             If Trim$(vLyDoGH) = "" Then
                 .Col = .ColLetterToNumber("I")
                 .Row = 85
                 .CellNote = vErrorCheckGH  'static
                 .BackColor = mErrorColor
                 iflagCheckGH = 1
             Else
                 .Col = .ColLetterToNumber("I")
                 .Row = 85
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckGH = 0
             End If
             ' Kiem tra thoi han gia han
             If Trim$(vThoiHanGH) = "" Or Trim$(vThoiHanGH) = "../../...." Then
                 .Col = .ColLetterToNumber("I")
                 .Row = 87
                 .CellNote = vErrorCheckThoiHanGH  'static
                 .BackColor = mErrorColor
                 iflagCheckThoiHanGH = 1
             Else
                 ' kiem tra han nop
                 arrDateGH = Split(vThoiHanGH, "/")
                 arrDateHanNop = Split(hanNopTKTNDN, "/")
                dThoiHanGH = DateSerial(CInt(arrDateGH(2)), CInt(arrDateGH(1)), CInt(arrDateGH(0)))
                dHanNop = DateSerial(CInt(arrDateHanNop(2)), CInt(arrDateHanNop(1)), CInt(arrDateHanNop(0)))
                If DateDiff("D", dThoiHanGH, dHanNop) > 0 Then
                    
                    .Col = .ColLetterToNumber("I")
                    .Row = 87
                    .CellNote = vErrorCheckHN  'static
                    .BackColor = mAlertColor
                Else
                    .Col = .ColLetterToNumber("I")
                    .Row = 87
                    .CellNote = ""  'static
                    .BackColor = mNonErrorColor
                 End If
             End If
             ' Kiem tra so thue
             If Val(vSoThueGH) <= 0 Then
                 .Col = .ColLetterToNumber("I")
                 .Row = 89
                 .CellNote = vErrorCheckSoThueGH
                 .BackColor = mErrorColor
                 iflagCheckSoThueGH = 1
             Else
                 .Col = .ColLetterToNumber("I")
                 .Row = 89
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckSoThueGH = 0
             End If
             
         Else
                ' Ly do gia han
                 .Col = .ColLetterToNumber("I")
                 .Row = 85
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckGH = 0
                 ' Thoi han gia han
                 .Col = .ColLetterToNumber("I")
                 .Row = 87
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckThoiHanGH = 0
                 ' So thue gia han
                 .Col = .ColLetterToNumber("I")
                 .Row = 89
                 .CellNote = ""  'static
                 .BackColor = mNonErrorColor
                 iflagCheckSoThueGH = 0
         End If
         
         ' Kiem tra den ngay nop cham
        .GetText .ColLetterToNumber("N"), 95, vDenNgayNC
        .GetText .ColLetterToNumber("J"), 95, vTuNgayNC
        If Trim(vDenNgayNC) <> "" And Trim(vDenNgayNC) <> "../../...." Then
                arrDateGH = Split(vDenNgayNC, "/")
                arrDateHanNop = Split(vTuNgayNC, "/")
                dThoiHanGH = DateSerial(CInt(arrDateGH(2)), CInt(arrDateGH(1)), CInt(arrDateGH(0)))
                dHanNop = DateSerial(CInt(arrDateHanNop(2)), CInt(arrDateHanNop(1)), CInt(arrDateHanNop(0)))
                If DateDiff("D", dThoiHanGH, dHanNop) > 0 Then
                    .Col = .ColLetterToNumber("N")
                    .Row = 95
                    .CellNote = vErrDenNgayNC1  'static
                    .BackColor = mErrorColor
                    iflagDenNgayNC = 1
                Else
                    arrDateGH = Split(vDenNgayNC, "/")
                    arrDateHanNop = Split(getDenNgayNC, "/")
                    dThoiHanGH = DateSerial(CInt(arrDateGH(2)), CInt(arrDateGH(1)), CInt(arrDateGH(0)))
                    dHanNop = DateSerial(CInt(arrDateHanNop(2)), CInt(arrDateHanNop(1)), CInt(arrDateHanNop(0)))
                    If DateDiff("D", dHanNop, dThoiHanGH) > 0 Then
                        .Col = .ColLetterToNumber("N")
                        .Row = 95
                        .CellNote = vErrDenNgayNC  'static
                        .BackColor = mAlertColor
                    Else
                        .Col = .ColLetterToNumber("N")
                        .Row = 95
                        .CellNote = ""  'static
                        .BackColor = mNonErrorColor
                    End If
                End If
        Else
            .Col = .ColLetterToNumber("N")
            .Row = 95
            .CellNote = ""  'static
            .BackColor = mNonErrorColor
        End If
         
         
         .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 46, IIf(iflagCheckGH = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 53, IIf(iflagCheckThoiHanGH = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 54, IIf(iflagCheckSoThueGH = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 47, IIf(iFlagCheckCSHachToan = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 48, IIf(iFlagCheckLienKet = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 92, IIf(iflagDenNgayNC = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 97, IIf(iFlagCheckCtA1 = 1, "0", "1")
        
                
        ' Phu Luc 03-1A
        .Sheet = 2
        blCheck_S03_1A = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
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
        
        ' Kiem tra validate cua PL 03-3C
        .Sheet = sheetPl03_3C_TNDN
        blCheck_S03_3C = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S03_3C = True Then
            checkPL03_3C
        End If
                
        ' Check validate PL 03-4
        mCurrentSheet = sheetPl03_4_TNDN
        .Sheet = sheetPl03_4_TNDN
        blCheck_S03_4 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
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
            
            ' Check validate sheet 03-7
            .Sheet = .SheetCount
            .GetText .ColLetterToNumber("E"), 33, errMST
            .GetText .ColLetterToNumber("E"), 38, errHinhThucLK
            
            .GetText .ColLetterToNumber("E"), 69, errCot5
            .GetText .ColLetterToNumber("E"), 83, errCot9
            .Sheet = sheetPl03_7_TNDN
            blCheck_S12 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
            If blCheck_S12 = True Then
                .Row = 31
                .Col = .ColLetterToNumber("B")
                strCol = "U~V~W~X~Y~Z~AA~AB~AC~AD~AE~AF~AG"
                arrCol = Split(strCol, "~")
                Do
                    iWhiteRow = 1
                    '.GetText .ColLetterToNumber("AH"), .Row, vCheckMST
                    'If Trim$(CStr(vCheckMST)) = "GB" Then
                       .GetText .ColLetterToNumber("D"), .Row, vTenLK
                       .GetText .ColLetterToNumber("I"), .Row, vDiaChi
                       .GetText .ColLetterToNumber("O"), .Row, vMST
                        
                        
                        ' Clear cell note
                        .Col = .ColLetterToNumber("O")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        
                        isCheck = False
                        For i = 0 To UBound(arrCol)
                            .GetText .ColLetterToNumber(arrCol(i)), .Row, vCheck(i)
                            
                            .Col = .ColLetterToNumber(arrCol(i))
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
                        
                        ' Kiem tra rang buoc check
                        If isCheck = False And iWhiteRow = 0 Then
                            For i = 0 To UBound(arrCol)
                                .Col = .ColLetterToNumber(arrCol(i))
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = errHinhThucLK
                                .BackColor = mErrorColor
                            Next
                            iFlagCheck = 1
                        Else
                            For i = 0 To UBound(arrCol)
                                .Col = .ColLetterToNumber(arrCol(i))
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            Next
                        End If
                
                    'End If
                    .Row = .Row + 1
                    .Col = .ColLetterToNumber("B")
                Loop Until .Text = "kk"
                
                
                strRowGroup = "19~3~4~3~3~4~3~3"
                strFinish = "aa~bb~cc~dd~ee~ff~gg~hh"
                arrRowGroup = Split(strRowGroup, "~")
                arrFinish = Split(strFinish, "~")
                
                iFlagCot5 = 0
                iFlagCot9 = 0
                
                
                For i = 0 To UBound(arrRowGroup)
                    .Col = .ColLetterToNumber("B")
                     .Row = .Row + Val(arrRowGroup(i))
                     Do
                         .GetText .ColLetterToNumber("M"), .Row, vCot5
                         .GetText .ColLetterToNumber("Y"), .Row, vCot9
                        If Val(vCot5) < 0 Then
                            .Col = .ColLetterToNumber("M")
                            .CellNote = errCot5
                            .BackColor = mErrorColor
                            iFlagCot5 = 1
                        Else
                            .Col = .ColLetterToNumber("M")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                      
                        If Val(vCot9) < 0 Then
                            .Col = .ColLetterToNumber("M")
                            .CellNote = errCot5
                            .BackColor = mErrorColor
                            iFlagCot5 = 1
                        Else
                            .Col = .ColLetterToNumber("M")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                      
                        .Col = .ColLetterToNumber("B")
                        .Row = .Row + 1
                     Loop Until .Text = arrFinish(i)
                 Next
                
                
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 33, IIf(iFlagMST = 1, "0", "1")
                .SetText .ColLetterToNumber("B"), 38, IIf(iFlagCheck = 1, "0", "1")
                
                .SetText .ColLetterToNumber("B"), 69, IIf(iFlagCot5 = 1, "0", "1")
                .SetText .ColLetterToNumber("B"), 83, IIf(iFlagCot9 = 1, "0", "1")
            End If
            
            
            
            ' PL 03-8
            .Sheet = .SheetCount
            .GetText .ColLetterToNumber("E"), 85, errTongTLPB
            .GetText .ColLetterToNumber("E"), 87, errTongSoThue
            
            .Sheet = sheetPl03_8_TNDN
            blCheck_S12 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
            If blCheck_S12 = True Then
                mCurrentSheet = .Sheet
                iFlagEmpty = 0
                iFlagTyLe = 0
                iFlagCheckMaST = 0
                iFlagCheckSoThue = 0
                iWhiteRow = 1
                
                ' lay so thue cua to khai
                .GetText .ColLetterToNumber("K"), 29, vCt5
                
                'check not empty ten doanh nghiep
                .Row = 41
                Do
                    .GetText .ColLetterToNumber("D"), .Row, vTenDn
                    .GetText .ColLetterToNumber("F"), .Row, vMST
                    .GetText .ColLetterToNumber("I"), .Row, vCQTCapCuc
                    .GetText .ColLetterToNumber("K"), .Row, vCqtChiCuc
                    .GetText .ColLetterToNumber("L"), .Row, vTyLe
                    .GetText .ColLetterToNumber("O"), .Row, vSoThueQ1
                    .GetText .ColLetterToNumber("P"), .Row, vSoThueQ2
                    .GetText .ColLetterToNumber("Q"), .Row, vSoThueQ3
                    .GetText .ColLetterToNumber("R"), .Row, vSoThueQ4
                    .GetText .ColLetterToNumber("U"), .Row, vSoThue
                    
                    If Trim$(vTenDn) <> "" Or Trim$(vMST) <> "" Or Trim$(vCQTCapCuc) <> "" Or Trim$(vCqtChiCuc) <> "" Or Val(Trim$(vTyLe)) <> 0 Or Trim$(vSoThueQ1) <> "0" _
                    Or Trim$(vSoThueQ2) <> "0" Or Trim$(vSoThueQ3) <> "0" Or Trim$(vSoThueQ4) <> "0" Then
                        iWhiteRow = 0
                    End If
    
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
                        .Col = .ColLetterToNumber("I")
                        .BackColor = mErrorColor
                        .CellNote = errEmty
                        iFlagEmpty = 1
                    Else
                        .Col = .ColLetterToNumber("I")
                        .BackColor = mNonErrorColor
                        .CellNote = ""
                    End If
                    
                    If (vCqtChiCuc = "" And iWhiteRow = 0) Then
                        .Col = .ColLetterToNumber("K")
                        .BackColor = mErrorColor
                        .CellNote = errEmty
                        iFlagEmpty = 1
                    Else
                        .Col = .ColLetterToNumber("K")
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
                        .Col = .ColLetterToNumber("L")
                        .Row = .Row + 1
                        .BackColor = mErrorColor
                        .CellNote = errTongTLPB
                        iFlagTyLe = 1
                Else
                        .Row = .Row + 1
                        .Col = .ColLetterToNumber("C")
                         vColor = .BackColor
                        .Col = .ColLetterToNumber("L")
                        .BackColor = vColor
                        .CellNote = ""
                End If
                ' Kiem tra so thue
                If (CDbl(vCt5) - vTongSoThue > 5 Or vTongSoThue - CDbl(vCt5) > 5) Then
                        .Col = .ColLetterToNumber("U")
                        .BackColor = mErrorColor
                        .CellNote = errTongSoThue
                        iFlagCheckSoThue = 1
                Else
                        .Col = .ColLetterToNumber("C")
                         vColor = .BackColor
                        .Col = .ColLetterToNumber("U")
                        .BackColor = vColor
                        .CellNote = ""
                End If
                
                
            End If
            
            .Sheet = .SheetCount
            If blCheck_S12 Then
                .SetText .ColLetterToNumber("B"), 85, IIf(iFlagTyLe = 1, "0", "1")
                .SetText .ColLetterToNumber("B"), 87, IIf(iFlagCheckSoThue = 1, "0", "1")
            Else
                .SetText .ColLetterToNumber("B"), 85, "1"
                .SetText .ColLetterToNumber("B"), 87, "1"
            End If
            
            .SetText .ColLetterToNumber("B"), 29, IIf(iFlagEmpty = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 30, IIf(iFlagCheckMaST = 1, "0", "1")
        
            ' PL 03-9 Phu luc phan bo thuy dien
            .Sheet = .SheetCount
            
            Dim vChon As Variant
            Dim vTenNhaMay, vCQTCapDiaPhuong, vTyLePhanBo, vSoThuePhaiNop As Variant
            Dim vMessRequire As Variant
            
            Dim flagSoThue As Integer
            
            .GetText .ColLetterToNumber("E"), 88, vMessRequire
            .GetText .ColLetterToNumber("E"), 89, errMST
            
            .Sheet = sheetPl03_9_TNDN
            blCheck_S12 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
            If blCheck_S12 = True Then
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
            .SetText .ColLetterToNumber("B"), 88, IIf(iFlagEmpty = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 89, IIf(iFlagCheckMaST = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 90, IIf(flagSoThue = 1, "0", "1")
            
            ' check PL 03-5
            Dim vErrCt10, vErrCt11a, vErrCt12, vErrCt9  As Variant
            Dim iFlagCt10, iFlagCt11a, iFlagCt9, iFlagCt12 As Integer
            Dim ct9, ct10, ct11a, ct11, ct12 As Variant
            Dim ctTongCot5 As Variant
            Dim blCheck_S2B As Boolean
            
            .Sheet = .SheetCount
            .GetText .ColLetterToNumber("E"), 93, vErrCt10
            .GetText .ColLetterToNumber("E"), 94, vErrCt12
            .GetText .ColLetterToNumber("E"), 95, vErrCt11a
            .GetText .ColLetterToNumber("E"), 96, vErrCt9
            
            .Sheet = sheetPl03_5_TNDN
            blCheck_S12 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
            blCheck_S2B = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_2B_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
            If blCheck_S12 = True Then
                .GetText .ColLetterToNumber("F"), 30, ctTongCot5
                .GetText .ColLetterToNumber("E"), 29, ct9
                .GetText .ColLetterToNumber("E"), 30, ct10
                .GetText .ColLetterToNumber("E"), 31, ct11
                .GetText .ColLetterToNumber("E"), 32, ct11a
                .GetText .ColLetterToNumber("E"), 33, ct12
                If blCheck_S2B = True Then
                    ' check ct10
                    If Val(ct9) > 0 Then
                        If Val(ct10) > Val(ctTongCot5) Then
                            .Row = 30
                            .Col = .ColLetterToNumber("E")
                            .CellNote = vErrCt10
                            .BackColor = mErrorColor
                            iFlagCt10 = 1
                        Else
                            .Row = 30
                            .Col = .ColLetterToNumber("E")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    Else
                        .Row = 30
                        .Col = .ColLetterToNumber("E")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                    ' check ct11a
'                    If Val(ct11) < 0 Then
'                        If Val(ct11a) > Val(ctTongCot5) Then
'                            .Row = 32
'                            .Col = .ColLetterToNumber("E")
'                            .CellNote = vErrCt11a
'                            .BackColor = mErrorColor
'                            iFlagCt11a = 1
'                        Else
'                            .Row = 32
'                            .Col = .ColLetterToNumber("E")
'                            .CellNote = ""
'                            .BackColor = mNonErrorColor
'                        End If
'                    Else
'                        .Row = 32
'                        .Col = .ColLetterToNumber("E")
'                        .CellNote = ""
'                        .BackColor = mNonErrorColor
'                    End If
                    
                    
                    
'                    'check ct10+ct11a = tong cot 5 03-2B/TNDN
'                    If Val(ct10) + Val(ct11a) <> Val(ctTongCot5) Then
'                        .Row = 30
'                        .Col = .ColLetterToNumber("E")
'                        .CellNote = vErrCt10_11a
'                        .BackColor = mErrorColor
'
'                        .Row = 32
'                        .Col = .ColLetterToNumber("E")
'                        .CellNote = vErrCt10_11a
'                        .BackColor = mErrorColor
'                        iFlagCt10_11a = 1
'                    Else
'                        .Row = 30
'                        .Col = .ColLetterToNumber("E")
'                        .CellNote = ""
'                        .BackColor = mNonErrorColor
'
'                        .Row = 32
'                        .Col = .ColLetterToNumber("E")
'                        .CellNote = ""
'                        .BackColor = mNonErrorColor
'                    End If
                    
                End If
                ' kiem tra chi tieu 12
                    If Val(ct11) >= 0 Then
                        If Val(ct12) <= Round(Val(ct11) * 10 / 100, 0) Then
                            .Row = 33
                            .Col = .ColLetterToNumber("E")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        Else
                            .Row = 33
                            .Col = .ColLetterToNumber("E")
                            .CellNote = vErrCt12
                            .BackColor = mErrorColor
                            iFlagCt12 = 1
                        End If
                    Else
                        .Row = 33
                        .Col = .ColLetterToNumber("E")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                ' Kiem tra B14 = chi tieu 09 tren PL 03-5
                    If Val(ct9) <> Val(ctB14) Then
                        .Row = 29
                        .Col = .ColLetterToNumber("E")
                        .CellNote = vErrCt9
                        .BackColor = mErrorColor
                        iFlagCt9 = 1
                    Else
                        .Row = 29
                        .Col = .ColLetterToNumber("E")
                        .CellNote = ""
                        .BackColor = mTempColor
                    End If
            End If
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 93, IIf(iFlagCt10 = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 94, IIf(iFlagCt12 = 1, "0", "1")
            '.SetText .ColLetterToNumber("B"), 95, IIf(iFlagCt11a = 1, "0", "1")
            .SetText .ColLetterToNumber("B"), 96, IIf(iFlagCt9 = 1, "0", "1")
            
        .EventEnabled(EventAllEvents) = True
    End With
    
End Sub

'Calling this function whenever we don't need this class anymore.
'It will do some special task with special Error/ special Data
'no parameter
Public Sub finish()
    ResetWarning

    Dim ASheet As Integer, SSheet As Integer
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim vTemp As Variant
    Dim iCol As Long, iRow As Long, i As Long
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
                  
        Set xmlCellNode = Nothing
        CheckDynamicError
        UpdateSheets
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
        
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

End Sub

Public Function ResetData() As Boolean
    Dim startRow As Long
    Dim endRow As Long, endCol As Long
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlNodeReset As MSXML.IXMLDOMNode
    With fps
        Set xmlCellNode = TAX_Utilities_v1.Data(0).nodeFromID(GetCellID(fps, .ColLetterToNumber("O"), 31))
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode.lastChild
        Set xmlCellNode = xmlCellNode.firstChild
        Set xmlCellNode = xmlCellNode.firstChild
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), endCol, endRow

         For Each xmlNodeReset In TAX_Utilities_v1.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            .Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            .Col = lCol
            .Row = lRow
            
            If (lRow < 30 Or lRow >= endRow) And mCurrentSheet = 1 Then
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
                TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
            End If
        Next
    End With
End Function
'Xu ly khi nut Nhap lai dc kich hoat
Public Sub SetData()
    Dim intCurrSheet As Integer
    Dim arrResul() As String
    Dim blCheck_S05 As Boolean
    
        intCurrSheet = fps.Sheet
        With fps
            If .ActiveSheet = 1 Then
                mCurrentSheet = 1
                .Sheet = mCurrentSheet
                ' set lai trang thai dung cho C11, C12 to khai
                .Col = .ColLetterToNumber("O")
                .Row = 61
                .CellNote = ""
                .BackColor = mNonErrorColor

                .Col = .ColLetterToNumber("O")
                .Row = 62
                .CellNote = ""
                .BackColor = mNonErrorColor
                
                .Col = .ColLetterToNumber("B")
                .Row = 103
                .Formula = "B102+1"
            End If
             'set du lieu cho 03-2A
            If fps.ActiveSheet = sheetPl03_2A_TNDN Then
                mCurrentSheet = sheetPl03_2A_TNDN
                fps.Sheet = mCurrentSheet
                'Ham tao cac truong nam co trong PL
                Gene_Year
            End If
    
             'set du lieu cho 03-2B
            If fps.ActiveSheet = sheetPl03_2B_TNDN Then
                mCurrentSheet = sheetPl03_2B_TNDN
                fps.Sheet = mCurrentSheet
                'Ham tao cac truong nam co trong PL
                Gene_Year
            End If

            fps.Sheet = intCurrSheet
            mCurrentSheet = intCurrSheet
        End With
End Sub
' Sinh cac truong du lieu nam co trong PL
Sub Gene_Year()
    Dim rowStart As Integer
    Dim rowFinish As Integer
    Dim colGen As String
    Dim i As Integer
    Dim j As Integer
    
    rowStart = 26
    rowFinish = 30
    colGen = "C"
    j = 5
        
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = .ActiveSheet
        For i = rowStart To rowFinish
            .Col = fps.ColLetterToNumber(colGen)
            .Row = i
            .Text = TAX_Utilities_v1.Year - j
            .Lock = True
            .TypeHAlign = TypeHAlignCenter
            UpdateCell fps, .Col, .Row, .Text
            j = j - 1
        Next i
        .EventEnabled(EventAllEvents) = True
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
    
    Set xmlCellNode = TAX_Utilities_v1.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
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
    
    Set xmlCellNode = TAX_Utilities_v1.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
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
    Set xmlCellNode = TAX_Utilities_v1.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
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
    
    If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
        With fps
            .Sheet = .ActiveSheet
            mCurrentSheet = .ActiveSheet
            iCol = .ActiveCol
            iRow = .ActiveRow
            GetCellSpan fps, iCol, iRow
    
            Select Case .Sheet
                Case sheetPl03_3A_TNDN ' Phu Luc 03-3A
                    mCurrentSheet = .Sheet
                    Set xmlCellNode = TAX_Utilities_v1.Data(.Sheet - 1).nodeFromID(GetCellID(fps, iCol, iRow))
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

                Case sheetPl03_3B_TNDN ' Phu Luc 03-3B
                    mCurrentSheet = .Sheet
                    Set xmlCellNode = TAX_Utilities_v1.Data(.Sheet - 1).nodeFromID(GetCellID(fps, iCol, iRow))
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
                Case sheetPl03_3C_TNDN ' Phu Luc 03-3C
                    mCurrentSheet = .Sheet
                Case sheetPl03_4_TNDN ' Phu Luc 03-4
                    mCurrentSheet = .Sheet
                    .Row = iRow
                    .Col = .ColLetterToNumber("AE")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AH")
                    UpdateCell fps, .Col, .Row, .value
                    .Col = .ColLetterToNumber("AN")
                    UpdateCell fps, .Col, .Row, .value
                Case sheetPl03_6_TNDN ' Phu Luc 03-05
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
                .Row = 103
                Do
                     .Text = str(i)
                     .Col = .ColLetterToNumber("B")
                     .Row = i + 103
                     i = i + 1
                 Loop Until .Text = "aa"
            End If
            
            
            If mCurrentSheet = sheetPl03_4_TNDN Then
                fps.EventEnabled(EventAllEvents) = False
                i = 1
                .Col = .ColLetterToNumber("C")
                .Row = 26
                Do
                     .Text = str(i)
                     .Col = .ColLetterToNumber("C")
                     .Row = i + 26
                     i = i + 1
                 Loop Until .Text = "aa"
                 fps.EventEnabled(EventAllEvents) = True
            End If
            
            ' PL 03-6
            If .ActiveSheet = sheetPl03_6_TNDN Then
                   fps.EventEnabled(EventAllEvents) = True
                    i = 1
                    .Col = .ColLetterToNumber("C")
                    .Row = 26
                    Do
                         .Text = i
                         i = i + 1
                         .Row = .Row + 1
                    Loop Until .Text = "aa"
            End If
            
            ' PL 03-7
            If .ActiveSheet = sheetPl03_7_TNDN Then
                    mCurrentSheet = sheetPl03_7_TNDN
                   fps.EventEnabled(EventAllEvents) = False
                    i = 1
                   .Row = 31
                    Do
                         .Text = i
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "kk"
                                       
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = .Row + 19
                    Do
                     .Text = "1.1." & Trim(str(i))
                     UpdateCell fps, .Col, .Row, .Text
                     .Row = .Row + 1
                     i = i + 1
                    Loop Until .Text = "aa"
                    
                   i = 1
                   .Row = .Row + 3
                    Do
                         .Text = "1.2." & Trim(str(i))
                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"


                   i = 1
                   .Row = .Row + 4
                    Do
                         .Text = "2.1." & Trim(str(i))
                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "cc"

                   i = 1
                   .Row = .Row + 3
                    Do
                         .Text = "2.2." & Trim(str(i))
                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "dd"


                   i = 1
                   .Row = .Row + 3
                    Do
                         .Text = "2.3." & Trim(str(i))
                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "ee"

                   i = 1
                   .Row = .Row + 4
                    Do
                         .Text = "2.4.1." & Trim(str(i))
                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "ff"

                   i = 1
                   .Row = .Row + 3
                    Do
                         .Text = "2.4.2." & Trim(str(i))
                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "gg"


                    i = 1
                   .Row = .Row + 3
                    Do
                         .Text = "2.5." & Trim(str(i))
                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "hh"

            End If
            
            ' PL 03-8
            If .ActiveSheet = sheetPl03_8_TNDN Then
                 i = 1
                .Col = .ColLetterToNumber("C")
                .Row = 41
                Do
                     .Text = i
                     i = i + 1
                     .Row = .Row + 1
                Loop Until .Text = "aa"
            End If
            
            ' set lock CQT phan bo
            If .ActiveSheet = sheetPl03_9_TNDN Then
                .Sheet = sheetPl03_9_TNDN
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
    
  
    
    Set xmlCellNode = TAX_Utilities_v1.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
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
    
    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               If (lRow <= 97) And (lRow >= 31) And lRow <> 83 Then
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
                End If
            Next
            
       ' sheet 2
       For i = 2 To .SheetCount - 1
            .Sheet = i
            mCurrentSheet = .Sheet
            blnSheet = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
            If blnSheet = True Then
                For Each xmlCellNodeData In TAX_Utilities_v1.Data(.Sheet - 1).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    If i = sheetPl03_3A_TNDN Then
                        If (lRow <= 56) And (lRow >= 47) Then
                            .GetText lCol, lRow, varTemp
                            UpdateCell fps, lCol, lRow, varTemp
                        End If
                    ElseIf i = sheetPl03_3B_TNDN Then
                        If (lRow <= 65) And (lRow >= 52) Then
                             .GetText lCol, lRow, varTemp
                             UpdateCell fps, lCol, lRow, varTemp
                        End If
                    ElseIf i = sheetPl03_3C_TNDN Then
                        If (lRow <= 44) And (lRow >= 40) Then
                            .GetText lCol, lRow, varTemp
                            UpdateCell fps, lCol, lRow, varTemp
                        End If
                    ElseIf i = sheetPl03_7_TNDN Then
                        If lRow < TAX_Utilities_v1.Data(.Sheet - 1).getElementsByTagName("cell").length - 13 Then
                            .GetText lCol, lRow, varTemp
                            UpdateCell fps, lCol, lRow, varTemp
                        End If
                    ElseIf i = sheetPl03_9_TNDN Then
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
            Set xmlNode = .nodeFromID("O_60")
            result = "CTC10_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("O_61")
            result = result & "~" & "CTC11_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("O_62")
            result = result & "~" & "CTC12_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("O_65")
            result = result & "~" & "CTC15_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("O_66")
            result = result & "~" & "CTC16_" & Val(GetAttribute(xmlNode, "Value"))
'            Set xmlNode = .nodeFromID("O_65")
'            result = result & "~" & "CTD_" & Val(GetAttribute(xmlNode, "Value"))
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
                .GetText .ColLetterToNumber("O"), 60, varTemp
                strCurrentValue = "CTC10_" & Trim(varTemp)
                .GetText .ColLetterToNumber("O"), 61, varTemp
                strCurrentValue = strCurrentValue & "~CTC11_" & Trim(varTemp)
                .GetText .ColLetterToNumber("O"), 62, varTemp
                strCurrentValue = strCurrentValue & "~CTC12_" & Trim(varTemp)
                .GetText .ColLetterToNumber("O"), 65, varTemp
                strCurrentValue = strCurrentValue & "~CTC15_" & Trim(varTemp)
                .GetText .ColLetterToNumber("O"), 66, varTemp
                strCurrentValue = strCurrentValue & "~CTC16_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("O"), 65, varTemp
'                strCurrentValue = strCurrentValue & "~CTD_" & Trim(varTemp)
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
                    
                    If Left(strMaCT, 4) = "C10_" Or Left(strMaCT, 4) = "C16_" Then
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

' Lay ve chi tieu E o to khai lan truoc
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
            Set xmlNode = .nodeFromID("O_67")
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
        .GetText .ColLetterToNumber("O"), 67, varTemp
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
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 26, errCT3_4
        
        
        .Sheet = sheetPl03_3C_TNDN
        mCurrentSheet = sheetPl03_3C_TNDN
        
        Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl03_3C_TNDN - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("D"), 24))
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
'        .Row = 50
'        .Lock = status
        
        ' chi tieu C3a
        .Row = 51
        .Lock = status
'        'chi tieu C8
'        .Row = 47
'        .Lock = status
'        ' chi tieu C9
'        .Row = 48
'        .Lock = status
'        ' chi tieu C10
'        .Row = 49
'        .Lock = status
    End With
End Sub

' Lock cac cell lay du lieu len tu phu luc
Public Sub unLockCellPL(strLoaiToKhai As String)
    Dim blCheck_S2A As Boolean, blCheck_S2B As Boolean, blCheck_S3A As Boolean, blCheck_S3B As Boolean
    Dim blCheck_S3C As Boolean, blCheck_S4 As Boolean, blCheck_S5 As Boolean
    
    blCheck_S2A = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_2A_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S2B = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_2B_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S3A = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_3A_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S3B = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_3B_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S3C = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_3C_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S4 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_4_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S5 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_5_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    If strLoaiToKhai = "TKCT" Then
         With fps
            .Sheet = 1
            ' lock chi tieu C3, C3a
            .Col = .ColLetterToNumber("O")
            'If blCheck_S2A Or blCheck_S2B Then
            If blCheck_S2A Then
                ' chi tieu C3
'                .Row = 50
'                .Lock = True
                ' chi tieu C3a
                .Row = 51
                .Lock = True
            Else
                ' chi tieu C3
'                .Row = 50
'                .Lock = False
                ' chi tieu C3a
                .Row = 51
                .Lock = False
            End If
            ' Lock chi tieu C11
            If blCheck_S3A Or blCheck_S3B Then
                .Row = 61
                .Lock = True
            Else
                .Row = 61
                .Lock = False
            End If
            ' Lock chi tieu c12
            If blCheck_S3A Or blCheck_S3B Or blCheck_S3C Then
                .Row = 62
                .Lock = True
            Else
                .Row = 62
                .Lock = False
            End If
            ' Chi tieu C15
            If blCheck_S4 Then
                .Row = 65
                .Lock = True
            Else
                .Row = 65
                .Lock = False
            End If
'            ' Chi tieu D
'            If blCheck_S5 Then
'                .Row = 66
'                .Lock = True
'            Else
'                .Row = 66
'                .Lock = False
'            End If
            
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
        yearTK = TAX_Utilities_v1.Year
        For i = 1 To 4
            strDataFileName1A = TAX_Utilities_v1.DataFolder & GetAttribute(GetValidityNode("11").childNodes(0), "DataFile") & "_0" & i & TAX_Utilities_v1.Year & ".xml"
            strDataFileName1B = TAX_Utilities_v1.DataFolder & GetAttribute(GetValidityNode("12").childNodes(0), "DataFile") & "_0" & i & TAX_Utilities_v1.Year & ".xml"
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
    End With
    GetValue32_31TK_Quy = arrResult
    Set xmlDomLastData = Nothing
    Set xmlNode = Nothing
    
    Exit Function
    
ErrorHandle:
    SaveErrorLog "cls001", "GetValue32_31TK_Quy", Err.number, Err.Description
End Function


' TOng hop gia tri len PL 03-8/TNDN
Public Sub tongHopPL03_8()
         ' Lay gia tri cua cac quy khac de set vao phu luc 05/TNDN
         Dim arrResul() As String
         Dim blCheck_S05 As Boolean
        arrResul = GetValue32_31TK_Quy
        blCheck_S05 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_8_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        With fps
            If blCheck_S05 = True Then
                fps.EventEnabled(EventAllEvents) = False
                .Sheet = sheetPl03_8_TNDN
                mCurrentSheet = .Sheet
                .Row = 41
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
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
                Case "KHBS"
                    blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl03_KHBS_TNDN - 1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(sheetPl03_KHBS_TNDN - 1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl03_KHBS_TNDN - 1).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl03_KHBS_TNDN - 1).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v1.Data(sheetPl03_KHBS_TNDN - 1).getElementsByTagName("Cell").length > 31 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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
    tuRowDL = 26
    denRowDL = 31
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "17_18_2~17_18_3~17_18_4~16_18_5~16_18_6~19_21_7~20_22_8~19_21_9~14_16_10~15_17_11~16_18_13~19_21_14~11_12_15~20_30_16"
    ' end
End Sub

' set co quan thue ra quyet dinh hoan thue
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
                        tenCucThue = tenCucThue + arrDanhsach(3) + Chr$(9)
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
            .Sheet = sheetPl03_8_TNDN 'set thong tin CQL
            '.EventEnabled(EventAllEvents) = False
            
            .Row = 41
            .Col = .ColLetterToNumber("I")
            .TypeComboBoxList = tenCucThue
            
            .Row = 41
            .Col = .ColLetterToNumber("Y")
            .TypeComboBoxList = maCucThue
            
            .Row = 41
            .Col = .ColLetterToNumber("K")
            .TypeComboBoxList = tenChiCucThue
    
            .Row = 41
            .Col = .ColLetterToNumber("Z")
            .TypeComboBoxList = maChiCucThue
            
            ' Phu luc 03_9
            .Sheet = sheetPl03_9_TNDN 'set thong tin CQL
            .Row = 42
            .Col = .ColLetterToNumber("P")
            .TypeComboBoxList = tenCucThue
            
            .Row = 42
            .Col = .ColLetterToNumber("AH")
            .TypeComboBoxList = maCucThue
            
            '.EventEnabled(EventAllEvents) = True
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
    
    
    Dim tenCucThue As String
    Dim maCucThue As String
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim intIndexCombo As Integer, i As Integer
    Dim iCol As Long, iRow As Long
  '  Dim checkDL As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        If .ActiveSheet = sheetPl03_8_TNDN Then
            .Sheet = sheetPl03_8_TNDN
            mCurrentSheet = sheetPl03_8_TNDN
            
            .Col = Col
            .Row = Row
            
            If Col = .ColLetterToNumber("I") Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
                intIndexCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("Y")
                .Row = Row
                .typeComboboxCurSel = intIndexCuc
                'Nhap lai Combo cot Q
                .Col = .ColLetterToNumber("Y")
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
                    .Col = .ColLetterToNumber("K")
                    .TypeComboBoxList = tenChiCucThue
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("Z")
                    .TypeComboBoxList = maChiCucThue
                    .Text = ""
                     UpdateCell fps, .Col, .Row, .Text
                   
                End If
            End If
                        
            If Col = .ColLetterToNumber("K") Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
                intIndexChiCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("Z")
                .Row = Row
                .typeComboboxCurSel = intIndexChiCuc
                'Nhap lai Combo cot Q
                .Col = .ColLetterToNumber("Z")
                UpdateCell fps, .Col, .Row, .Text
            End If
        ElseIf .ActiveSheet = sheetTK03_TNDN Then
            .Sheet = sheetTK03_TNDN
            mCurrentSheet = sheetTK03_TNDN
            .Col = Col
            .Row = Row
            
            If Col = .ColLetterToNumber("G") And Row = 14 Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot G va gia tri Index cot L
                intIndexChiCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("P")
                .Row = Row
                .typeComboboxCurSel = intIndexChiCuc
                UpdateCell fps, .Col, .Row, .Text
            End If
        ElseIf .ActiveSheet = sheetPl03_9_TNDN Then
            .Sheet = sheetPl03_9_TNDN
            mCurrentSheet = sheetPl03_9_TNDN
            .Col = Col
            .Row = Row
            
            
            If Col = .ColLetterToNumber("P") Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot G va gia tri Index cot L
                intIndexChiCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("AH")
                .Row = Row
                .typeComboboxCurSel = intIndexChiCuc
                UpdateCell fps, .Col, .Row, .Text
            End If
           
 
            End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub


Private Function getDenNgayNC() As String
    Dim strNgayBatDauNamTC As String
    Dim strDenNgayNC As String
    strNgayBatDauNamTC = GetNgayBatDauNamTaiChinh
    If strNgayBatDauNamTC = "01/01" Then
        strDenNgayNC = "31/03/" & Val(TAX_Utilities_v1.Year) + 1
    ElseIf strNgayBatDauNamTC = "01/04" Then
        strDenNgayNC = "30/06/" & Val(TAX_Utilities_v1.Year) + 1
    ElseIf strNgayBatDauNamTC = "01/07" Then
        strDenNgayNC = "30/09/" & Val(TAX_Utilities_v1.Year) + 1
    ElseIf strNgayBatDauNamTC = "01/10" Then
        strDenNgayNC = "31/12/" & Val(TAX_Utilities_v1.Year) + 1
    End If
    getDenNgayNC = strDenNgayNC
End Function


' ham set thong tin sau khi tai XML
Public Sub finishImportXML()
    Dim i As Integer
    Dim CQT As Variant
    Dim tempValue As Variant
    Dim strCol As String
    Dim arrCol() As String
    Dim strRow As String
    Dim arrRow() As String

    
    With fps
        .EventEnabled(EventAllEvents) = False
        ' TK 03/TNDN
         .Sheet = sheetTK03_TNDN
         mCurrentSheet = sheetTK03_TNDN
         strRow = "8~10~12"
         arrRow = Split(strRow, "~")
         For i = 0 To UBound(arrRow)
            .Row = arrRow(i)
            .GetText .ColLetterToNumber("D"), .Row, tempValue
            If tempValue = "1" Then
               .Col = .ColLetterToNumber("D")
               UpdateCell fps, .ColLetterToNumber("D"), .Row, "x"
            ElseIf tempValue = "0" Then
               .Col = .ColLetterToNumber("D")
               UpdateCell fps, .ColLetterToNumber("D"), .Row, ""
            End If
         Next
         
         .Row = 83
         .GetText .ColLetterToNumber("F"), .Row, tempValue
         If tempValue = "1" Then
           .Col = .ColLetterToNumber("F")
           UpdateCell fps, .ColLetterToNumber("F"), .Row, "x"
         ElseIf tempValue = "0" Then
           .Col = .ColLetterToNumber("F")
           UpdateCell fps, .ColLetterToNumber("F"), .Row, ""
         End If
         
        
         ' PL 03-3A
         If TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_3A_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0" Then
                .Sheet = sheetPl03_3A_TNDN
                mCurrentSheet = sheetPl03_3A_TNDN
                .Row = 25
                For i = 0 To 10
                     .GetText .ColLetterToNumber("C"), .Row, tempValue
                     If tempValue = "1" Then
                        .Col = .ColLetterToNumber("C")
                        UpdateCell fps, .ColLetterToNumber("C"), .Row, "x"
                     ElseIf tempValue = "0" Then
                        .Col = .ColLetterToNumber("C")
                        UpdateCell fps, .ColLetterToNumber("C"), .Row, ""
                     End If
                     .Row = .Row + 1
                Next
        End If
                 
         ' PL 03-3B
         If TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_3B_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0" Then
             .Sheet = sheetPl03_3B_TNDN
             mCurrentSheet = sheetPl03_3B_TNDN
             .Row = 27
             For i = 0 To 4
                  .GetText .ColLetterToNumber("D"), .Row, tempValue
                  If tempValue = "1" Then
                     .Col = .ColLetterToNumber("D")
                     UpdateCell fps, .ColLetterToNumber("D"), .Row, "x"
                  ElseIf tempValue = "0" Then
                     .Col = .ColLetterToNumber("D")
                     UpdateCell fps, .ColLetterToNumber("D"), .Row, ""
                  End If
                  .Row = .Row + 1
             Next
             
            .Row = 33
            .GetText .ColLetterToNumber("D"), .Row, tempValue
            If tempValue = "1" Then
               .Col = .ColLetterToNumber("D")
               UpdateCell fps, .ColLetterToNumber("D"), .Row, "x"
            ElseIf tempValue = "0" Then
               .Col = .ColLetterToNumber("D")
               UpdateCell fps, .ColLetterToNumber("D"), .Row, ""
            End If
              
              
            .Row = 34
            .GetText .ColLetterToNumber("D"), .Row, tempValue
            If tempValue = "1" Then
               .Col = .ColLetterToNumber("D")
               UpdateCell fps, .ColLetterToNumber("D"), .Row, "x"
            ElseIf tempValue = "0" Then
               .Col = .ColLetterToNumber("D")
               UpdateCell fps, .ColLetterToNumber("D"), .Row, ""
            End If
        End If
        ' PL 03-3C
        
        If TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_3C_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0" Then
            .Sheet = sheetPl03_3C_TNDN
            mCurrentSheet = sheetPl03_3C_TNDN
            strRow = "24~28~33~34"
            arrRow = Split(strRow, "~")
            For i = 0 To UBound(arrRow)
               .Row = arrRow(i)
               .GetText .ColLetterToNumber("D"), .Row, tempValue
               If tempValue = "1" Then
                  .Col = .ColLetterToNumber("D")
                  UpdateCell fps, .ColLetterToNumber("D"), .Row, "x"
               ElseIf tempValue = "0" Then
                  .Col = .ColLetterToNumber("D")
                  UpdateCell fps, .ColLetterToNumber("D"), .Row, ""
               End If
            Next
        End If
        
        
        If TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_7_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0" Then
             .Sheet = sheetPl03_7_TNDN
             mCurrentSheet = sheetPl03_7_TNDN
             strCol = "U~V~W~X~Y~Z~AA~AB~AC~AD~AE~AF~AG"
             arrCol = Split(strCol, "~")
             
             i = 1
            .Row = 31
            Do
                   For i = 0 To UBound(arrCol)
                        .GetText .ColLetterToNumber(arrCol(i)), .Row, tempValue
                        If tempValue = "1" Then
                            .Col = .ColLetterToNumber(arrCol(i))
                            UpdateCell fps, .ColLetterToNumber(arrCol(i)), .Row, "x"
                        ElseIf tempValue = "0" Then
                            .Col = .ColLetterToNumber(arrCol(i))
                            UpdateCell fps, .ColLetterToNumber(arrCol(i)), .Row, ""
                        End If
                   Next
                  .Col = .ColLetterToNumber("B")
                  .Text = i
                  .Row = 1 + .Row
            Loop Until .Text = "kk"
                                
                                        
            i = 1
             .Col = .ColLetterToNumber("B")
             .Row = .Row + 19
             Do
              .GetText .ColLetterToNumber("N"), .Row, tempValue
              If Len(Trim$(tempValue)) > 3 Then
                    .Col = .ColLetterToNumber("AJ")
                    .Text = Left(CStr(tempValue), 3)
                    UpdateCell fps, .Col, .Row, .Text
              End If
              
              .GetText .ColLetterToNumber("AB"), .Row, tempValue
              If Len(Trim$(tempValue)) > 3 Then
                    .Col = .ColLetterToNumber("AK")
                    .Text = Left(CStr(tempValue), 3)
                    UpdateCell fps, .Col, .Row, .Text
              End If
              
              .Col = .ColLetterToNumber("B")
               .Text = "1.1." & Trim(str(i))
               UpdateCell fps, .Col, .Row, .Text
              .Row = .Row + 1
             Loop Until .Text = "aa"
             
            .Row = .Row + 3
            i = 1
             Do
                    
                    .GetText .ColLetterToNumber("N"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AJ")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
                    
                    .GetText .ColLetterToNumber("AB"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AK")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
                    
                  .Col = .ColLetterToNumber("B")
                  .Text = "1.2." & Trim(str(i))
                  UpdateCell fps, .Col, .Row, .Text
                  .Row = 1 + .Row
             Loop Until .Text = "bb"
    
    
            .Row = .Row + 4
            i = 1
             Do
    
                   .GetText .ColLetterToNumber("N"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AJ")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
                    
                    .GetText .ColLetterToNumber("AB"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AK")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
    
                  .Col = .ColLetterToNumber("B")
                  .Text = "2.1." & Trim(str(i))
                  UpdateCell fps, .Col, .Row, .Text
                  .Row = 1 + .Row
             Loop Until .Text = "cc"
    
            .Row = .Row + 3
            i = 1
             Do
    
                   .GetText .ColLetterToNumber("N"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AJ")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
                    
                    .GetText .ColLetterToNumber("AB"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AK")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
                    
                  .Col = .ColLetterToNumber("B")
                  .Text = "2.2." & Trim(str(i))
                  UpdateCell fps, .Col, .Row, .Text
                  .Row = 1 + .Row
             Loop Until .Text = "dd"
    
    
            .Row = .Row + 3
            i = 1
             Do
    
                   .GetText .ColLetterToNumber("N"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AJ")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
                    
                    .GetText .ColLetterToNumber("AB"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AK")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
    
                  .Col = .ColLetterToNumber("B")
                  .Text = "2.3." & Trim(str(i))
                  UpdateCell fps, .Col, .Row, .Text
                  .Row = 1 + .Row
             Loop Until .Text = "ee"
    
            .Row = .Row + 4
            i = 1
             Do
                    
                   .GetText .ColLetterToNumber("N"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AJ")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
                    
                    .GetText .ColLetterToNumber("AB"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AK")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
                    
                  .Col = .ColLetterToNumber("B")
                  .Text = "2.4.1." & Trim(str(i))
                  UpdateCell fps, .Col, .Row, .Text
                  .Row = 1 + .Row
             Loop Until .Text = "ff"
    
    
            .Row = .Row + 3
            i = 1
             Do
                    
                   .GetText .ColLetterToNumber("N"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AJ")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
                    
                    .GetText .ColLetterToNumber("AB"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AK")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
                    
                  .Col = .ColLetterToNumber("B")
                  .Text = "2.4.2." & Trim(str(i))
                  UpdateCell fps, .Col, .Row, .Text
                  .Row = 1 + .Row
             Loop Until .Text = "gg"
    
    
            .Row = .Row + 3
            i = 1
             Do
                    
                   .GetText .ColLetterToNumber("N"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AJ")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
                    
                    .GetText .ColLetterToNumber("AB"), .Row, tempValue
                    If Len(Trim$(tempValue)) > 3 Then
                          .Col = .ColLetterToNumber("AK")
                          .Text = Left(CStr(tempValue), 3)
                          UpdateCell fps, .Col, .Row, .Text
                    End If
                    
                  .Col = .ColLetterToNumber("B")
                  .Text = "2.5." & Trim(str(i))
                  UpdateCell fps, .Col, .Row, .Text
                  .Row = 1 + .Row
             Loop Until .Text = "hh"
             ' end fill data
          End If
        
        
        If TAX_Utilities_v1.NodeValidity.childNodes(sheetPl03_9_TNDN - 1).Attributes.getNamedItem("Active").nodeValue <> "0" Then
            .Sheet = sheetPl03_9_TNDN
            mCurrentSheet = sheetPl03_9_TNDN
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
         End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub


'Public Sub setSttPL03_7()
'    Dim i As Long
'    With pGrid
'        If TAX_Utilities_v1.NodeValidity.childNodes(12).Attributes.getNamedItem("Active").nodeValue <> "0" Then
'            .Sheet = 13
'            pGrid.EventEnabled(EventAllEvents) = False
'
'             i = 1
'             .Col = .ColLetterToNumber("B")
'             .Row = 48
'             Do
'              .Text = "1.1." & Trim(str(i))
'              .Row = .Row + 1
'              i = i + 1
'             Loop Until .Text = "aa"
'
'            i = 1
'            .Row = .Row + 3
'             Do
'                  .Text = "1.2." & Trim(str(i))
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "bb"
'
'
'            i = 1
'            .Row = .Row + 4
'             Do
'                  .Text = "2.1." & Trim(str(i))
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "cc"
'
'            i = 1
'            .Row = .Row + 3
'             Do
'                  .Text = "2.2." & Trim(str(i))
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "dd"
'
'
'            i = 1
'            .Row = .Row + 3
'             Do
'                  .Text = "2.3." & Trim(str(i))
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "ee"
'
'            i = 1
'            .Row = .Row + 4
'             Do
'                  .Text = "2.4.1." & Trim(str(i))
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "ff"
'
'            i = 1
'            .Row = .Row + 3
'             Do
'                  .Text = "2.4.2." & Trim(str(i))
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "gg"
'
'
'             i = 1
'            .Row = .Row + 3
'             Do
'                  .Text = "2.5." & Trim(str(i))
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "hh"
'            pGrid.EventEnabled(EventAllEvents) = True
'        End If
'      End With
'End Sub


