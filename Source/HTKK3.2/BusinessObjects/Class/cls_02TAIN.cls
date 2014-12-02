VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_02TAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Thu Nhap Doanh Nghiep" interface sheets
'this Class is belong to TAX_Business_v1 project which will be compline to DLL
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
Dim mblKeyF5 As Boolean
Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
Dim arrData() As String
Public loaiKHBS As String
Public FlagThaydoi As String
Public Strloaitk As String
Public StrSolanBosung As String
Dim arrDanhSachKHBS() As String

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

    With fps
        SetDateFormat fps, 1, 51, .ColLetterToNumber("AG"), DDMMYYYY
        SetDateFormat fps, 2, 31, fps.ColLetterToNumber("BE"), DDMMYYYY
        SetDateFormat fps, 2, 46, fps.ColLetterToNumber("BG"), DDMMYYYY
        
        .Sheet = 1 '01
        .Row = 51
        .Col = fps.ColLetterToNumber("AG")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
    End With
    blChon = False
    'strDataFileName = GetCatalogueFileName(blChon)
    blChon = False
    If TAX_Utilities_v1.Year = "2008" Then
        strDataFileName = Replace(GetCatalogueTemplateFile, "Catalogue_052008_01", "CatalogueQTTN_2008_01")
    ElseIf TAX_Utilities_v1.Year = "2010" Then
        strDataFileName = Replace(GetCatalogueTemplateFile, "Catalogue_072010_01", "CatalogueQTTN_2010_01")
    ElseIf TAX_Utilities_v1.Year = "2014" Then
        strDataFileName = Replace(GetCatalogueTemplateFile, "Catalogue_022014_01", "CatalogueQTTN_2014_01")
    Else
        strDataFileName = GetCatalogueFileName(blChon)
    End If
    
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
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        .Col = .ColLetterToNumber("F")
        .Row = 40
        .TypeComboBoxList = strCombo 'CPab("Maple", 100) + Chr$(9) + CPab("Oak", 100) + Chr$(9) + CPab("Ash", 100)
        .Row = 44
        .TypeComboBoxList = strCombo
        'Combo Hien
        .Col = .ColLetterToNumber("G")
        .Row = 40
        .TypeComboBoxList = strComboHien
        .Row = 44
        .TypeComboBoxList = strComboHien
        .Col = .ColLetterToNumber("P")
        .Row = 40
        .TypeComboBoxList = strDonvitinh
        .Row = 44
        .TypeComboBoxList = strDonvitinh
    End With

    Set xmlNodeListCell = Nothing
    ' Lay dia diem khai thac tai nguyen
    If TAX_Utilities_v1.Month = "01" Then 'Thang 1 ve thang 12 nam truoc
        strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & "12" & Trim(str(CInt(TAX_Utilities_v1.Year) - 1)) & ".xml"
    Else
        'dntai comment ngay 29/08/2011
        'strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & Right("00" & Trim(str(CInt(TAX_Utilities_v1.Month) - 1)), 2) & TAX_Utilities_v1.Year & ".xml"
        'end
        strDataFileName = TAX_Utilities_v1.DataFolder & GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") & "_" & TAX_Utilities_v1.Year & ".xml"
    End If
    With fps
        .Sheet = 1
        mCurrentSheet = 1
        If xmlDomData.Load(strDataFileName) Then
                Set xmlNode = xmlDomData.nodeFromID(GetCellID(fps, fps.ColLetterToNumber("N"), 6))
                .Row = 6
                .Col = .ColLetterToNumber("N")
                strDDKTTN = GetAttribute(xmlNode, "Value")
                .Text = strDDKTTN
                'UpdateCell fps, .Col, .Row, strDDKTTN
        End If
    End With
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
Dim xmlNode As MSXML.IXMLDOMNode
Dim iCol As Long, iRow As Long, GTSua As String, tRow As Long
'bien dung de luu cac section
Dim iSection, countRow As Integer
Dim count1 As Long
    With fps
        mblKhoiDauToKhai = True
        .Sheet = .ActiveSheet
        mCurrentSheet = .ActiveSheet
        mblKhoiDau = False
        
        .Sheet = 1
        'set to khai bo xung hay chinh thuc
            If Strloaitk = "TKCT" Then
                .Col = .ColLetterToNumber("I")
                .Row = 17
                .Text = 1
                UpdateCell fps, .Col, .Row, .Text
                
                .Col = .ColLetterToNumber("M")
                .Row = 17
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
             ElseIf Strloaitk = "TKBS" Then
                .Col = .ColLetterToNumber("I")
                .Row = 17
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                
                 .Col = .ColLetterToNumber("M")
                .Row = 17
                .Text = StrSolanBosung
                UpdateCell fps, .Col, .Row, .Text
            End If
         'end
         
         ' set tu thang den thang
        .Col = .ColLetterToNumber("Q")
        .Row = 17
        .Text = TAX_Utilities_v1.FirstDay
        UpdateCell fps, .Col, .Row, .Text
        
        .Col = .ColLetterToNumber("R")
        .Row = 17
        .Text = TAX_Utilities_v1.LastDay
        UpdateCell fps, .Col, .Row, .Text
        
            iSection = 1
            i = 1
            .Col = .ColLetterToNumber("C")
            .Row = 40
            countRow = 40
            tRow = .Row
            Do
                .GetText .ColLetterToNumber("F"), .Row, strLTN
                .GetText .ColLetterToNumber("AD"), .Row, Col7
                .GetText .ColLetterToNumber("E"), .Row, strId
                .GetText .ColLetterToNumber("O"), .Row, strLTNCu
                
                GTSua = Mid(strLTNCu, 231, 1)
                
                If Col7 = 0 Then
                     If Trim(strId) = "1101" Or GTSua = "1" Then
                        .Col = .ColLetterToNumber("AA")
                        .Lock = False
                     Else
                        If Trim(Mid(strLTN, 221, 10)) <> "" Then
                            .Col = .ColLetterToNumber("AA")
                            .value = Trim(Mid(strLTN, 221, 10))
                            UpdateCell fps, .Col, .Row, .value
                            .Lock = True
                        End If
                     End If
                Else
                    .Col = .ColLetterToNumber("AA")
                    .value = 0
                    .Row = tRow
                    UpdateCell fps, .Col, .Row, .value
                    .Lock = True
                End If
'                .CellNote = ""
'                .BackColor = mNonErrorColor
                
                If strLTN = "" And strLTNCu <> "" Then
                    .Col = .ColLetterToNumber("F")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                    '.TypeComboBoxList = strCombo + strLTNCu + Chr$(9)
                    .TypeComboBoxString = strLTNCu
                    .Text = strLTNCu
                    UpdateCell fps, .Col, .Row, .Text

                    'Combo Hien
                    .Col = .ColLetterToNumber("G")
                    '.TypeComboBoxList = strComboHien + Trim(Mid(strLTN, 1, 200)) + Chr$(9)
                    .TypeComboBoxString = Trim(Mid(strLTNCu, 1, 200))
                    .Text = Trim(Mid(strLTNCu, 1, 200))
                    UpdateCell fps, .Col, .Row, .Text
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                    .ActiveSheet = 1
                    fps_Change .Col, .Row
                End If
                
                If strLTN = "" And strLTNCu = "" And strId <> "" Then
                     If arrTaiNguyen.Exists(strId) Then
                        .Col = .ColLetterToNumber("F")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        '.TypeComboBoxList = strCombo + strLTNCu + Chr$(9)
                        .Text = arrTaiNguyenGoc.Item(strId)
                        UpdateCell fps, .Col, .Row, .Text
    
                        'Combo Hien
                        .Col = .ColLetterToNumber("G")
                        '.TypeComboBoxList = strComboHien + Trim(Mid(strLTN, 1, 200)) + Chr$(9)
                        .value = Trim(Mid(arrTaiNguyenGoc.Item(strId), 1, 200))
                        .Text = Trim(Mid(arrTaiNguyenGoc.Item(strId), 1, 200))
                        UpdateCell fps, .Col, .Row, .Text
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        mblKhoiDau = True
                        .ActiveSheet = 1
                        fps_Change .Col, .Row
                    Else
                        .Col = .ColLetterToNumber("F")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        '.TypeComboBoxList = strCombo + strLTNCu + Chr$(9)
                        .TypeComboBoxString = arrTaiNguyenGoc.Item(strId)
                        .Text = arrTaiNguyenGoc.Item(strId)
                        UpdateCell fps, .Col, .Row, .Text
    
                        'Combo Hien
                        .Col = .ColLetterToNumber("G")
                        '.TypeComboBoxList = strComboHien + Trim(Mid(strLTN, 1, 200)) + Chr$(9)
                        .TypeComboBoxString = Trim(Mid(arrTaiNguyenGoc.Item(strId), 1, 200))
                        .Text = Trim(Mid(arrTaiNguyenGoc.Item(strId), 1, 200))
                        UpdateCell fps, .Col, .Row, .Text
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        mblKhoiDau = True
                        .ActiveSheet = 1
                        fps_Change .Col, .Row
                    End If
                End If
                .Col = .ColLetterToNumber("G")
                If .value <> "" Then
                    .Col = .ColLetterToNumber("P")
                    UpdateCell fps, .Col, .Row, .Text
                    .Lock = False
                End If
                'UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("C")
                ' kiem tra neu nam 2014 tro di cho phep sua cot 8
                If Val(TAX_Utilities_v1.Year) >= 2014 Then
                     .Col = .ColLetterToNumber("AH")
                    .Lock = False
                    
                    .Col = .ColLetterToNumber("C")
                End If
                ' end
      '          .Row = i + 40
                .Row = countRow + i
                tRow = .Row
                i = i + 1
                
                If iSection = 1 And .Text = "aa" Then
                    iSection = iSection + 1
                    .Row = .Row + 3
                    countRow = .Row
                    i = 1
                End If
             Loop Until .Text = "aa"

            .Sheet = 1
            mCurrentSheet = 1

            Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 40))

            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode.lastChild
            Set xmlCellNode = xmlCellNode.firstChild
            
            'set nguoi ky
            Set xmlCellNode = xmlCellNode.firstChild
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
            'set NV DLT
            Set xmlCellNode = xmlCellNode.nextSibling
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            .Row = iRow
            .Col = iCol
            .Formula = ""
            'set CC hanh nghe
            Set xmlCellNode = xmlCellNode.nextSibling
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            .Row = iRow
            .Col = iCol
            .Formula = ""
        'End If
        mblKhoiDauToKhai = False
        
        .Sheet = 1
'tinh so dong` them vao`
'            .Row = 13
'            .Col = .ColLetterToNumber("C")
'            Do While .value <> "aa"
'                count1 = count1 + 1
'                .Row = .Row + 1
'            Loop
        
        .Row = .MaxRows - 7
        .Col = .ColLetterToNumber("AG")
        If .value = "" Then
            SetDateFormat fps, 1, .Row, .ColLetterToNumber("AG"), DDMMYYYY
            '.Value = "" + Conversion.CStr(Day(Now)) + Conversion.CStr(Month(Now)) + Conversion.CStr(Year(Now))
            .Text = Format(Date, "dd/mm/yyyy")
            UpdateCell fps, .Col, .Row, .Text
        End If
        
'       If Not TAX_Utilities_v1.DataKHBS Is Nothing Then
'                Set xmlNode = TAX_Utilities_v1.Data(.SheetCount - 2).childNodes(2).firstChild
'                SetAttribute xmlNode, "DateKHBS", TAX_Utilities_v1.DateKHBS
'                SetAttribute xmlNode, "loaiKHBS", loaiKHBS
'                If loaiKHBS = "frmKHBS_BS" Then
'                    .SetText .ColLetterToNumber("B"), 5, GetAttribute(GetMessageCellById("0115"), "Msg") & _
'                    GetAttribute(GetMessageCellById("0117"), "Msg") & Mid(TAX_Utilities_v1.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 5, 4)
'                Else
'                    .SetText .ColLetterToNumber("B"), 5, GetAttribute(GetMessageCellById("0116"), "Msg") & _
'                    GetAttribute(GetMessageCellById("0117"), "Msg") & Mid(TAX_Utilities_v1.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_v1.DateKHBS, 5, 4)
'                End If
'            filldata
'          .EventEnabled(EventAllEvents) = False
'             .Sheet = 1
'            .Row = .MaxRows - 7
'            .Col = .ColLetterToNumber("AG")
'            .Formula = "Header!T5"
'            UpdateCell fps, .Col, .Row, .Text
'            .Sheet = 1
'            .Row = .MaxRows - 9
'            .Col = .ColLetterToNumber("AG")
'            .Text = Format(Date, "dd/mm/yyyy")
'            UpdateCell fps, .Col, .Row, .Text
'            .EventEnabled(EventAllEvents) = True
'         .ActiveSheet = .SheetCount - 1
'         .ActiveSheet = .SheetCount - 1
'       End If
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
Dim xmlNodeH15 As MSXML.IXMLDOMNode

    fps.EventEnabled(EventAllEvents) = False
    With fps
        
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("G") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("F")
            .typeComboboxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber("F")
            UpdateCell fps, .Col, .Row, .Text
             
            'Chon lai combo cua colF
            .Col = .ColLetterToNumber("E")
            .GetText .ColLetterToNumber("F"), Row, strId
            .Text = Trim(Mid(strId, 200, 10))
             UpdateCell fps, .Col, .Row, .Text
             .Col = .ColLetterToNumber("D")
             .Text = Trim(Left(strId, 200))
             UpdateCell fps, .Col, .Row, .Text
             .Col = .ColLetterToNumber("P")
             .Text = Trim(Mid(strId, 210, 10))
             UpdateCell fps, .Col, .Row, .Text
             If strId <> "" Then
                .Col = .ColLetterToNumber("O")
                .Text = strId
                UpdateCell fps, .Col, .Row, .Text
             End If
            .GetText .ColLetterToNumber("AD"), .Row, Col7
            
            GTSua = Mid(strId, 231, 1)
            
            .EventEnabled(EventChange) = True
            If Col7 = 0 Then
                .Col = .ColLetterToNumber("AA")
                If Trim(Mid(strId, 200, 10)) = "1101" Or GTSua = "1" Then
                    If mblKhoiDau Or mblKeyF5 Then
                        .Lock = False
                        mblKhoiDau = False
                    Else
                        .Col = .ColLetterToNumber("AA")
                        .value = Trim(Mid(strId, 221, 10))
                        .Col = .ColLetterToNumber("AA")
                        .Row = Row
                        UpdateCell fps, .Col, .Row, .value
                        .Lock = False
                        mblKhoiDau = False
                    End If
                Else
                    .Col = .ColLetterToNumber("AA")
                    .value = Trim(Mid(strId, 221, 10))
                    .Col = .ColLetterToNumber("AA")
                    .Row = Row
                    UpdateCell fps, .Col, .Row, .value
                    .Lock = True
                    mblKhoiDau = False
                End If
            Else
                .Col = .ColLetterToNumber("AA")
                .value = 0
                .Col = .ColLetterToNumber("AA")
                .Row = Row
                UpdateCell fps, .Col, .Row, 0
                .Lock = True
                .Col = .ColLetterToNumber("V")
                .value = 0
                .Col = .ColLetterToNumber("V")
                .Row = Row
                UpdateCell fps, .Col, .Row, 0
                mblKhoiDau = False
            End If
            .EventEnabled(EventChange) = False
        End If
        
        mblKeyF5 = False
        
        If Col = .ColLetterToNumber("F") Then
            .Col = .ColLetterToNumber("E")
            .GetText .ColLetterToNumber("F"), Row, strId
            .Text = Trim(Mid(strId, 200, 10))
             UpdateCell fps, .Col, .Row, .Text
             .Col = .ColLetterToNumber("D")
             .Text = Trim(Left(strId, 200))
             UpdateCell fps, .Col, .Row, .Text
             .Col = .ColLetterToNumber("P")
             .Text = Trim(Mid(strId, 210, 10))
             UpdateCell fps, .Col, .Row, .Text
            .GetText .ColLetterToNumber("AD"), .Row, Col7
            If strId <> "" Then
                .Col = .ColLetterToNumber("O")
                .Text = strId
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            GTSua = Mid(strId, 231, 1)
            
            If Col7 = 0 Then
                .Col = .ColLetterToNumber("AA")
                .value = Trim(Mid(strId, 221, 10))
                .Col = .ColLetterToNumber("AA")
                .Row = Row
                UpdateCell fps, .Col, .Row, .value
                If Trim(Mid(strId, 200, 10)) = "1101" Or GTSua = "1" Then
                    .Lock = False
                Else
                    .Lock = True
                End If
            Else
                .Col = .ColLetterToNumber("AA")
                .value = 0
                .Col = .ColLetterToNumber("AA")
                .Row = Row
                UpdateCell fps, .Col, .Row, 0
                
                GTSua = Mid(strId, 231, 1)
                
                If Trim(Mid(strId, 200, 10)) = "1101" Or GTSua = "1" Then
                    .Lock = False
                Else
                    .Lock = True
                End If
            End If
        End If
        
        .Col = Col
        If Col = .ColLetterToNumber("V") Then
            If Val(.value) > 0 Then
                .Col = .ColLetterToNumber("E")
                .GetText .ColLetterToNumber("F"), Row, strId
    
                'Cot 7 = 0
                .Col = .ColLetterToNumber("AD")
                .value = 0
                UpdateCell fps, .Col, .Row, .Text
                
                GTSua = Mid(strId, 231, 1)
    
                'Hien lai Thue suat
                .Col = .ColLetterToNumber("AA")
                If Trim(Mid(strId, 200, 10)) = "1101" Or GTSua = "1" Then
                    If Val(.value) = 0 Then
                        .Col = .ColLetterToNumber("AA")
                        .value = Trim(Mid(strId, 221, 10))
                        UpdateCell fps, .Col, .Row, .value
                        
                    End If
                    .Lock = False
                Else
                    .Col = .ColLetterToNumber("AA")
                    .value = Trim(Mid(strId, 221, 10))
                    If .value = vbNullString Then
                        .value = "0"
                    End If
                    UpdateCell fps, .Col, .Row, .value
                    .Lock = True
                End If
            End If
         End If

         If Col = .ColLetterToNumber("AD") Then
            If Val(.value) > 0 Then
                'Cot 5 = 0
                .Col = .ColLetterToNumber("V")
                .value = 0
                UpdateCell fps, .Col, .Row, .Text
    
                'Hien lai Thue suat an
                .Col = .ColLetterToNumber("AA")
                .value = 0
                UpdateCell fps, .Col, .Row, 0
                .Lock = True
            End If
         End If
        .GetText .ColLetterToNumber("F"), .Row, strLTN
        If Trim(Mid(strLTN, 221, 10)) <> "" Then
            .Col = .ColLetterToNumber("P")
            UpdateCell fps, .Col, .Row, .Text
            .Lock = False
        End If
        If mblKhoiDauToKhai = True Then
            .Row = Row
            .Col = .ColLetterToNumber("AH")
            UpdateCell fps, .ColLetterToNumber("AH"), Row, .value
            .Row = Row
            .Col = .ColLetterToNumber("AO")
            UpdateCell fps, .ColLetterToNumber("AO"), Row, .value
            
            Set txmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 40))

            Set txmlCellNode = txmlCellNode.parentNode
            Set txmlCellNode = txmlCellNode.parentNode
            Set txmlCellNode = txmlCellNode.parentNode.lastChild
            Set txmlCellNode = txmlCellNode.firstChild
            Set txmlCellNode = txmlCellNode.firstChild
            ParserCellID fps, GetAttribute(txmlCellNode, "CellID"), tCol, tRow
            .Row = tRow
            .Col = .ColLetterToNumber("AH")
            UpdateCell fps, .Col, .Row, .value
            .Col = .ColLetterToNumber("AK")
            UpdateCell fps, .Col, .Row, .value
            .Col = .ColLetterToNumber("AO")
            UpdateCell fps, .Col, .Row, .value
        End If
'KHBS
     If .ActiveSheet = .SheetCount - 1 Then
         Dim cValueOld As Variant
         Dim cellid As String
         Dim xmlCellNodeData As MSXML.IXMLDOMNode
        ' Dim i As Integer
             fps.EventEnabled(EventAllEvents) = False
              cCol = .Col
              cRow = .Row
              .Sheet = .SheetCount - 1
             .Col = .ColLetterToNumber("BD")
             .Row = .Row
             .GetText .ColLetterToNumber("BC"), .Row, varTemp

       If varTemp <> "" Then
          Dim arr() As String
            For i = 0 To UBound(arrDanhSachKHBS)
               If varTemp = Left(arrDanhSachKHBS(i), InStr(arrDanhSachKHBS(i), "~") - 1) Then
                   cellid = "AG_" & Right(arrDanhSachKHBS(i), Len(arrDanhSachKHBS(i)) - InStr(arrDanhSachKHBS(i), "~") - 2)
                  For Each xmlCellNodeData In TAX_Utilities_v1.DataKHBS.getElementsByTagName("Cell")
                       If GetAttribute(xmlCellNodeData, "CellID") = cellid Then
                           cValue = CDbl(cValue) + CDbl(GetAttribute(xmlCellNodeData, "Value"))
                            Exit For
                        End If
                   Next
               End If
            Next
            .Col = .ColLetterToNumber("BF")
            .Row = Row
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
            Set xmlNodeH15 = TAX_Utilities_v1.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
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

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim strId As Variant
Dim iCol As Long, iRow As Long
Dim intIndexCombo As Integer, i As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim itemp As Variant, temp As Variant
Dim cValue As Variant
With fps
   If .ActiveSheet = .SheetCount - 1 Then
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("BD") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("BC")
            .typeComboboxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("BE")
            UpdateCell fps, .Col, .Row, "8"
        End If
   
        Set xmlCellNode = TAX_Utilities_v1.Data(.SheetCount - 2).nodeFromID(GetCellID(fps, Col, Row))
        Set xmlCellNode = xmlCellNode.parentNode.parentNode.firstChild.firstChild
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        blGiaTriCu = False
        If Col = fps.ColLetterToNumber("BD") Then
            fps.Col = .ColLetterToNumber("BC")
            fps.Row = Row
            temp = fps.Text
            Set xmlCellsNode = xmlCellNode.parentNode
            i = 1
            .Col = .ColLetterToNumber("B")
            ParserCellID fps, "BD_9", iCol, iRow
            .Row = iRow
            Do
                 .Col = .ColLetterToNumber("BC")
                 itemp = .Text
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
                 .Col = .ColLetterToNumber("B")
                 .Row = i + iRow
                 i = i + 1
            Loop Until .Text = "bb"
       End If
     
    End If
    
End With
    
'    CellChange Col, Row
'    TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = True
End Sub


Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, iCol As Long, iRow As Long
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tCol As Long, tRow As Long
Dim strId As Variant, Col7 As Variant, intIndexCombo As Integer
'bien dung de luu section
Dim iSection, countRow As Integer
Dim NguoiKyCol As Long, NguoiKyRow As Long
    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
         If .ActiveSheet = 1 Then
                If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
                    'Cot 7 - Cap nhap lai tinh toan truoc khi an F5
                    If KeyCode = vbKeyF5 Then
                            If iCol = .ColLetterToNumber("G") Then
                                mblKeyF5 = True
                            End If
                            fps_Change iCol, iRow
                            .Row = iRow
                            .Col = .ColLetterToNumber("AH")
                            UpdateCell fps, .ColLetterToNumber("AH"), iRow, .value
                            .Row = iRow
                            .Col = .ColLetterToNumber("AO")
                            UpdateCell fps, .ColLetterToNumber("AO"), iRow, .value
                    End If
                    .Sheet = .ActiveSheet
                    mCurrentSheet = .ActiveSheet
                    'If mCurrentSheet = 16 Then
                    .Sheet = .ActiveSheet
                                
                    If iCol = .ColLetterToNumber("N") Then
                        Exit Sub
                    End If
                     .Sheet = 1
                    mCurrentSheet = 1
                     
                    Set xmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, iCol, iRow))
                 
                    Set xmlCellNode = xmlCellNode.parentNode
                    Set xmlCellNode = xmlCellNode.parentNode
                    If GetAttribute(xmlCellNode, "Dynamic") = "0" Then
                        Exit Sub
                    End If
                    
                    fps.EventEnabled(EventAllEvents) = True
                    i = 1
                    iSection = 1
                    .Col = .ColLetterToNumber("C")
                    .Row = 40
                    countRow = 40
                    Do
                         .Text = CStr(i)
                         .Col = .ColLetterToNumber("C")
                         .Row = countRow + i
                         i = i + 1
                         If iSection = 1 And .Text = "aa" Then
                            iSection = iSection + 1
                            .Row = .Row + 3
                            countRow = .Row
                            i = 1
                         End If
                    Loop Until .Text = "aa"
                    .SetActiveCell iCol, iRow
                    
                    If Val(GetAttribute(xmlCellNode, "MaxRows")) = xmlCellNode.childNodes.length Then
                        Exit Sub
                    End If
                    
                    .Row = iRow
                    .Col = iCol
                    
                     fps.EventEnabled(EventAllEvents) = False
                    If KeyCode = vbKeyF5 Then
                        .Col = .ColLetterToNumber("F")
                        .Row = .Row + 1
                        .TypeComboBoxList = strCombo 'CPab("Maple", 100) + Chr$(9) + CPab("Oak", 100) + Chr$(9) + CPab("Ash", 100)
                        'Combo Hien
                        .Col = .ColLetterToNumber("G")
                        '.Row = 13
                        .TypeComboBoxList = strComboHien '
                        .Col = .ColLetterToNumber("AA")
                        .value = 0
                        .Lock = True
                    End If
                    If KeyCode = vbKeyF6 Or KeyCode = vbKeyF5 Then
                        .Row = 10
                        .Col = .ColLetterToNumber("AO")
                        UpdateCell fps, .Col, .Row, .value
                        
                        Set txmlCellNode = TAX_Utilities_v1.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("G"), 40))
        
                        Set txmlCellNode = txmlCellNode.parentNode
                        Set txmlCellNode = txmlCellNode.parentNode
                        Set txmlCellNode = txmlCellNode.parentNode.lastChild
                        Set txmlCellNode = txmlCellNode.firstChild
                        Set txmlCellNode = txmlCellNode.firstChild
                        ParserCellID fps, GetAttribute(txmlCellNode, "CellID"), tCol, tRow
                        .Row = tRow
                        .Col = .ColLetterToNumber("AH")
                        UpdateCell fps, .Col, .Row, .value
                        .Col = .ColLetterToNumber("AK")
                        UpdateCell fps, .Col, .Row, .value
                        .Col = .ColLetterToNumber("AO")
                        UpdateCell fps, .Col, .Row, .value
                    End If
                
                     .SetActiveCell iCol, iRow
                
                End If
                fps.EventEnabled(EventAllEvents) = False
                
                If KeyCode = vbKeyDelete And iCol = .ColLetterToNumber("G") Then
                                  
                        .Row = iRow
                        .Col = iCol
                        .Text = ""
                        UpdateCell fps, .Col, .Row, .Text
                      
                        .Col = .ColLetterToNumber("F")
                        .Text = ""
                        UpdateCell fps, .Col, .Row, .Text
                         
                        'Chon lai combo cua colF
                        .Col = .ColLetterToNumber("E")
                        .Text = ""
                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("D")
                         .Text = ""
                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("P")
                         .Text = ""
                         UpdateCell fps, .Col, .Row, .Text
                         
                        .Col = .ColLetterToNumber("O")
                        .Text = ""
                        UpdateCell fps, .Col, .Row, .Text
                         
                        .GetText .ColLetterToNumber("AD"), .Row, Col7
                        .EventEnabled(EventChange) = True
                       
                        .Col = .ColLetterToNumber("AA")
                        .value = 0
                        .Col = .ColLetterToNumber("AA")
                        .Row = iRow
                        UpdateCell fps, .Col, .Row, 0
                End If
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
                         .Text = "8"
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 9
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 3
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("BE")
                         .Text = "8"
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
        
        fps.EventEnabled(EventAllEvents) = True
    End With
    
End Sub
Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim CellNguoiKy As String
    
    CellNguoiKy = GetAttribute(TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")(TAX_Utilities_v1.Data(0).getElementsByTagName("Cell").length - 5), "CellID")
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("AG") And Row = .MaxRows - 7 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                        .TypeHAlign = TypeHAlignLeft
                    Else
'                        .SetFocus
                        .SetActiveCell Col, Row
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
                      
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
   
    Dim vErrMsgAJ13 As Variant, vErrMsgG13 As Variant
    
    Dim ColAJ13 As Long, ColAG13 As Long, ColS13 As Long, ColG13 As Long, ColV13 As Long, ColAD13 As Long
    Dim ColAA13 As Long
    Dim ColLoiAJ13 As Long, ColLoiAG13 As Long, ColLoiG13 As Long
    
    Dim RowAJ13 As Long, RowAG13 As Long, RowG13 As Long
    Dim RowLoiAJ13 As Long, RowLoiAG13 As Long, RowLoiG13 As Long
    
    Dim GT8 As Variant, GT9 As Variant, GT2 As Variant, GT4 As Variant, GT5 As Variant
    Dim GT7 As Variant, GT6 As Variant, GT10 As Variant
    
    Dim XacDinhLoiAJ13 As Boolean
    Dim XacDinhLoiG13 As Boolean

    Dim i, iSection As Integer, iCurrentSheet As Integer
    
    Dim LoiTonTai As Boolean
    Dim errorName, errorDate As Variant
    Dim errLoiCT As Variant, LoiCT As Boolean
    'PL01_1
    Dim blCheck_S2 As Boolean
    Dim errCTieu, errMST As Variant
    Dim iWhiteRow As Integer
    Dim vMST, vCTieu, vSTT, vTyLe, vThueGTGT, vMSTtemp As Variant
    Dim iFlagCtieu As Integer, iFlagMST As Integer
    Dim strCheck As String
    
    With fps
    '.Visible = False
    '.ReDraw = False
    iCurrentSheet = .Sheet
    .Sheet = .SheetCount

    CheckErrorMST
    .Sheet = iCurrentSheet

            .EventEnabled(EventAllEvents) = False
            XacDinhLoiAJ13 = False
            XacDinhLoiG13 = False
            LoiCT = False
            
            ColAA13 = .ColLetterToNumber("AA")
            ColAG13 = .ColLetterToNumber("AH")
            ColG13 = .ColLetterToNumber("G")
            ColS13 = .ColLetterToNumber("S")
            ColV13 = .ColLetterToNumber("V")
            ColAD13 = .ColLetterToNumber("AD")
                       
            .Sheet = .SheetCount
            .GetText .ColLetterToNumber("E"), 18, errorName
            .GetText .ColLetterToNumber("E"), 19, errLoiCT
            .GetText .ColLetterToNumber("E"), 17, vErrMsgG13
            .GetText .ColLetterToNumber("E"), 20, errCTieu
            .GetText .ColLetterToNumber("E"), 21, errMST
        
           
            mCurrentSheet = 1
            .Sheet = 1
            .Row = 40
            i = 40
            iSection = 1
            
            LoiTonTai = False
            
            Do
                .GetText ColAG13, .Row, GT8
                .GetText ColG13, .Row, GT2
                .GetText ColS13, .Row, GT4
                .GetText ColV13, .Row, GT5
                .GetText ColAD13, .Row, GT7
                .GetText ColAA13, .Row, GT6
                .GetText .ColLetterToNumber("AK"), .Row, GT9
                .GetText .ColLetterToNumber("AO"), .Row, GT10
                 
                 If GT2 = "" And (GT4 > 0 Or GT5 > 0 Or GT7 > 0 Or GT6 > 0) Then
                        .Col = ColG13
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgG13 'static
                        .BackColor = mErrorColor
                        If XacDinhLoiG13 = False And LoiTonTai = False Then
                            ColLoiG13 = ColG13
                            RowLoiG13 = .Row
                            LoiTonTai = True
                            XacDinhLoiG13 = True
                        End If
                 Else
                         .Col = ColG13
                         .CellNote = ""
                         .BackColor = mNonErrorColor
                 End If
                 
                 If Val(GT9) > Val(GT8) Then
                    .Col = .ColLetterToNumber("AK")
                    .BackColor = mErrorColor
                    .CellNote = errLoiCT
                    LoiCT = True
                 Else
                    .Col = .ColLetterToNumber("AK")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                 End If
                 
                 .Row = i + 1
                 .Col = .ColLetterToNumber("C")
                 i = i + 1
                 If iSection = 1 And .Text = "aa" Then
                    iSection = iSection + 1
                    .Row = .Row + 3
                    i = .Row
                 End If
            Loop Until .Text = "aa"
'check error date and name
'        .Col = .ColLetterToNumber("AF")
'        .Row = 15 + i
'        If .Value = "" Then
'            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'            .CellNote = errorName
'            .BackColor = mErrorColor
'            LoiTonTai = True
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 18, "0"
'            .Sheet = 1
'        Else
'            .CellNote = ""
'            .BackColor = mNonErrorColor
'            LoiTonTai = False
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 18, "1"
'            .Sheet = 1
'        End If
'        .Row = 17 + i
'        If .Value = "" Then
'            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'            .CellNote = errorDate
'            .BackColor = mErrorColor
'            LoiTonTai = True
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 19, "0"
'            .Sheet = 1
'        Else
'            .CellNote = ""
'            .BackColor = mNonErrorColor
'            LoiTonTai = False
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 19, "1"
'            .Sheet = 1
'        End If
            Dim strFormula As String
        
        If XacDinhLoiG13 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 17, IIf(XacDinhLoiG13, "0", "1")
            
            .Sheet = 1
            strFormula = "'" & .SheetName & "'!" & Trim(.ColNumberToLetter(CLng(ColLoiG13))) & Trim(str(RowLoiG13))
            .Sheet = .SheetCount
            .Col = .ColLetterToNumber("C")
            .Row = 17
            .Formula = strFormula
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 17, "1"
        End If
        
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 19, IIf(LoiCT = True, "0", "1")
        
'        .Sheet = 2
'         blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'         If blCheck_S2 = True Then
'
'                .Col = .ColLetterToNumber("B")
'                .Row = 41
'                Do
'                      iWhiteRow = 1
'                      .GetText .ColLetterToNumber("B"), .Row, vSTT
'                      .GetText .ColLetterToNumber("C"), .Row, vCTieu
'                      .GetText .ColLetterToNumber("L"), .Row, vMST
'                      .GetText .ColLetterToNumber("R"), .Row, vTyLe
'                      .GetText .ColLetterToNumber("U"), .Row, vThueGTGT
'                      If Trim(vSTT) <> "" Or Trim(vCTieu) <> "" Or Trim(vMST) <> "" Or Val(Trim(vTyLe)) <> 0 Or Trim(vThueGTGT) <> "0" Then
'                        iWhiteRow = 0
'                      End If
'                      ' Reset lai cac cellnote
'                      ' Ctieu
'                       .Col = .ColLetterToNumber("C")
'                        .CellNote = ""
'                        .BackColor = mNonErrorColor
'                      ' MST
'                       .Col = .ColLetterToNumber("L")
'                        .CellNote = ""
'                        .BackColor = mNonErrorColor
'                      ' End reset
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
'                                If Len(vMST) = Len(vMSTtemp) Or Mid$(vMSTtemp, 11, 1) = "-" Then
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
'
'                     .Col = .ColLetterToNumber("B")
'                     .Row = .Row + 1
'                Loop Until .Text = "aa"
'                .Sheet = .SheetCount
'                .SetText .ColLetterToNumber("B"), 20, IIf(iFlagCtieu = 1, "0", "1")
'                .SetText .ColLetterToNumber("B"), 21, IIf(iFlagMST = 1, "0", "1")
'        End If
        
        
        
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
             Case "01_1"
                blnNullValue = True
                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
                (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 5)
                strValue1 = GetAttribute(xmlCellNode, "Value")
                If (Not IsNullValue(strValue1)) Then
                     blnNullValue = False
                End If
                
                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If
                
             Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 1)
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

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()
Dim i As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = 1
        i = 1
        .SetText .ColLetterToNumber("C"), 39, 0
        .Col = .ColLetterToNumber("C")
        .Row = 40
        .Formula = "C39+1"
        .SetText .ColLetterToNumber("C"), 43, 0
        .Col = .ColLetterToNumber("C")
        .Row = 44
        .Formula = "C43+1"
        .EventEnabled(EventAllEvents) = True
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
                    strComboName = strComboName + arrDanhsach(2) + Chr$(9)
                    strComboMCT = strComboMCT + "8" + Chr$(9)
                End If
            End If
        Next
    End If
    
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
        .Text = "8"
        .Lock = True
        .Row = 13
        .TypeComboBoxList = strComboMCT
        .Text = "8"
        .Lock = True
'Cho phep nhap am cot 4, 5, 7 to khai bo sung
        .Sheet = 1
        .Row = 13
        .Col = .ColLetterToNumber("S")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("V")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("AD")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        
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
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateLastKHBSCell fps, lCol, lRow, varTemp
            Next
       Else
            For Each xmlCell In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
               If Left(GetAttribute(xmlCell, "CellID"), 2) = "BC" Then
                    strId = Trim(Mid(GetAttribute(xmlCell, "Value"), 200, 10))
                        
                    For Each xmlCellGoc In TAX_Utilities_v1.DataKHBS.getElementsByTagName("Cell")
                        If Left(GetAttribute(xmlCellGoc, "CellID"), 1) = "E" Then
                            If GetAttribute(xmlCellGoc, "Value") = strId Then
                            'update secsion
                               Set xmlCellUpdate = seachCellNode(TAX_Utilities_v1.Data(0).getElementsByTagName("Cell"), GetAttribute(xmlCellGoc, "CellID"))
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
     If Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "S" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "V" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 2) = "AD" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "AH" Then
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
        ReDim arrDanhSachKHBS(TAX_Utilities_v1.Data(0).getElementsByTagName("Section")(1).childNodes.length - 1)
        i = 0
        For Each xmlCellNodeData In TAX_Utilities_v1.DataKHBS.getElementsByTagName("Cell")
           If Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "F" Then
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

'1
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
            Set xmlNode = .nodeFromID("AO_17")
            result = result & "CT10_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_02TAIN", "getValueTK", Err.number, Err.Description
End Function

'2
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
                .GetText .ColLetterToNumber("AO"), 17, varTemp
                strCurrentValue = "CT10_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("AO"), 17, varTemp
'                strCurrentValue = strCurrentValue & "~CT10_" & Trim(varTemp)
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
                'If Left(strMaCT, 3) = "21_" Or Left(strMaCT, 3) = "23_" Or Left(strMaCT, 3) = "28a" Or Left(strMaCT, 3) = "30_" Or Left(strMaCT, 3) = "31_" Then
                If Left(strMaCT, 3) = "8_" Or Left(strMaCT, 3) = "10_" Then
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

'3
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
            Set xmlNode = .nodeFromID("AO_17")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_02TAIN", "getValueCTDC", Err.number, Err.Description
End Function

'4
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("AO"), 17, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function

'5
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
    SaveErrorLog "02_TAIN", "getMapCT", Err.number, Err.Description
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
                    Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
                        (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                    Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                    Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length > 20 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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


' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 11
    denRowDL = 18
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "20_30_2"
    ' end
End Sub


'
Public Sub getNhomHH(ByVal maNhom As String, ByRef tenNhom As String)
    Dim arrDanhsach() As String
    Dim strDataFileName As String
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim blChon As Boolean

        blChon = False
        If TAX_Utilities_v1.Year = "2008" Then
            strDataFileName = Replace(GetCatalogueTemplateFile, "Catalogue_052008_01", "CatalogueQTTN_2008_01")
        ElseIf TAX_Utilities_v1.Year = "2010" Then
            strDataFileName = Replace(GetCatalogueTemplateFile, "Catalogue_072010_01", "CatalogueQTTN_2010_01")
        ElseIf TAX_Utilities_v1.Year = "2014" Then
            strDataFileName = Replace(GetCatalogueTemplateFile, "Catalogue_022014_01", "CatalogueQTTN_2014_01")
        Else
            strDataFileName = GetCatalogueFileName(blChon)
        End If
    
       If xmlDomData.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If maNhom = arrDanhsach(1) Then
                            tenNhom = arrDanhsach(2)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub


Public Sub getRowDanhMuc(ByVal maNhom As String, ByRef tenNhom As String)
    Dim arrDanhsach() As String
    Dim strDataFileName As String
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim blChon As Boolean
        blChon = False
        If TAX_Utilities_v1.Year = "2008" Then
            strDataFileName = Replace(GetCatalogueTemplateFile, "Catalogue_052008_01", "CatalogueQTTN_2008_01")
        ElseIf TAX_Utilities_v1.Year = "2010" Then
            strDataFileName = Replace(GetCatalogueTemplateFile, "Catalogue_072010_01", "CatalogueQTTN_2010_01")
        ElseIf TAX_Utilities_v1.Year = "2014" Then
            strDataFileName = Replace(GetCatalogueTemplateFile, "Catalogue_022014_01", "CatalogueQTTN_2014_01")
        Else
            strDataFileName = GetCatalogueFileName(blChon)
        End If
    
       If xmlDomData.Load(GetAbsolutePath(strDataFileName)) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        If maNhom = arrDanhsach(1) Then
                            tenNhom = CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 10) + CPab(arrDanhsach(5), 1)
                            Exit Sub
                        End If
                End If
            Next
        End If
End Sub


' ham set thong tin sau khi tai XML
Public Sub finishImportXML()
    Dim maNhom As Variant
    Dim tenNhom As String
    Dim tempValue As Variant
    Dim iSection  As Long
     With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
         mCurrentSheet = .Sheet
        .Row = 40
        .Col = .ColLetterToNumber("C")
        iSection = 1
        Do
           
            .GetText .ColLetterToNumber("E"), .Row, maNhom
            getRowDanhMuc maNhom, tenNhom
            
            .Col = .ColLetterToNumber("F")
            .Text = Trim(tenNhom)
            UpdateCell fps, .Col, .Row, .Text
            
            If tenNhom <> "" Then
                .Col = .ColLetterToNumber("O")
                .Text = Trim(tenNhom)
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            ' lay ten danh tai nguyen
            .GetText .ColLetterToNumber("G"), .Row, tempValue
            .Col = .ColLetterToNumber("D")
            .Text = tempValue
            
            
            .Row = .Row + 1
            .Col = .ColLetterToNumber("C")
             
            If iSection = 1 And .Text = "aa" Then
                iSection = iSection + 1
                .Row = .Row + 3
            End If
        Loop Until .Text = "aa"
    
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
