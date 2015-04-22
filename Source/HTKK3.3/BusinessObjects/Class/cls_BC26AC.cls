VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_BC26AC"
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
Private totalRow2ABH As Long
Public numRowInsert As Integer
'dhdang sua check ton dau ky
'ngay 18-05
Dim result_ton() As String
Dim arrCellInRow() As String
Dim strTemp_ton As String

Private Const Col_D = "D"
Private Const Col_E = "E"
Private Const Col_F = "F"
Private Const Col_G = "G"
Private Const Col_H = "H"
Private Const Col_I = "I"
Private Const Col_J = "J"
Private Const Col_K = "K"
Private Const Col_L = "L"
Private Const Col_M = "M"
Private Const Col_N = "N"
Private Const Col_O = "O"
Private Const Col_Q = "Q"
Private Const Col_R = "R"
Private Const Col_S = "S"
Private Const Col_T = "T"
Private Const Col_U = "U"
Private Const Col_V = "V"
Private Const Col_W = "W"
Private Const Col_Y = "Y"
Private Const Col_Z = "Z"
Private Const Col_AA = "AA"
Private Const Col_AB = "AB"
Private Const Col_AC = "AC"
Private Const Col_AD = "AD"

Public strTkThangQuy As String

'Private strCheckTon As String
'Private strCheckTonKhoang As String
Private loaiHoaDonS1() As LOAI_HOA_DON

'Dinh nghia type luu du  HD
Public Type HOA_DON
    tenHD As String
    mauSoHD As String
    kyHieuHD As String
    tuSo As Double
    denso As Double
    isDelete As Boolean
End Type

'Dinh nghia type luu du lieu bang HD
Public Type LOAI_HOA_DON
    loaiHD As String
    hoaDon() As HOA_DON
End Type
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    With fps

        Dim xmlDocument As New MSXML.DOMDocument
        Dim xmlNode As MSXML.IXMLDOMNode
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_loai_HD.xml"))
        Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
        Dim strQuanHeNNTID As String
        Dim strQuanHeNNT As String
        Dim strDataFileName As String
        Dim arrDanhsach() As String
        Dim strComboHien As String
        Dim strCombo As String
        Dim strCombo1 As String
        Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
        Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
       strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_loai_HD.xml")
        ' Lay danh muc loai hoa don
        ' 15/11/2010
        If xmlDomData.Load(strDataFileName) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    strCombo = strCombo + CPab(arrDanhsach(0), 10) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(2), 200) + Chr$(9)
'                    strCombo = strCombo + CPab(arrDanhsach(1), 10) + Chr$(9)
                    strComboHien = strComboHien + arrDanhsach(1) + Chr$(9)
                End If
            Next
        End If
       .Sheet = 1
'        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
'        For Each xmlNode In xmlNodeListMap
'            strQuanHeNNT = strQuanHeNNT + GetAttribute(xmlNode, "Value") + Chr$(9)
'        Next
'        strQuanHeNNTID = "01GTKT" + Chr$(9) + "02GTTT" + Chr$(9) + "06HDXK" + Chr$(9) + "07KPTQ" + Chr$(9) + "03XKNB" + Chr$(9) + "04HGDL" + Chr$(9) + "01/" + Chr$(9) + "02/" + Chr$(9) + "01/" + Chr$(9) + "02/" + Chr$(9) + "01/" + Chr$(9) + "02/" + Chr$(9)
        .Row = 22
        .Col = .ColLetterToNumber("C")
        
        .TypeComboBoxList = strComboHien
        .Text = strQuanHeNNT
        .Col = .ColLetterToNumber(Col_AA)
        .TypeComboBoxList = strCombo
        'SetDateFormat fps, 1, 22, .ColLetterToNumber(Col_G), DDMMYYYY
        ' SetKyHieuHDFormat fps, 1, 22, .ColLetterToNumber(Col_F)
         
        SetDateFormat fps, 1, 32, .ColLetterToNumber(Col_N), DDMMYYYY
        .Row = 32
        .Col = .ColLetterToNumber(Col_N)
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
       
   ' vttoan
   ' sua theo yeu cau: O chuyen dia diem thanh combobox
        Dim tenCucThue As Variant
        Dim maCucThue As Variant
        Dim tenChiCucThue As Variant
        Dim maChiCucThue As Variant
        If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
                For Each xmlNode In xmlNodeListCell
                    If GetAttribute(xmlNode, "Value") <> "" Then
                        arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                        
                        If arrDanhsach(0) = "1" Then
                            tenCucThue = tenCucThue + arrDanhsach(3) + "{" + arrDanhsach(1) + "}" + Chr$(9)
                            maCucThue = maCucThue + arrDanhsach(1) + Chr$(9)
                        End If
                        'If arrDanhsach(0) = "0" Then
                            tenChiCucThue = tenChiCucThue + arrDanhsach(3) + "                                                                                                                                                             {" + arrDanhsach(1) + "}" + Chr$(9)
                            maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
                        'End If
                    End If
                Next
            Set xmlDomData = Nothing
            Set xmlNodeListCell = Nothing
            Set xmlNode = Nothing
        End If
    .Sheet = 2
            .Row = 54
            .Col = .ColLetterToNumber("M")
            .TypeComboBoxList = tenChiCucThue
            'SetKyHieuHDFormat fps, 2, 46, .ColLetterToNumber("S")
            'SetKyHieuHDFormat fps, 2, 50, .ColLetterToNumber("S")
            SetDateFormat fps, 2, 58, .ColLetterToNumber(Col_R), DDMMYYYY
            .Row = 58
            .Col = .ColLetterToNumber(Col_R)
            .Text = Format(Date, "dd/mm/yyyy")
            .TypeHAlign = TypeHAlignLeft
            ' Set gia tri ten loai hoa don
            .Col = .ColLetterToNumber("C")
            .Row = 46
            .TypeComboBoxList = strComboHien

            ' Set gia tri cho combo an
            .Col = .ColLetterToNumber("AI")
            .Row = 46
            .TypeComboBoxList = strCombo

            .Col = .ColLetterToNumber("C")
            .Row = 50
            .TypeComboBoxList = strComboHien

            ' Set gia tri cho combo an
            .Col = .ColLetterToNumber("AI")
            .Row = 50
            .TypeComboBoxList = strCombo

        .Sheet = 3
            'SetKyHieuHDFormat fps, 3, 45, .ColLetterToNumber("S")
            SetDateFormat fps, 3, 51, .ColLetterToNumber(Col_R), DDMMYYYY
            .Row = 51
            .Col = .ColLetterToNumber(Col_R)
            .Text = Format(Date, "dd/mm/yyyy")
            .TypeHAlign = TypeHAlignLeft
            ' Set gia tri ten loai hoa don
            .Col = .ColLetterToNumber("C")
            .Row = 45
            .TypeComboBoxList = strComboHien

            ' Set gia tri cho combo an
            .Col = .ColLetterToNumber("AI")
            .Row = 45
            .TypeComboBoxList = strCombo
            
    End With
End Sub


'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim i As Integer, j As Integer
'    Dim result() As String
'    Dim arrCellInRow() As String
'    Dim strTemp As String
    result_ton = loadTonCuoiKy
    
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim iCol As Long, iRow As Long
    
    Dim tmpSheet As Integer
    
    With fps
        .Sheet = 1
        .EventEnabled(EventAllEvents) = False
        'dhdang sua loi o check "ky bao cao cc" va "chuyen dia diem"
        If TAX_Utilities_v2.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0" Then
                .Row = 14
                .Col = .ColLetterToNumber("B")
                .Text = "X"
                .Lock = True
                
                .Row = 14
                .Col = .ColLetterToNumber("G")
                .Text = ""
                .Lock = True
            ElseIf TAX_Utilities_v2.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0" Then
                .Row = 14
                .Col = .ColLetterToNumber("B")
                .Text = ""
                .Lock = True
                
                .Row = 14
                .Col = .ColLetterToNumber("G")
                .Text = "X"
                .Lock = True
            Else
                .Row = 14
                .Col = .ColLetterToNumber("B")
                .Text = ""
                .Lock = True
                
                .Row = 14
                .Col = .ColLetterToNumber("G")
                .Text = ""
                .Lock = True
        End If
        ' Set format ngay ky
        If numRowInsert >= 0 And TAX_Utilities_v2.CheckNewDataBC26 = True Then
            .Row = 22
            For i = 0 To numRowInsert
                ' set du lieu ton dau ky
                strTemp_ton = result_ton(i)
                mCurrentSheet = 1
                arrCellInRow = Split(strTemp_ton, "~")
                'Mau so
                .Col = .ColLetterToNumber(Col_E)
                .SetText .Col, .Row, arrCellInRow(0)
                UpdateCell fps, .Col, .Row, .Text
                ' Ky hieu
                .Col = .ColLetterToNumber(Col_F)
                .SetText .Col, .Row, arrCellInRow(1)
                UpdateCell fps, .Col, .Row, .Text
                ' tu so
                .Col = .ColLetterToNumber(Col_H)
                .SetText .Col, .Row, arrCellInRow(2)
                UpdateCell fps, .Col, .Row, .Text
                ' den so
                .Col = .ColLetterToNumber(Col_I)
                .SetText .Col, .Row, arrCellInRow(3)
                UpdateCell fps, .Col, .Row, .Text
                ' Tong so
                .Col = .ColLetterToNumber(Col_G)
                .SetText .Col, .Row, Val(arrCellInRow(3)) - Val(arrCellInRow(2)) + 1
                UpdateCell fps, .Col, .Row, .Text
                ' strCheck
                .Col = .ColLetterToNumber(Col_AA)
                .SetText .Col, .Row, arrCellInRow(4)
                UpdateCell fps, .Col, .Row, .Text
                ' ten HD
                .Col = .ColLetterToNumber(Col_D)
                .SetText .Col, .Row, arrCellInRow(5)
                UpdateCell fps, .Col, .Row, .Text
                'Ma HD
                .Col = .ColLetterToNumber("C")
                .SetText .Col, .Row, arrCellInRow(6)
                UpdateCell fps, .Col, .Row, .Text
                'luu trang thai ton
                .Col = .ColLetterToNumber("AC")
                .SetText .Col, .Row, "1"
                UpdateCell fps, .Col, .Row, .Text
                
                .Row = .Row + 1
            Next i
        End If

        
        
        i = 1
        .Col = .ColLetterToNumber("B")
        .Row = 22
        Do
'             .Col = .ColLetterToNumber("C")
'             .Text = "0"
             
             .Col = .ColLetterToNumber("B")
             .Row = i + 22
             totalRow2ABH = i
             i = i + 1
        Loop Until UCase(.Text) = "AA"
        
        'dhdang sua theo yc tten ngay 25/01
        .Row = totalRow2ABH + 29
        .Col = .ColLetterToNumber(Col_N)
        If Trim(.Text) = "" Then
            .SetText .ColLetterToNumber(Col_N), totalRow2ABH + 29, GetNguoiKy()
        End If
        
         'Xoa cong thuc cua o Chu ky
        Set xmlCellNode = TAX_Utilities_v2.Data(0).nodeFromID(GetCellID(fps, .ColLetterToNumber("D"), 22))
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode.lastChild
        Set xmlCellNode = xmlCellNode.firstChild
        Set xmlCellNode = xmlCellNode.firstChild
        Set xmlCellNode = xmlCellNode.nextSibling
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        .Row = iRow
        .Col = iCol
        .Formula = ""
        
        tmpSheet = mCurrentSheet
        mCurrentSheet = 1
        .Sheet = 1
        ' set loai tk thang/ quy
        If strTkThangQuy = "TK_THANG" Then
            .Col = .ColLetterToNumber("H")
            .Row = 13
            .Text = "0"
            UpdateCell fps, .Col, .Row, .Text
        ElseIf strTkThangQuy = "TK_QUY" Then
            .Col = .ColLetterToNumber("H")
            .Row = 13
            .Text = "1"
            UpdateCell fps, .Col, .Row, .Text
        End If
            'end
            
            
            
        .Sheet = 1
        .ActiveSheet = .Sheet
        .Row = 4
        .Col = .ColLetterToNumber("D")
        .CellType = CellTypeButton
        .TypeButtonTextColor = RGB(0, 0, 255)
        .TypeButtonColor = &H8000000F
        .TypeButtonAlign = TypeButtonAlignRight
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
            
            
            
        .Sheet = 2
        .ActiveSheet = .Sheet
        .Row = 2
        .Col = .ColLetterToNumber("B")
        .CellType = CellTypeButton
        .TypeButtonTextColor = RGB(0, 0, 255)
        .TypeButtonColor = &H8000000F
        .TypeButtonAlign = TypeButtonAlignRight
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
        
        
        .Sheet = 3
        .ActiveSheet = .Sheet
        .Row = 2
        .Col = .ColLetterToNumber("B")
        .CellType = CellTypeButton
        .TypeButtonTextColor = RGB(0, 0, 255)
        .TypeButtonColor = &H8000000F
        .TypeButtonAlign = TypeButtonAlignRight
        .Font = "Tahoma"
        .FontSize = 8
        .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
            
            
        mCurrentSheet = tmpSheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Function GetValidatedCellIndex(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long) As Integer
    
    
End Function

Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    fps.EventEnabled(EventAllEvents) = False
    If (Col = fps.ColLetterToNumber("B") And Row = 14) Or (Col = fps.ColLetterToNumber(Col_G) And Row = 14) Then
           If fps.CellType = CellTypeCheckBox Then
                fps.Sheet = fps.ActiveSheet
                mCurrentSheet = fps.Sheet
                fps.Row = Row
                fps.Col = Col
                UpdateCell fps, fps.Col, fps.Row, IIf(fps.value = 1, "x", vbNullString) 'fps.Value
        End If
    End If
    fps.EventEnabled(EventAllEvents) = True
    CellChange Col, Row
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
     Dim strId As Variant
    Dim intIndexCombo As Integer

    fps.EventEnabled(EventAllEvents) = False
    With fps
    If .ActiveSheet = 1 Then
        
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("C") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber(Col_AA)
            .typeComboboxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber(Col_AA)
            UpdateCell_sheet fps, 1, .Col, .Row, .Text
             
            
            .GetText .ColLetterToNumber(Col_AA), .Row, strId
            
            
            
            .Col = .ColLetterToNumber(Col_AA)
            If Trim(Left(strId, 1)) = "0" Then
                .SetText .ColLetterToNumber(Col_E), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_E), .Row, .Text
                .Col = .ColLetterToNumber(Col_E)
                .TypeMaxEditLen = 13
                .Col = .ColLetterToNumber(Col_D)
                .SetText .ColLetterToNumber(Col_D), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_D), .Row, .Text
            ElseIf Trim(Left(strId, 1)) = "3" Then
                .SetText .ColLetterToNumber(Col_E), .Row, ""
                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_E), .Row, .Text
                .Col = .ColLetterToNumber(Col_E)
                .TypeMaxEditLen = 20
                
                .Col = .ColLetterToNumber(Col_D)
                .SetText .ColLetterToNumber(Col_D), .Row, ""
                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_D), .Row, .Text
            ElseIf Trim(Left(strId, 1)) = "4" Then
                .SetText .ColLetterToNumber(Col_E), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_E), .Row, .Text
                .Col = .ColLetterToNumber(Col_E)
                .TypeMaxEditLen = 13
                .Col = .ColLetterToNumber(Col_D)
                .SetText .ColLetterToNumber(Col_D), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_D), .Row, .Text
            Else
                .SetText .ColLetterToNumber(Col_E), .Row, Trim(Mid(strId, 11, 3))
                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_E), .Row, .Text
                .Col = .ColLetterToNumber(Col_E)
                .TypeMaxEditLen = 20
                .Col = .ColLetterToNumber(Col_D)
                .SetText .ColLetterToNumber(Col_D), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_D), .Row, .Text
            End If
            .SetText .ColLetterToNumber(Col_AC), .Row, Trim(Left(strId, 1))
            UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_AC), .Row, .Text
            
            .EventEnabled(EventChange) = True
         End If
    ElseIf .ActiveSheet = 2 Then
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("C") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("AI")
            .typeComboboxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber("AI")
            UpdateCell_sheet fps, 2, .Col, .Row, .Text
             
            
            .GetText .ColLetterToNumber("AI"), .Row, strId

            
            .Col = .ColLetterToNumber(Col_O)
            If Trim(Left(strId, 1)) = "0" Then
                .SetText .ColLetterToNumber(Col_O), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell_sheet fps, 2, .ColLetterToNumber(Col_O), .Row, .Text
                .Col = .ColLetterToNumber(Col_O)
                .TypeMaxEditLen = 13
                .Col = .ColLetterToNumber(Col_D)
                .SetText .ColLetterToNumber(Col_D), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 2, .ColLetterToNumber(Col_D), .Row, .Text
            ElseIf Trim(Left(strId, 1)) = "3" Then
                .SetText .ColLetterToNumber(Col_O), .Row, ""
                UpdateCell_sheet fps, 2, .ColLetterToNumber(Col_O), .Row, .Text
                .Col = .ColLetterToNumber(Col_O)
                .TypeMaxEditLen = 20
                .Col = .ColLetterToNumber(Col_D)
                .SetText .ColLetterToNumber(Col_D), .Row, ""
                UpdateCell_sheet fps, 2, .ColLetterToNumber(Col_D), .Row, .Text
            ElseIf Trim(Left(strId, 1)) = "4" Then
                .SetText .ColLetterToNumber(Col_O), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell_sheet fps, 2, .ColLetterToNumber(Col_O), .Row, .Text
                .Col = .ColLetterToNumber(Col_O)
                .TypeMaxEditLen = 13
                .Col = .ColLetterToNumber(Col_D)
                .SetText .ColLetterToNumber(Col_D), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 2, .ColLetterToNumber(Col_D), .Row, .Text
            Else
                .SetText .ColLetterToNumber(Col_O), .Row, Trim(Mid(strId, 11, 3))
                UpdateCell_sheet fps, 2, .ColLetterToNumber(Col_O), .Row, .Text
                .Col = .ColLetterToNumber(Col_O)
                .TypeMaxEditLen = 20
                .Col = .ColLetterToNumber(Col_D)
                .SetText .ColLetterToNumber(Col_D), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 2, .ColLetterToNumber(Col_D), .Row, .Text
            End If
            .EventEnabled(EventChange) = True
         End If
    ElseIf .ActiveSheet = 3 Then
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("C") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("AI")
            .typeComboboxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber("AI")
            UpdateCell_sheet fps, 3, .Col, .Row, .Text
             
            
            .GetText .ColLetterToNumber("AI"), .Row, strId
            
            
            
            .Col = .ColLetterToNumber(Col_O)
            If Trim(Left(strId, 1)) = "0" Then
                .SetText .ColLetterToNumber(Col_O), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell_sheet fps, 3, .ColLetterToNumber(Col_O), .Row, .Text
                .Col = .ColLetterToNumber(Col_O)
                .TypeMaxEditLen = 13
                .Col = .ColLetterToNumber(Col_D)
                .SetText .ColLetterToNumber(Col_D), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 3, .ColLetterToNumber(Col_D), .Row, .Text
            ElseIf Trim(Left(strId, 1)) = "3" Then
                .SetText .ColLetterToNumber(Col_O), .Row, ""
                UpdateCell_sheet fps, 3, .ColLetterToNumber(Col_O), .Row, .Text
                .Col = .ColLetterToNumber(Col_O)
                .TypeMaxEditLen = 20
                .Col = .ColLetterToNumber(Col_D)
                .SetText .ColLetterToNumber(Col_D), .Row, ""
                UpdateCell_sheet fps, 3, .ColLetterToNumber(Col_D), .Row, .Text
            ElseIf Trim(Left(strId, 1)) = "4" Then
                .SetText .ColLetterToNumber(Col_O), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell_sheet fps, 3, .ColLetterToNumber(Col_O), .Row, .Text
                .Col = .ColLetterToNumber(Col_O)
                .TypeMaxEditLen = 13
                .Col = .ColLetterToNumber(Col_D)
                .SetText .ColLetterToNumber(Col_D), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 3, .ColLetterToNumber(Col_D), .Row, .Text
            Else
                .SetText .ColLetterToNumber(Col_O), .Row, Trim(Mid(strId, 11, 3))
                UpdateCell_sheet fps, 3, .ColLetterToNumber(Col_O), .Row, .Text
                .Col = .ColLetterToNumber(Col_O)
                .TypeMaxEditLen = 20
                .Col = .ColLetterToNumber(Col_D)
                .SetText .ColLetterToNumber(Col_D), .Row, Trim(Mid(strId, 21))
                UpdateCell_sheet fps, 3, .ColLetterToNumber(Col_D), .Row, .Text
            End If
            .EventEnabled(EventChange) = True
         End If
    End If
    End With
   fps.EventEnabled(EventAllEvents) = True

End Sub

Public Sub kiemTraDuLieuImport()
   
End Sub


Private Sub fps_KeyPress(KeyAscii As Integer)
    With fps
        If .ActiveSheet = 1 Then
            If .Col = .ColLetterToNumber(Col_H) Or .Col = .ColLetterToNumber(Col_I) Or .Col = .ColLetterToNumber(Col_J) Or .Col = .ColLetterToNumber(Col_K) Or .Col = .ColLetterToNumber(Col_L) Or .Col = .ColLetterToNumber(Col_M) Or .Col = .ColLetterToNumber("X") Or .Col = .ColLetterToNumber(Col_Y) Then
                If (KeyAscii < 48 And KeyAscii <> 3 And KeyAscii <> 22 And KeyAscii <> 13 And KeyAscii <> 8) Or KeyAscii > 57 Then
                    KeyAscii = 0
                End If
            End If
            If .Col = .ColLetterToNumber(Col_R) Or .Col = .ColLetterToNumber(Col_T) Or .Col = .ColLetterToNumber(Col_V) Then
                If (KeyAscii < 48 And KeyAscii <> 3 And KeyAscii <> 22 And KeyAscii <> 13 And KeyAscii <> 45 And KeyAscii <> 8) Or (KeyAscii > 57 And KeyAscii <> 59) Then
                    KeyAscii = 0
                End If
            End If
        ElseIf .ActiveSheet = 2 Then
            If .Col = .ColLetterToNumber(Col_V) Or .Col = .ColLetterToNumber(Col_Y) Then
                If (KeyAscii < 48 And KeyAscii <> 3 And KeyAscii <> 22 And KeyAscii <> 13 And KeyAscii <> 8) Or KeyAscii > 57 Then
                    KeyAscii = 0
                End If
            End If
        ElseIf .ActiveSheet = 3 Then
            If .Col = .ColLetterToNumber(Col_V) Or .Col = .ColLetterToNumber(Col_Y) Then
                If (KeyAscii < 48 And KeyAscii <> 3 And KeyAscii <> 22 And KeyAscii <> 13 And KeyAscii <> 8) Or KeyAscii > 57 Then
                    KeyAscii = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, j As Integer, iCol As Long, iRow As Long
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tCol As Long, tRow As Long
    With fps
        .EventEnabled(EventAllEvents) = False
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 1 Then
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 22
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 22
                         totalRow2ABH = i
                         i = i + 1
                    Loop Until UCase(.Text) = "AA"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    
                    .SetActiveCell iCol, iRow
            ElseIf .ActiveSheet = 2 Then
                    Dim sec As Variant
                     i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 46
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 46
                         sec = i
                         i = i + 1
                    Loop Until UCase(.Text) = "AA"
                    
                    j = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 49 + sec
                    Do
                         .Text = str(j)
                         .Col = .ColLetterToNumber("B")
                         .Row = j + 49 + sec
                         'totalRow2ABH = i
                         j = j + 1
                    Loop Until .Text = "bb"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    
                    .SetActiveCell iCol, iRow
            ElseIf .ActiveSheet = 3 Then
                     i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 45
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 45
                         totalRow2ABH = i
                         i = i + 1
                    Loop Until UCase(.Text) = "AA"
                    
                    .SetActiveCell iCol, iRow
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
    Dim varMST As Variant, varMST5A As Variant, varMST5B As Variant
    Dim sumThuNhap5A As Variant, sumThuNhap5B As Variant
    Dim tuSo As Variant
    Dim denso As Variant
    Dim TongSo As Variant
    Dim strSoLuong As String
    Dim strTuso1 As Variant, strDenso1 As Variant, strTuso2 As Variant, strDenso2 As Variant, strCt10 As Variant, strCt11 As Variant, strCt12 As Variant, strCt5 As Variant, strCt20 As Variant, strCt21 As Variant
    Dim str() As String
    Dim soluongTonCK As Double
    
    Dim countRowDynamic As Variant
    Dim countRowDynamic1_PL1 As Variant
    Dim countRowDynamic2_PL1 As Variant
    Dim strId As Variant
    Dim strLoai As Variant
    Dim strTemp As Variant
    
    Dim vSoSD, vSoXoa, vSoMat, vSoHuy As Variant
    
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .ActiveSheet
        
            If .ActiveSheet = 1 Then
                    ' Chuyen ky hieu HD thanh chu cai in hoa
                    'dhdang sua fomat ky hieu hoa don
                    If Col = .ColLetterToNumber(Col_F) Then
                         .GetText Col, Row, varTemp
                         .SetText Col, Row, FormatKyHieu(CStr(varTemp))
                         .Col = Col
                         .Row = Row
                        UpdateCell fps, .Col, .Row, .Text
                    End If
                    ''''
                    If Col = .ColLetterToNumber(Col_R) Or Col = .ColLetterToNumber(Col_T) Or Col = .ColLetterToNumber(Col_T) Or Col = .ColLetterToNumber(Col_V) Or Col = .ColLetterToNumber("P") Then
                            .SetText .ColLetterToNumber("X"), Row, ""
                            .SetText .ColLetterToNumber(Col_Y), Row, ""
                            .SetText .ColLetterToNumber(Col_Z), Row, "0"
                    End If
                   
                   ' Format dinh dang tu so den so co 7 ky tu
                    If Col = .ColLetterToNumber(Col_L) Or Col = .ColLetterToNumber(Col_M) Or Col = .ColLetterToNumber(Col_H) Or Col = .ColLetterToNumber(Col_I) Or Col = .ColLetterToNumber(Col_J) Or Col = .ColLetterToNumber(Col_K) Or Col = .ColLetterToNumber("P") Or Col = .ColLetterToNumber(Col_R) Or Col = .ColLetterToNumber(Col_T) Or Col = .ColLetterToNumber(Col_V) Then
                        If Col = .ColLetterToNumber(Col_H) Or Col = .ColLetterToNumber(Col_I) Or Col = .ColLetterToNumber(Col_J) Or Col = .ColLetterToNumber(Col_K) Then
                             .GetText Col, Row, varTemp
                             .SetText Col, Row, FormatSoHD(CStr(varTemp))
                             UpdateCell fps, Col, Row, .Text
                        End If
                        ' Kiem tra neu so SD, xoa, mat, huy deu bang 0 thi set [10],[11] =trang
                        .GetText .ColLetterToNumber("P"), Row, vSoSD
                        .GetText .ColLetterToNumber("Q"), Row, vSoXoa
                        .GetText .ColLetterToNumber("S"), Row, vSoMat
                        .GetText .ColLetterToNumber("U"), Row, vSoHuy
                        
                        'Ho tro set tu so [10]
                        .GetText .ColLetterToNumber(Col_H), Row, strTuso1
                        .GetText .ColLetterToNumber(Col_I), Row, strDenso1
                        .GetText .ColLetterToNumber(Col_J), Row, strTuso2
                        .GetText .ColLetterToNumber(Col_K), Row, strDenso2
                        str = getTusoDenso(CStr(strTuso1), CStr(strDenso1), CStr(strTuso2), CStr(strDenso2))
                        If str(0) <> "-1" Then
                            ' Neu tong so xoa, sd, mat, huy =0 thi khong ho tro tinh [10][11]
                            If Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0 Then
                                .SetText .ColLetterToNumber(Col_L), Row, FormatSoHD(str(0))
                                UpdateCell fps, .ColLetterToNumber(Col_L), Row, .Text
                            Else
                                .SetText .ColLetterToNumber(Col_L), Row, ""
                                UpdateCell fps, .ColLetterToNumber(Col_L), Row, .Text
                            End If
                            .SetText .ColLetterToNumber(Col_G), Row, FormatSoHD(str(2))
                             UpdateCell fps, .ColLetterToNumber(Col_G), Row, .Text
                        End If
                        ' Ho tro tinh chi tieu [11]
                        .GetText .ColLetterToNumber(Col_M), Row, strCt11
                        .GetText .ColLetterToNumber(Col_N), Row, strCt12
                        .GetText .ColLetterToNumber(Col_G), Row, strCt5
                        If Val(strCt5) = Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0) Then
                            ' [5] =[12] ,  [13] + [14] + [16] + [18] <> 0
                            soluongTonCK = Val(strCt5) - Val(strCt12)
                            .SetText .ColLetterToNumber("X"), Row, ""  'FormatSoHD(Val(0))
                            UpdateCell fps, .ColLetterToNumber("X"), Row, .Text
                            .SetText .ColLetterToNumber(Col_Y), Row, ""   'FormatSoHD(Val(0) + soluongTonCK)
                            UpdateCell fps, .ColLetterToNumber(Col_Y), Row, .Text
                            .SetText .ColLetterToNumber(Col_Z), Row, soluongTonCK
                            UpdateCell fps, .ColLetterToNumber(Col_Z), Row, .Text
                        ElseIf Val(strCt5) > Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0) Then
                            soluongTonCK = Val(strCt5) - Val(strCt12)
                            .SetText .ColLetterToNumber("X"), Row, FormatSoHD(Val(strCt11) + 1)
                            UpdateCell fps, .ColLetterToNumber("X"), Row, .Text
                            .SetText .ColLetterToNumber(Col_Y), Row, FormatSoHD(Val(strCt11) + soluongTonCK)
                            UpdateCell fps, .ColLetterToNumber(Col_Y), Row, .Text
                            .SetText .ColLetterToNumber(Col_Z), Row, soluongTonCK
                            UpdateCell fps, .ColLetterToNumber(Col_Z), Row, .Text
                        ElseIf Val(strCt5) > Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) = 0) And (Trim(strTuso1) = "" And Trim(strDenso1) = "") Then
                             soluongTonCK = Val(strCt5)
                            .SetText .ColLetterToNumber("X"), Row, FormatSoHD(Val(strTuso2))
                            UpdateCell fps, .ColLetterToNumber("X"), Row, .Text
                            .SetText .ColLetterToNumber(Col_Y), Row, FormatSoHD(Val(strDenso2))
                            UpdateCell fps, .ColLetterToNumber(Col_Y), Row, .Text
                            .SetText .ColLetterToNumber(Col_Z), Row, soluongTonCK
                            UpdateCell fps, .ColLetterToNumber(Col_Z), Row, .Text
                        ElseIf Val(strCt5) > Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) = 0) And (Trim(strTuso2) = "" And Trim(strDenso2) = "") Then
                             soluongTonCK = Val(strCt5)
                            .SetText .ColLetterToNumber("X"), Row, FormatSoHD(Val(strTuso1))
                            UpdateCell fps, .ColLetterToNumber("X"), Row, .Text
                            .SetText .ColLetterToNumber(Col_Y), Row, FormatSoHD(Val(strDenso1))
                            UpdateCell fps, .ColLetterToNumber(Col_Y), Row, .Text
                            .SetText .ColLetterToNumber(Col_Z), Row, soluongTonCK
                            UpdateCell fps, .ColLetterToNumber(Col_Z), Row, .Text
                        Else
                             soluongTonCK = Val(strCt5)
                            .SetText .ColLetterToNumber("X"), Row, FormatSoHD(Val(strTuso1))
                            UpdateCell fps, .ColLetterToNumber("X"), Row, .Text
                            .SetText .ColLetterToNumber(Col_Y), Row, FormatSoHD(Val(strDenso2))
                            UpdateCell fps, .ColLetterToNumber(Col_Y), Row, .Text
                            .SetText .ColLetterToNumber(Col_Z), Row, soluongTonCK
                            UpdateCell fps, .ColLetterToNumber(Col_Z), Row, .Text
                        End If
                    End If
                'Tinh so hoa don mat, xoa, huy
                    If Col = .ColLetterToNumber(Col_R) Or Col = .ColLetterToNumber(Col_T) Or Col = .ColLetterToNumber(Col_V) Or Col = .ColLetterToNumber("P") Then
                        .GetText Col, Row, varTemp
                        TongSo = GetSoLuongHD(CStr(varTemp))
                        If Col = .ColLetterToNumber(Col_R) Then
                            .SetText .ColLetterToNumber(Col_Q), Row, TongSo
                            UpdateCell fps, .ColLetterToNumber(Col_Q), Row, .Text
                        End If
                        
                        If Col = .ColLetterToNumber(Col_T) Then
                            .SetText .ColLetterToNumber(Col_S), Row, TongSo
                            UpdateCell fps, .ColLetterToNumber(Col_S), Row, .Text
                        End If
                        
                        If Col = .ColLetterToNumber(Col_V) Then
                            .SetText .ColLetterToNumber(Col_U), Row, TongSo
                            UpdateCell fps, .ColLetterToNumber(Col_U), Row, .Text
                        End If
                        
                        
                        ' Ho tro tinh den so [11]
                        .GetText .ColLetterToNumber(Col_L), Row, strCt10
                        .GetText .ColLetterToNumber(Col_N), Row, strCt12
                        strCt11 = Val(strCt10) + Val(strCt12) - 1
                        
                        ' Neu tong so xoa, sd, mat, huy =0 thi khong ho tro tinh [10][11]
                        If Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0 Then
                            .SetText .ColLetterToNumber(Col_M), Row, FormatSoHD(CStr(strCt11))
                            UpdateCell fps, .ColLetterToNumber(Col_M), Row, .Text
                            If str(0) <> "-1" Then
                                .SetText .ColLetterToNumber(Col_L), Row, FormatSoHD(str(0))
                                UpdateCell fps, .ColLetterToNumber(Col_L), Row, .Text
                            End If
                        Else
                            .SetText .ColLetterToNumber(Col_L), Row, ""
                            UpdateCell fps, .ColLetterToNumber(Col_L), Row, .Text
                            
                            .SetText .ColLetterToNumber(Col_M), Row, ""
                            UpdateCell fps, .ColLetterToNumber(Col_M), Row, .Text
                        End If
                                                        
                        .GetText .ColLetterToNumber(Col_G), Row, strCt5
                        If Val(strCt5) > Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0) Then
                            soluongTonCK = Val(strCt5) - Val(strCt12)
                            .SetText .ColLetterToNumber("X"), Row, FormatSoHD(Val(strCt11) + 1)
                            UpdateCell fps, .ColLetterToNumber("X"), Row, .Text
                            .SetText .ColLetterToNumber(Col_Y), Row, FormatSoHD(Val(strCt11) + soluongTonCK)
                            UpdateCell fps, .ColLetterToNumber(Col_Y), Row, .Text
                            .SetText .ColLetterToNumber(Col_Z), Row, soluongTonCK
                            UpdateCell fps, .ColLetterToNumber(Col_Z), Row, .Text
                        End If
                    End If
        
                   If Col = .ColLetterToNumber("P") Then
                        ' Ho tro tinh den so [11]
                        .GetText .ColLetterToNumber(Col_L), Row, strCt10
                        .GetText .ColLetterToNumber(Col_N), Row, strCt12
                        .GetText .ColLetterToNumber(Col_G), Row, strCt5
                        strCt11 = Val(strCt10) + Val(strCt12) - 1
                        
                        ' Neu tong so xoa, sd, mat, huy =0 thi khong ho tro tinh [10][11]
                        If Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0 Then
                            .SetText .ColLetterToNumber(Col_M), Row, FormatSoHD(CStr(strCt11))
                            UpdateCell fps, .ColLetterToNumber(Col_M), Row, .Text
                            If str(0) <> "-1" Then
                                .SetText .ColLetterToNumber(Col_L), Row, FormatSoHD(str(0))
                                UpdateCell fps, .ColLetterToNumber(Col_L), Row, .Text
                            End If
                        Else
                            .SetText .ColLetterToNumber(Col_L), Row, ""
                            UpdateCell fps, .ColLetterToNumber(Col_L), Row, .Text
                            
                            .SetText .ColLetterToNumber(Col_M), Row, ""
                            UpdateCell fps, .ColLetterToNumber(Col_M), Row, .Text
                        End If
                        
                        If Val(strCt5) > Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0) Then
                            soluongTonCK = Val(strCt5) - Val(strCt12)
                            .SetText .ColLetterToNumber("X"), Row, FormatSoHD(Val(strCt11) + 1)
                            UpdateCell fps, .ColLetterToNumber("X"), Row, .Text
                            .SetText .ColLetterToNumber(Col_Y), Row, FormatSoHD(Val(strCt11) + soluongTonCK)
                            UpdateCell fps, .ColLetterToNumber(Col_Y), Row, .Text
                            .SetText .ColLetterToNumber(Col_Z), Row, soluongTonCK
                            UpdateCell fps, .ColLetterToNumber(Col_Z), Row, .Text
                        End If
                    End If
                  ' Tinh toan lai cac gia tri
                  If Col = .ColLetterToNumber("H") Or Col = .ColLetterToNumber("I") Or Col = .ColLetterToNumber("J") Or Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("P") Or Col = .ColLetterToNumber("R") Or Col = .ColLetterToNumber("T") Or Col = .ColLetterToNumber("V") Then
                        .GetText .ColLetterToNumber("G"), Row, strCt5
                        .GetText .ColLetterToNumber("N"), Row, strCt12
                        .GetText .ColLetterToNumber("P"), Row, vSoSD
                        .GetText .ColLetterToNumber("Q"), Row, vSoXoa
                        .GetText .ColLetterToNumber("S"), Row, vSoMat
                        .GetText .ColLetterToNumber("U"), Row, vSoHuy
                        If Val(strCt5) + Val(strCt12) + Val(vSoSD) + Val(vSoXoa) + Val(vSoHuy) + Val(vSoMat) = 0 Then
                            soluongTonCK = 0
                             .SetText .ColLetterToNumber("X"), Row, ""
                            UpdateCell fps, .ColLetterToNumber("X"), Row, .Text
                            .SetText .ColLetterToNumber(Col_Y), Row, ""
                            UpdateCell fps, .ColLetterToNumber(Col_Y), Row, .Text
                            .SetText .ColLetterToNumber(Col_Z), Row, soluongTonCK
                            UpdateCell fps, .ColLetterToNumber(Col_Z), Row, .Text
                        ElseIf Val(strCt5) = Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0) Then
                            ' [5] =[12] ,  [13] + [14] + [16] + [18] <> 0
                            soluongTonCK = Val(strCt5) - Val(strCt12)
                            .SetText .ColLetterToNumber("X"), Row, ""  'FormatSoHD(Val(0))
                            UpdateCell fps, .ColLetterToNumber("X"), Row, .Text
                            .SetText .ColLetterToNumber(Col_Y), Row, ""   'FormatSoHD(Val(0) + soluongTonCK)
                            UpdateCell fps, .ColLetterToNumber(Col_Y), Row, .Text
                            .SetText .ColLetterToNumber(Col_Z), Row, soluongTonCK
                            UpdateCell fps, .ColLetterToNumber(Col_Z), Row, .Text
                        End If
                  End If
                   ' end
                   
                  ' Chuyen cac ky hieu mau so HD thanh chu in hoa
                  If Col = .ColLetterToNumber(Col_E) Then
                        .GetText .ColLetterToNumber("AA"), Row, strId
                        strLoai = Trim(Left(strId, 1))
                        
                        .GetText Col, Row, varTemp
                        .SetText Col, Row, IIf(strLoai = "3", varTemp, UCase(varTemp))
                        UpdateCell fps, Col, Row, .Text
                  End If
            ElseIf .ActiveSheet = 2 Then
                    ' Dem so dong dong sec 1
                    .Row = 46
                    Do
                        countRowDynamic1_PL1 = countRowDynamic1_PL1 + 1
                        .Col = .ColLetterToNumber("B")
                        .Row = .Row + 1
                    Loop Until UCase(.Text) = "AA"
                    .Row = 49 + countRowDynamic1_PL1
                    Do
                        countRowDynamic2_PL1 = countRowDynamic2_PL1 + 1
                        .Col = .ColLetterToNumber("B")
                        .Row = .Row + 1
                    Loop Until .Text = "bb"
                ' check ngay ky
                    If Col = .ColLetterToNumber(Col_R) And Row = countRowDynamic1_PL1 + countRowDynamic2_PL1 + 51 Then
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
                        UpdateCell_sheet fps, 2, Col, Row, .Text
                    End If
                  ' Format dinh dang tu so den so co 7 ky tu
                    If Col = .ColLetterToNumber(Col_V) Or Col = .ColLetterToNumber(Col_Y) Then
                         .GetText Col, Row, varTemp
                         .SetText Col, Row, FormatSoHD(CStr(varTemp))
                         .Col = Col
                         .Row = Row
                        UpdateCell_sheet fps, 2, .Col, .Row, .Text
                   ' Ho tro tinh so luong
                        .GetText .ColLetterToNumber(Col_V), Row, tuSo
                        .GetText .ColLetterToNumber(Col_Y), Row, denso
                        If Trim(tuSo) <> "" Or Trim(denso) <> "" Then
                            .SetText .ColLetterToNumber(Col_AA), Row, Val(denso) - Val(tuSo) + 1
                            .Col = .ColLetterToNumber(Col_AA)
                            .Row = Row
                            UpdateCell_sheet fps, 2, .Col, .Row, .Text
                        End If
                        
                    End If
                    
                  ' Chuyen ky hieu HD thanh chu cai in hoa
                    If Col = .ColLetterToNumber(Col_S) Then
                         .GetText Col, Row, varTemp
                         .SetText Col, Row, FormatKyHieu(CStr(varTemp))
                         .Col = Col
                         .Row = Row
                        UpdateCell_sheet fps, 2, .Col, .Row, .Text
                    End If
                   'Kiem tra cau truc cua ky hieu neu sai set cot an =1 dung =0
                    If Col = .ColLetterToNumber(Col_S) Then
                         .GetText Col, Row, varTemp
                         .SetText .ColLetterToNumber("AH"), Row, CheckSoHD(CStr(varTemp), strLoai)
                         .Col = .ColLetterToNumber("AH")
                         .Row = Row
                        UpdateCell_sheet fps, 2, .Col, .Row, .Text
                    End If
                   'Kiem tra cau truc cua Mau so
                   If Col = .ColLetterToNumber(Col_O) Then
                        .GetText .ColLetterToNumber("AI"), Row, strId
                        strLoai = Trim(Left(strId, 1))
                        If strLoai = "0" Then
                            strTemp = Trim(Mid(strId, 11, 6))
                        ElseIf strLoai = "4" Then
                            strTemp = Trim(Mid(strId, 11, 6))
                        Else
                            strTemp = Trim(Mid(strId, 11, 3))
                        End If
                        .GetText Col, Row, varTemp
                        .SetText Col, Row, IIf(strLoai = "3", varTemp, UCase(varTemp))
                         .SetText .ColLetterToNumber("AJ"), Row, CheckMauSoHD(CStr(varTemp), CStr(strLoai), CStr(strTemp))
                         .Col = .ColLetterToNumber("AJ")
                         .Row = Row
                        UpdateCell_sheet fps, 2, .Col, .Row, .Text
                    End If
            ElseIf .ActiveSheet = 3 Then
                Dim countRowDynamic1 As Integer
                    ' Dem so dong dong
                    .Row = 45
                    Do
                        countRowDynamic1 = countRowDynamic1 + 1
                        .Col = .ColLetterToNumber("B")
                        .Row = .Row + 1
                    Loop Until UCase(.Text) = "AA"
                    
                ' check ngay ky
                    If Col = .ColLetterToNumber(Col_R) And Row = countRowDynamic1 + 50 Then
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
                        UpdateCell_sheet fps, 3, .Col, .Row, .Text
                    End If
                  ' Format dinh dang tu so den so co 7 ky tu
                    If Col = .ColLetterToNumber(Col_V) Or Col = .ColLetterToNumber(Col_Y) Then
                         .GetText Col, Row, varTemp
                         .SetText Col, Row, FormatSoHD(CStr(varTemp))
                         .Col = Col
                         .Row = Row
                        UpdateCell_sheet fps, 3, .Col, .Row, .Text
                   ' Ho tro tinh so luong
                        .GetText .ColLetterToNumber(Col_V), Row, tuSo
                        .GetText .ColLetterToNumber(Col_Y), Row, denso
                     If Trim(tuSo) <> "" Or Trim(denso) <> "" Then
                         .SetText .ColLetterToNumber(Col_AA), Row, Val(denso) - Val(tuSo) + 1
                         .Col = .ColLetterToNumber(Col_AA)
                         .Row = Row
                         UpdateCell_sheet fps, 3, .Col, .Row, .Text
                     End If
                        
                    End If
                    
                  ' Chuyen ky hieu HD thanh chu cai in hoa
                    If Col = .ColLetterToNumber(Col_S) Then
                         .GetText Col, Row, varTemp
                         .SetText Col, Row, FormatKyHieu(CStr(varTemp))
                         .Col = Col
                         .Row = Row
                        UpdateCell_sheet fps, 3, .Col, .Row, .Text
                    End If
                   'Kiem tra cau truc cua ky hieu neu sai set cot an =1 dung =0
                    If Col = .ColLetterToNumber(Col_S) Then
                         .GetText Col, Row, varTemp
                         .SetText .ColLetterToNumber("AH"), Row, CheckSoHD(CStr(varTemp), strLoai)
                         .Col = .ColLetterToNumber("AH")
                         .Row = Row
                        UpdateCell_sheet fps, 3, .Col, .Row, .Text
                    End If
                   'Kiem tra cau truc cua Mau so
                   If Col = .ColLetterToNumber(Col_O) Then
                        .GetText .ColLetterToNumber("AI"), Row, strId
                        strLoai = Trim(Left(strId, 1))
                        If strLoai = "0" Then
                            strTemp = Trim(Mid(strId, 11, 6))
                        ElseIf strLoai = "4" Then
                            strTemp = Trim(Mid(strId, 11, 6))
                        Else
                            strTemp = Trim(Mid(strId, 11, 3))
                        End If
                        .GetText Col, Row, varTemp
                        .SetText Col, Row, IIf(strLoai = "3", varTemp, UCase(varTemp))
                         .SetText .ColLetterToNumber("AJ"), Row, CheckMauSoHD(CStr(varTemp), CStr(strLoai), CStr(strTemp))
                         .Col = .ColLetterToNumber("AJ")
                         .Row = Row
                        UpdateCell_sheet fps, 3, .Col, .Row, .Text
                    End If
            End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
    
End Sub

'Calling this function whenever we don't need this class anymore.
'No parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
    'If flgLoadToKhai = True Then Exit Sub
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        
        SetData
        ' 4/1/2011 dhdang sua
        UpdateSheets
        CheckDynamicError 'Set Exception Error on cells of interface
            
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
        
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    Dim result() As String
    With fps
            clearCellErorrHDR
            checkSheet1
            If TAX_Utilities_v2.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0" Then
                checkSheet2
            End If
            If TAX_Utilities_v2.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0" Then
                checkSheet3
            End If
    End With
    'result = loadTonCuoiKy
    
End Sub

Private Sub checkSheet1()

    Dim strCheck As String
    Dim errMauSoHD As Variant, errKyHieuHD As Variant, errSoLuong As Variant, errSoLienHD As Variant, errSoLienHDDT As Variant
    Dim errKhoangGN As Variant, errLengthHD As Variant, errEmty As Variant, errNgayKy As Variant, errNgayKyHL As Variant, errNgayHuy As Variant, errTrungKhoang As Variant, errLienKhoang As Variant, errSoXoa As Variant, errSaiDD As Variant
    Dim errSLSudung As Variant, errSLXoa As Variant, errSLMat As Variant, errSLHuy As Variant, errSLSD_Xoa_Mat_Huy As Variant
    Dim strTenHD As Variant, strMauSoHD As Variant, strKyHieuHD As Variant, strTuSo6 As Variant, strDenSo7 As Variant, strTuso8 As Variant, strDenso9 As Variant, strTuSo10 As Variant, strDenSo11 As Variant, strTuSo20 As Variant, strDenSo21 As Variant, errSLTonDK As Variant
    Dim strSLSuDung As Variant, strSLMat As Variant, strSLhuy As Variant, strSLXoa As Variant, strTongSo5 As Variant, strTongSo12 As Variant

    Dim strTenHD1 As Variant, strMauSoHD1 As Variant, strKyHieuHD1 As Variant, strTuso1 As Variant, strDenso1 As Variant
    Dim strSoLuong As Variant, strNguoiKy As Variant, strNguoiLapBieu As Variant, strNgayKy As Variant, strNgayHuy As Variant
    
    Dim iFlagTenHD As Integer, iFlagMauSoHD As Integer, iFlagKyHieu As Integer, iFlagSoLienHD As Integer, iFlagEmty As Integer
    Dim iWhiteRow As Integer, iFlagLengHD As Integer, iFlagSoLuong As Integer, iFlagNgayKy As Integer, iFlagNgayKyHL As Integer, iFlagNgayHuy As Integer
    Dim iFlagKhoangTSDSTrung As Integer, iFocusFlag As Integer, iFlagLienKhoang As Integer, iFlagSoXoa As Integer, iFlagSaiDD As Integer
    Dim iFlagSLSudung As Integer, iFlagSLXoa As Integer, iFlagSLMat As Integer, iFlagSLHuy As Integer, iFlagSD_Xoa_Mat_Huy As Integer
    
    Dim strId As Variant, strLoai As String, strTemp As String, result As String, strTon As Variant
    Dim i As Integer, countDynamicRow As Integer, j As Integer, k As Integer, t As Integer
    
    Dim checkXoa As Boolean, checkMat As Boolean, checkHuy As Boolean, blCheck_S2 As Boolean, blCheck_S3 As Boolean
    
    Dim countSoTon As Double
    Dim blnTemp As Boolean
    
    Dim vKyBCTuNgay As Variant
    
    ' group theo tung nhom
    Dim arrGroupHD() As String
    Dim strGroupHD As String
    
    Dim str() As String
    Dim strLietKe() As String
    
    Dim strTrung67() As String, strTrung89() As String, strTrung2021() As String, strCheckTrung() As String
    
    blCheck_S2 = IIf(TAX_Utilities_v2.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S3 = IIf(TAX_Utilities_v2.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    
    strTrung67 = checkTrung(fps, 22, Col_H, Col_I, "B", Col_D, Col_E, Col_F, "", False, 1)
    strTrung89 = checkTrung(fps, 22, Col_J, Col_K, "B", Col_D, Col_E, Col_F, "", False, 1)
    strTrung2021 = checkTrung(fps, 22, "X", "Y", "B", Col_D, Col_E, Col_F, "", False, 1)
    
    
    ' Clear cell error
    clearCellErorr (22)
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = .SheetCount
        ' Lay thong bao loi tu sheet hearder
        .GetText .ColLetterToNumber(Col_E), 13, errEmty
        .GetText .ColLetterToNumber(Col_E), 14, errTrungKhoang
        .GetText .ColLetterToNumber(Col_E), 15, errKyHieuHD
        .GetText .ColLetterToNumber(Col_E), 16, errSoLuong
        .GetText .ColLetterToNumber(Col_E), 17, errMauSoHD
        .GetText .ColLetterToNumber(Col_E), 18, errSoLienHD
        .GetText .ColLetterToNumber(Col_E), 19, errNgayKy
        .GetText .ColLetterToNumber(Col_E), 20, errLengthHD
        .GetText .ColLetterToNumber(Col_E), 21, errLienKhoang
        .GetText .ColLetterToNumber(Col_E), 22, errSoXoa
        .GetText .ColLetterToNumber(Col_E), 23, errSaiDD
        .GetText .ColLetterToNumber(Col_E), 24, errSLSudung
        .GetText .ColLetterToNumber(Col_E), 25, errSLSD_Xoa_Mat_Huy
        .GetText .ColLetterToNumber(Col_E), 26, errSLXoa
        .GetText .ColLetterToNumber(Col_E), 27, errSLMat
        .GetText .ColLetterToNumber(Col_E), 28, errSLHuy
        .GetText .ColLetterToNumber(Col_E), 49, errSLTonDK
        
        .GetText .ColLetterToNumber(Col_E), 52, errNgayKyHL
        
        .GetText .ColLetterToNumber(Col_E), 53, errSoLienHDDT
        
        .GetText .ColLetterToNumber("T"), 6, vKyBCTuNgay
        
        .Sheet = 1
        
        '
        iFlagEmty = 0
        iFlagTenHD = 0
        iFlagMauSoHD = 0
        iFlagKyHieu = 0
        iFlagSoLienHD = 0
        iFlagEmty = 0
        iFlagLengHD = 0
        iFlagNgayKy = 0
        iFlagNgayKyHL = 0
        iFlagLienKhoang = 0
        iFlagSoXoa = 0
        iFlagSaiDD = 0
        iFlagSLSudung = 0
        iFlagSD_Xoa_Mat_Huy = 0
        iFlagSLXoa = 0
        iFlagSLMat = 0
        iFlagSLHuy = 0
        
        
        checkXoa = False
        checkMat = False
        checkHuy = False
        
        .Row = 22
        Do
            countDynamicRow = countDynamicRow + 1
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
            ' Check cho den het cac dong co du lieu thi thoi
        Loop Until UCase(.Text) = Col_AA
        
        .Row = 22
        Do
            j = .Row
            iWhiteRow = 1
            '.GetText .ColLetterToNumber(Col_D), .Row, strMaHD
            .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
            .GetText .ColLetterToNumber(Col_E), .Row, strMauSoHD
            .GetText .ColLetterToNumber(Col_F), .Row, strKyHieuHD
            .GetText .ColLetterToNumber(Col_H), .Row, strTuSo6
            .GetText .ColLetterToNumber(Col_I), .Row, strDenSo7
            .GetText .ColLetterToNumber(Col_J), .Row, strTuso8
            .GetText .ColLetterToNumber(Col_K), .Row, strDenso9
            .GetText .ColLetterToNumber("X"), .Row, strTuSo20
            .GetText .ColLetterToNumber("Y"), .Row, strDenSo21
            .GetText .ColLetterToNumber("P"), .Row, strSLSuDung
            .GetText .ColLetterToNumber(Col_R), .Row, strSLXoa
            .GetText .ColLetterToNumber(Col_T), .Row, strSLMat
            .GetText .ColLetterToNumber(Col_V), .Row, strSLhuy
            .GetText .ColLetterToNumber(Col_L), .Row, strTuSo10
            .GetText .ColLetterToNumber(Col_M), .Row, strDenSo11
            .GetText .ColLetterToNumber(Col_G), .Row, strTongSo5
            .GetText .ColLetterToNumber(Col_N), .Row, strTongSo12
            
            .GetText .ColLetterToNumber(Col_AA), .Row, strId
            .GetText .ColLetterToNumber("AC"), .Row, strTon
            
            
            ' set gia tri cho check KH NN/XC
            strMauSoHD_01GTKT = Trim(Mid(strId, 11, 6))
            ' end
            'dhdang sua check ton dau ky
            'ngay 18-05
            '--------------------------------------------Begin
            
            If numRowInsert >= 0 And TAX_Utilities_v2.CheckNewDataBC26 = True And Val(strTon) = 1 Then
                   countSoTon = 0
                   For t = 0 To numRowInsert
                        ' set du lieu ton dau ky
                        strTemp_ton = result_ton(t)
                        arrCellInRow = Split(strTemp_ton, "~")
                       ' lay ve thong tin tu so den so va cac truong hop
                            If strMauSoHD = arrCellInRow(0) And strKyHieuHD = arrCellInRow(1) And strTenHD = arrCellInRow(5) Then
                            ' truong hop chi co so ton
'                                If arrCellInRow(2) <> strTuSo6 Then
'                                    .Col = .ColLetterToNumber(Col_H)
'                                    .CellNote = errSLTonDK  'static
'                                    .BackColor = mAlertColor
'                                End If
'                                If arrCellInRow(3) <> strDenSo7 Then
'                                    .Col = .ColLetterToNumber(Col_I)
'                                    .CellNote = errSLTonDK  'static
'                                    .BackColor = mAlertColor
'                                End If
                                If arrCellInRow(2) = strTuSo6 And arrCellInRow(3) = strDenSo7 Then
                                    countSoTon = countSoTon + 1
                                End If
                            End If
                    Next t
                    
                    If countSoTon = 0 Then
                        .Col = .ColLetterToNumber(Col_H)
                        .CellNote = errSLTonDK  'static
                        .BackColor = mAlertColor
                        .Col = .ColLetterToNumber(Col_I)
                        .CellNote = errSLTonDK  'static
                        .BackColor = mAlertColor
                    Else
                        .Col = .ColLetterToNumber(Col_H)
                        .CellNote = ""  'static
                        .BackColor = mNonErrorColor
                        .Col = .ColLetterToNumber(Col_I)
                        .CellNote = ""  'static
                        .BackColor = mNonErrorColor
                    End If
                End If

            '---------------------------------------------End
            
            
            'check mau so hoa don
           If Trim(strTenHD) <> "" Or Trim(strMauSoHD) <> "" Or Trim(strKyHieuHD) <> "" Or Trim(strId) <> "" Or Trim(strTuSo6) <> "" _
           Or Trim(strDenSo7) <> "" Or Trim(strTuso8) <> "" Or Trim(strDenso9) <> "" Or Trim(strSLXoa) <> "" Or Trim(strSLMat) <> "" Or Trim(strSLhuy) <> "" Or Val(strSLSuDung) <> 0 Then
                ' Khong phai dong trang
                iWhiteRow = 0
            End If
            'Ten HD
            If Trim(strTenHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber(Col_D)
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            End If
            ' Mau so HD
            If Trim(strMauSoHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber(Col_E)
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                strLoai = Trim(Left(strId, 1))
                If strLoai = "0" Then
                    strTemp = Trim(Mid(strId, 11, 6))
                ElseIf strLoai = "4" Then
                    strTemp = Trim(Mid(strId, 11, 6))
                Else
                    strTemp = Trim(Mid(strId, 11, 3))
                End If
                result = CheckMauSoHD(CStr(strMauSoHD), strLoai, strTemp)
                If result = "1" And iWhiteRow = 0 Then
                    .Col = .ColLetterToNumber(Col_E)
                    .CellNote = errMauSoHD  'static
                    .BackColor = mErrorColor
                    iFlagMauSoHD = 1
                ElseIf result = "2" And iWhiteRow = 0 Then
                    If Right$(Trim$(strKyHieuHD), 1) = "E" Then
                        If CheckSoLienHDDT(CStr(strMauSoHD), strLoai, strTemp, CStr(strKyHieuHD)) = "3" Then
                             .Col = .ColLetterToNumber(Col_E)
                            .CellNote = errSoLienHDDT  'static
'                            .BackColor = mErrorColor
'                            iFlagSoLienHD = 2
                             .BackColor = mAlertColor
                        Else
                            .Col = .ColLetterToNumber(Col_E)
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    Else
                        .Col = .ColLetterToNumber(Col_E)
                        .CellNote = errSoLienHD  'static
                        .BackColor = mErrorColor
                        iFlagSoLienHD = 1
                    End If
                Else
                    If Right$(Trim$(strKyHieuHD), 1) = "E" Then
                        If CheckSoLienHDDT(CStr(strMauSoHD), strLoai, strTemp, CStr(strKyHieuHD)) = "3" Then
                             .Col = .ColLetterToNumber(Col_E)
                            .CellNote = errSoLienHDDT  'static
'                            .BackColor = mErrorColor
'                            iFlagSoLienHD = 2
                             .BackColor = mAlertColor
                        Else
                            .Col = .ColLetterToNumber(Col_E)
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    Else
                        .Col = .ColLetterToNumber(Col_E)
                        .CellNote = ""  'static
                        .BackColor = mNonErrorColor
                    End If
                End If
            End If
            
            'Kiem tra cau truc cua ky hieu
            If strKyHieuHD = "" And iWhiteRow = 0 Then
                 .Col = .ColLetterToNumber(Col_F)
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                result = CheckSoHD(CStr(strKyHieuHD), strLoai)
                If result = "1" And iWhiteRow = 0 Then
                    .Col = .ColLetterToNumber(Col_F)
                    .CellNote = errKyHieuHD  'static
                    .BackColor = mErrorColor
                    iFlagKyHieu = 1
                ElseIf result = "2" And iWhiteRow = 0 Then
                    .Col = .ColLetterToNumber(Col_F)
                    .CellNote = errLengthHD  'static
                    .BackColor = mErrorColor
                    iFlagLengHD = 1
                End If
            End If
            
            ' kiem tra tu so den so
            ' reset cac cell error
            ' chi tieu [6],[7], [8], [9]
            
            If iWhiteRow = 0 Then
                ' lay ve thong tin tu so den so va cac truong hop
                str = checkSoTonSoPH(CStr(strTuSo6), CStr(strDenSo7), CStr(strTuso8), CStr(strDenso9))
                If str(0) = "1" Then
                ' truong hop chi co so ton
                    If str(1) = "1" Then
                        .Col = .ColLetterToNumber(Col_H)
                        .CellNote = errEmty  'static
                        .BackColor = mErrorColor
                        iFlagEmty = 1
                    ElseIf str(1) = "2" Then
                        .Col = .ColLetterToNumber(Col_I)
                        .CellNote = errEmty  'static
                        .BackColor = mErrorColor
                        iFlagEmty = 1
                    ElseIf str(1) = "3" Then
                        ' tu so
                        .Col = .ColLetterToNumber(Col_H)
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        ' den so
                        .Col = .ColLetterToNumber(Col_I)
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        iFlagSoLuong = 1
                    ElseIf str(1) = "4" Then
                        ' kiem tra so hoa don xoa
                        blnTemp = isDuLieuHL(CStr(strSLXoa))
                        If strSLXoa <> "" And blnTemp Then
                            If strSLXoa <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                                If blnTemp Then
                                    ' du lieu hop le moi check
                                    result = checkHDXoa(CStr(strTuSo10), CStr(strDenSo11), CStr(strSLXoa))
                                    If result = "-1" Then
                                        .Col = .ColLetterToNumber(Col_R)
                                        .CellNote = errSoXoa  'static
                                        .BackColor = mErrorColor
                                        iFlagSoXoa = 1
                                        checkXoa = True
                                    End If
                                Else
                                        .Col = .ColLetterToNumber(Col_R)
                                        .CellNote = errSaiDD  'static
                                        .BackColor = mErrorColor
                                        iFlagSaiDD = 1
                                End If
                            End If
                        ElseIf strSLXoa <> "" And Not blnTemp Then
                            .Col = .ColLetterToNumber(Col_R)
                            .CellNote = errSaiDD  'static
                            .BackColor = mErrorColor
                            iFlagSaiDD = 1
                        End If
                        ' kiem tra so mat
                        blnTemp = isDuLieuHL(CStr(strSLMat))
                        If strSLMat <> "" And blnTemp Then
                            If strSLMat <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                                If blnTemp Then
                                    result = checkHDXoa(CStr(strTuSo10), CStr(strDenSo11), CStr(strSLMat))
                                    If result = "-1" Then
                                        .Col = .ColLetterToNumber(Col_T)
                                        .CellNote = errSoXoa  'static
                                        .BackColor = mErrorColor
                                        iFlagSoXoa = 1
                                        checkMat = True
                                    End If
                                Else
                                        .Col = .ColLetterToNumber(Col_T)
                                        .CellNote = errSaiDD  'static
                                        .BackColor = mErrorColor
                                        iFlagSaiDD = 1
                                End If
                            End If
                        ElseIf strSLMat <> "" And Not blnTemp Then
                            .Col = .ColLetterToNumber(Col_T)
                            .CellNote = errSaiDD  'static
                            .BackColor = mErrorColor
                            iFlagSaiDD = 1
                        End If
                        ' kiem tra so huy
                        blnTemp = isDuLieuHL(CStr(strSLhuy))
                        If strSLhuy <> "" And blnTemp Then
                            If strSLhuy <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                                If blnTemp Then
                                    result = checkHDXoa(CStr(strTuSo10), CStr(strDenSo11), CStr(strSLhuy))
                                    If result = "-1" Then
                                        .Col = .ColLetterToNumber(Col_V)
                                        .CellNote = errSoXoa  'static
                                        .BackColor = mErrorColor
                                        iFlagSoXoa = 1
                                        checkHuy = True
                                    End If
                                Else
                                        .Col = .ColLetterToNumber(Col_V)
                                        .CellNote = errSaiDD  'static
                                        .BackColor = mErrorColor
                                        iFlagSaiDD = 1
                                End If
                            End If
                         ElseIf strSLhuy <> "" And Not blnTemp Then
                            .Col = .ColLetterToNumber(Col_V)
                            .CellNote = errSaiDD  'static
                            .BackColor = mErrorColor
                            iFlagSaiDD = 1
                         End If
                        ' kiem tra trung khoang xoa, mat , huy
                        ' 1 trung Xoa, Mat
                        ' 2 trung Xoa, Huy
                        ' 3 trung Mat, huy
                        ' 4 trung ca 3
                        If (strSLXoa <> "" Or strSLMat <> "" Or strSLhuy <> "") And checkXoa = False And checkHuy = False And checkMat = False Then
                            result = checkTrungXoa(CStr(strSLXoa), CStr(strSLMat), CStr(strSLhuy))
                            If result = "1" Then
                                .Col = .ColLetterToNumber(Col_R)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_T)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                iFlagSoXoa = 1
                            ElseIf result = "2" Then
                                .Col = .ColLetterToNumber(Col_R)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_V)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                iFlagSoXoa = 1
                            ElseIf result = "3" Then
                                .Col = .ColLetterToNumber(Col_T)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_V)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                iFlagSoXoa = 1
                            ElseIf result = "4" Then
                                .Col = .ColLetterToNumber(Col_T)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_R)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_V)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                iFlagSoXoa = 1
                            
                            End If
                        End If
                    End If
                ElseIf str(0) = "2" Then
                ' truong hop chi co so phat hanh
                    If str(2) = "1" Then
                        .Col = .ColLetterToNumber(Col_J)
                        .CellNote = errEmty  'static
                        .BackColor = mErrorColor
                        iFlagEmty = 1
                    ElseIf str(2) = "2" Then
                        .Col = .ColLetterToNumber(Col_K)
                        .CellNote = errEmty  'static
                        .BackColor = mErrorColor
                        iFlagEmty = 1
                    ElseIf str(2) = "3" Then
                        ' tu so
                        .Col = .ColLetterToNumber(Col_J)
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        ' den so
                        .Col = .ColLetterToNumber(Col_K)
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        iFlagSoLuong = 1
                    ElseIf str(2) = "4" Then
                        ' kiem tra so hoa don xoa
                        blnTemp = isDuLieuHL(CStr(strSLXoa))
                        If strSLXoa <> "" And blnTemp Then
                            If strSLXoa <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                                If blnTemp Then
                                    ' du lieu hop le moi check
                                    result = checkHDXoa(CStr(strTuSo10), CStr(strDenSo11), CStr(strSLXoa))
                                    If result = "-1" Then
                                        .Col = .ColLetterToNumber(Col_R)
                                        .CellNote = errSoXoa  'static
                                        .BackColor = mErrorColor
                                        iFlagSoXoa = 1
                                        checkXoa = True
                                    End If
                                Else
                                        .Col = .ColLetterToNumber(Col_R)
                                        .CellNote = errSaiDD  'static
                                        .BackColor = mErrorColor
                                        iFlagSaiDD = 1
                                End If
                            End If
                        ElseIf strSLXoa <> "" And Not blnTemp Then
                            .Col = .ColLetterToNumber(Col_R)
                            .CellNote = errSaiDD  'static
                            .BackColor = mErrorColor
                            iFlagSaiDD = 1
                        End If
                        ' kiem tra so mat
                        blnTemp = isDuLieuHL(CStr(strSLMat))
                        If strSLMat <> "" And blnTemp Then
                            If strSLMat <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                                If blnTemp Then
                                    result = checkHDXoa(CStr(strTuSo10), CStr(strDenSo11), CStr(strSLMat))
                                    If result = "-1" Then
                                        .Col = .ColLetterToNumber(Col_T)
                                        .CellNote = errSoXoa  'static
                                        .BackColor = mErrorColor
                                        iFlagSoXoa = 1
                                        checkMat = True
                                    End If
                                Else
                                        .Col = .ColLetterToNumber(Col_T)
                                        .CellNote = errSaiDD  'static
                                        .BackColor = mErrorColor
                                        iFlagSaiDD = 1
                                End If
                            End If
                        ElseIf strSLMat <> "" And Not blnTemp Then
                            .Col = .ColLetterToNumber(Col_T)
                            .CellNote = errSaiDD  'static
                            .BackColor = mErrorColor
                            iFlagSaiDD = 1
                        End If
                        ' kiem tra so huy
                        blnTemp = isDuLieuHL(CStr(strSLhuy))
                        If strSLhuy <> "" And blnTemp Then
                            If strSLhuy <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                                If blnTemp Then
                                    result = checkHDXoa(CStr(strTuSo10), CStr(strDenSo11), CStr(strSLhuy))
                                    If result = "-1" Then
                                        .Col = .ColLetterToNumber(Col_V)
                                        .CellNote = errSoXoa  'static
                                        .BackColor = mErrorColor
                                        iFlagSoXoa = 1
                                        checkHuy = True
                                    End If
                                Else
                                        .Col = .ColLetterToNumber(Col_V)
                                        .CellNote = errSaiDD  'static
                                        .BackColor = mErrorColor
                                        iFlagSaiDD = 1
                                End If
                            End If
                       ElseIf strSLhuy <> "" And Not blnTemp Then
                            .Col = .ColLetterToNumber(Col_V)
                            .CellNote = errSaiDD  'static
                            .BackColor = mErrorColor
                            iFlagSaiDD = 1
                       End If
                       ' kiem tra trung khoang xoa, mat , huy
                        ' 1 trung Xoa, Mat
                        ' 2 trung Xoa, Huy
                        ' 3 trung Mat, huy
                        ' 4 trung ca 3
                        If (strSLXoa <> "" Or strSLMat <> "" Or strSLhuy <> "") And checkXoa = False And checkHuy = False And checkMat = False Then
                            result = checkTrungXoa(CStr(strSLXoa), CStr(strSLMat), CStr(strSLhuy))
                            If result = "1" Then
                                .Col = .ColLetterToNumber(Col_R)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_T)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                iFlagSoXoa = 1
                            ElseIf result = "2" Then
                                .Col = .ColLetterToNumber(Col_R)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_V)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                iFlagSoXoa = 1
                            ElseIf result = "3" Then
                                .Col = .ColLetterToNumber(Col_T)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_V)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                iFlagSoXoa = 1
                            ElseIf result = "4" Then
                                .Col = .ColLetterToNumber(Col_T)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_R)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_V)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                iFlagSoXoa = 1
                            
                            End If
                        End If
                    End If
                ElseIf str(0) = "3" Then
                ' truong hop co ca so ton va so phat hanh
                    '6,7
                    If str(1) = "1" Then
                        .Col = .ColLetterToNumber(Col_H)
                        .CellNote = errEmty  'static
                        .BackColor = mErrorColor
                        iFlagEmty = 1
                    ElseIf str(1) = "2" Then
                        .Col = .ColLetterToNumber(Col_I)
                        .CellNote = errEmty  'static
                        .BackColor = mErrorColor
                        iFlagEmty = 1
                    ElseIf str(1) = "3" Then
                        ' tu so
                        .Col = .ColLetterToNumber(Col_H)
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        ' den so
                        .Col = .ColLetterToNumber(Col_I)
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        iFlagSoLuong = 1
                    End If
                    ' 8,9
                    If str(2) = "1" Then
                        .Col = .ColLetterToNumber(Col_J)
                        .CellNote = errEmty  'static
                        .BackColor = mErrorColor
                        iFlagEmty = 1
                    ElseIf str(2) = "2" Then
                        .Col = .ColLetterToNumber(Col_K)
                        .CellNote = errEmty  'static
                        .BackColor = mErrorColor
                        iFlagEmty = 1
                    ElseIf str(2) = "3" Then
                        ' tu so
                        .Col = .ColLetterToNumber(Col_J)
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        ' den so
                        .Col = .ColLetterToNumber(Col_K)
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        iFlagSoLuong = 1
                    End If
                    ' check lien khoang
                    If str(3) = "1" And strDenSo7 <> "0000000" Then
                        ' den so 7
                        .Col = .ColLetterToNumber(Col_I)
                        .CellNote = errLienKhoang  'static
                        .BackColor = mErrorColor
                        ' tu so 8
                        .Col = .ColLetterToNumber(Col_J)
                        .CellNote = errLienKhoang  'static
                        .BackColor = mErrorColor
                        iFlagLienKhoang = 1
                    End If
                    '
                    If str(2) = "4" And str(1) = "4" Then
                        ' kiem tra so hoa don xoa
                        blnTemp = isDuLieuHL(CStr(strSLXoa))
                        If strSLXoa <> "" And blnTemp Then
                            If strSLXoa <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                                If blnTemp Then
                                    ' du lieu hop le moi check
                                    result = checkHDXoa(CStr(strTuSo10), CStr(strDenSo11), CStr(strSLXoa))
                                    If result = "-1" Then
                                        .Col = .ColLetterToNumber(Col_R)
                                        .CellNote = errSoXoa  'static
                                        .BackColor = mErrorColor
                                        iFlagSoXoa = 1
                                        checkXoa = True
                                    End If
                                Else
                                        .Col = .ColLetterToNumber(Col_R)
                                        .CellNote = errSaiDD  'static
                                        .BackColor = mErrorColor
                                        iFlagSaiDD = 1
                                End If
                            End If
                        ElseIf strSLXoa <> "" And Not blnTemp Then
                            .Col = .ColLetterToNumber(Col_R)
                            .CellNote = errSaiDD  'static
                            .BackColor = mErrorColor
                            iFlagSaiDD = 1
                        End If
                        ' kiem tra so mat
                        blnTemp = isDuLieuHL(CStr(strSLMat))
                        If strSLMat <> "" And blnTemp Then
                            If strSLMat <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                                If blnTemp Then
                                    result = checkHDXoa(CStr(strTuSo10), CStr(strDenSo11), CStr(strSLMat))
                                    If result = "-1" Then
                                        .Col = .ColLetterToNumber(Col_T)
                                        .CellNote = errSoXoa  'static
                                        .BackColor = mErrorColor
                                        iFlagSoXoa = 1
                                        checkMat = True
                                    End If
                                Else
                                        .Col = .ColLetterToNumber(Col_T)
                                        .CellNote = errSaiDD  'static
                                        .BackColor = mErrorColor
                                        iFlagSaiDD = 1
                                End If
                            End If
                        ElseIf strSLMat <> "" And Not blnTemp Then
                            .Col = .ColLetterToNumber(Col_T)
                            .CellNote = errSaiDD  'static
                            .BackColor = mErrorColor
                            iFlagSaiDD = 1
                        End If
                        ' kiem tra so huy
                        blnTemp = isDuLieuHL(CStr(strSLhuy))
                        If strSLhuy <> "" And blnTemp Then
                            If strSLhuy <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                                If blnTemp Then
                                    result = checkHDXoa(CStr(strTuSo10), CStr(strDenSo11), CStr(strSLhuy))
                                    If result = "-1" Then
                                        .Col = .ColLetterToNumber(Col_V)
                                        .CellNote = errSoXoa  'static
                                        .BackColor = mErrorColor
                                        iFlagSoXoa = 1
                                        checkHuy = True
                                    End If
                                Else
                                        .Col = .ColLetterToNumber(Col_V)
                                        .CellNote = errSaiDD  'static
                                        .BackColor = mErrorColor
                                        iFlagSaiDD = 1
                                End If
                            End If
                       ElseIf strSLhuy <> "" And Not blnTemp Then
                            .Col = .ColLetterToNumber(Col_V)
                            .CellNote = errSaiDD  'static
                            .BackColor = mErrorColor
                            iFlagSaiDD = 1
                       End If
                       ' kiem tra trung khoang xoa, mat , huy
                        ' 1 trung Xoa, Mat
                        ' 2 trung Xoa, Huy
                        ' 3 trung Mat, huy
                        ' 4 trung ca 3
                        If (strSLXoa <> "" Or strSLMat <> "" Or strSLhuy <> "") And checkXoa = False And checkHuy = False And checkMat = False Then
                            result = checkTrungXoa(CStr(strSLXoa), CStr(strSLMat), CStr(strSLhuy))
                            If result = "1" Then
                                .Col = .ColLetterToNumber(Col_R)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_T)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                iFlagSoXoa = 1
                            ElseIf result = "2" Then
                                .Col = .ColLetterToNumber(Col_R)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_V)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                iFlagSoXoa = 1
                            ElseIf result = "3" Then
                                .Col = .ColLetterToNumber(Col_T)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_V)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                iFlagSoXoa = 1
                            ElseIf result = "4" Then
                                .Col = .ColLetterToNumber(Col_T)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_R)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                .Col = .ColLetterToNumber(Col_V)
                                .CellNote = errSoXoa  'static
                                .BackColor = mErrorColor
                                iFlagSoXoa = 1
                            
                            End If
                        End If
                    End If
                End If
                ' kiem tra so xoa nam trong khoang so ton hoac so phat hanh
                
             End If
                ' Kiem tra so luong hd su dung phai nho hon so mua, ton [13] <=[5]
             If Val(strSLSuDung) > Val(strTongSo5) And Trim(strSLSuDung) <> "" Then
                    .Col = .ColLetterToNumber("P")
                    .CellNote = errSLSudung  'static
                    .BackColor = mErrorColor
                    iFlagSLSudung = 1
             End If
                ' kiem tra so luong hd xoa
             If Trim(strSLXoa) <> "" Then
                If Val(GetSoLuongHD(CStr(strSLXoa))) > Val(strTongSo5) Then
                       .Col = .ColLetterToNumber(Col_R)
                       .CellNote = errSLXoa  'static
                       .BackColor = mErrorColor
                       iFlagSLXoa = 1
                End If
             End If
                ' kiem tra so luong hd mat
             If Trim(strSLMat) <> "" Then
                If Val(GetSoLuongHD(CStr(strSLMat))) > Val(strTongSo5) Then
                       .Col = .ColLetterToNumber(Col_T)
                       .CellNote = errSLMat  'static
                       .BackColor = mErrorColor
                       iFlagSLMat = 1
                End If
             End If
                ' kiem tra so luong hd huy
             If Trim(strSLhuy) <> "" Then
                If Val(GetSoLuongHD(CStr(strSLhuy))) > Val(strTongSo5) Then
                       .Col = .ColLetterToNumber(Col_V)
                       .CellNote = errSLHuy  'static
                       .BackColor = mErrorColor
                       iFlagSLHuy = 1
                End If
             End If
                ' kiem tra so luong sd, mat, huy , xoa [12] <=[5]
             If Val(strTongSo12) > Val(strTongSo5) Then
                    If Trim(strSLSuDung) <> "" Then
                        .Col = .ColLetterToNumber("P")
                        .CellNote = errSLSD_Xoa_Mat_Huy  'static
                        .BackColor = mErrorColor
                    End If
                    If Trim(strSLXoa) <> "" Then
                        .Col = .ColLetterToNumber(Col_R)
                        .CellNote = errSLSD_Xoa_Mat_Huy  'static
                        .BackColor = mErrorColor
                    End If
                    If Trim(strSLMat) <> "" Then
                        .Col = .ColLetterToNumber(Col_T)
                        .CellNote = errSLSD_Xoa_Mat_Huy  'static
                        .BackColor = mErrorColor
                    End If
                    If Trim(strSLhuy) <> "" Then
                        .Col = .ColLetterToNumber(Col_V)
                        .CellNote = errSLSD_Xoa_Mat_Huy  'static
                        .BackColor = mErrorColor
                    End If
                    iFlagSD_Xoa_Mat_Huy = 1
             End If
                ' Kiem tra trung nhau tu so den so
             If Trim(strTuSo6) <> "" And Trim(strDenSo7) <> "" Then
                If UBound(strTrung67) > 0 Then
                    For i = 0 To UBound(strTrung67)
                        If j = strTrung67(i) Then
                            .Col = .ColLetterToNumber(Col_H)
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = errTrungKhoang     'static
                            .BackColor = mErrorColor
                            
                             .Col = .ColLetterToNumber(Col_I)
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = errTrungKhoang     'static
                            .BackColor = mErrorColor
                            iFlagKhoangTSDSTrung = 1
                            Exit For
                        End If
                    Next i
                End If
             End If
             
             If Trim(strTuso8) <> "" And Trim(strDenso9) <> "" Then
                If UBound(strTrung89) > 0 Then
                    For i = 0 To UBound(strTrung89)
                        If j = strTrung89(i) Then
                            .Col = .ColLetterToNumber(Col_J)
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = errTrungKhoang     'static
                            .BackColor = mErrorColor
                            
                             .Col = .ColLetterToNumber(Col_K)
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = errTrungKhoang     'static
                            .BackColor = mErrorColor
                            iFlagKhoangTSDSTrung = 1
                            Exit For
                        End If
                    Next i
                End If
             End If
                
                    '--------------------------------------------------
             If Trim(strTuSo20) <> "" And Trim(strDenSo21) <> "" Then
                If UBound(strTrung2021) > 0 Then
                    For i = 0 To UBound(strTrung2021)
                        If j = strTrung2021(i) Then
                            .Col = .ColLetterToNumber("X")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = errTrungKhoang     'static
                            .BackColor = mErrorColor
                            
                             .Col = .ColLetterToNumber("Y")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = errTrungKhoang     'static
                            .BackColor = mErrorColor
                            iFlagKhoangTSDSTrung = 1
                            Exit For
                        End If
                    Next i
                End If
             End If
                
                    '--------------------------------------------------
            .Row = j
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
            ' Check cho den het cac dong co du lieu thi thoi
        Loop Until UCase(.Text) = Col_AA
        
         ' Kiem tra du lieu nguoi dai dien
         .GetText .ColLetterToNumber(Col_N), countDynamicRow + 29, strNguoiKy
         .GetText .ColLetterToNumber(Col_E), countDynamicRow + 29, strNguoiLapBieu
         .GetText .ColLetterToNumber(Col_N), countDynamicRow + 31, strNgayKy
         ' kiem tra nguoi ky bat buoc nhap
         If Trim(strNguoiKy) = "" Then
            .Row = countDynamicRow + 29
            .Col = .ColLetterToNumber(Col_N)
            .CellNote = errEmty  'static
            .BackColor = mErrorColor
            iFlagEmty = 1
         Else
            .Row = countDynamicRow + 29
            .Col = .ColLetterToNumber(Col_N)
            .CellNote = ""
            .BackColor = mNonErrorColor
         End If
         'kiem tra ngay ky
         If Trim(strNgayKy) = "" Or Trim(strNgayKy) = "../../...." Then
            .Row = countDynamicRow + 31
            .Col = .ColLetterToNumber(Col_N)
            .CellNote = errEmty  'static
            .BackColor = mErrorColor
            iFlagEmty = 1
         Else
                strNgayKy = DateSerial(CInt(Mid$(strNgayKy, 7, 4)), CInt(Mid$(strNgayKy, 4, 2)), CInt(Mid$(strNgayKy, 1, 2)))
                If strNgayKy > Now Then
                    .Col = .ColLetterToNumber(Col_N)
                    .Row = countDynamicRow + 31
                    .CellNote = errNgayKy  'static
                    .BackColor = mErrorColor
                    iFlagNgayKy = 1
                ElseIf strNgayKy < DateSerial(CInt(Mid$(vKyBCTuNgay, 7, 4)), CInt(Mid$(vKyBCTuNgay, 4, 2)), CInt(Mid$(vKyBCTuNgay, 1, 2))) Then
                    .Col = .ColLetterToNumber(Col_N)
                    .Row = countDynamicRow + 31
                    .CellNote = errNgayKyHL  'static
                    .BackColor = mErrorColor
                    iFlagNgayKyHL = 1
                    
                Else
                    .Row = countDynamicRow + 31
                    .Col = .ColLetterToNumber(Col_N)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
         End If
        
        
         ' group so hoa don vao thanh tung nhom
         ' Tam thoi check sau se sua lai doan nay
         .Row = 22
         .Col = .ColLetterToNumber("B")
         Dim arrGroupValue() As String  ' Luu thong tin cua tung dong check
         Dim coutCheck, countCheckTrung As Long
         Dim arrGroupHDNew As String
         Dim vitriGhep As Long
         'Dim isTrung As Boolean
         Dim strGroupHDTS_DS As String
         Dim arrGroupHDTS_DS() As String
         Dim arrGroupValueTS_DS() As String
         Dim strTuSo, strDenso As Variant
         Dim strTempTS_DS As String
         Dim arrGroupHDNewTS_DS As String
         
         If blCheck_S2 = True Or blCheck_S3 = True Then
             Do
                .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
                .GetText .ColLetterToNumber(Col_E), .Row, strMauSoHD
                .GetText .ColLetterToNumber(Col_F), .Row, strKyHieuHD
                .GetText .ColLetterToNumber(Col_Z), .Row, strSoLuong
                
     
                
                If .Row = 22 Then
                    strGroupHD = Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~0"
                    'strGroupHDTS_DS = Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~" & getListHDTS_DS(CStr(strTuSo), CStr(strDenso))
                    strGroupHDTS_DS = Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~"
                Else
                    arrGroupHD = Split(strGroupHD, "###")
                    coutCheck = 0
                    countCheckTrung = 0
                    For i = 0 To UBound(arrGroupHD)
                        coutCheck = coutCheck + 1
                        arrGroupValue = Split(arrGroupHD(i), "~")
                        If Trim(strTenHD) = arrGroupValue(0) And Trim(strMauSoHD) = arrGroupValue(1) And Trim(strKyHieuHD) = arrGroupValue(2) Then
                            Exit For
                        Else
                            countCheckTrung = countCheckTrung + 1
                        End If
                    Next i
                    If coutCheck = countCheckTrung Then
                        strGroupHD = strGroupHD & "###" & Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~0"
                        'strGroupHDTS_DS = strGroupHDTS_DS & "###" & Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~" & getListHDTS_DS(CStr(strTuSo), CStr(strDenso))
                         strGroupHDTS_DS = strGroupHDTS_DS & "###" & Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~"
                    End If
                End If
                .Row = .Row + 1
             Loop Until .Text = "aa"
             
             'test
             Dim hoaDon As HOA_DON
             Dim tgHoaDon As HOA_DON
             Dim arrHoaDon() As HOA_DON
             Dim arrTempHoaDon() As HOA_DON
             Dim loaiHoaDon() As LOAI_HOA_DON
             Dim index As Long
             Dim vitri As Long
             .Row = 22
             .Col = .ColLetterToNumber("B")
             index = 0
             arrGroupHD = Split(strGroupHD, "###")
             ReDim Preserve loaiHoaDon(UBound(arrGroupHD))
             For i = 0 To UBound(arrGroupHD)
                ReDim Preserve arrHoaDon(0)
                loaiHoaDon(i).loaiHD = arrGroupHD(i)
                loaiHoaDon(i).hoaDon = arrHoaDon
             Next i
             
             Do
                .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
                .GetText .ColLetterToNumber(Col_E), .Row, strMauSoHD
                .GetText .ColLetterToNumber(Col_F), .Row, strKyHieuHD
                .GetText .ColLetterToNumber("X"), .Row, strTuSo
                .GetText .ColLetterToNumber("Y"), .Row, strDenso
                hoaDon.tenHD = CStr(strTenHD)
                hoaDon.mauSoHD = CStr(strMauSoHD)
                hoaDon.kyHieuHD = CStr(strKyHieuHD)
                hoaDon.tuSo = Val(CStr(strTuSo))
                hoaDon.denso = Val(CStr(strDenso))
                'If UBound(loaiHoaDon) > 0 Then
                For i = 0 To UBound(loaiHoaDon)
                    If loaiHoaDon(i).loaiHD = hoaDon.tenHD & "~" & hoaDon.mauSoHD & "~" & hoaDon.kyHieuHD & "~0" Then
                        arrTempHoaDon = loaiHoaDon(i).hoaDon
                        If Trim(arrTempHoaDon(0).mauSoHD) = "" Then
                            vitri = 0
                        Else
                            vitri = UBound(arrTempHoaDon) + 1
                            ReDim Preserve arrTempHoaDon(vitri)
                        End If
                        arrTempHoaDon(vitri) = hoaDon
                        loaiHoaDon(i).hoaDon = arrTempHoaDon
                    End If
                Next i
    '            Else
    '
    '            End If
                .Row = .Row + 1
             Loop Until .Text = "aa"
             ' Sap xep theo thu tu tang dan trong tung nhom hoa don
             'If UBound(loaiHoaDon) > 0 Then
            For i = 0 To UBound(loaiHoaDon)
                arrTempHoaDon = loaiHoaDon(i).hoaDon
                If UBound(arrTempHoaDon) > 0 Then
                    For j = 0 To UBound(arrTempHoaDon) - 1
                        For k = j + 1 To UBound(arrTempHoaDon)
                            If arrTempHoaDon(j).tuSo > arrTempHoaDon(k).tuSo Then
                                tgHoaDon = arrTempHoaDon(j)
                                arrTempHoaDon(j) = arrTempHoaDon(k)
                                arrTempHoaDon(k) = tgHoaDon
                            End If
                        Next k
                    Next j
                End If
                loaiHoaDon(i).hoaDon = arrTempHoaDon
            Next i
             'End If
             ' Tim cac khoang lien nhau trong cung mot nhom
             Dim tmpArrHoaDon() As HOA_DON
             'If UBound(loaiHoaDon) > 0 Then
            For i = 0 To UBound(loaiHoaDon)
                arrTempHoaDon = loaiHoaDon(i).hoaDon
                If UBound(arrTempHoaDon) > 0 Then
                    For j = 0 To UBound(arrTempHoaDon) - 1
                        For k = j + 1 To UBound(arrTempHoaDon)
                            If arrTempHoaDon(j).denso + 1 = arrTempHoaDon(k).tuSo Then
                               arrTempHoaDon(j).isDelete = True
                               arrTempHoaDon(k).tuSo = arrTempHoaDon(j).tuSo
                               Exit For
                            End If
                        Next k
                    Next j
                End If
                loaiHoaDon(i).hoaDon = arrTempHoaDon
            Next i
             'End If
             
             loaiHoaDonS1 = loaiHoaDon
         End If
         
         ' end test
         ' Tinh tong cac so hoa don cung 1 loai
         ' Kiem tra neu co phu luc moi check so luong HD ben phu luc va to khai
'         If blCheck_S2 = True Or blCheck_S3 = True Then
'            .Row = 22
'            .Col = .ColLetterToNumber("B")
'             Do
'                .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
'                .GetText .ColLetterToNumber(Col_E), .Row, strMauSoHD
'                .GetText .ColLetterToNumber(Col_F), .Row, strKyHieuHD
'                .GetText .ColLetterToNumber(Col_Z), .Row, strSoLuong
'
'                .GetText .ColLetterToNumber("X"), .Row, strTuSo
'                .GetText .ColLetterToNumber("Y"), .Row, strDenso
'
'                arrGroupHD = Split(strGroupHD, "###")
'                arrGroupHDNew = ""
'
'                arrGroupHDTS_DS = Split(strGroupHDTS_DS, "###")
'                arrGroupHDNewTS_DS = ""
'                If UBound(arrGroupHD) > 0 Then
'                    For i = 0 To UBound(arrGroupHD)
'                        arrGroupValue = Split(arrGroupHD(i), "~")
'                        '
'                        arrGroupValueTS_DS = Split(arrGroupHDTS_DS(i), "~")
'                        If Trim(strTenHD) = arrGroupValue(0) And Trim(strMauSoHD) = arrGroupValue(1) And Trim(strKyHieuHD) = arrGroupValue(2) Then
'                            arrGroupValue(3) = Val(arrGroupValue(3)) + Val(strSoLuong)
'                            arrGroupHDNew = arrGroupValue(0) & "~" & arrGroupValue(1) & "~" & arrGroupValue(2) & "~" & arrGroupValue(3)
'                            vitriGhep = i
'                            ' ghep lai cac gia tri hd voi loai va ky hieu
'                            'strTempTS_DS = strTempTS_DS & getListHDTS_DS(CStr(strTuSo), CStr(strDenso))
'                            arrGroupValueTS_DS(3) = arrGroupValueTS_DS(3) & getListHDTS_DS(CStr(strTuSo), CStr(strDenso))
'                            arrGroupHDNewTS_DS = Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~" & arrGroupValueTS_DS(3)
'                            Exit For
'                        End If
'                    Next i
'                    ' ghep lai chuoi moi khi da cong du lieu
'                    strGroupHD = ""
'                    strGroupHDTS_DS = ""
'                    For i = 0 To UBound(arrGroupHD)
'                        If i = vitriGhep Then
'                            strGroupHD = strGroupHD & "###" & arrGroupHDNew
'                            '
'                            strGroupHDTS_DS = strGroupHDTS_DS & "###" & arrGroupHDNewTS_DS
'                        Else
'                            strGroupHD = strGroupHD & "###" & arrGroupHD(i)
'                            '
'                            strGroupHDTS_DS = strGroupHDTS_DS & "###" & arrGroupHDTS_DS(i)
'                        End If
'                    Next i
'                    strGroupHD = Mid$(strGroupHD, 4)
'                    '
'                    strGroupHDTS_DS = Mid$(strGroupHDTS_DS, 4)
'                Else
'                    arrGroupValue = Split(arrGroupHD(0), "~")
'                    arrGroupValue(3) = Val(arrGroupValue(3)) + Val(strSoLuong)
'                    strGroupHD = arrGroupValue(0) & "~" & arrGroupValue(1) & "~" & arrGroupValue(2) & "~" & arrGroupValue(3)
'
'                    ' Tinh lai gia tri cua tu so den so
'                    arrGroupValueTS_DS = Split(arrGroupHDTS_DS(0), "~")
'                    'strTempTS_DS = strTempTS_DS & getListHDTS_DS(CStr(strTuSo), CStr(strDenso))
'                    arrGroupValueTS_DS(3) = arrGroupValueTS_DS(3) & getListHDTS_DS(CStr(strTuSo), CStr(strDenso))
'                    strGroupHDTS_DS = Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~" & arrGroupValueTS_DS(3)
'                End If
'
'
'                .Row = .Row + 1
'            Loop Until .Text = "aa"
'         End If
'         strCheckTon = strGroupHD
'         strCheckTonKhoang = strGroupHDTS_DS
      
         ' set lai trang thai cua HDR
         .Sheet = .SheetCount
         .SetText .ColLetterToNumber("B"), 13, IIf(iFlagEmty = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 14, IIf(iFlagKhoangTSDSTrung = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 15, IIf(iFlagKyHieu = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 16, IIf(iFlagSoLuong = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 17, IIf(iFlagMauSoHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 18, IIf(iFlagSoLienHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 19, IIf(iFlagNgayKy = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 20, IIf(iFlagLengHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 21, IIf(iFlagLienKhoang = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 22, IIf(iFlagSoXoa = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 23, IIf(iFlagSaiDD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 24, IIf(iFlagSLSudung = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 25, IIf(iFlagSD_Xoa_Mat_Huy = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 26, IIf(iFlagSLXoa = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 27, IIf(iFlagSLMat = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 28, IIf(iFlagSLHuy = 1, "0", "1")
        '
        .SetText .ColLetterToNumber("B"), 52, IIf(iFlagNgayKyHL = 1, "0", "1")
        
        .SetText .ColLetterToNumber("B"), 53, IIf(iFlagSoLienHD = 2, "0", "1")
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
Private Sub checkSheet2()
    Dim strCheck As String
    Dim errMauSoHD As Variant, errKyHieuHD As Variant, errSoLuong As Variant, errSoLienHD As Variant, errTon As Variant, errTonTK As Variant, errSoLienHDDT As Variant
    Dim errKhoangGN As Variant, errLengthHD As Variant, errEmty As Variant, errNgayKy As Variant, errNgayHuy As Variant, errTrungKhoang As Variant
    Dim strTenHD As Variant, strMauSoHD As Variant, strKyHieuHD As Variant, strTuSo As Variant, strDenso As Variant
    Dim strTenHD1 As Variant, strMauSoHD1 As Variant, strKyHieuHD1 As Variant, strTuso1 As Variant, strDenso1 As Variant
    Dim strSoLuong As Variant, strNguoiKy As Variant, strNguoiLapBieu As Variant, strNgayKy As Variant, strNgayHuy As Variant
    Dim iFlagTenHD As Integer, iFlagMauSoHD As Integer, iFlagKyHieu As Integer, iFlagSoLienHD As Integer, iFlagEmty As Integer
    Dim iWhiteRow As Integer, iFlagLengHD As Integer, iFlagSoLuong As Integer, iFlagNgayKy As Integer, iFlagNgayHuy As Integer, iFlagTon As Integer, iFlagTonTK As Integer
    Dim iFlagKhoangTSDSTrung As Integer, iFocusFlag As Integer
    
    Dim strId As Variant, strLoai As String, strTemp As String, result As String
    Dim i As Integer, countDynamicRow As Integer, j As Integer, k As Integer
    
    ' group theo tung nhom
    Dim arrGroupHD() As String
    Dim strGroupHD As String
    Dim arrGroupHD_TK() As String
    Dim arrGroupValue_TK() As String
    
    ' Danh sach cac hoa don ton tren sheet 1
'    Dim arrGroupHD_Ton() As String
'    Dim arrHD_Ton() As String
'
    
    Dim str() As String
    Dim str1 As String
    ' tra ve mang cac row bi trung khoang
    str = checkTrung_01(fps, 46, Col_V, Col_Y, "B", "C", Col_O, Col_S, "", False, 2)
    ' check trung ton
    ' str1 = checkTon(fps, 46, Col_V, Col_Y, "B", "D", Col_O, Col_S, "", False, 2)
    'arrGroupHD_Ton = Split(strCheckTonKhoang, "###")
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber(Col_E), 32, errMauSoHD
        .GetText .ColLetterToNumber(Col_E), 33, errKyHieuHD
        .GetText .ColLetterToNumber(Col_E), 34, errSoLienHD
        .GetText .ColLetterToNumber(Col_E), 35, errKhoangGN
        .GetText .ColLetterToNumber(Col_E), 36, errLengthHD
        .GetText .ColLetterToNumber(Col_E), 37, errSoLuong
        .GetText .ColLetterToNumber(Col_E), 38, errEmty
        .GetText .ColLetterToNumber(Col_E), 39, errNgayHuy
        .GetText .ColLetterToNumber(Col_E), 40, errNgayKy
        .GetText .ColLetterToNumber(Col_E), 41, errTrungKhoang

        .GetText .ColLetterToNumber(Col_E), 50, errTon
        .GetText .ColLetterToNumber(Col_E), 48, errTonTK
        
        .GetText .ColLetterToNumber(Col_E), 54, errSoLienHDDT
        
        'set flag bang 0
        iFlagTenHD = 0
        iFlagMauSoHD = 0
        iFlagKyHieu = 0
        iFlagSoLienHD = 0
        iFlagEmty = 0
        iFlagSoLuong = 0
        iFlagNgayKy = 0
        iFlagNgayHuy = 0
        iFlagKhoangTSDSTrung = 0
        
 
        
        ' dem so row dynamic
        countDynamicRow = 0
        ' Kiem tra tren sheet 1
        .Sheet = 2
               ' re set cac cellnote
        .Row = 46
        .Col = .ColLetterToNumber("B")
        Do
            .Col = .ColLetterToNumber(Col_AA)
            '.Col = .ColLetterToNumber("U")
            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            .CellNote = ""     'static
            .BackColor = mNonErrorColor
            ' cot tu so
            .Col = .ColLetterToNumber(Col_V)
            .CellNote = ""     'static
            .BackColor = mNonErrorColor
            ' cot den so
            .Col = .ColLetterToNumber(Col_Y)
            .CellNote = ""     'static
            .BackColor = mNonErrorColor
            ' cot ten hoa don
            .Col = .ColLetterToNumber(Col_D)
            .CellNote = ""
            .BackColor = mNonErrorColor
            
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until UCase(.Text) = "AA"
        
        
        
        
        .Col = .ColLetterToNumber("B")
        .Row = 46
        'i = 1
        Do
            j = .Row
            countDynamicRow = countDynamicRow + 1
             ' Lay cac gia tri cua 1 row
             iWhiteRow = 1
            .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
            .GetText .ColLetterToNumber(Col_O), .Row, strMauSoHD
            .GetText .ColLetterToNumber(Col_S), .Row, strKyHieuHD
            .GetText .ColLetterToNumber(Col_V), .Row, strTuSo
            .GetText .ColLetterToNumber(Col_Y), .Row, strDenso
            .GetText .ColLetterToNumber(Col_AA), .Row, strSoLuong
            .GetText .ColLetterToNumber("AI"), .Row, strId
            
            ' set gia tri cho check KH NN/XC
            strMauSoHD_01GTKT = Trim(Mid(strId, 11, 6))
            ' end
'
            If Trim(strTenHD) <> "" Or Trim(strMauSoHD) <> "" Or Trim(strKyHieuHD) <> "" Or Trim(strId) <> "" Or Trim(strTuSo) <> "" Or Trim(strDenso) <> "" Then
                ' Khong phai dong trang
                iWhiteRow = 0
            End If
            If Trim(strTenHD) <> "" Or Trim(strMauSoHD) <> "" Then
                ' Khong phai dong trang
                iWhiteRow = 0
            End If
            'Ten HD
            If Trim(strTenHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber(Col_D)
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
'            Else
'                .Col = .ColLetterToNumber(Col_D)
'                .CellNote = ""
'                .BackColor = mNonErrorColor
            End If
            ' Mau so HD
            If Trim(strMauSoHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber(Col_O)
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                strLoai = Trim(Left(strId, 1))
                If strLoai = "0" Then
                    strTemp = Trim(Mid(strId, 11, 6))
                ElseIf strLoai = "4" Then
                    strTemp = Trim(Mid(strId, 11, 6))
                Else
                    strTemp = Trim(Mid(strId, 11, 3))
                End If
                result = CheckMauSoHD(CStr(strMauSoHD), strLoai, strTemp)
                If result = "1" And iWhiteRow = 0 Then
                    .Col = .ColLetterToNumber(Col_O)
                    .CellNote = errMauSoHD  'static
                    .BackColor = mErrorColor
                    iFlagMauSoHD = 1
                ElseIf result = "2" And iWhiteRow = 0 Then
                     If Right$(Trim$(strKyHieuHD), 1) = "E" Then
                        If CheckSoLienHDDT(CStr(strMauSoHD), strLoai, strTemp, CStr(strKyHieuHD)) = "3" Then
                             .Col = .ColLetterToNumber(Col_O)
                            .CellNote = errSoLienHDDT  'static
'                            .BackColor = mErrorColor
'                            iFlagSoLienHD = 2
                             .BackColor = mAlertColor
                        Else
                            .Col = .ColLetterToNumber(Col_O)
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    Else
                        .Col = .ColLetterToNumber(Col_O)
                        .CellNote = errSoLienHD  'static
                        .BackColor = mErrorColor
                        iFlagSoLienHD = 1
                    End If
                Else
                    If Right$(Trim$(strKyHieuHD), 1) = "E" Then
                        If CheckSoLienHDDT(CStr(strMauSoHD), strLoai, strTemp, CStr(strKyHieuHD)) = "3" Then
                             .Col = .ColLetterToNumber(Col_O)
                            .CellNote = errSoLienHDDT  'static
'                            .BackColor = mErrorColor
'                            iFlagSoLienHD = 2
                             .BackColor = mAlertColor
                        Else
                            .Col = .ColLetterToNumber(Col_O)
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    Else
                        .Col = .ColLetterToNumber(Col_O)
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                End If
            End If
            ' check ky hieu
            If Trim(strKyHieuHD) = "" And iWhiteRow = 1 Then
                .Col = .ColLetterToNumber(Col_S)
                .CellNote = ""  'static
                .BackColor = mNonErrorColor
                'iFlagEmty = 1
            ElseIf Trim(strKyHieuHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber(Col_S)
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                result = CheckSoHD(CStr(strKyHieuHD), strLoai)
                If result = "1" Then
                    .Col = .ColLetterToNumber(Col_S)
                    .CellNote = errKyHieuHD  'static
                    .BackColor = mErrorColor
                    iFlagKyHieu = 1
                ElseIf result = "2" Then
                    .Col = .ColLetterToNumber(Col_S)
                    .CellNote = errLengthHD  'static
                    .BackColor = mErrorColor
                    iFlagLengHD = 1
                Else
                    .Col = .ColLetterToNumber(Col_S)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            End If
            ' kiem tra tu so den so
            If iWhiteRow = 0 Then
                If Trim(strTuSo) = "" And Trim(strDenso) = "" Then
                    ' tu so
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    ' den so
                    .Col = .ColLetterToNumber(Col_Y)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                ElseIf Trim(strTuSo) = "" And Trim(strDenso) <> "" Then
                    ' tu so
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                    ' den so
                    .Col = .ColLetterToNumber(Col_Y)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                ElseIf Trim(strTuSo) <> "" And Trim(strDenso) = "" Then
                    .Col = .ColLetterToNumber(Col_Y)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                    ' tu so
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNote = ""
                    .BackColor = mNonErrorColor

                Else
                    If (Val(strDenso) - Val(strTuSo) >= 0 And Val(strDenso) > 0) Or (Val(strTuSo) = 0 And Val(strDenso) = 0) Then
                        ' tu so
                        .Col = .ColLetterToNumber(Col_V)
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        ' den so
                        .Col = .ColLetterToNumber(Col_Y)
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    Else
                        .Col = .ColLetterToNumber(Col_V)
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        .Col = .ColLetterToNumber(Col_Y)
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        iFlagSoLuong = 1
                    End If
                End If
            Else
                ' tu so
                .Col = .ColLetterToNumber(Col_V)
                .CellNote = ""
                .BackColor = mNonErrorColor
                ' den so
                .Col = .ColLetterToNumber(Col_Y)
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            ' kiem tra trung khoang HD
            If UBound(str) > 0 And iWhiteRow = 0 Then
                For i = 0 To UBound(str)
                    If j = str(i) Then
                        .Col = .ColLetterToNumber(Col_V)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errTrungKhoang     'static
                        .BackColor = mErrorColor
                        
                         .Col = .ColLetterToNumber(Col_Y)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errTrungKhoang     'static
                        .BackColor = mErrorColor
                        iFlagKhoangTSDSTrung = 1
                        Exit For
                    End If
                Next i
            End If
            ' end check trung
            
                'dhdang so luong ton cuoi ky
            ' kiem tra trung khoang HD
            'If InStr(str1, j) <> 0 And iWhiteRow = 0 Then
'            Dim strListHD As String
'            Dim arrHD_PL2() As String
'            Dim arrHD_BC() As String
'            Dim idxCheckTon As Double
'            Dim isCheck As Boolean
'            strListHD = getListHDTS_DS(CStr(strTuSo), CStr(strDenso))
'            arrHD_PL2 = Split(Mid(strListHD, 1, Len(strListHD) - 1), ";")
'            isCheck = False
'            For i = 0 To UBound(arrHD_PL2)
'                For idxCheckTon = 0 To UBound(arrGroupHD_Ton)
'                    arrHD_BC = Split(arrGroupHD_Ton(idxCheckTon), "~")
'                    If iWhiteRow = 0 And Trim(strTenHD) = arrHD_BC(0) And Trim(strMauSoHD) = arrHD_BC(1) And Trim(strKyHieuHD) = arrHD_BC(2) Then
'                        'For i = 0 To UBound(str1)
'                            'If j = str1(i) Then
'                              If InStr(";" & arrHD_BC(3), ";" & arrHD_PL2(i) & ";") = 0 Then
'                                    .Col = .ColLetterToNumber(Col_V)
'                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                    .CellNote = errTonTK     'static
'                                    .BackColor = mErrorColor
'
'                                     .Col = .ColLetterToNumber(Col_Y)
'                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                    .CellNote = errTonTK     'static
'                                    .BackColor = mErrorColor
'                                    iFlagTonTK = 1
'                                    isCheck = True
'                                    Exit For
'                               End If
'                        'Next i
'                    End If
'                Next idxCheckTon
'                If isCheck = True Then Exit For
'            Next i
            ' test
            Dim arrHoaDonS1() As HOA_DON
            Dim isTrongKhoang As Boolean ' Kiem tra tu so den so thuoc khoang 20,21
            Dim isNotCheckTrongKhoang As Boolean
            Dim isHoaDonS1 As Boolean  ' la hoa don ben TK
            Dim isExitFor As Boolean  ' thoat khoi vong lap
            
            'If UBound(loaiHoaDonS1) > 0 Then
            isTrongKhoang = False
            isHoaDonS1 = False
            isExitFor = False
            For i = 0 To UBound(loaiHoaDonS1)
                If loaiHoaDonS1(i).loaiHD = Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~0" And iWhiteRow = 0 Then
                    isHoaDonS1 = True
                    arrHoaDonS1 = loaiHoaDonS1(i).hoaDon
                    'If UBound(arrHoaDonS1) > 0 Then
                    For k = 0 To UBound(arrHoaDonS1)
                        If Val(strTuSo) >= arrHoaDonS1(k).tuSo And Val(strDenso) <= arrHoaDonS1(k).denso And arrHoaDonS1(k).isDelete = False Then
                            isTrongKhoang = True
                            isExitFor = True
                            Exit For
                        End If
                    Next k
                    'End If
                End If
                If isExitFor = True Then
                    Exit For
                End If
            Next i
            'End If
            ' end test
            If iWhiteRow = 0 Then
                If isHoaDonS1 = True Then
'                    .Col = .ColLetterToNumber(Col_V)
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = ""     'static
'                    .BackColor = mNonErrorColor
'
'                     .Col = .ColLetterToNumber(Col_Y)
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = ""     'static
'                    .BackColor = mNonErrorColor
                    If isNotCheckTrongKhoang = False Then
                        If isTrongKhoang = True Then
'                            .Col = .ColLetterToNumber(Col_V)
'                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                            .CellNote = ""     'static
'                            .BackColor = mNonErrorColor
'
'                             .Col = .ColLetterToNumber(Col_Y)
'                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                            .CellNote = ""     'static
'                            .BackColor = mNonErrorColor
                        Else
            
                            .Col = .ColLetterToNumber(Col_V)
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = errTonTK     'static
                            .BackColor = mErrorColor
        
                             .Col = .ColLetterToNumber(Col_Y)
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = errTonTK     'static
                            .BackColor = mErrorColor
                            iFlagTonTK = 1
                        End If
                     End If
                Else
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = errTon     'static
                    .BackColor = mErrorColor

                     .Col = .ColLetterToNumber(Col_Y)
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = errTon     'static
                    .BackColor = mErrorColor
                    iFlagTon = 1
                End If
            Else
'                .Col = .ColLetterToNumber(Col_V)
'                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                .CellNote = ""     'static
'                .BackColor = mNonErrorColor
'
'                 .Col = .ColLetterToNumber(Col_Y)
'                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                .CellNote = ""     'static
'                .BackColor = mNonErrorColor
            End If
            
            

            .Row = j
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")

            If UCase(.Text) = "AA" Then
                isNotCheckTrongKhoang = True
                .Row = .Row + 3
            End If
        Loop Until .Text = "bb"
                
        ' Kiem tra du lieu nguoi dai dien
         .GetText .ColLetterToNumber(Col_R), countDynamicRow + 54, strNguoiKy
         .GetText .ColLetterToNumber(Col_D), countDynamicRow + 54, strNguoiLapBieu
         .GetText .ColLetterToNumber(Col_R), countDynamicRow + 56, strNgayKy
         .GetText .ColLetterToNumber(Col_O), 8, strNgayHuy
         ' kiem tra nguoi ky bat buoc nhap
         If Trim(strNguoiKy) = "" Then
            .Row = countDynamicRow + 54
            .Col = .ColLetterToNumber(Col_R)
            .CellNote = errEmty  'static
            .BackColor = mErrorColor
            iFlagEmty = 1
         Else
            .Row = countDynamicRow + 54
            .Col = .ColLetterToNumber(Col_R)
            .CellNote = ""
            .BackColor = mNonErrorColor
         End If
         'kiem tra ngay ky
         If Trim(strNgayKy) = "" Or Trim(strNgayKy) = "../../...." Then
            .Row = countDynamicRow + 56
            .Col = .ColLetterToNumber(Col_R)
            .CellNote = errEmty  'static
            .BackColor = mErrorColor
            iFlagEmty = 1
         Else
                strNgayKy = DateSerial(CInt(Mid$(strNgayKy, 7, 4)), CInt(Mid$(strNgayKy, 4, 2)), CInt(Mid$(strNgayKy, 1, 2)))
                If strNgayKy > Now Then
                    .Col = .ColLetterToNumber(Col_R)
                    .Row = countDynamicRow + 56
                    .CellNote = errNgayKy  'static
                    .BackColor = mErrorColor
                    iFlagNgayKy = 1
                Else
                    .Row = countDynamicRow + 56
                    .Col = .ColLetterToNumber(Col_R)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
         End If
         
         ' group so hoa don vao thanh tung nhom
         
'         .Row = 46
'         .Col = .ColLetterToNumber("B")
'         Dim arrGroupValue() As String  ' Luu thong tin cua tung dong check
'         Dim coutCheck, countCheckTrung As Long
'         Dim arrGroupHDNew As String
'         Dim vitriGhep As Long
'
'         Dim strGroupHDTS_DS As String
'         'Dim isTrung As Boolean
'
'         Do
'            .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
'            .GetText .ColLetterToNumber(Col_O), .Row, strMauSoHD
'            .GetText .ColLetterToNumber(Col_S), .Row, strKyHieuHD
'            .GetText .ColLetterToNumber(Col_V), .Row, strTuSo
'            .GetText .ColLetterToNumber(Col_Y), .Row, strDenso
'            .GetText .ColLetterToNumber(Col_AA), .Row, strSoLuong
'            .GetText .ColLetterToNumber("AI"), .Row, strId
'            If .Row = 46 Then
'                strGroupHD = Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~0"
'                strGroupHDTS_DS = Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & getListHDTS_DS(CStr(strTuSo), CStr(strDenso))
'            Else
'                arrGroupHD = Split(strGroupHD, "###")
'                coutCheck = 0
'                countCheckTrung = 0
'                For i = 0 To UBound(arrGroupHD)
'                    coutCheck = coutCheck + 1
'                    arrGroupValue = Split(arrGroupHD(i), "~")
'                    If Trim(strTenHD) = arrGroupValue(0) And Trim(strMauSoHD) = arrGroupValue(1) And Trim(strKyHieuHD) = arrGroupValue(2) Then
'                        Exit For
'                    Else
'                        countCheckTrung = countCheckTrung + 1
'                    End If
'                Next i
'                If coutCheck = countCheckTrung Then
'                    strGroupHD = strGroupHD & "###" & Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~0"
'                    strGroupHDTS_DS = strGroupHDTS_DS & "###" & Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & getListHDTS_DS(CStr(strTuSo), CStr(strDenso))
'                End If
'            End If
'            .Row = .Row + 1
'         Loop Until .Text = "aa"
'         ' Tinh tong cac so hoa don cung 1 loai
'            .Row = 46
'            .Col = .ColLetterToNumber("B")
'             Do
'            .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
'            .GetText .ColLetterToNumber(Col_O), .Row, strMauSoHD
'            .GetText .ColLetterToNumber(Col_S), .Row, strKyHieuHD
'            .GetText .ColLetterToNumber(Col_V), .Row, strTuSo
'            .GetText .ColLetterToNumber(Col_Y), .Row, strDenso
'            .GetText .ColLetterToNumber(Col_AA), .Row, strSoLuong
'            .GetText .ColLetterToNumber("AI"), .Row, strId
'            arrGroupHD = Split(strGroupHD, "###")
'            arrGroupHDNew = ""
'            If UBound(arrGroupHD) > 0 Then
'                For i = 0 To UBound(arrGroupHD)
'                    arrGroupValue = Split(arrGroupHD(i), "~")
'                    If Trim(strTenHD) = arrGroupValue(0) And Trim(strMauSoHD) = arrGroupValue(1) And Trim(strKyHieuHD) = arrGroupValue(2) Then
'                        arrGroupValue(3) = Val(arrGroupValue(3)) + Val(strSoLuong)
'                        arrGroupHDNew = arrGroupValue(0) & "~" & arrGroupValue(1) & "~" & arrGroupValue(2) & "~" & arrGroupValue(3)
'                        vitriGhep = i
'                        Exit For
'                    End If
'                Next i
'                ' ghep lai chuoi moi khi da cong du lieu
'                strGroupHD = ""
'                For i = 0 To UBound(arrGroupHD)
'                    If i = vitriGhep Then
'                        strGroupHD = strGroupHD & "###" & arrGroupHDNew
'                    Else
'                        strGroupHD = strGroupHD & "###" & arrGroupHD(i)
'                    End If
'                Next i
'                strGroupHD = Mid$(strGroupHD, 4)
'            Else
'                arrGroupValue = Split(arrGroupHD(0), "~")
'                arrGroupValue(3) = Val(arrGroupValue(3)) + Val(strSoLuong)
'                strGroupHD = arrGroupValue(0) & "~" & arrGroupValue(1) & "~" & arrGroupValue(2) & "~" & arrGroupValue(3)
'            End If
'
'
'            .Row = .Row + 1
'         Loop Until .Text = "aa"
'
'         ' Lay ve group cua sheet 1
'         Dim isNotExistBC As Boolean
'         arrGroupHD_TK = Split(strCheckTon, "###")
'         arrGroupHD = Split(strGroupHD, "###")
'        ' If UBound(arrGroupHD) > 0 Then
'            For i = 0 To UBound(arrGroupHD)
'                arrGroupValue = Split(arrGroupHD(i), "~")
'                isNotExistBC = False
'                For j = 0 To UBound(arrGroupHD_TK)
'                    arrGroupValue_TK = Split(arrGroupHD_TK(j), "~")
'                    If arrGroupValue(0) = arrGroupValue_TK(0) And arrGroupValue(1) = arrGroupValue_TK(1) And arrGroupValue(2) = arrGroupValue_TK(2) Then
'                        isNotExistBC = True ' co ton tai trong bao cao
'                        If arrGroupValue(3) <> arrGroupValue_TK(3) Then
'                            .Row = 46
'                            .Col = .ColLetterToNumber("B")
'                            Do
'                                .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
'                                .GetText .ColLetterToNumber(Col_O), .Row, strMauSoHD
'                                .GetText .ColLetterToNumber(Col_S), .Row, strKyHieuHD
'                                If arrGroupValue(0) = Trim(strTenHD) And arrGroupValue(1) = Trim(strMauSoHD) And arrGroupValue(2) = Trim(strKyHieuHD) Then
'                                    .Col = .ColLetterToNumber(Col_AA)
'                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                    .CellNote = errTon     'static
'                                    .BackColor = mErrorColor
'                                    iFlagTon = 1
'                                 End If
'                                 .Row = .Row + 1
'                                 .Col = .ColLetterToNumber("B")
'                            Loop Until UCase(.Text) = "AA"
'                            Exit For ' Thoat vong lap trong
'                        End If
'                    Else
'
''                        .Row = 46
''                        .Col = .ColLetterToNumber("B")
''                        Do
''                            .Col = .ColLetterToNumber(Col_AA)
''                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                            .CellNote = ""     'static
''                            .BackColor = mNonErrorColor
''                            iFlagTon = 0
''                            .Row = .Row + 1
''                            .Col = .ColLetterToNumber("B")
''                        Loop Until UCase(.Text) = "AA"
'                    End If
'                Next j
'                ' Neu truong hop khong ton tai trong bao cao thi se thong bao loi
'                If isNotExistBC = False Then
'                        .Row = 46
'                        .Col = .ColLetterToNumber("B")
'                        Do
'                            .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
'                            .GetText .ColLetterToNumber(Col_O), .Row, strMauSoHD
'                            .GetText .ColLetterToNumber(Col_S), .Row, strKyHieuHD
'                            If arrGroupValue(0) = Trim(strTenHD) And arrGroupValue(1) = Trim(strMauSoHD) And arrGroupValue(2) = Trim(strKyHieuHD) Then
'                                .Col = .ColLetterToNumber(Col_AA)
'                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                .CellNote = errTon     'static
'                                .BackColor = mErrorColor
'                                iFlagTon = 1
'                             End If
'                             .Row = .Row + 1
'                             .Col = .ColLetterToNumber("B")
'                        Loop Until UCase(.Text) = "AA"
'                End If
'            Next i
'         Else
'
'         End If
         
         ' So sanh ton tren group sheet 1 va group sheet 2
         'for i = 0 to Ubound(
         
         ' end
        
        ' set lai trang thai cua HDR
          .Sheet = .SheetCount
         .SetText .ColLetterToNumber("B"), 37, IIf(iFlagEmty = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 32, IIf(iFlagMauSoHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 34, IIf(iFlagSoLienHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 33, IIf(iFlagKyHieu = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 35, IIf(iFlagLengHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 36, IIf(iFlagSoLuong = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 38, IIf(iFlagNgayHuy = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 39, IIf(iFlagNgayKy = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 40, IIf(iFlagKhoangTSDSTrung = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 50, IIf(iFlagTon = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 48, IIf(iFlagTonTK = 1, "0", "1")
         
         .SetText .ColLetterToNumber("B"), 54, IIf(iFlagSoLienHD = 2, "0", "1")
        ' CheckErrorMST
        '.ReDraw = True
        '.Visible = True
        .EventEnabled(EventAllEvents) = True
    End With
    
End Sub
Private Sub checkSheet3()
    Dim strCheck As String
    Dim errMauSoHD As Variant, errKyHieuHD As Variant, errSoLuong As Variant, errSoLienHD As Variant, errSoLienHDDT As Variant
    Dim errKhoangGN As Variant, errLengthHD As Variant, errEmty As Variant, errNgayKy As Variant, errNgayHuy As Variant, errTrungKhoang As Variant, errTon As Variant, errTonTK As Variant
    Dim strTenHD As Variant, strMauSoHD As Variant, strKyHieuHD As Variant, strTuSo As Variant, strDenso As Variant
    Dim strTenHD1 As Variant, strMauSoHD1 As Variant, strKyHieuHD1 As Variant, strTuso1 As Variant, strDenso1 As Variant
    Dim strSoLuong As Variant, strNguoiKy As Variant, strNguoiLapBieu As Variant, strNgayKy As Variant, strNgayHuy As Variant
    Dim iFlagTenHD As Integer, iFlagMauSoHD As Integer, iFlagKyHieu As Integer, iFlagSoLienHD As Integer, iFlagEmty As Integer
    Dim iWhiteRow As Integer, iFlagLengHD As Integer, iFlagSoLuong As Integer, iFlagNgayKy As Integer, iFlagNgayHuy As Integer
    Dim iFlagKhoangTSDSTrung As Integer, iFocusFlag As Integer, iFlagTon As Integer, iFlagTonTK As Integer
    
    Dim strId As Variant, strLoai As String, strTemp As String, result As String
    Dim i As Integer, countDynamicRow As Integer, j As Integer, k As Integer
    
    ' group theo tung nhom
    Dim arrGroupHD() As String
    Dim strGroupHD As String
    Dim arrGroupHD_TK() As String
    Dim arrGroupValue_TK() As String
    
    ' Danh sach cac hoa don ton tren sheet 1
    Dim arrGroupHD_Ton() As String
    Dim arrHD_Ton() As String
    
    Dim str() As String
    Dim str1 As String
    ' tra ve mang cac row bi trung khoang
    str = checkTrung(fps, 45, Col_V, Col_Y, "B", "D", Col_O, Col_S, "", False, 3)
    
    'str1 = checkTon(fps, 45, Col_V, Col_Y, "B", "D", Col_O, Col_S, "", False, 3)
    'arrGroupHD_Ton = Split(strCheckTonKhoang, "###")
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber(Col_E), 32, errMauSoHD
        .GetText .ColLetterToNumber(Col_E), 33, errKyHieuHD
        .GetText .ColLetterToNumber(Col_E), 34, errSoLienHD
        .GetText .ColLetterToNumber(Col_E), 35, errKhoangGN
        .GetText .ColLetterToNumber(Col_E), 36, errLengthHD
        .GetText .ColLetterToNumber(Col_E), 37, errSoLuong
        .GetText .ColLetterToNumber(Col_E), 38, errEmty
        .GetText .ColLetterToNumber(Col_E), 39, errNgayHuy
        .GetText .ColLetterToNumber(Col_E), 40, errNgayKy
        .GetText .ColLetterToNumber(Col_E), 41, errTrungKhoang
        .GetText .ColLetterToNumber(Col_E), 48, errTon
        .GetText .ColLetterToNumber(Col_E), 51, errTonTK
        
        .GetText .ColLetterToNumber(Col_E), 55, errSoLienHDDT
        
        'set flag bang 0
        iFlagTenHD = 0
        iFlagMauSoHD = 0
        iFlagKyHieu = 0
        iFlagSoLienHD = 0
        iFlagEmty = 0
        iFlagSoLuong = 0
        iFlagNgayKy = 0
        iFlagNgayHuy = 0
        iFlagKhoangTSDSTrung = 0
        iFlagTon = 0
        
        ' dem so row dynamic
        countDynamicRow = 0
        ' Kiem tra tren sheet 1
        .Sheet = 3
        
        ' reset lai ca cell error
        .Row = 45
        .Col = .ColLetterToNumber("B")
        Do
            .Col = .ColLetterToNumber(Col_AA)
            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            .CellNote = ""     'static
            .BackColor = mNonErrorColor
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until UCase(.Text) = "AA"

        
        .Col = .ColLetterToNumber("B")
        .Row = 45
        'i = 1
        Do
            j = .Row
            countDynamicRow = countDynamicRow + 1
             ' Lay cac gia tri cua 1 row
             iWhiteRow = 1
            .GetText .ColLetterToNumber("D"), .Row, strTenHD
            .GetText .ColLetterToNumber("O"), .Row, strMauSoHD
            .GetText .ColLetterToNumber("S"), .Row, strKyHieuHD
            .GetText .ColLetterToNumber("V"), .Row, strTuSo
            .GetText .ColLetterToNumber("Y"), .Row, strDenso
            .GetText .ColLetterToNumber("AA"), .Row, strSoLuong
            .GetText .ColLetterToNumber("AI"), .Row, strId
            
            ' set gia tri cho check KH NN/XC
            strMauSoHD_01GTKT = Trim(Mid(strId, 11, 6))
            ' end
            
            If Trim(strTenHD) <> "" Or Trim(strMauSoHD) <> "" Or Trim(strKyHieuHD) <> "" Or Trim(strId) <> "" Or Trim(strTuSo) <> "" Or Trim(strDenso) <> "" Then
                ' Khong phai dong trang
                iWhiteRow = 0
            End If
            
            'Ten HD
            If Trim(strTenHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber(Col_D)
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                .Col = .ColLetterToNumber(Col_D)
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            ' Mau so HD
            If Trim(strMauSoHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber(Col_O)
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                strLoai = Trim(Left(strId, 1))
                If strLoai = "0" Then
                    strTemp = Trim(Mid(strId, 11, 6))
                ElseIf strLoai = "4" Then
                    strTemp = Trim(Mid(strId, 11, 6))
                Else
                    strTemp = Trim(Mid(strId, 11, 3))
                End If
                result = CheckMauSoHD(CStr(strMauSoHD), strLoai, strTemp)
                If result = "1" And iWhiteRow = 0 Then
                    .Col = .ColLetterToNumber(Col_O)
                    .CellNote = errMauSoHD  'static
                    .BackColor = mErrorColor
                    iFlagMauSoHD = 1
                ElseIf result = "2" And iWhiteRow = 0 Then
                     If Right$(Trim$(strKyHieuHD), 1) = "E" Then
                        If CheckSoLienHDDT(CStr(strMauSoHD), strLoai, strTemp, CStr(strKyHieuHD)) = "3" Then
                             .Col = .ColLetterToNumber(Col_O)
                            .CellNote = errSoLienHDDT  'static
'                            .BackColor = mErrorColor
'                            iFlagSoLienHD = 2
                             .BackColor = mAlertColor
                        Else
                            .Col = .ColLetterToNumber(Col_O)
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    Else
                        .Col = .ColLetterToNumber(Col_O)
                        .CellNote = errSoLienHD  'static
                        .BackColor = mErrorColor
                        iFlagSoLienHD = 1
                    End If
                Else
                    If Right$(Trim$(strKyHieuHD), 1) = "E" Then
                        If CheckSoLienHDDT(CStr(strMauSoHD), strLoai, strTemp, CStr(strKyHieuHD)) = "3" Then
                             .Col = .ColLetterToNumber(Col_O)
                            .CellNote = errSoLienHDDT  'static
'                            .BackColor = mErrorColor
'                            iFlagSoLienHD = 2
                             .BackColor = mAlertColor
                        Else
                            .Col = .ColLetterToNumber(Col_O)
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                    Else
                        .Col = .ColLetterToNumber(Col_O)
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                End If
            End If
            ' check ky hieu
            If Trim(strKyHieuHD) = "" And iWhiteRow = 1 Then
                .Col = .ColLetterToNumber(Col_S)
                .CellNote = ""  'static
                .BackColor = mNonErrorColor
                'iFlagEmty = 1
            ElseIf Trim(strKyHieuHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber(Col_S)
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                result = CheckSoHD(CStr(strKyHieuHD), strLoai)
                If result = "1" Then
                    .Col = .ColLetterToNumber(Col_S)
                    .CellNote = errKyHieuHD  'static
                    .BackColor = mErrorColor
                    iFlagKyHieu = 1
                ElseIf result = "2" Then
                    .Col = .ColLetterToNumber(Col_S)
                    .CellNote = errLengthHD  'static
                    .BackColor = mErrorColor
                    iFlagLengHD = 1
                Else
                    .Col = .ColLetterToNumber(Col_S)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            End If
            ' kiem tra tu so den so
            If iWhiteRow = 0 Then
                If Trim(strTuSo) = "" And Trim(strDenso) = "" Then
                    ' tu so
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    ' den so
                    .Col = .ColLetterToNumber(Col_Y)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                ElseIf Trim(strTuSo) = "" And Trim(strDenso) <> "" Then
                    ' tu so
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                    ' den so
                    .Col = .ColLetterToNumber(Col_Y)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                ElseIf Trim(strTuSo) <> "" And Trim(strDenso) = "" Then
                    .Col = .ColLetterToNumber(Col_Y)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                    ' tu so
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNote = ""
                    .BackColor = mNonErrorColor

                Else
                    If (Val(strDenso) - Val(strTuSo) >= 0 And Val(strDenso) > 0) Or (Val(strTuSo) = 0 And Val(strDenso) = 0) Then
                        ' tu so
                        .Col = .ColLetterToNumber(Col_V)
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        ' den so
                        .Col = .ColLetterToNumber(Col_Y)
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    Else
                        .Col = .ColLetterToNumber(Col_V)
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        .Col = .ColLetterToNumber(Col_Y)
                        .CellNote = errSoLuong  'static
                        .BackColor = mErrorColor
                        iFlagSoLuong = 1
                    End If
                End If
            Else
                ' tu so
                .Col = .ColLetterToNumber(Col_V)
                .CellNote = ""
                .BackColor = mNonErrorColor
                ' den so
                .Col = .ColLetterToNumber(Col_Y)
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            
            ' kiem tra trung khoang HD
            If UBound(str) > 0 And iWhiteRow = 0 Then
                For i = 0 To UBound(str)
                    If j = str(i) Then
                        .Col = .ColLetterToNumber(Col_V)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errTrungKhoang     'static
                        .BackColor = mErrorColor
                        
                         .Col = .ColLetterToNumber(Col_Y)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errTrungKhoang     'static
                        .BackColor = mErrorColor
                        iFlagKhoangTSDSTrung = 1
                        Exit For
                    End If
                Next i
            End If
            'dhdang so luong ton cuoi ky
            ' kiem tra trung khoang HD
'           Dim strListHD As String
'            Dim arrHD_PL2() As String
'            Dim arrHD_BC() As String
'            Dim idxCheckTon As Double
'            Dim isCheck As Boolean
'            strListHD = getListHDTS_DS(CStr(strTuSo), CStr(strDenso))
'            arrHD_PL2 = Split(Mid(strListHD, 1, Len(strListHD) - 1), ";")
'            isCheck = False
'            For i = 0 To UBound(arrHD_PL2)
'                For idxCheckTon = 0 To UBound(arrGroupHD_Ton)
'                    arrHD_BC = Split(arrGroupHD_Ton(idxCheckTon), "~")
'                    If iWhiteRow = 0 And Trim(strTenHD) = arrHD_BC(0) And Trim(strMauSoHD) = arrHD_BC(1) And Trim(strKyHieuHD) = arrHD_BC(2) Then
'                        'For i = 0 To UBound(str1)
'                            'If j = str1(i) Then
'                              If InStr(";" & arrHD_BC(3), ";" & arrHD_PL2(i) & ";") = 0 Then
'                                    .Col = .ColLetterToNumber(Col_V)
'                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                    .CellNote = errTonTK     'static
'                                    .BackColor = mErrorColor
'
'                                     .Col = .ColLetterToNumber(Col_Y)
'                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                    .CellNote = errTonTK     'static
'                                    .BackColor = mErrorColor
'                                    iFlagTonTK = 1
'                                    isCheck = True
'                                    Exit For
'                               End If
'                        'Next i
'                    End If
'                Next idxCheckTon
'                If isCheck = True Then Exit For
'            Next i
             
            
'            If InStr(str1, j) <> 0 And iWhiteRow = 0 Then
'                'For i = 0 To UBound(str1)
'                    'If j = str1(i) Then
'                        .Col = .ColLetterToNumber(Col_V)
'                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                        .CellNote = errTon     'static
'                        .BackColor = mErrorColor
'
'                         .Col = .ColLetterToNumber(Col_Y)
'                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                        .CellNote = errTon     'static
'                        .BackColor = mErrorColor
'                        iFlagTon = 1
'                        'Exit For
'                    'End If
'                'Next i
'            End If
            ' Test
            Dim arrHoaDonS1() As HOA_DON
            Dim isTrongKhoang As Boolean ' Kiem tra tu so den so thuoc khoang 20,21
            Dim isNotCheckTrongKhoang As Boolean
            Dim isHoaDonS1 As Boolean  ' la hoa don ben TK
            Dim isExitFor As Boolean  ' thoat khoi vong lap
            
            'If UBound(loaiHoaDonS1) > 0 Then
            isHoaDonS1 = False
            isTrongKhoang = False
            isExitFor = False
            
            For i = 0 To UBound(loaiHoaDonS1)
                If loaiHoaDonS1(i).loaiHD = Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~0" And iWhiteRow = 0 Then
                    isHoaDonS1 = True
                    arrHoaDonS1 = loaiHoaDonS1(i).hoaDon
                    'If UBound(arrHoaDonS1) > 0 Then
                    For k = 0 To UBound(arrHoaDonS1)
                        If Val(strTuSo) >= arrHoaDonS1(k).tuSo And Val(strDenso) <= arrHoaDonS1(k).denso And arrHoaDonS1(k).isDelete = False Then
                            isTrongKhoang = True
                            isExitFor = True
                            Exit For
                        End If
                    Next k
                    'End If
                End If
                If isExitFor = True Then
                    Exit For
                End If
            Next i
            'End If
            ' end test
            If iWhiteRow = 0 Then
                If isHoaDonS1 = True Then
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = ""     'static
                    .BackColor = mNonErrorColor

                     .Col = .ColLetterToNumber(Col_Y)
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = ""     'static
                    .BackColor = mNonErrorColor
                    If isTrongKhoang = True Then
                        .Col = .ColLetterToNumber(Col_V)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = ""     'static
                        .BackColor = mNonErrorColor
    
                         .Col = .ColLetterToNumber(Col_Y)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = ""     'static
                        .BackColor = mNonErrorColor
                    Else
        
                        .Col = .ColLetterToNumber(Col_V)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errTon     'static
                        .BackColor = mErrorColor
    
                         .Col = .ColLetterToNumber(Col_Y)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errTon     'static
                        .BackColor = mErrorColor
                        iFlagTon = 1
                    End If
                Else
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = errTonTK     'static
                    .BackColor = mErrorColor

                     .Col = .ColLetterToNumber(Col_Y)
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = errTonTK     'static
                    .BackColor = mErrorColor
                    iFlagTonTK = 1
                End If
            Else
                .Col = .ColLetterToNumber(Col_V)
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""     'static
                .BackColor = mNonErrorColor

                 .Col = .ColLetterToNumber(Col_Y)
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""     'static
                .BackColor = mNonErrorColor
            End If
            
            ' End test
            ' end check trung
            .Row = j
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until UCase(.Text) = "AA"
                
        ' Kiem tra du lieu nguoi dai dien
         .GetText .ColLetterToNumber(Col_R), countDynamicRow + 48, strNguoiKy
         .GetText .ColLetterToNumber(Col_D), countDynamicRow + 48, strNguoiLapBieu
         .GetText .ColLetterToNumber(Col_R), countDynamicRow + 50, strNgayKy
         ' kiem tra nguoi ky bat buoc nhap
         If Trim(strNguoiKy) = "" Then
            .Row = countDynamicRow + 48
            .Col = .ColLetterToNumber(Col_R)
            .CellNote = errEmty  'static
            .BackColor = mErrorColor
            iFlagEmty = 1
         Else
            .Row = countDynamicRow + 48
            .Col = .ColLetterToNumber(Col_R)
            .CellNote = ""
            .BackColor = mNonErrorColor
         End If
         'kiem tra ngay ky
         If Trim(strNgayKy) = "" Or Trim(strNgayKy) = "../../...." Then
            .Row = countDynamicRow + 50
            .Col = .ColLetterToNumber(Col_R)
            .CellNote = errEmty  'static
            .BackColor = mErrorColor
            iFlagEmty = 1
         Else
                strNgayKy = DateSerial(CInt(Mid$(strNgayKy, 7, 4)), CInt(Mid$(strNgayKy, 4, 2)), CInt(Mid$(strNgayKy, 1, 2)))
                If strNgayKy > Now Then
                    .Col = .ColLetterToNumber(Col_R)
                    .Row = countDynamicRow + 50
                    .CellNote = errNgayKy  'static
                    .BackColor = mErrorColor
                    iFlagNgayKy = 1
                Else
                    .Row = countDynamicRow + 50
                    .Col = .ColLetterToNumber(Col_R)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
         End If
         
              ' group so hoa don vao thanh tung nhom
         
'         .Row = 45
'         .Col = .ColLetterToNumber("B")
'         Dim arrGroupValue() As String  ' Luu thong tin cua tung dong check
'         Dim coutCheck, countCheckTrung As Long
'         Dim arrGroupHDNew As String
'         Dim vitriGhep As Long
'         'Dim isTrung As Boolean
'
'         Do
'            .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
'            .GetText .ColLetterToNumber(Col_O), .Row, strMauSoHD
'            .GetText .ColLetterToNumber(Col_S), .Row, strKyHieuHD
'            .GetText .ColLetterToNumber(Col_V), .Row, strTuSo
'            .GetText .ColLetterToNumber(Col_Y), .Row, strDenso
'            .GetText .ColLetterToNumber(Col_AA), .Row, strSoLuong
'            .GetText .ColLetterToNumber("AI"), .Row, strId
'            If .Row = 45 Then
'                strGroupHD = Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~0"
'            Else
'                arrGroupHD = Split(strGroupHD, "###")
'                coutCheck = 0
'                countCheckTrung = 0
'                For i = 0 To UBound(arrGroupHD)
'                    coutCheck = coutCheck + 1
'                    arrGroupValue = Split(arrGroupHD(i), "~")
'                    If Trim(strTenHD) = arrGroupValue(0) And Trim(strMauSoHD) = arrGroupValue(1) And Trim(strKyHieuHD) = arrGroupValue(2) Then
'                        Exit For
'                    Else
'                        countCheckTrung = countCheckTrung + 1
'                    End If
'                Next i
'                If coutCheck = countCheckTrung Then
'                    strGroupHD = strGroupHD & "###" & Trim(strTenHD) & "~" & Trim(strMauSoHD) & "~" & Trim(strKyHieuHD) & "~0"
'                End If
'            End If
'            .Row = .Row + 1
'         Loop Until .Text = "aa"
'         ' test
'          Dim hoaDon As HOA_DON
'         Dim tgHoaDon As HOA_DON
'         Dim arrHoaDon() As HOA_DON
'         Dim arrTempHoaDon() As HOA_DON
'         Dim loaiHoaDon() As LOAI_HOA_DON
'         Dim index As Long
'         Dim vitri As Long
'         .Row = 45
'         .Col = .ColLetterToNumber("B")
'         index = 0
'         arrGroupHD = Split(strGroupHD, "###")
'         ReDim Preserve loaiHoaDon(UBound(arrGroupHD))
'         For i = 0 To UBound(arrGroupHD)
'            ReDim Preserve arrHoaDon(0)
'            loaiHoaDon(i).loaiHD = arrGroupHD(i)
'            loaiHoaDon(i).hoaDon = arrHoaDon
'         Next i
'
'         Do
'            .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
'            .GetText .ColLetterToNumber(Col_O), .Row, strMauSoHD
'            .GetText .ColLetterToNumber(Col_S), .Row, strKyHieuHD
'            .GetText .ColLetterToNumber(Col_V), .Row, strTuSo
'            .GetText .ColLetterToNumber(Col_Y), .Row, strDenso
'            hoaDon.tenHD = CStr(strTenHD)
'            hoaDon.mauSoHD = CStr(strMauSoHD)
'            hoaDon.kyHieuHD = CStr(strKyHieuHD)
'            hoaDon.tuSo = Val(CStr(strTuSo))
'            hoaDon.denSo = Val(CStr(strDenso))
'            'If UBound(loaiHoaDon) > 0 Then
'            For i = 0 To UBound(loaiHoaDon)
'                If loaiHoaDon(i).loaiHD = hoaDon.tenHD & "~" & hoaDon.mauSoHD & "~" & hoaDon.kyHieuHD & "~0" Then
'                    arrTempHoaDon = loaiHoaDon(i).hoaDon
'                    If Trim(arrTempHoaDon(0).mauSoHD) = "" Then
'                        vitri = 0
'                    Else
'                        vitri = UBound(arrTempHoaDon) + 1
'                        ReDim Preserve arrTempHoaDon(vitri)
'                    End If
'                    arrTempHoaDon(vitri) = hoaDon
'                    loaiHoaDon(i).hoaDon = arrTempHoaDon
'                End If
'            Next i
'            'End If
'            .Row = .Row + 1
'         Loop Until .Text = "aa"
         
         ' end test
         
         
         
'         ' Tinh tong cac so hoa don cung 1 loai
'            .Row = 45
'            .Col = .ColLetterToNumber("B")
'             Do
'            .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
'            .GetText .ColLetterToNumber(Col_O), .Row, strMauSoHD
'            .GetText .ColLetterToNumber(Col_S), .Row, strKyHieuHD
'            .GetText .ColLetterToNumber(Col_V), .Row, strTuSo
'            .GetText .ColLetterToNumber(Col_Y), .Row, strDenso
'            .GetText .ColLetterToNumber(Col_AA), .Row, strSoLuong
'            .GetText .ColLetterToNumber("AI"), .Row, strId
'            arrGroupHD = Split(strGroupHD, "###")
'            arrGroupHDNew = ""
'            If UBound(arrGroupHD) > 0 Then
'                For i = 0 To UBound(arrGroupHD)
'                    arrGroupValue = Split(arrGroupHD(i), "~")
'                    If Trim(strTenHD) = arrGroupValue(0) And Trim(strMauSoHD) = arrGroupValue(1) And Trim(strKyHieuHD) = arrGroupValue(2) Then
'                        arrGroupValue(3) = Val(arrGroupValue(3)) + Val(strSoLuong)
'                        arrGroupHDNew = arrGroupValue(0) & "~" & arrGroupValue(1) & "~" & arrGroupValue(2) & "~" & arrGroupValue(3)
'                        vitriGhep = i
'                        Exit For
'                    End If
'                Next i
'                ' ghep lai chuoi moi khi da cong du lieu
'                strGroupHD = ""
'                For i = 0 To UBound(arrGroupHD)
'                    If i = vitriGhep Then
'                        strGroupHD = strGroupHD & "###" & arrGroupHDNew
'                    Else
'                        strGroupHD = strGroupHD & "###" & arrGroupHD(i)
'                    End If
'                Next i
'                strGroupHD = Mid$(strGroupHD, 4)
'            Else
'                arrGroupValue = Split(arrGroupHD(0), "~")
'                arrGroupValue(3) = Val(arrGroupValue(3)) + Val(strSoLuong)
'                strGroupHD = arrGroupValue(0) & "~" & arrGroupValue(1) & "~" & arrGroupValue(2) & "~" & arrGroupValue(3)
'            End If
'
'
'            .Row = .Row + 1
'         Loop Until .Text = "aa"
'
'         ' Lay ve group cua sheet 1
'          Dim isNotExistBC As Boolean
'         arrGroupHD_TK = Split(strCheckTon, "###")
'         arrGroupHD = Split(strGroupHD, "###")
'        ' If UBound(arrGroupHD) > 0 Then
'            For i = 0 To UBound(arrGroupHD)
'                arrGroupValue = Split(arrGroupHD(i), "~")
'                isNotExistBC = False
'                For j = 0 To UBound(arrGroupHD_TK)
'                    arrGroupValue_TK = Split(arrGroupHD_TK(j), "~")
'                    If arrGroupValue(0) = arrGroupValue_TK(0) And arrGroupValue(1) = arrGroupValue_TK(1) And arrGroupValue(2) = arrGroupValue_TK(2) Then
'                        isNotExistBC = True
'                        If arrGroupValue(3) <> arrGroupValue_TK(3) Then
'                            .Row = 45
'                            .Col = .ColLetterToNumber("B")
'                            Do
'                                .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
'                                .GetText .ColLetterToNumber(Col_O), .Row, strMauSoHD
'                                .GetText .ColLetterToNumber(Col_S), .Row, strKyHieuHD
'                                If arrGroupValue(0) = Trim(strTenHD) And arrGroupValue(1) = Trim(strMauSoHD) And arrGroupValue(2) = Trim(strKyHieuHD) Then
'                                    .Col = .ColLetterToNumber(Col_AA)
'                                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                    .CellNote = errTonTK     'static
'                                    .BackColor = mErrorColor
'                                    iFlagTonTK = 1
'                                 End If
'                                 .Row = .Row + 1
'                                 .Col = .ColLetterToNumber("B")
'                            Loop Until UCase(.Text) = "AA"
'                            Exit For ' Thoat vong lap trong
'                        End If
'                    Else
''                        .Row = 45
''                        .Col = .ColLetterToNumber("B")
''                        Do
''                            .Col = .ColLetterToNumber(Col_AA)
''                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
''                            .CellNote = ""     'static
''                            .BackColor = mNonErrorColor
''                            iFlagTonTK = 0
''                            .Row = .Row + 1
''                            .Col = .ColLetterToNumber("B")
''                        Loop Until UCase(.Text) = "AA"
'                    End If
'                Next j
'                              ' Neu truong hop khong ton tai trong bao cao thi se thong bao loi
'                If isNotExistBC = False Then
'                        .Row = 45
'                        .Col = .ColLetterToNumber("B")
'                        Do
'                            .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
'                            .GetText .ColLetterToNumber(Col_O), .Row, strMauSoHD
'                            .GetText .ColLetterToNumber(Col_S), .Row, strKyHieuHD
'                            If arrGroupValue(0) = Trim(strTenHD) And arrGroupValue(1) = Trim(strMauSoHD) And arrGroupValue(2) = Trim(strKyHieuHD) Then
'                                .Col = .ColLetterToNumber(Col_AA)
'                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                                .CellNote = errTonTK     'static
'                                .BackColor = mErrorColor
'                                iFlagTonTK = 1
'                             End If
'                             .Row = .Row + 1
'                             .Col = .ColLetterToNumber("B")
'                        Loop Until UCase(.Text) = "AA"
'                End If
'          Next i
                
        ' set lai trang thai cua HDR
          .Sheet = .SheetCount
         .SetText .ColLetterToNumber("B"), 37, IIf(iFlagEmty = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 32, IIf(iFlagMauSoHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 34, IIf(iFlagSoLienHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 33, IIf(iFlagKyHieu = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 35, IIf(iFlagLengHD = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 36, IIf(iFlagSoLuong = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 38, IIf(iFlagNgayHuy = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 39, IIf(iFlagNgayKy = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 40, IIf(iFlagKhoangTSDSTrung = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 48, IIf(iFlagTon = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 51, IIf(iFlagTonTK = 1, "0", "1")
         
         .SetText .ColLetterToNumber("B"), 55, IIf(iFlagSoLienHD = 2, "0", "1")
        ' CheckErrorMST
        '.ReDraw = True
        '.Visible = True
        .EventEnabled(EventAllEvents) = True
    End With
    
End Sub
Private Sub clearCheckSheet1()
    
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
        .GetText .ColLetterToNumber(Col_E), 12, vError1
        .GetText .ColLetterToNumber(Col_E), 13, vError2
        .GetText .ColLetterToNumber(Col_E), 14, vError3
        
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
    For Each xmlNodeValid In TAX_Utilities_v2.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "BC26AC"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
                    If lCol = fps.ColLetterToNumber(Col_D) Or _
                       lCol = fps.ColLetterToNumber(Col_E) Then
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
            Case "01_1"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_v2.Data(1).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
                    If lCol = fps.ColLetterToNumber(Col_O) Or _
                       lCol = fps.ColLetterToNumber(Col_S) Then
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
            Case "01_2"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_v2.Data(2).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
                    If lCol = fps.ColLetterToNumber(Col_O) Or _
                       lCol = fps.ColLetterToNumber(Col_S) Then
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
End Sub
Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean
    Dim totalCell  As Long
    Dim countRowReset As Long
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    totalCell = TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length
    countRowReset = (totalCell - 7) / 25
    With fps
        For Each xmlNodeReset In TAX_Utilities_v2.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            .Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            If (lRow < 22 Or lRow > 22 + countRowReset) And mCurrentSheet = 1 Then
            Else
                .Col = lCol
                .Row = lRow
                
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
                TAX_Utilities_v2.AdjustData(mCurrentSheet - 1) = True
            End If
        Next
    End With
End Function

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()
Dim isheet As Integer
'    With fps
'        .EventEnabled(EventAllEvents) = False
'         isheet = mCurrentSheet
'         mCurrentSheet = 1
'        .Sheet = 1
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 15
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 16
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 17
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 17
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 18
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 19
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 20
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 21
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 22
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 23
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'        .EventEnabled(EventAllEvents) = True
'        mCurrentSheet = isheet
'        .ActiveSheet = mCurrentSheet
'    End With
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
    Dim IsUpdate As Boolean
    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               If (lCol = .ColLetterToNumber("B") Or lCol = .ColLetterToNumber(Col_G)) And lRow = 14 Then
                    IsUpdate = True
                    ' tam thoi xu ly de ghi gia tri xuong file xml
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
               Else
                    IsUpdate = False
               End If
               If Not IsUpdate Then
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
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
    Dim maHD    As Variant
    Dim kyHieuHD As Variant
    Dim tuSo As Variant
    Dim denso As Variant
    
    Dim strTemp As String
    
    Dim strId As Variant
    Dim strLoai As Variant
    
    Dim tenHD  As String
    Dim varTemp As Variant
    Dim arrTemp() As String
    
    Dim vSoSD, vSoXoa, vSoMat, vSoHuy As Variant

    Dim TongSo As Variant
    Dim strSoLuong As String
    Dim strTuso1 As Variant, strDenso1 As Variant, strTuso2 As Variant, strDenso2 As Variant, strCt10 As Variant, strCt11 As Variant, strCt12 As Variant, strCt5 As Variant, strCt20 As Variant, strCt21 As Variant
    Dim str() As String
    Dim soluongTonCK As Double
    Dim i As Long
With fps
fps.EventEnabled(EventAllEvents) = False
           ' set TK
            If .ActiveSheet = 1 Then
                   .Row = 22
                    i = 1
                    Do
                        .GetText .ColLetterToNumber("E"), .Row, maHD
                         GetMaHoaDon maHD, tenHD
                         If tenHD <> "" Then
                            arrTemp = Split(tenHD, "###")
                            If UBound(arrTemp) > 0 Then
                                .Col = .ColLetterToNumber("C")
                                .Text = arrTemp(0)
                                UpdateCell fps, .Col, .Row, .Text
                                .Col = .ColLetterToNumber("AA")
                                .Text = CStr(arrTemp(1))
                                UpdateCell fps, .Col, .Row, .Text
                                
                                .Col = .ColLetterToNumber("AC")
                                .SetText .Col, .Row, "0"
                                UpdateCell fps, .Col, .Row, .Text
                            End If
                            
                            '
                             'dhdang sua fomat ky hieu hoa don
                             .GetText .ColLetterToNumber(Col_F), .Row, varTemp
                             .SetText .ColLetterToNumber(Col_F), .Row, FormatKyHieu(CStr(varTemp))
                             .Col = .ColLetterToNumber(Col_F)
                            UpdateCell_sheet fps, 1, .Col, .Row, .Text
                           
                           ' Format dinh dang tu so den so co 7 ky tu
                            ' cot H
                            .GetText .ColLetterToNumber(Col_H), .Row, varTemp
                            .SetText .ColLetterToNumber(Col_H), .Row, FormatSoHD(CStr(varTemp))
                            .Col = .ColLetterToNumber(Col_H)
                            UpdateCell_sheet fps, 1, .Col, .Row, .Text
                            ' cot I
                            .GetText .ColLetterToNumber(Col_I), .Row, varTemp
                            .SetText .ColLetterToNumber(Col_I), .Row, FormatSoHD(CStr(varTemp))
                            .Col = .ColLetterToNumber(Col_I)
                            UpdateCell_sheet fps, 1, .Col, .Row, .Text
                            ' cot J
                            .GetText .ColLetterToNumber(Col_J), .Row, varTemp
                            .SetText .ColLetterToNumber(Col_J), .Row, FormatSoHD(CStr(varTemp))
                            .Col = .ColLetterToNumber(Col_J)
                            UpdateCell_sheet fps, 1, .Col, .Row, .Text
                            ' cot K
                            .GetText .ColLetterToNumber(Col_K), .Row, varTemp
                            .SetText .ColLetterToNumber(Col_K), .Row, FormatSoHD(CStr(varTemp))
                            .Col = .ColLetterToNumber(Col_K)
                            UpdateCell_sheet fps, 1, .Col, .Row, .Text
                            
                            ' Kiem tra neu so SD, xoa, mat, huy deu bang 0 thi set [10],[11] =trang
                            .GetText .ColLetterToNumber("P"), .Row, vSoSD
                            .GetText .ColLetterToNumber("Q"), .Row, vSoXoa
                            .GetText .ColLetterToNumber("S"), .Row, vSoMat
                            .GetText .ColLetterToNumber("U"), .Row, vSoHuy
                            
                            'Ho tro set tu so [10]
                            .GetText .ColLetterToNumber(Col_H), .Row, strTuso1
                            .GetText .ColLetterToNumber(Col_I), .Row, strDenso1
                            .GetText .ColLetterToNumber(Col_J), .Row, strTuso2
                            .GetText .ColLetterToNumber(Col_K), .Row, strDenso2
                            str = getTusoDenso(CStr(strTuso1), CStr(strDenso1), CStr(strTuso2), CStr(strDenso2))
                            If str(0) <> "-1" Then
                                ' Neu tong so xoa, sd, mat, huy =0 thi khong ho tro tinh [10][11]
                                If Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0 Then
                                    .SetText .ColLetterToNumber(Col_L), .Row, FormatSoHD(str(0))
                                    .Col = .ColLetterToNumber(Col_L)
                                    UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_L), .Row, .Text
                                Else
                                    .SetText .ColLetterToNumber(Col_L), .Row, ""
                                    .Col = .ColLetterToNumber(Col_L)
                                    UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_L), .Row, .Text
                                End If
                                .SetText .ColLetterToNumber(Col_G), .Row, FormatSoHD(str(2))
                                .Col = .ColLetterToNumber(Col_G)
                                 UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_G), .Row, .Text
                            End If
                            ' Ho tro tinh chi tieu [11]
                            .GetText .ColLetterToNumber(Col_M), .Row, strCt11
                            .GetText .ColLetterToNumber(Col_N), .Row, strCt12
                            .GetText .ColLetterToNumber(Col_G), .Row, strCt5
                            If Val(strCt5) = Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0) Then
                                ' [5] =[12] ,  [13] + [14] + [16] + [18] <> 0
                                soluongTonCK = Val(strCt5) - Val(strCt12)
                                .SetText .ColLetterToNumber("X"), .Row, ""  'FormatSoHD(Val(0))
                                .Col = .ColLetterToNumber("X")
                                UpdateCell_sheet fps, 1, .ColLetterToNumber("X"), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Y), .Row, ""   'FormatSoHD(Val(0) + soluongTonCK)
                                .Col = .ColLetterToNumber(Col_Y)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Y), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Z), .Row, soluongTonCK
                                .Col = .ColLetterToNumber(Col_Z)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Z), .Row, .Text
                            ElseIf Val(strCt5) > Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0) Then
                                soluongTonCK = Val(strCt5) - Val(strCt12)
                                .SetText .ColLetterToNumber("X"), .Row, FormatSoHD(Val(strCt11) + 1)
                                .Col = .ColLetterToNumber("X")
                                UpdateCell_sheet fps, 1, .ColLetterToNumber("X"), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Y), .Row, FormatSoHD(Val(strCt11) + soluongTonCK)
                                .Col = .ColLetterToNumber(Col_Y)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Y), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Z), .Row, soluongTonCK
                                .Col = .ColLetterToNumber(Col_Z)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Z), .Row, .Text
                            ElseIf Val(strCt5) > Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) = 0) And (Trim(strTuso1) = "" And Trim(strDenso1) = "") Then
                                 soluongTonCK = Val(strCt5)
                                .SetText .ColLetterToNumber("X"), .Row, FormatSoHD(Val(strTuso2))
                                .Col = .ColLetterToNumber("X")
                                UpdateCell_sheet fps, 1, .ColLetterToNumber("X"), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Y), .Row, FormatSoHD(Val(strDenso2))
                                .Col = .ColLetterToNumber(Col_Y)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Y), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Z), .Row, soluongTonCK
                                .Col = .ColLetterToNumber(Col_Z)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Z), .Row, .Text
                            ElseIf Val(strCt5) > Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) = 0) And (Trim(strTuso2) = "" And Trim(strDenso2) = "") Then
                                 soluongTonCK = Val(strCt5)
                                .SetText .ColLetterToNumber("X"), .Row, FormatSoHD(Val(strTuso1))
                                .Col = .ColLetterToNumber("X")
                                UpdateCell_sheet fps, 1, .ColLetterToNumber("X"), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Y), .Row, FormatSoHD(Val(strDenso1))
                                .Col = .ColLetterToNumber(Col_Y)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Y), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Z), .Row, soluongTonCK
                                .Col = .ColLetterToNumber(Col_Z)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Z), .Row, .Text
                            Else
                                 soluongTonCK = Val(strCt5)
                                .SetText .ColLetterToNumber("X"), .Row, FormatSoHD(Val(strTuso1))
                                .Col = .ColLetterToNumber("X")
                                UpdateCell_sheet fps, 1, .ColLetterToNumber("X"), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Y), .Row, FormatSoHD(Val(strDenso2))
                                .Col = .ColLetterToNumber(Col_Y)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Y), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Z), .Row, soluongTonCK
                                .Col = .ColLetterToNumber(Col_Z)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Z), .Row, .Text
                            End If
                            
                        'Tinh so hoa don mat, xoa, huy
                            ' Tinh so xoa
                            .GetText .ColLetterToNumber(Col_R), .Row, varTemp
                            TongSo = GetSoLuongHD(CStr(varTemp))
                            .SetText .ColLetterToNumber(Col_Q), .Row, TongSo
                            .Col = .ColLetterToNumber(Col_Q)
                            UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Q), .Row, .Text
                            
                            ' Tinh so mat
                            .GetText .ColLetterToNumber(Col_T), .Row, varTemp
                            TongSo = GetSoLuongHD(CStr(varTemp))
                            .SetText .ColLetterToNumber(Col_S), .Row, TongSo
                            .Col = .ColLetterToNumber(Col_S)
                            UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_S), .Row, .Text
                            
                            ' Tinh so huy
                            .GetText .ColLetterToNumber(Col_V), .Row, varTemp
                            TongSo = GetSoLuongHD(CStr(varTemp))
                            .SetText .ColLetterToNumber(Col_U), .Row, TongSo
                            .Col = .ColLetterToNumber(Col_U)
                            UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_U), .Row, .Text
                            
                            
                            
                            ' Ho tro tinh den so [11]
                            .GetText .ColLetterToNumber(Col_L), .Row, strCt10
                            .GetText .ColLetterToNumber(Col_N), .Row, strCt12
                            strCt11 = Val(strCt10) + Val(strCt12) - 1
                            
                            ' Neu tong so xoa, sd, mat, huy =0 thi khong ho tro tinh [10][11]
                            If Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0 Then
                                .SetText .ColLetterToNumber(Col_M), .Row, FormatSoHD(CStr(strCt11))
                                .Col = .ColLetterToNumber(Col_M)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_M), .Row, .Text
                                If str(0) <> "-1" Then
                                    .SetText .ColLetterToNumber(Col_L), .Row, FormatSoHD(str(0))
                                    .Col = .ColLetterToNumber(Col_L)
                                    UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_L), .Row, .Text
                                End If
                            Else
                                .SetText .ColLetterToNumber(Col_L), .Row, ""
                                .Col = .ColLetterToNumber(Col_L)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_L), .Row, .Text
                                
                                .SetText .ColLetterToNumber(Col_M), .Row, ""
                                .Col = .ColLetterToNumber(Col_M)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_M), .Row, .Text
                            End If
                                                            
                            .GetText .ColLetterToNumber(Col_G), .Row, strCt5
                            If Val(strCt5) > Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0) Then
                                soluongTonCK = Val(strCt5) - Val(strCt12)
                                .SetText .ColLetterToNumber("X"), .Row, FormatSoHD(Val(strCt11) + 1)
                                .Col = .ColLetterToNumber("X")
                                UpdateCell_sheet fps, 1, .ColLetterToNumber("X"), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Y), .Row, FormatSoHD(Val(strCt11) + soluongTonCK)
                                .Col = .ColLetterToNumber(Col_Y)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Y), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Z), .Row, soluongTonCK
                                .Col = .ColLetterToNumber(Col_Z)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Z), .Row, .Text
                            End If
                            
                                        
                            ' Ho tro tinh den so [11]
                            .GetText .ColLetterToNumber(Col_L), .Row, strCt10
                            .GetText .ColLetterToNumber(Col_N), .Row, strCt12
                            .GetText .ColLetterToNumber(Col_G), .Row, strCt5
                            strCt11 = Val(strCt10) + Val(strCt12) - 1
                            
                            ' Neu tong so xoa, sd, mat, huy =0 thi khong ho tro tinh [10][11]
                            If Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0 Then
                                .SetText .ColLetterToNumber(Col_M), .Row, FormatSoHD(CStr(strCt11))
                                .Col = .ColLetterToNumber(Col_M)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_M), .Row, .Text
                                If str(0) <> "-1" Then
                                    .SetText .ColLetterToNumber(Col_L), .Row, FormatSoHD(str(0))
                                    .Col = .ColLetterToNumber(Col_L)
                                    UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_L), .Row, .Text
                                End If
                            Else
                                .SetText .ColLetterToNumber(Col_L), .Row, ""
                                .Col = .ColLetterToNumber(Col_L)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_L), .Row, .Text
                                
                                .SetText .ColLetterToNumber(Col_M), .Row, ""
                                .Col = .ColLetterToNumber(Col_M)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_M), .Row, .Text
                            End If
                            
                            If Val(strCt5) > Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0) Then
                                soluongTonCK = Val(strCt5) - Val(strCt12)
                                .SetText .ColLetterToNumber("X"), .Row, FormatSoHD(Val(strCt11) + 1)
                                .Col = .ColLetterToNumber("X")
                                UpdateCell_sheet fps, 1, .ColLetterToNumber("X"), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Y), .Row, FormatSoHD(Val(strCt11) + soluongTonCK)
                                .Col = .ColLetterToNumber(Col_Y)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Y), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Z), .Row, soluongTonCK
                                .Col = .ColLetterToNumber(Col_Z)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Z), .Row, .Text
                            End If
                            
                          ' Tinh toan lai cac gia tri
                            .GetText .ColLetterToNumber("G"), .Row, strCt5
                            .GetText .ColLetterToNumber("N"), .Row, strCt12
                            .GetText .ColLetterToNumber("P"), .Row, vSoSD
                            .GetText .ColLetterToNumber("Q"), .Row, vSoXoa
                            .GetText .ColLetterToNumber("S"), .Row, vSoMat
                            .GetText .ColLetterToNumber("U"), .Row, vSoHuy
                            If Val(strCt5) + Val(strCt12) + Val(vSoSD) + Val(vSoXoa) + Val(vSoHuy) + Val(vSoMat) = 0 Then
                                soluongTonCK = 0
                                 .SetText .ColLetterToNumber("X"), .Row, ""
                                 .Col = .ColLetterToNumber("X")
                                UpdateCell_sheet fps, 1, .ColLetterToNumber("X"), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Y), .Row, ""
                                .Col = .ColLetterToNumber(Col_Y)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Y), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Z), .Row, soluongTonCK
                                .Col = .ColLetterToNumber(Col_Z)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Z), .Row, .Text
                            ElseIf Val(strCt5) = Val(strCt12) And (Val(vSoHuy) + Val(vSoMat) + Val(vSoXoa) + Val(vSoSD) > 0) Then
                                ' [5] =[12] ,  [13] + [14] + [16] + [18] <> 0
                                soluongTonCK = Val(strCt5) - Val(strCt12)
                                .SetText .ColLetterToNumber("X"), .Row, ""  'FormatSoHD(Val(0))
                                .Col = .ColLetterToNumber("X")
                                UpdateCell_sheet fps, 1, .ColLetterToNumber("X"), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Y), .Row, ""   'FormatSoHD(Val(0) + soluongTonCK)
                                .Col = .ColLetterToNumber(Col_Y)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Y), .Row, .Text
                                .SetText .ColLetterToNumber(Col_Z), .Row, soluongTonCK
                                .Col = .ColLetterToNumber(Col_Z)
                                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_Z), .Row, .Text
                            End If
                           ' end
                           
                          ' Chuyen cac ky hieu mau so HD thanh chu in hoa
                            .GetText .ColLetterToNumber("AA"), .Row, strId
                            strLoai = Trim(Left(strId, 1))
                            
                            .GetText .ColLetterToNumber(Col_E), .Row, varTemp
                            .SetText .ColLetterToNumber(Col_E), .Row, IIf(strLoai = "3", varTemp, UCase(varTemp))
                            .Col = .ColLetterToNumber(Col_E)
                            UpdateCell_sheet fps, 1, .Col, .Row, .Text
                          
                         End If
                         
                         
                         
                         
                         .Col = .ColLetterToNumber("B")
                         .Text = i
                          i = i + 1
                         .Row = .Row + 1
                    Loop Until .Text = "aa"

            End If
            ' set PL 3.12
            If .ActiveSheet = 2 Then
                    maHD = ""
                    tenHD = ""
                    mCurrentSheet = .Sheet
                    .Row = 46
                    i = 1
                    Do
                        .GetText .ColLetterToNumber("O"), .Row, maHD
                         GetMaHoaDon maHD, tenHD
                         If tenHD <> "" Then
                            arrTemp = Split(tenHD, "###")
                            If UBound(arrTemp) > 0 Then
                                .Col = .ColLetterToNumber("C")
                                .Text = arrTemp(0)
                                UpdateCell fps, .Col, .Row, .Text
                                .Col = .ColLetterToNumber("AI")
                                .Text = CStr(arrTemp(1))
                                UpdateCell fps, .Col, .Row, .Text
                                
'                                .Col = .ColLetterToNumber("AJ")
'                                .Text = CheckMauSoHD(CStr(maHD), Left$(CStr(arrTemp(1)), 1), arrTemp(0))
'                                 UpdateCell fps, .Col, .Row, .Text
                            End If
                            .GetText .ColLetterToNumber("AI"), .Row, strId
                            strLoai = Trim(Left(strId, 1))
                            'Ho tro dinh dang 7 ky tu
                             ' tu so
                             .GetText .ColLetterToNumber(Col_V), .Row, varTemp
                             .SetText .ColLetterToNumber(Col_V), .Row, FormatSoHD(CStr(varTemp))
                             .Col = .ColLetterToNumber(Col_V)
                             UpdateCell_sheet fps, 2, .Col, .Row, .Text
                             ' den so
                             .GetText .ColLetterToNumber(Col_Y), .Row, varTemp
                             .SetText .ColLetterToNumber(Col_Y), .Row, FormatSoHD(CStr(varTemp))
                             .Col = .ColLetterToNumber(Col_Y)
                             UpdateCell_sheet fps, 2, .Col, .Row, .Text
                             
                              ' Ho tro tinh so luong
                            .GetText .ColLetterToNumber(Col_V), .Row, tuSo
                            .GetText .ColLetterToNumber(Col_Y), .Row, denso
                            If Trim(tuSo) <> "" Or Trim(denso) <> "" Then
                                .SetText .ColLetterToNumber(Col_AA), .Row, Val(denso) - Val(tuSo) + 1
                                .Col = .ColLetterToNumber(Col_AA)
                                UpdateCell_sheet fps, 2, .Col, .Row, .Text
                            End If
                             
                         End If
                         
                         ' chuyen hoa don thanh ky tu in hoa
                         .GetText .ColLetterToNumber(Col_S), .Row, varTemp
                         .SetText .ColLetterToNumber(Col_S), .Row, FormatKyHieu(CStr(varTemp))
                         .Col = .ColLetterToNumber(Col_S)
                         UpdateCell_sheet fps, 2, .Col, .Row, .Text
                         
                          'Kiem tra cau truc cua ky hieu neu sai set cot an =1 dung =0
                         .GetText .ColLetterToNumber(Col_S), .Row, varTemp
                         .SetText .ColLetterToNumber("AH"), .Row, CheckSoHD(CStr(varTemp), strLoai)
                         .Col = .ColLetterToNumber("AH")
                         UpdateCell_sheet fps, 2, .Col, .Row, .Text
                
                          'Kiem tra cau truc cua Mau so
                        .GetText .ColLetterToNumber("AI"), .Row, strId
                        strLoai = Trim(Left(strId, 1))
                        If strLoai = "0" Then
                            strTemp = Trim(Mid(strId, 11, 6))
                        ElseIf strLoai = "4" Then
                            strTemp = Trim(Mid(strId, 11, 6))
                        Else
                            strTemp = Trim(Mid(strId, 11, 3))
                        End If
                        .GetText .ColLetterToNumber(Col_O), .Row, varTemp
                        .SetText .ColLetterToNumber(Col_O), .Row, IIf(strLoai = "3", varTemp, UCase(varTemp))
                         .SetText .ColLetterToNumber("AJ"), .Row, CheckMauSoHD(CStr(varTemp), CStr(strLoai), CStr(strTemp))
                         .Col = .ColLetterToNumber("AJ")
                        UpdateCell_sheet fps, 2, .Col, .Row, .Text
                    
                    
                         .Col = .ColLetterToNumber("B")
                         .Text = i
                          i = i + 1
                         .Row = .Row + 1
                    Loop Until .Text = "aa"
                    
                    maHD = ""
                    tenHD = ""
                    .Row = .Row + 3
                    i = 1
                    Do
                        .GetText .ColLetterToNumber("O"), .Row, maHD
                         GetMaHoaDon maHD, tenHD
                         If tenHD <> "" Then
                            arrTemp = Split(tenHD, "###")
                            If UBound(arrTemp) > 0 Then
                                .Col = .ColLetterToNumber("C")
                                .Text = arrTemp(0)
                                UpdateCell fps, .Col, .Row, .Text
                                .Col = .ColLetterToNumber("AI")
                                .Text = CStr(arrTemp(1))
                                UpdateCell fps, .Col, .Row, .Text
                                
        '                        .Col = .ColLetterToNumber("AJ")
        '                        .Text = Left$(CStr(arrTemp(1)), 1)
        '                         UpdateCell fps, .Col, .Row, .Text
                            End If
                            
                            .GetText .ColLetterToNumber("AI"), .Row, strId
                            strLoai = Trim(Left(strId, 1))
                            
                            'Ho tro dinh dang 7 ky tu
                             ' tu so
                             .GetText .ColLetterToNumber(Col_V), .Row, varTemp
                             .SetText .ColLetterToNumber(Col_V), .Row, FormatSoHD(CStr(varTemp))
                             .Col = .ColLetterToNumber(Col_V)
                             UpdateCell_sheet fps, 2, .Col, .Row, .Text
                             ' den so
                             .GetText .ColLetterToNumber(Col_Y), .Row, varTemp
                             .SetText .ColLetterToNumber(Col_Y), .Row, FormatSoHD(CStr(varTemp))
                             .Col = .ColLetterToNumber(Col_Y)
                             UpdateCell_sheet fps, 2, .Col, .Row, .Text
                             
                              ' Ho tro tinh so luong
                            .GetText .ColLetterToNumber(Col_V), .Row, tuSo
                            .GetText .ColLetterToNumber(Col_Y), .Row, denso
                            If Trim(tuSo) <> "" Or Trim(denso) <> "" Then
                                .SetText .ColLetterToNumber(Col_AA), .Row, Val(denso) - Val(tuSo) + 1
                                .Col = .ColLetterToNumber(Col_AA)
                                UpdateCell_sheet fps, 2, .Col, .Row, .Text
                            End If
                            
                            ' chuyen hoa don thanh ky tu in hoa
                           .GetText .ColLetterToNumber(Col_S), .Row, varTemp
                           .SetText .ColLetterToNumber(Col_S), .Row, FormatKyHieu(CStr(varTemp))
                           .Col = .ColLetterToNumber(Col_S)
                            UpdateCell_sheet fps, 2, .Col, .Row, .Text
                         
                            'Kiem tra cau truc cua ky hieu neu sai set cot an =1 dung =0
                            .GetText .ColLetterToNumber(Col_S), .Row, varTemp
                            .SetText .ColLetterToNumber("AH"), .Row, CheckSoHD(CStr(varTemp), strLoai)
                            .Col = .ColLetterToNumber("AH")
                            UpdateCell_sheet fps, 2, .Col, .Row, .Text
                            
                            'Kiem tra cau truc cua Mau so
                            .GetText .ColLetterToNumber("AI"), .Row, strId
                            strLoai = Trim(Left(strId, 1))
                            If strLoai = "0" Then
                                strTemp = Trim(Mid(strId, 11, 6))
                            ElseIf strLoai = "4" Then
                                strTemp = Trim(Mid(strId, 11, 6))
                            Else
                                strTemp = Trim(Mid(strId, 11, 3))
                            End If
                            .GetText .ColLetterToNumber(Col_O), .Row, varTemp
                            .SetText .ColLetterToNumber(Col_O), .Row, IIf(strLoai = "3", varTemp, UCase(varTemp))
                             .SetText .ColLetterToNumber("AJ"), .Row, CheckMauSoHD(CStr(varTemp), CStr(strLoai), CStr(strTemp))
                             .Col = .ColLetterToNumber("AJ")
                            UpdateCell_sheet fps, 2, .Col, .Row, .Text
                             
                         End If
                         
                         
                         .Col = .ColLetterToNumber("B")
                         .Text = i
                          i = i + 1
                         .Row = .Row + 1
                    Loop Until .Text = "bb"

            End If
            
            ' set PL 3.10
            If .ActiveSheet = 3 Then
              maHD = ""
              tenHD = ""
              mCurrentSheet = .Sheet
                .Row = 45
                i = 1
                Do
                    .GetText .ColLetterToNumber("O"), .Row, maHD
                     GetMaHoaDon maHD, tenHD
                     If tenHD <> "" Then
                        arrTemp = Split(tenHD, "###")
                        If UBound(arrTemp) > 0 Then
                            .Col = .ColLetterToNumber("AI")
                            .Text = CStr(arrTemp(1))
                            UpdateCell fps, .Col, .Row, .Text
                        End If
                        
                        'Ho tro format dinh dang tu so
                        .GetText .ColLetterToNumber(Col_V), .Row, varTemp
                        .SetText .ColLetterToNumber(Col_V), .Row, FormatSoHD(CStr(varTemp))
                        .Col = .ColLetterToNumber(Col_V)
                        UpdateCell_sheet fps, 3, .ColLetterToNumber(Col_V), .Row, .Text
                        ' den so
                        .GetText .ColLetterToNumber(Col_Y), .Row, varTemp
                        .SetText .ColLetterToNumber(Col_Y), .Row, FormatSoHD(CStr(varTemp))
                        .Col = .ColLetterToNumber(Col_Y)
                        UpdateCell_sheet fps, 3, .ColLetterToNumber(Col_Y), .Row, .Text
                        
                        ' Ho tro tinh so luong
                        .GetText .ColLetterToNumber(Col_V), .Row, tuSo
                        .GetText .ColLetterToNumber(Col_Y), .Row, denso
                          If Trim(tuSo) <> "" Or Trim(denso) <> "" Then
                              .SetText .ColLetterToNumber(Col_AA), .Row, Val(denso) - Val(tuSo) + 1
                              .Col = .ColLetterToNumber(Col_AA)
                              UpdateCell_sheet fps, 3, .Col, .Row, .Text
                          End If
                          
                          ' Ho tro format ky hieu
                          ' Chuyen ky hieu HD thanh chu cai in hoa
                             .GetText .ColLetterToNumber(Col_S), .Row, varTemp
                             .SetText .ColLetterToNumber(Col_S), .Row, FormatKyHieu(CStr(varTemp))
                             .Col = .ColLetterToNumber(Col_S)
                            UpdateCell_sheet fps, 3, .Col, .Row, .Text
                             
                            'Kiem tra cau truc cua ky hieu neu sai set cot an =1 dung =0
                             .GetText .ColLetterToNumber(Col_S), .Row, varTemp
                             .SetText .ColLetterToNumber("AH"), .Row, CheckSoHD(CStr(varTemp), strLoai)
                             .Col = .ColLetterToNumber("AH")
                            UpdateCell_sheet fps, 3, .Col, .Row, .Text
                            'Kiem tra cau truc cua Mau so
                            .GetText .ColLetterToNumber("AI"), .Row, strId
                            strLoai = Trim(Left(strId, 1))
                            If strLoai = "0" Then
                                strTemp = Trim(Mid(strId, 11, 6))
                            ElseIf strLoai = "4" Then
                                strTemp = Trim(Mid(strId, 11, 6))
                            Else
                                strTemp = Trim(Mid(strId, 11, 3))
                            End If
                            .GetText .ColLetterToNumber(Col_O), .Row, varTemp
                            .SetText .ColLetterToNumber(Col_O), .Row, IIf(strLoai = "3", varTemp, UCase(varTemp))
                             .SetText .ColLetterToNumber("AJ"), .Row, CheckMauSoHD(CStr(varTemp), CStr(strLoai), CStr(strTemp))
                             .Col = .ColLetterToNumber("AJ")
                            UpdateCell_sheet fps, 3, .Col, .Row, .Text
                     End If
                     .Col = .ColLetterToNumber("B")
                     .Text = i
                     i = i + 1
                     .Row = .Row + 1
                Loop Until .Text = "aa"

            End If
        fps.EventEnabled(EventAllEvents) = True
        End With
End Sub

Private Sub fps_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
    Dim sumLaoDong As Variant, sumThuNhap As Variant, sumLDKhauTru As Variant, sumTNKhauTru As Variant
    Dim sumSoThueTNDaKhauTru As Variant
    Dim varHoTen As Variant, varMST As Variant
    Dim varThuNhap As Variant, varSoThueTNDaKhauTru As Variant
    
    'If flgLoadToKhai = True Then Exit Sub
    
    If NewSheet = 1 Then
        kiemTraDuLieuImport
        With fps
            'Because of unclearly-rised events, this function is here for turning it off
            'at the end of this sub, this event will be turned on
            .EventEnabled(EventAllEvents) = False
            
            .Sheet = 1
                    'dhdang sua loi o check "ky bao cao cc" va "chuyen dia diem"
                If TAX_Utilities_v2.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0" Then
                        .Row = 14
                        .Col = .ColLetterToNumber("B")
                        .Text = "X"
                        .Lock = True
                        
                        .Row = 14
                        .Col = .ColLetterToNumber("G")
                        .Text = ""
                        .Lock = True
                    ElseIf TAX_Utilities_v2.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0" Then
                        .Row = 14
                        .Col = .ColLetterToNumber("B")
                        .Text = ""
                        .Lock = True
                        
                        .Row = 14
                        .Col = .ColLetterToNumber("G")
                        .Text = "X"
                        .Lock = True
                    Else
                        .Row = 14
                        .Col = .ColLetterToNumber("B")
                        .Text = ""
                        .Lock = True
                        
                        .Row = 14
                        .Col = .ColLetterToNumber("G")
                        .Text = ""
                        .Lock = True
                End If
                
                        
            .EventEnabled(EventAllEvents) = True
        End With
    End If
End Sub
Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
'    Dim intIndexCombo As Integer
'    Dim temp As String
'    With fps
'        If .ActiveSheet = 1 Then
'            If Col = .ColLetterToNumber(Col_D) And .Text <> "" Then
'                intIndexCombo = .TypeComboBoxCurSel
'                 .Row = Row
'                 .Col = .ColLetterToNumber(Col_W)
'                 .TypeComboBoxCurSel = intIndexCombo
'                UpdateCell fps, .Col, .Row, .Value
'                temp = .Text
'                .Col = .ColLetterToNumber(Col_E)
'               .Text = temp
'            End If
'        End If
'    End With
End Sub

' Ham kiem tra ct [6],[7],[8],[9]
' phan tu 0: =1 chi co so ton, =2 chi co phat hanh, =3 co ca so ton va phat hanh
' phan tu 1: = 1 tuSo1 trang, =2 denSo1 trang, =3 tuSo1>denSo1 , =4 thoa man
' phan tu 2: = 1 tuSo2 trang, =2 denSo2 trang, =3 tuSo2>denSo2 , =4 thoa man
' phan tu 3 : = 1 denSo1 +1 <> tuSo2
Public Function checkSoTonSoPH(Tuso1 As String, denso1 As String, Tuso2 As String, denso2 As String) As String()
      Dim str(5) As String
      
      If (Tuso1 <> "" Or denso1 <> "") And (Tuso2 = "" And denso2 = "") Then
      ' truong hop chi co so ton
            str(0) = "1"
            If Tuso1 <> "" And denso1 = "" Then
                str(1) = "2"
            ElseIf Tuso1 = "" And denso1 <> "" Then
                str(1) = "1"
            Else
                If Val(denso1) - Val(Tuso1) < 0 Then
                    str(1) = "3"
                Else
                    str(1) = "4"
                End If
            End If
      ElseIf (Tuso1 = "" And denso1 = "") And (Tuso2 <> "" Or denso2 <> "") Then
      ' truong hop chi co so phat hanh
            str(0) = "2"
            If Tuso2 <> "" And denso2 = "" Then
                str(2) = "2"
            ElseIf Tuso2 = "" And denso2 <> "" Then
                str(2) = "1"
            Else
                If Val(denso2) - Val(Tuso2) < 0 Then
                    str(2) = "3"
                Else
                    str(2) = "4"
                End If
            End If
      ElseIf (Tuso1 <> "" Or denso1 <> "") And (Tuso2 <> "" Or denso2 <> "") Then
      ' truong hop co ca so ton va so phat hanh
            str(0) = "3"
            ' kiem tra loi so ton
            If Tuso1 <> "" And denso1 = "" Then
                str(1) = "2"
            ElseIf Tuso1 = "" And denso1 <> "" Then
                str(1) = "1"
            Else
                If Val(denso1) - Val(Tuso1) < 0 Then
                    str(1) = "3"
                Else
                    str(1) = "4"
                End If
            End If
            ' kiem tra loi so phat hanh
            If Tuso2 <> "" And denso2 = "" Then
                str(2) = "2"
            ElseIf Tuso2 = "" And denso2 <> "" Then
                str(2) = "1"
            Else
                If Val(denso2) - Val(Tuso2) < 0 Then
                    str(2) = "3"
                Else
                    str(2) = "4"
                End If
            End If
            If str(2) = "4" And str(1) = "4" Then
                If Val(Tuso2) - Val(denso1) <> 1 And Tuso1 <> "" And denso1 <> "" And Tuso2 <> "" And denso2 <> "" Then
                    str(3) = 1
                End If
            End If
      Else
          str(0) = "0"
      End If
      checkSoTonSoPH = str
End Function

' Ham kiem tra so hoa don xoa, huy, mat co thoa man khong
' tra ve gia tri =-1 loi
Public Function checkHDXoa(Tuso1 As String, denso1 As String, strCheck As String) As String
    Dim strLietKe() As String
    Dim i As Integer
    'test tuning toc do
    Dim arrKhoang() As String
    ' End test tuning
    'strLietKe = getListHD(strCheck)
    ' test
    If isDuLieuHL(strCheck) Then
        If InStr(strCheck, ";") > 0 Then
            strLietKe = Split(strCheck, ";")
            For i = 0 To UBound(strLietKe)
                If InStr(strLietKe(i), "-") > 0 Then
                    arrKhoang = Split(strLietKe(i), "-")
                    If Val(arrKhoang(0)) < Val(Tuso1) Or Val(arrKhoang(1)) > Val(denso1) Then
                        checkHDXoa = -1
                        Exit Function
                    End If
                Else
                    If Val(strLietKe(i)) < Val(Tuso1) Or Val(strLietKe(i)) > Val(denso1) Then
                        checkHDXoa = -1
                        Exit Function
                    End If
                End If
            Next i
        ElseIf InStr(strCheck, "-") > 0 Then
            arrKhoang = Split(strCheck, "-")
            If Val(arrKhoang(0)) < Val(Tuso1) Or Val(arrKhoang(1)) > Val(denso1) Then
                checkHDXoa = -1
                Exit Function
            End If
        Else
            If Val(strCheck) < Val(Tuso1) Or Val(strCheck) > Val(denso1) Then
                checkHDXoa = -1
                Exit Function
            End If
        End If
    End If
    ' end test
'    For i = 0 To UBound(strLietKe)
'        If Val(strLietKe(i)) < Val(Tuso1) Or Val(strLietKe(i)) > Val(denso1) Then
'            checkHDXoa = -1
'            Exit Function
'        End If
'    Next i
    If isTrungSoHD(strCheck) = True Then
        checkHDXoa = -1
        Exit Function
    End If
    checkHDXoa = "0"
End Function

' Ham kiem tra trung so hoa don trong xoa, huy, mat
Public Function isTrungSoHD(str As String) As Boolean
    Dim result As Boolean
    Dim arrstr() As String
    Dim arrstr1() As String
    Dim i As Integer, j As Integer
    Dim strTemp As String
    result = False
    ' Test tuning
    Dim arrKhoang() As String, arrKhoang1() As String
    Dim arrLietKe() As String
    Dim tSo1 As Double, tSo2 As Double, dSo1 As Double, dSo2 As Double
    
    If isDuLieuHL(str) Then
        If InStr(str, ";") > 0 Then
            arrLietKe = Split(str, ";")
            For i = 0 To UBound(arrLietKe)
                ' Neu la khoang thi so sanh khoang
                If InStr(arrLietKe(i), "-") Then
                    arrKhoang = Split(arrLietKe(i), "-")
                    tSo1 = Val(arrKhoang(0))
                    dSo1 = Val(arrKhoang(1))
                    For j = i + 1 To UBound(arrLietKe)
                        ' Neu la khoang thi so sanh khoang
                        If InStr(arrLietKe(j), "-") Then
                            arrKhoang1 = Split(arrLietKe(j), "-")
                            tSo2 = arrKhoang1(0)
                            dSo2 = arrKhoang1(1)
                            If (((tSo1 - tSo2) * (dSo1 - tSo2)) <= 0 Or ((tSo1 - dSo2) * (dSo1 - dSo2)) <= 0 Or ((tSo1 > tSo2) And (dSo1 < dSo2))) Then
                                isTrungSoHD = True
                                Exit Function
                            End If
                        Else
                        ' So HD le
                            If Val(arrLietKe(j)) >= tSo1 And Val(arrLietKe(j)) <= dSo1 Then
                                isTrungSoHD = True
                                Exit Function
                            End If
                        End If
                    Next j
                Else
                ' So HD le
                    For j = i + 1 To UBound(arrLietKe)
                        ' Neu la khoang thi so sanh khoang
                        If InStr(arrLietKe(j), "-") Then
                            arrKhoang1 = Split(arrLietKe(j), "-")
                            tSo2 = arrKhoang1(0)
                            dSo2 = arrKhoang1(1)
                            If tSo2 <= Val(arrLietKe(i)) And dSo2 >= Val(arrLietKe(i)) Then
                                isTrungSoHD = True
                                Exit Function
                            End If
                        Else
                        ' So HD le
                            If Val(arrLietKe(j)) = Val(arrLietKe(i)) Then
                                isTrungSoHD = True
                                Exit Function
                            End If
                        End If
                    Next j
                End If
            Next i
'        ElseIf InStr(str, "-") > 0 Then
'
'        Else
'
        End If
    End If
    isTrungSoHD = result
End Function

' Ham kiem tra trung khoang giua xoa, huy , mat
' chi xet khi co 2 phan tro len
' 1 trung Xoa, Mat
' 2 trung Xoa, Huy
' 3 trung Mat, huy
' 4 trung ca 3
Public Function checkTrungXoa(strXoa As String, strMat As String, strHuy As String) As String
    Dim result As String
    Dim arrXoa() As String, arrMat() As String, arrHuy() As String
    Dim strTemp As String, strTemp1 As String
    Dim isMatHuy As Boolean, isMatXoa As Boolean, isXoaHuy As Boolean
    Dim i As Integer
    ' chi co Xoa, mat
    result = "0"
    If strXoa <> "" And strMat <> "" And strHuy = "" Then
'        arrXoa = getListHD(strXoa)
'        arrMat = getListHD(strMat)
'        For i = 0 To UBound(arrXoa)
'            strTemp = strTemp & "~" & arrXoa(i)
'        Next i
'        'strTemp = Mid$(strTemp, 2)
'        strTemp = strTemp & "~"
'        For i = 0 To UBound(arrMat)
'            If InStr(1, strTemp, "~" & arrMat(i) & "~") > 0 Then
'                checkTrungXoa = "1"
'                Exit Function
'            End If
'        Next i
    ' Test tuning
       If isTrungSoHD(strXoa & ";" & strMat) Then
            checkTrungXoa = "1"
            Exit Function
       End If
    ' end test tuning
    ' chi co xoa huy
    ElseIf strXoa <> "" And strMat = "" And strHuy <> "" Then
'        arrXoa = getListHD(strXoa)
'        arrHuy = getListHD(strHuy)
'        For i = 0 To UBound(arrXoa)
'            strTemp = strTemp & "~" & arrXoa(i)
'        Next i
'        'strTemp = Mid$(strTemp, 2)
'        strTemp = strTemp & "~"
'        For i = 0 To UBound(arrHuy)
'            If InStr(1, strTemp, "~" & arrHuy(i) & "~") > 0 Then
'                checkTrungXoa = "2"
'                Exit Function
'            End If
'        Next i
    ' Test tuning
       If isTrungSoHD(strXoa & ";" & strHuy) Then
            checkTrungXoa = "2"
            Exit Function
       End If
    ' end test tuning
    ' chi co mat huy
    ElseIf strXoa = "" And strMat <> "" And strHuy <> "" Then
'        arrHuy = getListHD(strHuy)
'        arrMat = getListHD(strMat)
'        For i = 0 To UBound(arrHuy)
'            strTemp = strTemp & "~" & arrHuy(i)
'        Next i
'        'strTemp = Mid$(strTemp, 2)
'        strTemp = strTemp & "~"
'        For i = 0 To UBound(arrMat)
'            If InStr(1, strTemp, "~" & arrMat(i) & "~") > 0 Then
'                checkTrungXoa = "3"
'                Exit Function
'            End If
'        Next i
    ' Test tuning
       If isTrungSoHD(strMat & ";" & strHuy) Then
            checkTrungXoa = "3"
            Exit Function
       End If
    ' end test tuning
    ElseIf strXoa <> "" And strMat <> "" And strHuy <> "" Then
'        arrHuy = getListHD(strHuy)
'        arrMat = getListHD(strMat)
'        arrXoa = getListHD(strXoa)
        ' kiem tra 3 loai trung
'        For i = 0 To UBound(arrXoa)
'            strTemp = strTemp & "~" & arrXoa(i)
'        Next i
'        'strTemp = Mid$(strTemp, 2)
'        strTemp = strTemp & "~"
'        ' kiem tra voi so hd xoa
'        For i = 0 To UBound(arrMat)
'            If InStr(1, strTemp, "~" & arrMat(i) & "~") > 0 Then
'                isMatXoa = True
'                Exit For
'            End If
'        Next i
       ' Test tuning
       If isTrungSoHD(strXoa & ";" & strMat) Then
            isMatXoa = True
            'Exit Function
       End If
      ' end test tuning
      ' Kiem tra xoa huy
       ' Test tuning
       If isTrungSoHD(strXoa & ";" & strHuy) Then
            isXoaHuy = True
            'Exit Function
       End If
      ' end test tuning
'        For i = 0 To UBound(arrHuy)
'            If InStr(1, strTemp, "~" & arrHuy(i) & "~") > 0 Then
'                isXoaHuy = True
'                Exit For
'            End If
'        Next i
        ' kiem tra mat va huy
'        For i = 0 To UBound(arrMat)
'            strTemp1 = strTemp1 & "~" & arrMat(i)
'        Next i
'        'strTemp1 = Mid$(strTemp1, 2)
'        strTemp1 = strTemp1 & "~"
'        For i = 0 To UBound(arrHuy)
'            If InStr(1, strTemp1, "~" & arrHuy(i) & "~") > 0 Then
'                isMatHuy = True
'                Exit For
'            End If
'        Next i
       ' Test tuning
       ' kiem tra mat va huy
       If isTrungSoHD(strMat & ";" & strHuy) Then
            isMatHuy = True
            'Exit Function
       End If
      ' end test tuning

        If isMatHuy = True And isMatXoa = False And isXoaHuy = False Then
            result = "3"
        ElseIf isMatXoa = True And isMatHuy = False And isXoaHuy = False Then
            result = "1"
        ElseIf isXoaHuy = True And isMatHuy = False And isMatXoa = False Then
            result = "2"
        ElseIf (isXoaHuy = True And isMatHuy = True And isMatXoa = True) Or (isXoaHuy = False And isMatHuy = True And isMatXoa = True) Or (isXoaHuy = True And isMatHuy = False And isMatXoa = True) Or (isXoaHuy = True And isMatHuy = True And isMatXoa = False) Then
            result = "4"
        Else
            result = "0"
        End If
    Else
        result = "0"
    End If
    checkTrungXoa = result
End Function


' Ham chuyen du lieu khoang thanh du lieu liet ke
' Tam thoi khong dung ham nay 07/11/2011 (toc do cham)
'Public Function getListHD(str As String) As String()
'    Dim strTemp As String
'    Dim strLietKe() As String
'    Dim strKhoang() As String
'    Dim strResult() As String
'    Dim j As Double
'    Dim flag1 As Boolean, flag2 As Boolean
'    Dim i As Integer
'    flag1 = False
'    flag2 = False
'    If InStr(1, str, ";") > 0 Then
'        flag1 = True  ' co danh sach liet ke
'    End If
'    If InStr(1, str, "-") > 0 Then
'        flag2 = True  ' co danh sach khoang
'    End If
'    ' chi nhap theo liet ke
'    If flag1 = True And flag2 = False Then
'        strResult = Split(str, ";")
'    ' truong hop theo khoang hop le
'    ElseIf flag1 = False And flag2 = True Then
'        strKhoang = Split(str, "-")
'        strTemp = ""
'        For j = Val(strKhoang(0)) To Val(strKhoang(1))
'            strTemp = strTemp & "~" & j
'        Next j
'        strTemp = Mid$(strTemp, 2)
'        strResult = Split(strTemp, "~")
'    ' truong hop nhap ca khoang va liet ke
'    ElseIf flag1 = True And flag2 = True Then
'        strLietKe = Split(str, ";")
'        For i = 0 To UBound(strLietKe)
'             If InStr(1, strLietKe(i), "-") > 0 Then
'                strKhoang = Split(strLietKe(i), "-")
'                For j = Val(strKhoang(0)) To Val(strKhoang(1))
'                    strTemp = strTemp & "~" & j
'                Next j
'             Else
'                strTemp = strTemp & "~" & strLietKe(i)
'             End If
'        Next i
'        strTemp = Mid$(strTemp, 2)
'        strResult = Split(strTemp, "~")
'    Else
'        ReDim strResult(0)
'        strResult(0) = str
'    End If
'    getListHD = strResult
'End Function

' kiem tra du lieu hop le
Public Function isDuLieuHL(str As String) As Boolean
    Dim flag1 As Boolean
    Dim flag2 As Boolean
    Dim result As Boolean
    Dim arrLietKe() As String
    Dim arrKhoang() As String
    Dim arrKhoang1() As String
    Dim i As Integer, j As Integer
    result = True
    flag1 = False
    flag2 = False
    If Trim(str) = "" Then
        isDuLieuHL = True
        Exit Function
    End If
    
    If InStr(1, str, ";") > 0 Then
        flag1 = True  ' co danh sach liet ke
    End If
    If InStr(1, str, "-") > 0 Then
        flag2 = True  ' co danh sach khoang
        ' Kiem tra dau "-" ma khong co ";"
        arrKhoang1 = Split(str, "-")
        For i = 1 To UBound(arrKhoang1) - 1
            If InStr(1, arrKhoang1(i), ";") <= 0 Then
                isDuLieuHL = False
                Exit Function
            End If
        Next i
    End If
    
    If flag1 = True And flag2 = False Then
    ' truong hop chi nhap liet ke
        arrLietKe = Split(str, ";")
        For i = 0 To UBound(arrLietKe)
            ' Kiem tra null
            If IsNullValue_ac(arrLietKe(i)) Then
                 result = False
                 Exit For
            End If
            ' Kiem tra dau cach trong so hoa don
            If InStr(1, arrLietKe(i), " ") > 0 Then
                 result = False
                 Exit For
            End If
            ' Kiem tra so hoa don co phai la so hay khong
            If IsNullValue(arrLietKe(i)) = False Then
                If IsNumeric(arrLietKe(i)) = False Then
                    result = False
                    Exit For
                Else
                    If InStr(1, arrLietKe(i), ",") > 0 Or InStr(1, arrLietKe(i), ".") > 0 Then
                        result = False
                        Exit For
                    End If
                End If
            End If
            
        Next i
    ElseIf flag1 = False And flag2 = True Then
    ' truong hop chi nhap khoang
        arrKhoang = Split(str, "-")
        For i = 0 To UBound(arrKhoang)
            ' Kiem tra gia tri null
            If IsNullValue_ac(arrKhoang(i)) Then
                result = False
                Exit For
            End If
            ' Kiem tra dau cach trong so hoa don
            If InStr(1, arrKhoang(i), " ") > 0 Then
                 result = False
                 Exit For
            End If
            ' Kiem tra so hoa don co phai la so hay khong
            If IsNullValue(arrKhoang(i)) = False Then
                If IsNumeric(arrKhoang(i)) = False Then
                    result = False
                    Exit For
                Else
                    If InStr(1, arrKhoang(i), ",") > 0 Or InStr(1, arrKhoang(i), ".") > 0 Then
                        result = False
                        Exit For
                    End If
                End If
            End If
        Next i
    ElseIf flag1 = True And flag2 = True Then
    ' truong hop chi nhap liet ke va khoagn
        arrLietKe = Split(str, ";")
        For i = 0 To UBound(arrLietKe)
            If IsNullValue_ac(arrLietKe(i)) Then
                result = False
                Exit For
            Else
                ' kiem tra xem la phan tu hay khoang
                If InStr(1, arrLietKe(i), "-") > 0 Then
                    arrKhoang = Split(arrLietKe(i), "-")
                    For j = 0 To UBound(arrKhoang)
                        ' Kiem tra so hoa don co null khong
                        If IsNullValue_ac(arrKhoang(j)) Then
                            result = False
                            Exit For
                        End If
                        ' Kiem tra dau cach trong so hoa don
                        If InStr(1, arrKhoang(j), " ") > 0 Then
                             result = False
                             Exit For
                        End If
                        ' Kiem tra so hoa don co phai la so hay khong
                        If IsNullValue(arrKhoang(j)) = False Then
                            If IsNumeric(arrKhoang(j)) = False Then
                                result = False
                                Exit For
                            Else
                                If InStr(1, arrKhoang(j), ",") > 0 Or InStr(1, arrKhoang(j), ".") > 0 Then
                                    result = False
                                    Exit For
                                End If
                            End If
                        End If
                    Next j
                End If
            End If
        Next i
    Else
        If str = "" Then
            result = False
        End If
        ' Kiem tra so hoa don co phai la so khong
        If IsNumeric(str) = False Then
            result = False
        Else
            If InStr(1, str, ",") > 0 Or InStr(1, str, ".") > 0 Then
                result = False
            End If
        End If
        ' Kiem tra dau cach
        If InStr(1, str, " ") > 0 Then
            result = False
        End If
    End If
    isDuLieuHL = result
End Function

' Ham get ve so luong HD
' loi tra ve 0
Public Function GetSoLuongHD(str As String) As String
    Dim result As String
    Dim arrLietKe() As String
    'test tuning
    Dim flag1 As Boolean, flag2 As Boolean
    Dim temp As Double
    Dim strResult() As String
    Dim strKhoang() As String
    Dim strLietKe() As String
    Dim i As Integer, j As Integer
    ' end test tuning
    If isDuLieuHL(str) And Trim(str) <> "" Then
        ' test
        flag1 = False
        flag2 = False
        If InStr(1, str, ";") > 0 Then
            flag1 = True  ' co danh sach liet ke
        End If
        If InStr(1, str, "-") > 0 Then
            flag2 = True  ' co danh sach khoang
        End If
        ' chi nhap theo liet ke
        If flag1 = True And flag2 = False Then
            strResult = Split(str, ";")
            GetSoLuongHD = UBound(strResult) + 1
            Exit Function
        ' truong hop theo khoang hop le
        ElseIf flag1 = False And flag2 = True Then
            strKhoang = Split(str, "-")
            GetSoLuongHD = Val(strKhoang(1)) - Val(strKhoang(0)) + 1
            Exit Function
        ' truong hop nhap ca khoang va liet ke
        ElseIf flag1 = True And flag2 = True Then
            strLietKe = Split(str, ";")
            temp = 0
            For i = 0 To UBound(strLietKe)
                 If InStr(1, strLietKe(i), "-") > 0 Then
                    strKhoang = Split(strLietKe(i), "-")
                    temp = temp + Val(strKhoang(1)) - Val(strKhoang(0)) + 1
                 Else
                    temp = temp + 1
                 End If
            Next i
            GetSoLuongHD = temp
            Exit Function
        Else
            GetSoLuongHD = 1
            Exit Function
        End If
    
        ' End test
        
'        arrLietKe = getListHD(str)
'        result = UBound(arrLietKe) + 1
    Else
        result = "0"
    End If
    GetSoLuongHD = result
End Function

' ham xoa cell errorr
Private Sub clearCellErorr(startRow As Long)
    With fps
        .Row = startRow
        .Sheet = 1
        Do
            ' ten Hoa don
            .Col = .ColLetterToNumber(Col_D)
            .CellNote = ""
            .BackColor = mNonErrorColor
            ' Mau so HD
            .Col = .ColLetterToNumber(Col_E)
            .CellNote = ""
            .BackColor = mNonErrorColor
            ' Ky hieu HD
            .Col = .ColLetterToNumber(Col_F)
            .CellNote = ""
            .BackColor = mNonErrorColor
            ' Tu so 6
            .Col = .ColLetterToNumber(Col_H)
            .CellNote = ""
            .BackColor = mNonErrorColor
            ' Den so 7
            .Col = .ColLetterToNumber(Col_I)
            .CellNote = ""
            .BackColor = mNonErrorColor
            ' Tu so 8
            .Col = .ColLetterToNumber(Col_J)
            .CellNote = ""
            .BackColor = mNonErrorColor
            ' Den so 9
            .Col = .ColLetterToNumber(Col_K)
            .CellNote = ""
            .BackColor = mNonErrorColor
            
            ' Tu so 20
            .Col = .ColLetterToNumber("X")
            .CellNote = ""
            .BackColor = mNonErrorColor
            ' Den so 21
            .Col = .ColLetterToNumber("Y")
            .CellNote = ""
            .BackColor = mNonErrorColor
            ' So da SD
            .Col = .ColLetterToNumber("P")
            .CellNote = ""
            .BackColor = mNonErrorColor
            ' So xoa
            .Col = .ColLetterToNumber(Col_R)
            .CellNote = ""
            .BackColor = mNonErrorColor
            ' So Mat
            .Col = .ColLetterToNumber(Col_T)
            .CellNote = ""
            .BackColor = mNonErrorColor
            ' So Huy
            .Col = .ColLetterToNumber(Col_V)
            .CellNote = ""
            .BackColor = mNonErrorColor
            ' next
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until UCase(.Text) = Col_AA
    End With
End Sub

' Ham load du lieu ton cuoi ky truoc
Public Function loadTonCuoiKy() As String()
    Dim xmlDomLastData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlSectionsTemplate As MSXML.IXMLDOMNode
    Dim xmlSectionTemplate As MSXML.IXMLDOMNode
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    'Dim strOriginDataFileName As String
    Dim strDataFileName As String, str43Target As String
    Dim strTempValue As String, strValue As String
    Dim result() As String
    Dim isInsert As Boolean
    Dim iDom As Integer, lCtrl As Integer, i As Integer, slTonCuoiKy As Double
    
    Dim MY As Date
    
    iDom = 0
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        MY = DateSerial(TAX_Utilities_v2.Year, getThangDauKy(TAX_Utilities_v2.ThreeMonths), 1)
        
        MY = DateAdd(Col_M, -3, MY)
        MY = DateSerial(Year(MY), getKyBC(Month(MY)), 1)
        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(GetValidityNode("68", Format(Month(MY), "0#"), , Format(Year(MY), "####")).childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        '***************************************
        If .Load(strDataFileName) = True Then
            Set xmlSectionsTemplate = .getElementsByTagName("Sections")(0)
            For lCtrl = 1 To xmlSectionsTemplate.childNodes.length - 2
                Set xmlSectionTemplate = xmlSectionsTemplate.childNodes(lCtrl)
                For i = 0 To xmlSectionTemplate.childNodes.length - 1
                    Set xmlNodeCells = xmlSectionTemplate.childNodes(i)
                    Set xmlNodeCell = xmlNodeCells.childNodes(21)  ' so luong
                    slTonCuoiKy = Val(GetAttribute(xmlNodeCell, "Value"))
                    strTempValue = ""
                    If slTonCuoiKy > 0 Then
                        Set xmlNodeCell = xmlNodeCells.childNodes(2)  ' Mau so
                        strTempValue = GetAttribute(xmlNodeCell, "Value")
                        Set xmlNodeCell = xmlNodeCells.childNodes(3)  ' Ky hieu
                        strTempValue = strTempValue & "~" & GetAttribute(xmlNodeCell, "Value")
                        Set xmlNodeCell = xmlNodeCells.childNodes(19)  ' so ton tu so
                        strTempValue = strTempValue & "~" & GetAttribute(xmlNodeCell, "Value")
                        Set xmlNodeCell = xmlNodeCells.childNodes(20)  ' so ton den so
                        strTempValue = strTempValue & "~" & GetAttribute(xmlNodeCell, "Value")
                        Set xmlNodeCell = xmlNodeCells.childNodes(22)  ' so strCheck
                        strTempValue = strTempValue & "~" & GetAttribute(xmlNodeCell, "Value")
                        Set xmlNodeCell = xmlNodeCells.childNodes(1)  ' so ten
                        strTempValue = strTempValue & "~" & GetAttribute(xmlNodeCell, "Value")
                        Set xmlNodeCell = xmlNodeCells.childNodes(0)  ' Ma hoa don
                        strTempValue = strTempValue & "~" & GetAttribute(xmlNodeCell, "Value")
                        isInsert = True
                        strValue = strValue & "###" & strTempValue
                    End If
                Next i
            Next lCtrl
            If isInsert = True Then
                strValue = Mid$(strValue, 4)
                result = Split(strValue, "###")
                numRowInsert = UBound(result)
            Else
                numRowInsert = -1
            End If
        Else
            numRowInsert = -1
        End If
    End With
    loadTonCuoiKy = result
End Function

' Lay ve thang dau tien cua ky
Private Function getThangDauKy(strPeriod As String) As Integer
    On Error GoTo ErrHandle
    Select Case strPeriod
        Case "1"
            getThangDauKy = 1
        Case "2"
            getThangDauKy = 4
        Case "3"
            getThangDauKy = 7
        Case "4"
            getThangDauKy = 10
    End Select
    Exit Function
ErrHandle:
    SaveErrorLog "bc26", "getThangDauKy", Err.number, Err.Description
End Function
' get ve ky
Private Function getKyBC(thang As Integer) As String
    On Error GoTo ErrHandle
    Select Case thang
        Case 1
            getKyBC = "01"
        Case 4
            getKyBC = "02"
        Case 7
            getKyBC = "03"
        Case 10
            getKyBC = "04"
    End Select
    Exit Function
ErrHandle:
    SaveErrorLog "bc26", "getKyBC", Err.number, Err.Description
End Function
Public Function GetNguoiKy() As String
    Dim xmlDomHeader As New MSXML.DOMDocument
    
    xmlDomHeader.Load GetAbsolutePath(TAX_Utilities_v2.DataFolder & "Header_01.xml")
    GetNguoiKy = GetAttribute(xmlDomHeader.nodeFromID("C_26"), "Value")
    
    Set xmlDomHeader = Nothing
End Function


' Ham tra ve danh sach cac hoa don
Function getListHDTS_DS(tuSo As String, denso As String) As String
    Dim i As Double
    Dim strResult As String
    For i = Val(tuSo) To Val(denso)
        strResult = strResult & i & ";"
    Next i
    getListHDTS_DS = strResult
End Function


' Xoa het trang thai check loi tren sheet hdr
Private Sub clearCellErorrHDR()
    With fps
        Dim i As Long
        .Sheet = .SheetCount
        For i = 12 To 49
            .SetText .ColLetterToNumber("B"), i, "1"
        Next i
    End With
End Sub


' ham set thong tin sau khi tai XML
Public Sub finishImportXML()
    Dim maHD    As Variant
    Dim tenHD  As String
    Dim arrTemp() As String
    
    With fps
        .EventEnabled(EventAllEvents) = False
          .Sheet = 1
          mCurrentSheet = .Sheet
            .Row = 22
            Do
                .GetText .ColLetterToNumber("E"), .Row, maHD
                 GetMaHoaDon maHD, tenHD
                 If tenHD <> "" Then
                    arrTemp = Split(tenHD, "###")
                    If UBound(arrTemp) > 0 Then
                        .Col = .ColLetterToNumber("C")
                        .Text = arrTemp(0)
                        UpdateCell fps, .Col, .Row, .Text
                        .Col = .ColLetterToNumber("AA")
                        .Text = CStr(arrTemp(1))
                        UpdateCell fps, .Col, .Row, .Text
                        
                        .Col = .ColLetterToNumber("AC")
                        .SetText .Col, .Row, "0"
                        UpdateCell fps, .Col, .Row, .Text
                    End If
                 End If
                 
                 
                 .Col = .ColLetterToNumber("B")
                 .Row = .Row + 1
            Loop Until .Text = "aa"
            
            
            .Sheet = 2
            If .SheetVisible = True Then
              maHD = ""
              tenHD = ""
              mCurrentSheet = .Sheet
                .Row = 46
                Do
                    .GetText .ColLetterToNumber("O"), .Row, maHD
                     GetMaHoaDon maHD, tenHD
                     If tenHD <> "" Then
                        arrTemp = Split(tenHD, "###")
                        If UBound(arrTemp) > 0 Then
                            .Col = .ColLetterToNumber("C")
                            .Text = arrTemp(0)
                            UpdateCell fps, .Col, .Row, .Text
                            .Col = .ColLetterToNumber("AI")
                            .Text = CStr(arrTemp(1))
                            UpdateCell fps, .Col, .Row, .Text
                            
                            .Col = .ColLetterToNumber("AJ")
                            .Text = CheckMauSoHD(CStr(maHD), Left$(CStr(arrTemp(1)), 1), arrTemp(0))
                             UpdateCell fps, .Col, .Row, .Text
                        End If
                     End If
                     
                     
                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                Loop Until .Text = "aa"
                
                maHD = ""
                tenHD = ""
                .Row = .Row + 3
                Do
                    .GetText .ColLetterToNumber("O"), .Row, maHD
                     GetMaHoaDon maHD, tenHD
                     If tenHD <> "" Then
                        arrTemp = Split(tenHD, "###")
                        If UBound(arrTemp) > 0 Then
                            .Col = .ColLetterToNumber("C")
                            .Text = arrTemp(0)
                            UpdateCell fps, .Col, .Row, .Text
                            .Col = .ColLetterToNumber("AI")
                            .Text = CStr(arrTemp(1))
                            UpdateCell fps, .Col, .Row, .Text
                            
    '                        .Col = .ColLetterToNumber("AJ")
    '                        .Text = Left$(CStr(arrTemp(1)), 1)
    '                         UpdateCell fps, .Col, .Row, .Text
                        End If
                     End If
                     
                     
                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                Loop Until .Text = "bb"
                
             End If
             
             
             .Sheet = 3
             If .SheetVisible = True Then
              maHD = ""
              tenHD = ""
              mCurrentSheet = .Sheet
                .Row = 45
                Do
                    .GetText .ColLetterToNumber("O"), .Row, maHD
                     GetMaHoaDon maHD, tenHD
                     If tenHD <> "" Then
                        arrTemp = Split(tenHD, "###")
                        If UBound(arrTemp) > 0 Then
                            .Col = .ColLetterToNumber("C")
                            .Text = arrTemp(0)
                            UpdateCell fps, .Col, .Row, .Text
                            .Col = .ColLetterToNumber("AI")
                            .Text = CStr(arrTemp(1))
                            UpdateCell fps, .Col, .Row, .Text
                            
    '                        .Col = .ColLetterToNumber("AJ")
    '                        .Text = Left$(CStr(arrTemp(1)), 1)
    '                         UpdateCell fps, .Col, .Row, .Text
                        End If
                     End If
                     
                     
                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                Loop Until .Text = "aa"
             End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
