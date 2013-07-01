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
'this Class is belong to TAX_Business project which will be compline to DLL

Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrSolanBosung As String
Public flgLoadToKhai As Boolean
Private totalRow2ABH As Long
Public numRowInsert As Integer
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
                    strCombo = strCombo + CPab(arrDanhsach(0), 10) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(0), 200) + Chr$(9)
'                    strCombo = strCombo + CPab(arrDanhsach(1), 10) + Chr$(9)
                    strComboHien = strComboHien + arrDanhsach(2) + Chr$(9)
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
        .Col = .ColLetterToNumber(Col_D)
        
        .TypeComboBoxList = strComboHien
        .Text = strQuanHeNNT
        .Col = .ColLetterToNumber(Col_AA)
        .TypeComboBoxList = strCombo
        'SetDateFormat fps, 1, 22, .ColLetterToNumber(Col_G), DDMMYYYY
         'SetKyHieuHDFormat fps, 1, 22, .ColLetterToNumber(Col_F)
        
        ' Set format ngay ky
        SetDateFormat fps, 1, 32, .ColLetterToNumber(Col_N), DDMMYYYY
        .Row = 32
        .Col = .ColLetterToNumber(Col_N)
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
    'dhdag them phu luc 2 va 3
        .Sheet = 2

            SetDateFormat fps, 2, 56, .ColLetterToNumber(Col_R), DDMMYYYY
            .Row = 56
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
    Dim result() As String
    Dim arrCellInRow() As String
    Dim strTemp As String
    result = loadTonCuoiKy
    
    With fps
        .Sheet = 1
        .EventEnabled(EventAllEvents) = False
        
        If numRowInsert >= 0 And TAX_Utilities_New.CheckNewDataBC26 = True Then
            .Row = 22
            For i = 0 To numRowInsert
                ' set du lieu ton dau ky
                strTemp = result(i)
                mCurrentSheet = 1
                arrCellInRow = Split(strTemp, "~")
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
                .Row = .Row + 1
            Next i
        End If

        
        
        i = 1
        .Col = .ColLetterToNumber("B")
        .Row = 22
        Do
             .Col = .ColLetterToNumber("C")
             .Text = "0"
             
             .Col = .ColLetterToNumber("B")
             .Row = i + 22
             totalRow2ABH = i
             i = i + 1
        Loop Until .Text = "AA"
        
        'dhdang sua theo yc chi en ngay 25/01
        .Row = totalRow2ABH + 29
        .Col = .ColLetterToNumber(Col_N)
        If Trim(.Text) = "" Then
            .SetText .ColLetterToNumber(Col_N), totalRow2ABH + 29, GetNguoiKy()
        End If
        
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
                UpdateCell fps, fps.Col, fps.Row, IIf(fps.Value = 1, "x", vbNullString) 'fps.Value
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
         If Col = .ColLetterToNumber(Col_D) And .Text <> "" Then
            intIndexCombo = .TypeComboBoxCurSel
            .Col = .ColLetterToNumber(Col_AA)
            .TypeComboBoxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber(Col_AA)
            UpdateCell_sheet fps, 1, .Col, .Row, .Text
             
            
            .GetText .ColLetterToNumber(Col_AA), .Row, strId
            
            
            
            .Col = .ColLetterToNumber(Col_AA)
            If Trim(Left(strId, 1)) = "0" Then
                .SetText .ColLetterToNumber(Col_E), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_E), .Row, .Text
            Else
                .SetText .ColLetterToNumber(Col_E), .Row, Trim(Mid(strId, 11, 3))
                UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_E), .Row, .Text
            End If
            .SetText .ColLetterToNumber(Col_AC), .Row, Trim(Left(strId, 1))
            UpdateCell_sheet fps, 1, .ColLetterToNumber(Col_AC), .Row, .Text
            
            .EventEnabled(EventChange) = True
         End If
    ElseIf .ActiveSheet = 2 Then
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("C") And .Text <> "" Then
            intIndexCombo = .TypeComboBoxCurSel
            .Col = .ColLetterToNumber("AI")
            .TypeComboBoxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber("AI")
            UpdateCell_sheet fps, 2, .Col, .Row, .Text
             
            
            .GetText .ColLetterToNumber("AI"), .Row, strId

            
            .Col = .ColLetterToNumber(Col_AA)
            If Trim(Left(strId, 1)) = "0" Then
                .SetText .ColLetterToNumber(Col_O), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell_sheet fps, 2, .ColLetterToNumber(Col_O), .Row, .Text
            Else
                .SetText .ColLetterToNumber(Col_O), .Row, Trim(Mid(strId, 11, 3))
                UpdateCell_sheet fps, 2, .ColLetterToNumber(Col_O), .Row, .Text
            End If
            .EventEnabled(EventChange) = True
         End If
    ElseIf .ActiveSheet = 3 Then
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("C") And .Text <> "" Then
            intIndexCombo = .TypeComboBoxCurSel
            .Col = .ColLetterToNumber("AI")
            .TypeComboBoxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber("AI")
            UpdateCell_sheet fps, 3, .Col, .Row, .Text
             
            
            .GetText .ColLetterToNumber("AI"), .Row, strId
            
            
            
            .Col = .ColLetterToNumber(Col_AA)
            If Trim(Left(strId, 1)) = "0" Then
                .SetText .ColLetterToNumber(Col_O), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell_sheet fps, 3, .ColLetterToNumber(Col_O), .Row, .Text
            Else
                .SetText .ColLetterToNumber(Col_O), .Row, Trim(Mid(strId, 11, 3))
                UpdateCell_sheet fps, 3, .ColLetterToNumber(Col_O), .Row, .Text
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
                If (KeyAscii < 48 And KeyAscii <> 3 And KeyAscii <> 22 And KeyAscii <> 13) Or KeyAscii > 57 Then
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
                If (KeyAscii < 48 And KeyAscii <> 3 And KeyAscii <> 22 And KeyAscii <> 13) Or KeyAscii > 57 Then
                    KeyAscii = 0
                End If
            End If
        ElseIf .ActiveSheet = 3 Then
            If .Col = .ColLetterToNumber(Col_V) Or .Col = .ColLetterToNumber(Col_Y) Then
                If (KeyAscii < 48 And KeyAscii <> 3 And KeyAscii <> 22 And KeyAscii <> 13) Or KeyAscii > 57 Then
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
                    Loop Until .Text = "AA"
                    
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
                    Loop Until .Text = "AA"
                    
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
                    Loop Until .Text = "AA"
                    
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
    Dim denSo As Variant
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
    
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .ActiveSheet
            If .ActiveSheet = 1 Then
                    ' Chuyen ky hieu HD thanh chu cai in hoa
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
                        'Ho tro set tu so [10]
                        .GetText .ColLetterToNumber(Col_H), Row, strTuso1
                        .GetText .ColLetterToNumber(Col_I), Row, strDenso1
                        .GetText .ColLetterToNumber(Col_J), Row, strTuso2
                        .GetText .ColLetterToNumber(Col_K), Row, strDenso2
                        str = getTusoDenso(CStr(strTuso1), CStr(strDenso1), CStr(strTuso2), CStr(strDenso2))
                        If str(0) <> "-1" Then
                             .SetText .ColLetterToNumber(Col_L), Row, FormatSoHD(str(0))
                             UpdateCell fps, .ColLetterToNumber(Col_L), Row, .Text
                             .SetText .ColLetterToNumber(Col_G), Row, FormatSoHD(str(2))
                             UpdateCell fps, .ColLetterToNumber(Col_G), Row, .Text
                        End If
                        ' Ho tro tinh chi tieu [11]
                        .GetText .ColLetterToNumber(Col_M), Row, strCt11
                        .GetText .ColLetterToNumber(Col_N), Row, strCt12
                        .GetText .ColLetterToNumber(Col_G), Row, strCt5
                        If Val(strCt5) > Val(strCt12) Then
                            soluongTonCK = Val(strCt5) - Val(strCt12)
                            .SetText .ColLetterToNumber("X"), Row, FormatSoHD(Val(strCt11) + 1)
                            UpdateCell fps, .ColLetterToNumber("X"), Row, .Text
                            .SetText .ColLetterToNumber(Col_Y), Row, FormatSoHD(Val(strCt11) + soluongTonCK)
                            UpdateCell fps, .ColLetterToNumber(Col_Y), Row, .Text
                            .SetText .ColLetterToNumber(Col_Z), Row, soluongTonCK
                            UpdateCell fps, .ColLetterToNumber(Col_Z), Row, .Text
                        End If
                    End If
                'Tinh so hoa don mat, xoa, huy
                    If Col = .ColLetterToNumber(Col_R) Or Col = .ColLetterToNumber(Col_T) Or Col = .ColLetterToNumber(Col_V) Then
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
                        .SetText .ColLetterToNumber(Col_M), Row, FormatSoHD(CStr(strCt11))
                        UpdateCell fps, .ColLetterToNumber(Col_M), Row, .Text
                                                        
                        .GetText .ColLetterToNumber(Col_G), Row, strCt5
                        If Val(strCt5) > Val(strCt12) Then
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
                        .SetText .ColLetterToNumber(Col_M), Row, FormatSoHD(CStr(strCt11))
                        UpdateCell fps, .ColLetterToNumber(Col_M), Row, .Text
                        
                        If Val(strCt5) > Val(strCt12) Then
                            soluongTonCK = Val(strCt5) - Val(strCt12)
                            .SetText .ColLetterToNumber("X"), Row, FormatSoHD(Val(strCt11) + 1)
                            UpdateCell fps, .ColLetterToNumber("X"), Row, .Text
                            .SetText .ColLetterToNumber(Col_Y), Row, FormatSoHD(Val(strCt11) + soluongTonCK)
                            UpdateCell fps, .ColLetterToNumber(Col_Y), Row, .Text
                            .SetText .ColLetterToNumber(Col_Z), Row, soluongTonCK
                            UpdateCell fps, .ColLetterToNumber(Col_Z), Row, .Text
                        End If
                    End If
                   
                  ' Chuyen cac ky hieu mau so HD thanh chu in hoa
                  If Col = .ColLetterToNumber(Col_E) Then
                        .GetText Col, Row, varTemp
                        .SetText Col, Row, UCase$(CStr(varTemp))
                        UpdateCell fps, Col, Row, .Text
                  End If
            ElseIf .ActiveSheet = 2 Then
                    ' Dem so dong dong sec 1
                    .Row = 46
                    Do
                        countRowDynamic1_PL1 = countRowDynamic1_PL1 + 1
                        .Col = .ColLetterToNumber("B")
                        .Row = .Row + 1
                    Loop Until .Text = "AA"
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
                        .GetText .ColLetterToNumber(Col_Y), Row, denSo
                        
                        .SetText .ColLetterToNumber(Col_AA), Row, Val(denSo) - Val(tuSo) + 1
                        .Col = .ColLetterToNumber(Col_AA)
                        .Row = Row
                        UpdateCell_sheet fps, 2, .Col, .Row, .Text
                        
                        
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
                         .SetText .ColLetterToNumber("AH"), Row, CheckSoHD(CStr(varTemp))
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
                        Else
                            strTemp = Trim(Mid(strId, 11, 3))
                        End If
                        .GetText Col, Row, varTemp
                        .SetText Col, Row, UCase(varTemp)
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
                    Loop Until .Text = "AA"
                    
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
                        .GetText .ColLetterToNumber(Col_Y), Row, denSo
                        
                        .SetText .ColLetterToNumber(Col_AA), Row, Val(denSo) - Val(tuSo) + 1
                        .Col = .ColLetterToNumber(Col_AA)
                        .Row = Row
                       UpdateCell_sheet fps, 3, .Col, .Row, .Text
                        
                        
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
                         .SetText .ColLetterToNumber("AH"), Row, CheckSoHD(CStr(varTemp))
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
                        Else
                            strTemp = Trim(Mid(strId, 11, 3))
                        End If
                        .GetText Col, Row, varTemp
                        .SetText Col, Row, UCase(varTemp)
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
            checkSheet1
            If TAX_Utilities_New.NodeValidity.childNodes(.SheetCount - 4).Attributes.getNamedItem("Active").nodeValue <> "0" Then
                checkSheet2
            End If
            If TAX_Utilities_New.NodeValidity.childNodes(.SheetCount - 3).Attributes.getNamedItem("Active").nodeValue <> "0" Then
                checkSheet3
            End If
    End With
    'result = loadTonCuoiKy
    
End Sub

Private Sub checkSheet1()

    Dim strCheck As String
    Dim errMauSoHD As Variant, errKyHieuHD As Variant, errSoLuong As Variant, errSoLienHD As Variant
    Dim errKhoangGN As Variant, errLengthHD As Variant, errEmty As Variant, errNgayKy As Variant, errNgayHuy As Variant, errTrungKhoang As Variant, errLienKhoang As Variant, errSoXoa As Variant, errSaiDD As Variant
    Dim errSLSudung As Variant, errSLXoa As Variant, errSLMat As Variant, errSLHuy As Variant, errSLSD_Xoa_Mat_Huy As Variant
    Dim strTenHD As Variant, strMauSoHD As Variant, strKyHieuHD As Variant, strTuSo6 As Variant, strDenSo7 As Variant, strTuso8 As Variant, strDenso9 As Variant, strTuSo10 As Variant, strDenSo11 As Variant
    Dim strSLSuDung As Variant, strSLMat As Variant, strSLhuy As Variant, strSLXoa As Variant, strTongSo5 As Variant, strTongSo12 As Variant

    Dim strTenHD1 As Variant, strMauSoHD1 As Variant, strKyHieuHD1 As Variant, strTuso1 As Variant, strDenso1 As Variant
    Dim strSoLuong As Variant, strNguoiKy As Variant, strNguoiLapBieu As Variant, strNgayKy As Variant, strNgayHuy As Variant
    Dim iFlagTenHD As Integer, iFlagMauSoHD As Integer, iFlagKyHieu As Integer, iFlagSoLienHD As Integer, iFlagEmty As Integer
    Dim iWhiteRow As Integer, iFlagLengHD As Integer, iFlagSoLuong As Integer, iFlagNgayKy As Integer, iFlagNgayHuy As Integer
    Dim iFlagKhoangTSDSTrung As Integer, iFocusFlag As Integer, iFlagLienKhoang As Integer, iFlagSoXoa As Integer, iFlagSaiDD As Integer
    Dim iFlagSLSudung As Integer, iFlagSLXoa As Integer, iFlagSLMat As Integer, iFlagSLHuy As Integer, iFlagSD_Xoa_Mat_Huy As Integer
    
    Dim strId As Variant, strLoai As String, strTemp As String, result As String
    Dim i As Integer, countDynamicRow As Integer, j As Integer, k As Integer
    
    Dim checkXoa As Boolean, checkMat As Boolean, checkHuy As Boolean
    
    Dim str() As String
    Dim strLietKe() As String
    
    Dim strTrung67() As String, strTrung89() As String, strCheckTrung() As String
    strTrung67 = checkTrung(fps, 22, Col_H, Col_I, "B", Col_D, Col_E, Col_F, "", False, 1)
    strTrung89 = checkTrung(fps, 22, Col_J, Col_K, "B", Col_D, Col_E, Col_F, "", False, 1)
    
    
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
            .GetText .ColLetterToNumber(Col_D), .Row, strTenHD
            .GetText .ColLetterToNumber(Col_E), .Row, strMauSoHD
            .GetText .ColLetterToNumber(Col_F), .Row, strKyHieuHD
            .GetText .ColLetterToNumber(Col_H), .Row, strTuSo6
            .GetText .ColLetterToNumber(Col_I), .Row, strDenSo7
            .GetText .ColLetterToNumber(Col_J), .Row, strTuso8
            .GetText .ColLetterToNumber(Col_K), .Row, strDenso9
            .GetText .ColLetterToNumber("P"), .Row, strSLSuDung
            .GetText .ColLetterToNumber(Col_R), .Row, strSLXoa
            .GetText .ColLetterToNumber(Col_T), .Row, strSLMat
            .GetText .ColLetterToNumber(Col_V), .Row, strSLhuy
            .GetText .ColLetterToNumber(Col_L), .Row, strTuSo10
            .GetText .ColLetterToNumber(Col_M), .Row, strDenSo11
            .GetText .ColLetterToNumber(Col_G), .Row, strTongSo5
            .GetText .ColLetterToNumber(Col_N), .Row, strTongSo12
            
            .GetText .ColLetterToNumber(Col_AA), .Row, strId
            'check mau so hoa don
           If Trim(strTenHD) <> "" Or Trim(strMauSoHD) <> "" Or Trim(strKyHieuHD) <> "" Then
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
                    .Col = .ColLetterToNumber(Col_E)
                    .CellNote = errSoLienHD  'static
                    .BackColor = mErrorColor
                    iFlagSoLienHD = 1
                End If
            End If
            
            'Kiem tra cau truc cua ky hieu
            If strKyHieuHD = "" And iWhiteRow = 0 Then
                 .Col = .ColLetterToNumber(Col_F)
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                result = CheckSoHD(CStr(strKyHieuHD))
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
                        If strSLXoa <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                            If isDuLieuHL(CStr(strSLXoa)) Then
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
                        ' kiem tra so mat
                        If strSLMat <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                            If isDuLieuHL(CStr(strSLMat)) Then
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
                        ' kiem tra so huy
                        If strSLhuy <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                            If isDuLieuHL(CStr(strSLhuy)) Then
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
                        If strSLXoa <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                            If isDuLieuHL(CStr(strSLXoa)) Then
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
                        ' kiem tra so mat
                        If strSLMat <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                            If isDuLieuHL(CStr(strSLMat)) Then
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
                        ' kiem tra so huy
                        If strSLhuy <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                            If isDuLieuHL(CStr(strSLhuy)) Then
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
                    If str(3) = "1" Then
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
                        If strSLXoa <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                            If isDuLieuHL(CStr(strSLXoa)) Then
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
                        ' kiem tra so mat
                        If strSLMat <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                            If isDuLieuHL(CStr(strSLMat)) Then
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
                        ' kiem tra so huy
                        If strSLhuy <> "" And strTuSo10 <> "" And strDenSo11 <> "" Then
                            If isDuLieuHL(CStr(strSLhuy)) Then
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
                Else
                    .Row = countDynamicRow + 31
                    .Col = .ColLetterToNumber(Col_N)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
         End If
        
      
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
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
Private Sub checkSheet2()
    Dim strCheck As String
    Dim errMauSoHD As Variant, errKyHieuHD As Variant, errSoLuong As Variant, errSoLienHD As Variant
    Dim errKhoangGN As Variant, errLengthHD As Variant, errEmty As Variant, errNgayKy As Variant, errNgayHuy As Variant, errTrungKhoang As Variant
    Dim strTenHD As Variant, strMauSoHD As Variant, strKyHieuHD As Variant, strTuSo As Variant, strDenSo As Variant
    Dim strTenHD1 As Variant, strMauSoHD1 As Variant, strKyHieuHD1 As Variant, strTuso1 As Variant, strDenso1 As Variant
    Dim strSoLuong As Variant, strNguoiKy As Variant, strNguoiLapBieu As Variant, strNgayKy As Variant, strNgayHuy As Variant
    Dim iFlagTenHD As Integer, iFlagMauSoHD As Integer, iFlagKyHieu As Integer, iFlagSoLienHD As Integer, iFlagEmty As Integer
    Dim iWhiteRow As Integer, iFlagLengHD As Integer, iFlagSoLuong As Integer, iFlagNgayKy As Integer, iFlagNgayHuy As Integer
    Dim iFlagKhoangTSDSTrung As Integer, iFocusFlag As Integer
    
    Dim strId As Variant, strLoai As String, strTemp As String, result As String
    Dim i As Integer, countDynamicRow As Integer, j As Integer, k As Integer
    
    Dim str() As String
    ' tra ve mang cac row bi trung khoang
    str = checkTrung_01(fps, 46, Col_V, Col_Y, "B", "C", Col_O, Col_S, "", False, 2)
    
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
        .Col = .ColLetterToNumber("B")
        .Row = 46
        'i = 1
        Do
            j = .Row
            countDynamicRow = countDynamicRow + 1
             ' Lay cac gia tri cua 1 row
             iWhiteRow = 1
            .GetText .ColLetterToNumber("C"), .Row, strTenHD
            .GetText .ColLetterToNumber(Col_O), .Row, strMauSoHD
            .GetText .ColLetterToNumber(Col_S), .Row, strKyHieuHD
            .GetText .ColLetterToNumber(Col_V), .Row, strTuSo
            .GetText .ColLetterToNumber(Col_Y), .Row, strDenSo
            .GetText .ColLetterToNumber(Col_AA), .Row, strSoLuong
            .GetText .ColLetterToNumber("AI"), .Row, strId
'
'            If Trim(strTenHD) <> "" Or Trim(strMauSoHD) <> "" Or Trim(strKyHieuHD) <> "" Or Trim(strTuSo) <> "" Or Trim(strDenSo) <> "" Or Trim(strSoLuong) <> "0" Then
'                ' Khong phai dong trang
'                iWhiteRow = 0
'            End If
            If Trim(strTenHD) <> "" Or Trim(strMauSoHD) <> "" Then
                ' Khong phai dong trang
                iWhiteRow = 0
            End If
            'Ten HD
            If Trim(strTenHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("C")
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                .Col = .ColLetterToNumber("C")
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
                    .Col = .ColLetterToNumber(Col_O)
                    .CellNote = errSoLienHD  'static
                    .BackColor = mErrorColor
                    iFlagSoLienHD = 1
                Else
                    .Col = .ColLetterToNumber(Col_O)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                
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
                result = CheckSoHD(CStr(strKyHieuHD))
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
                If Trim(strTuSo) = "" And Trim(strDenSo) = "" Then
                    ' tu so
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    ' den so
                    .Col = .ColLetterToNumber(Col_Y)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                ElseIf Trim(strTuSo) = "" And Trim(strDenSo) <> "" Then
                    ' tu so
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                    ' den so
                    .Col = .ColLetterToNumber(Col_Y)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                ElseIf Trim(strTuSo) <> "" And Trim(strDenSo) = "" Then
                    .Col = .ColLetterToNumber(Col_Y)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                    ' tu so
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNote = ""
                    .BackColor = mNonErrorColor

                Else
                    If Val(strDenSo) - Val(strTuSo) >= 0 And Val(strDenSo) > 0 Then
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
            .Row = j
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
                
            If .Text = "AA" Then
                .Row = .Row + 3
            End If
        Loop Until .Text = "bb"
                
        ' Kiem tra du lieu nguoi dai dien
         .GetText .ColLetterToNumber(Col_R), countDynamicRow + 52, strNguoiKy
         .GetText .ColLetterToNumber(Col_D), countDynamicRow + 52, strNguoiLapBieu
         .GetText .ColLetterToNumber(Col_R), countDynamicRow + 54, strNgayKy
         .GetText .ColLetterToNumber(Col_O), 8, strNgayHuy
         ' kiem tra nguoi ky bat buoc nhap
         If Trim(strNguoiKy) = "" Then
            .Row = countDynamicRow + 52
            .Col = .ColLetterToNumber(Col_R)
            .CellNote = errEmty  'static
            .BackColor = mErrorColor
            iFlagEmty = 1
         Else
            .Row = countDynamicRow + 52
            .Col = .ColLetterToNumber(Col_R)
            .CellNote = ""
            .BackColor = mNonErrorColor
         End If
         'kiem tra ngay ky
         If Trim(strNgayKy) = "" Or Trim(strNgayKy) = "../../...." Then
            .Row = countDynamicRow + 54
            .Col = .ColLetterToNumber(Col_R)
            .CellNote = errEmty  'static
            .BackColor = mErrorColor
            iFlagEmty = 1
         Else
                strNgayKy = DateSerial(CInt(Mid$(strNgayKy, 7, 4)), CInt(Mid$(strNgayKy, 4, 2)), CInt(Mid$(strNgayKy, 1, 2)))
                If strNgayKy > Now Then
                    .Col = .ColLetterToNumber(Col_R)
                    .Row = countDynamicRow + 54
                    .CellNote = errNgayKy  'static
                    .BackColor = mErrorColor
                    iFlagNgayKy = 1
                Else
                    .Row = countDynamicRow + 54
                    .Col = .ColLetterToNumber(Col_R)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
         End If

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
         
        ' CheckErrorMST
        '.ReDraw = True
        '.Visible = True
        .EventEnabled(EventAllEvents) = True
    End With
    
End Sub
Private Sub checkSheet3()
    Dim strCheck As String
    Dim errMauSoHD As Variant, errKyHieuHD As Variant, errSoLuong As Variant, errSoLienHD As Variant
    Dim errKhoangGN As Variant, errLengthHD As Variant, errEmty As Variant, errNgayKy As Variant, errNgayHuy As Variant, errTrungKhoang As Variant, errTon As Variant
    Dim strTenHD As Variant, strMauSoHD As Variant, strKyHieuHD As Variant, strTuSo As Variant, strDenSo As Variant
    Dim strTenHD1 As Variant, strMauSoHD1 As Variant, strKyHieuHD1 As Variant, strTuso1 As Variant, strDenso1 As Variant
    Dim strSoLuong As Variant, strNguoiKy As Variant, strNguoiLapBieu As Variant, strNgayKy As Variant, strNgayHuy As Variant
    Dim iFlagTenHD As Integer, iFlagMauSoHD As Integer, iFlagKyHieu As Integer, iFlagSoLienHD As Integer, iFlagEmty As Integer
    Dim iWhiteRow As Integer, iFlagLengHD As Integer, iFlagSoLuong As Integer, iFlagNgayKy As Integer, iFlagNgayHuy As Integer
    Dim iFlagKhoangTSDSTrung As Integer, iFocusFlag As Integer, iFlagTon As Integer
    
    Dim strId As Variant, strLoai As String, strTemp As String, result As String
    Dim i As Integer, countDynamicRow As Integer, j As Integer, k As Integer
    
    Dim str() As String
    Dim str1 As String
    ' tra ve mang cac row bi trung khoang
    str = checkTrung(fps, 45, Col_V, Col_Y, "B", "C", Col_O, Col_S, "", False, 3)
    
    str1 = checkTon(fps, 45, Col_V, Col_Y, "B", "C", Col_O, Col_S, "", False, 3)
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
        .GetText .ColLetterToNumber(Col_E), 45, errTon

        
        
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
        .Col = .ColLetterToNumber("B")
        .Row = 45
        'i = 1
        Do
            j = .Row
            countDynamicRow = countDynamicRow + 1
             ' Lay cac gia tri cua 1 row
             iWhiteRow = 1
            .GetText .ColLetterToNumber("C"), .Row, strTenHD
            .GetText .ColLetterToNumber(Col_O), .Row, strMauSoHD
            .GetText .ColLetterToNumber(Col_S), .Row, strKyHieuHD
            .GetText .ColLetterToNumber(Col_V), .Row, strTuSo
            .GetText .ColLetterToNumber(Col_Y), .Row, strDenSo
            .GetText .ColLetterToNumber(Col_AA), .Row, strSoLuong
            .GetText .ColLetterToNumber("AI"), .Row, strId
            
            If Trim(strTenHD) <> "" Or Trim(strMauSoHD) <> "" Or Trim(strKyHieuHD) <> "" Then
                ' Khong phai dong trang
                iWhiteRow = 0
            End If
            
            'Ten HD
            If Trim(strTenHD) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("C")
                .CellNote = errEmty  'static
                .BackColor = mErrorColor
                iFlagEmty = 1
            Else
                .Col = .ColLetterToNumber("C")
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
                    .Col = .ColLetterToNumber(Col_O)
                    .CellNote = errSoLienHD  'static
                    .BackColor = mErrorColor
                    iFlagSoLienHD = 1
                Else
                    .Col = .ColLetterToNumber(Col_O)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                
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
                result = CheckSoHD(CStr(strKyHieuHD))
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
                If Trim(strTuSo) = "" And Trim(strDenSo) = "" Then
                    ' tu so
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    ' den so
                    .Col = .ColLetterToNumber(Col_Y)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                ElseIf Trim(strTuSo) = "" And Trim(strDenSo) <> "" Then
                    ' tu so
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                    ' den so
                    .Col = .ColLetterToNumber(Col_Y)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                ElseIf Trim(strTuSo) <> "" And Trim(strDenSo) = "" Then
                    .Col = .ColLetterToNumber(Col_Y)
                    .CellNote = errEmty  'static
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                    ' tu so
                    .Col = .ColLetterToNumber(Col_V)
                    .CellNote = ""
                    .BackColor = mNonErrorColor

                Else
                    If Val(strDenSo) - Val(strTuSo) >= 0 And Val(strDenSo) > 0 Then
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
            If InStr(str1, j) = 0 And iWhiteRow = 0 Then
                'For i = 0 To UBound(str1)
                    'If j = str1(i) Then
                        .Col = .ColLetterToNumber(Col_V)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errTon     'static
                        .BackColor = mErrorColor
                        
                         .Col = .ColLetterToNumber(Col_Y)
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errTon     'static
                        .BackColor = mErrorColor
                        iFlagTon = 1
                        'Exit For
                    'End If
                'Next i
            End If
            
            ' end check trung
            .Row = j
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "AA"
                
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
         .SetText .ColLetterToNumber("B"), 45, IIf(iFlagTon = 1, "0", "1")
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
    For Each xmlNodeValid In TAX_Utilities_New.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "BC26AC"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
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
            For Each xmlCellNodeData In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               If (lCol = .ColLetterToNumber("B") Or lCol = .ColLetterToNumber(Col_G)) And lRow = 14 Then
                    IsUpdate = True
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
            
            .Sheet = 1 'Bang ke 2A
            
'            .Row = 17
'
'            sumLaoDong = 0
'            sumThuNhap = 0
'            sumLDKhauTru = 0
'            sumTNKhauTru = 0
'            sumSoThueTNDaKhauTru = 0
'
'            ' Lay cac thong tin sau khi da tong hop tren bang ke
'            .GetText .ColLetterToNumber(Col_J), .Row, sumLaoDong
'            .GetText .ColLetterToNumber(Col_F), .Row, sumThuNhap
'            .GetText .ColLetterToNumber(Col_K), .Row, sumLDKhauTru
'            .GetText .ColLetterToNumber(Col_L), .Row, sumTNKhauTru
'            .GetText .ColLetterToNumber(Col_I), .Row, sumSoThueTNDaKhauTru
                        
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
    strLietKe = getListHD(strCheck)
    For i = 0 To UBound(strLietKe)
        If Val(strLietKe(i)) < Val(Tuso1) Or Val(strLietKe(i)) > Val(denso1) Then
            checkHDXoa = -1
            Exit Function
        End If
    Next i
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
    Dim i As Integer
    Dim strTemp As String
    result = False
    If isDuLieuHL(str) Then
        arrstr = getListHD(str)
        For i = 0 To UBound(arrstr)
            If Not (InStr(1, strTemp, arrstr(i)) > 0) Then
                strTemp = strTemp & "~" & arrstr(i)
            End If
        Next i
        strTemp = Mid$(strTemp, 2)
        arrstr1 = Split(strTemp, "~")
        If UBound(arrstr) <> UBound(arrstr1) Then
            result = True
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
    If strXoa <> "" And strMat <> "" And strHuy = "" Then
        arrXoa = getListHD(strXoa)
        arrMat = getListHD(strMat)
        For i = 0 To UBound(arrXoa)
            strTemp = strTemp & "~" & arrXoa(i)
        Next i
        strTemp = Mid$(strTemp, 2)
        For i = 0 To UBound(arrMat)
            If InStr(1, strTemp, arrMat(i)) > 0 Then
                checkTrungXoa = "1"
                Exit Function
            End If
        Next i
    ' chi co xoa huy
    ElseIf strXoa <> "" And strMat = "" And strHuy <> "" Then
        arrXoa = getListHD(strXoa)
        arrHuy = getListHD(strHuy)
        For i = 0 To UBound(arrXoa)
            strTemp = strTemp & "~" & arrXoa(i)
        Next i
        strTemp = Mid$(strTemp, 2)
        For i = 0 To UBound(arrHuy)
            If InStr(1, strTemp, arrHuy(i)) > 0 Then
                checkTrungXoa = "2"
                Exit Function
            End If
        Next i
    ' chi co mat huy
    ElseIf strXoa = "" And strMat <> "" And strHuy <> "" Then
        arrHuy = getListHD(strHuy)
        arrMat = getListHD(strMat)
        For i = 0 To UBound(arrHuy)
            strTemp = strTemp & "~" & arrHuy(i)
        Next i
        strTemp = Mid$(strTemp, 2)
        For i = 0 To UBound(arrMat)
            If InStr(1, strTemp, arrMat(i)) > 0 Then
                checkTrungXoa = "3"
                Exit Function
            End If
        Next i
    ElseIf strXoa <> "" And strMat <> "" And strHuy <> "" Then
        arrHuy = getListHD(strHuy)
        arrMat = getListHD(strMat)
        arrXoa = getListHD(strXoa)
        ' kiem tra 3 loai trung
        For i = 0 To UBound(arrXoa)
            strTemp = strTemp & "~" & arrXoa(i)
        Next i
        strTemp = Mid$(strTemp, 2)
        ' kiem tra voi so hd xoa
        For i = 0 To UBound(arrMat)
            If InStr(1, strTemp, arrMat(i)) > 0 Then
                isMatXoa = True
                Exit For
            End If
        Next i
        For i = 0 To UBound(arrHuy)
            If InStr(1, strTemp, arrHuy(i)) > 0 Then
                isXoaHuy = True
                Exit For
            End If
        Next i
        ' kiem tra mat va huy
        For i = 0 To UBound(arrMat)
            strTemp1 = strTemp1 & "~" & arrMat(i)
        Next i
        strTemp1 = Mid$(strTemp1, 2)
        For i = 0 To UBound(arrHuy)
            If InStr(1, strTemp1, arrHuy(i)) > 0 Then
                isMatHuy = True
                Exit For
            End If
        Next i
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
Public Function getListHD(str As String) As String()
    Dim strTemp As String
    Dim strLietKe() As String
    Dim strKhoang() As String
    Dim strResult() As String
    Dim j As Double
    Dim flag1 As Boolean, flag2 As Boolean
    Dim i As Integer
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
    ' truong hop theo khoang hop le
    ElseIf flag1 = False And flag2 = True Then
        strKhoang = Split(str, "-")
        strTemp = ""
        For j = Val(strKhoang(0)) To Val(strKhoang(1))
            strTemp = strTemp & "~" & j
        Next j
        strTemp = Mid$(strTemp, 2)
        strResult = Split(strTemp, "~")
    ' truong hop nhap ca khoang va liet ke
    ElseIf flag1 = True And flag2 = True Then
        strLietKe = Split(str, ";")
        For i = 0 To UBound(strLietKe)
             If InStr(1, strLietKe(i), "-") > 0 Then
                strKhoang = Split(strLietKe(i), "-")
                For j = Val(strKhoang(0)) To Val(strKhoang(1))
                    strTemp = strTemp & "~" & j
                Next j
             Else
                strTemp = strTemp & "~" & strLietKe(i)
             End If
        Next i
        strTemp = Mid$(strTemp, 2)
        strResult = Split(strTemp, "~")
    Else
        ReDim strResult(0)
        strResult(0) = str
    End If
    getListHD = strResult
End Function

' kiem tra du lieu hop le
Public Function isDuLieuHL(str As String) As Boolean
    Dim flag1 As Boolean
    Dim flag2 As Boolean
    Dim result As Boolean
    Dim arrLietKe() As String
    Dim arrKhoang() As String
    Dim i As Integer, j As Integer
    result = True
    flag1 = False
    flag2 = False
    If InStr(1, str, ";") > 0 Then
        flag1 = True  ' co danh sach liet ke
    End If
    If InStr(1, str, "-") > 0 Then
        flag2 = True  ' co danh sach khoang
    End If
    
    If flag1 = True And flag2 = False Then
    ' truong hop chi nhap liet ke
        arrLietKe = Split(str, ";")
        For i = 0 To UBound(arrLietKe)
            If IsNullValue(arrLietKe(i)) Then
                 result = False
                 Exit For
            End If
        Next i
    ElseIf flag1 = False And flag2 = True Then
    ' truong hop chi nhap khoang
        arrKhoang = Split(str, "-")
        For i = 0 To UBound(arrKhoang)
            If IsNullValue(arrKhoang(i)) Then
                result = False
                Exit For
            End If
        Next i
    ElseIf flag1 = True And flag2 = True Then
    ' truong hop chi nhap liet ke va khoagn
        arrLietKe = Split(str, ";")
        For i = 0 To UBound(arrLietKe)
            If IsNullValue(arrLietKe(i)) Then
                result = False
                Exit For
            Else
                ' kiem tra xem la phan tu hay khoang
                If InStr(1, arrLietKe(i), "-") > 0 Then
                    arrKhoang = Split(arrLietKe(i), "-")
                    For j = 0 To UBound(arrKhoang)
                        If IsNullValue(arrKhoang(j)) Then
                            result = False
                            Exit For
                        End If
                    Next j
                End If
            End If
        Next i
    Else
        If str = "" Then
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
    If isDuLieuHL(str) Then
        arrLietKe = getListHD(str)
        result = UBound(arrLietKe) + 1
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
        MY = DateSerial(TAX_Utilities_New.Year, getThangDauKy(TAX_Utilities_New.ThreeMonths), 1)
        
        MY = DateAdd(Col_M, -3, MY)
        MY = DateSerial(Year(MY), getKyBC(Month(MY)), 1)
        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(GetValidityNode("68", Format(Month(MY), "0#"), , Format(Year(MY), "####")).childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        '***************************************
        If .Load(strDataFileName) = True Then
            Set xmlSectionsTemplate = .getElementsByTagName("Sections")(0)
            For lCtrl = 1 To xmlSectionsTemplate.childNodes.length - 2
                Set xmlSectionTemplate = xmlSectionsTemplate.childNodes(lCtrl)
                For i = 0 To xmlSectionTemplate.childNodes.length - 1
                    Set xmlNodeCells = xmlSectionTemplate.childNodes(i)
                    Set xmlNodeCell = xmlNodeCells.childNodes(20)  ' so luong
                    slTonCuoiKy = Val(GetAttribute(xmlNodeCell, "Value"))
                    strTempValue = ""
                    If slTonCuoiKy > 0 Then
                        Set xmlNodeCell = xmlNodeCells.childNodes(1)  ' Mau so
                        strTempValue = GetAttribute(xmlNodeCell, "Value")
                        Set xmlNodeCell = xmlNodeCells.childNodes(2)  ' Ky hieu
                        strTempValue = strTempValue & "~" & GetAttribute(xmlNodeCell, "Value")
                        Set xmlNodeCell = xmlNodeCells.childNodes(18)  ' so ton tu so
                        strTempValue = strTempValue & "~" & GetAttribute(xmlNodeCell, "Value")
                        Set xmlNodeCell = xmlNodeCells.childNodes(19)  ' so ton den so
                        strTempValue = strTempValue & "~" & GetAttribute(xmlNodeCell, "Value")
                        Set xmlNodeCell = xmlNodeCells.childNodes(21)  ' so strCheck
                        strTempValue = strTempValue & "~" & GetAttribute(xmlNodeCell, "Value")
                        Set xmlNodeCell = xmlNodeCells.childNodes(0)  ' so ten
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
    
    xmlDomHeader.Load GetAbsolutePath(TAX_Utilities_New.DataFolder & "Header_01.xml")
    GetNguoiKy = GetAttribute(xmlDomHeader.nodeFromID("C_26"), "Value")
    
    Set xmlDomHeader = Nothing
End Function
