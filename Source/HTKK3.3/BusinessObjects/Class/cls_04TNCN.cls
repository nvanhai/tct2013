VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_04TNCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "To khai quyet toan thu nhap ca nhan mau 04-TNCN" interface sheets
'this Class is belong to TAX_Business_v2 project which will be compline to DLL

Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    With fps
        SetDateFormat fps, 1, 69, .ColLetterToNumber("X"), DDMMYYYY
        SetDateFormat fps, 1, 74, .ColLetterToNumber("X"), DDMMYYYY
        
        SetDateFormat fps, 1, 81, .ColLetterToNumber("P"), DDMMYYYY
        .Sheet = 1 'To khai 01/TNCN
        .Row = 81
        .Col = .ColLetterToNumber("P")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
       .Sheet = 2 'PL 04-1/TNCN
        .ActiveSheet = .Sheet
        .Row = 1
        .Col = .ColLetterToNumber("B")
        .CellType = CellTypeButton
        .TypeButtonTextColor = RGB(0, 0, 255)
        .TypeButtonColor = &H8000000F
        .TypeButtonAlign = TypeButtonAlignRight
        .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")

    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
'    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
'    'PL 04-2-TNCN
'    Dim Row_PL_2_D5 As Long, Col_PL_2_D5 As Long, Row_PL_2_F7 As Long, Col_PL_2_F7 As Long
'    With fps
'        .Sheet = 3
'        mCurrentSheet = .Sheet
'        Set xmlCellNode = TAX_Utilities_v2.Data(3 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("D"), 5))
'        Set xmlCellsNode = xmlCellNode.parentNode
'
'        Do
'            ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(3, 5, .ColLetterToNumber("D")) - 1), "CellID"), Col_PL_2_D5, Row_PL_2_D5
'            ValidateDateError fps, 3, Row_PL_2_D5, Col_PL_2_D5, yyyy
'
'            ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(GetValidatedCellIndex(3, 7, .ColLetterToNumber("F")) - 1), "CellID"), Col_PL_2_F7, Row_PL_2_F7
'            ValidateDateError fps, 3, Row_PL_2_F7, Col_PL_2_F7, yyyy
'
'            Set xmlCellsNode = xmlCellsNode.nextSibling
'        Loop Until xmlCellsNode Is Nothing
'    End With
     
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

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    UpdateSheets
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
                    fps.EventEnabled(EventAllEvents) = False
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 69
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 69
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 4
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
            End If
            If .ActiveSheet = 2 Then
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
                   .Row = .Row + 4
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
            End If
            
        End If
  
      fps.EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    
    With fps
    mCurrentSheet = .Sheet
    .EventEnabled(EventAllEvents) = False
        If .ActiveSheet = 1 Then
            Dim count1 As Integer
                .Row = 69
                .Col = .ColLetterToNumber("B")
                Do While .value <> "bb"
                    count1 = count1 + 1
                    .Row = .Row + 1
                Loop
            
            If Col = .ColLetterToNumber("P") And Row = 81 + count1 - 6 Then
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
        
                If Col = .ColLetterToNumber("X") Then
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
    
    'CellChange Col, Row
End Sub

Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
    With fps
    
    
    End With
End Sub

'Calling this function whenever we don't need this class anymore.
'No parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        
        SetData
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
    Dim MST1 As Variant, MST2 As Variant, MST3 As Variant, MST4 As Variant
    Dim MST5 As Variant, MST6 As Variant, MST7 As Variant, MST8 As Variant
    Dim MST9 As Variant, MST10 As Variant, MST11 As Variant, MST12 As Variant, MST13 As Variant
    
    Dim strCheck As String
    Dim iCurrentSheet As Integer
    
    Dim blCheck_S1 As Boolean
    Dim blCheck_S2 As Boolean
    Dim v_S2 As Boolean
    
    Dim i As Integer, j As Integer
    Dim iWhiteRow As Integer
    Dim vMST As Variant, vErrMsgMST As Variant, vLastRow As Variant
    Dim vTenDC As Variant, vTNCT As Variant, vTPN As Variant
    Dim iFlagvMst As Integer, iFocusFlag As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String
    
    Dim sumTong, sumNguoiKoQuyetToanVN, sumNguoiKoQuyetToanNN, sumNguoiQuyetToanVN, sumNguoiQuyetToanNN As Long
    
    Dim startRow As Integer
    
    sumTong = 0
    sumNguoiKoQuyetToanVN = 0
    sumNguoiKoQuyetToanNN = 0
    sumNguoiQuyetToanVN = 0
    sumNguoiQuyetToanNN = 0
    
    With fps
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        CheckErrorMST
        
        
        
        .GetText .ColLetterToNumber("E"), 15, vErrMsgMST
        blCheck_S1 = True
        
        .Sheet = 1 'Check ma so thue tren to khai 04/TNCN muc (B/ Phan chi tiet)
        
        For i = 69 To .MaxRows
            iWhiteRow = 69
            .Row = i
            .Col = .ColLetterToNumber("H")
            .GetText .ColLetterToNumber("H"), .Row, vMST
            .GetText .ColLetterToNumber("B"), .Row, vLastRow
            If UCase(vLastRow) = "BB" Then Exit For ' Check cho den het muc B phan chi tiet la thoat luon
            If Trim(vMST) <> vbNullString Then
'                vMST = Replace(vMST, " ", "")
'                vMST = Replace(vMST, "-", "")
                 vMST = Left(vMST, 10)
               strCheck = ""
               If Len(vMST) = 10 Or Len(vMST) = 13 Then
                    strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
               End If
               If Mid(strCheck, 2, 1) <> "0" Then
                    .Col = .ColLetterToNumber("H")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMST   'static
                    .BackColor = mErrorColor
                    iFlagvMst = 1
                    blCheck_S1 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                .Col = .ColLetterToNumber("H")
                .CellNote = ""
                .BackColor = mNonErrorColor
                End If
            Else
                .Col = .ColLetterToNumber("H")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            
        Next
        
        ' Tinh tong so nguoi nop thue khong duoc quyet toan tai doanh nghiep (Nguoi Viet Nam) tren To khi 04/TNCN
        startRow = 69 ' Bat dau dem so ca nhan khong thuc hien quyet toan tai co quan chi tra
        For i = startRow To .MaxRows
            .GetText .ColLetterToNumber("B"), i, vLastRow
            .GetText .ColLetterToNumber("D"), i, vTenDC
            .GetText .ColLetterToNumber("L"), i, vTNCT   ' Thu nhap chiu thue
            If UCase(vLastRow) = "AA" Then Exit For ' Check cho den het muc B phan chi tiet tong so nguoi Viet nam ko quyet toan tai co quan chi tra
            ' Neu chua den dong cuoi cung kiem tra Ten, Thu nhap chiu thue neu co du lieu thi moi duoc tinh la mot nguoi
            If (vTenDC <> "") And (vTNCT <> "") Then
                sumNguoiKoQuyetToanVN = sumNguoiKoQuyetToanVN + 1
            End If
            ' Tang so dong len 1, de tinh cho truong hop sau
            startRow = startRow + 1
        Next
        ' Tinh tong so nguoi nop thue khong duoc quyet toan tai doanh nghiep (Nguoi Nuoc ngoai Nam) tren To khi 04/TNCN
        startRow = startRow + 4 ' Bat dau dem so ca nhan khong thuc hien quyet toan tai co quan chi tra
        For i = startRow To .MaxRows
            .GetText .ColLetterToNumber("B"), i, vLastRow
            .GetText .ColLetterToNumber("D"), i, vTenDC
            .GetText .ColLetterToNumber("L"), i, vTNCT   ' Thu nhap chiu thue
            If UCase(vLastRow) = "BB" Then Exit For ' Check cho den het muc B phan chi tiet tong so nguoi Nuoc ngoai ko quyet toan tai co quan chi tra
            
            ' Neu chua den dong cuoi cung kiem tra Ten, Thu nhap chiu thue neu co du lieu thi moi duoc tinh la mot nguoi
            If (vTenDC <> "") And (vTNCT <> "") Then
                sumNguoiKoQuyetToanNN = sumNguoiKoQuyetToanNN + 1
            End If
            ' Tang so dong len 1, de tinh cho truong hop sau
            startRow = startRow + 1
        Next
        
        
        ' Kiem tra PL 04-1/TNCN neu co
        .Sheet = 2
        v_S2 = IIf(TAX_Utilities_v2.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        blCheck_S2 = True
        If v_S2 = True Then
            For j = 7 To .MaxRows
                iWhiteRow = 7
                .Row = j
                .Col = .ColLetterToNumber("E")
                .GetText .ColLetterToNumber("E"), .Row, vMST
                .GetText .ColLetterToNumber("B"), .Row, vLastRow
                If UCase(vLastRow) = "BB" Then Exit For ' Check cho den het muc phu luc 04-1/TNCN
                If Trim(vMST) <> vbNullString Then
'                    vMST = Replace(vMST, " ", "")
'                    vMST = Replace(vMST, "-", "")
                     vMST = Left(vMST, 10)
                   strCheck = ""
                   If Len(vMST) = 10 Or Len(vMST) = 13 Then
                        strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                   End If
                   If Mid(strCheck, 2, 1) <> "0" Then
                        .Col = .ColLetterToNumber("E")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgMST   'static
                        .BackColor = mErrorColor
                        iFlagvMst = 1
                        blCheck_S2 = False
                        'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
                    Else
                        .Col = .ColLetterToNumber("E")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                Else
                    .Col = .ColLetterToNumber("E")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            Next
        
            ' Tinh tong so nguoi nop thue khong duoc quyet toan tai doanh nghiep (Nguoi Viet Nam) tren To khi 04/TNCN
            startRow = 9 ' Bat dau dem so ca nhan khong thuc hien quyet toan tai co quan chi tra
            For i = startRow To .MaxRows
                .GetText .ColLetterToNumber("B"), i, vLastRow
                .GetText .ColLetterToNumber("C"), i, vTenDC
                .GetText .ColLetterToNumber("G"), i, vTNCT   ' Thu nhap chiu thue
                .GetText .ColLetterToNumber("M"), i, vTPN    ' Thue phai nop
                
                If UCase(vLastRow) = "AA" Then Exit For ' Check cho den het muc B phan chi tiet tong so nguoi Viet nam ko quyet toan tai co quan chi tra
                ' Neu chua den dong cuoi cung kiem tra Ten, Thu nhap chiu thue neu co du lieu thi moi duoc tinh la mot nguoi
                If (vTenDC <> "") And (vTNCT <> "") And (vTPN <> "") Then
                    sumNguoiQuyetToanVN = sumNguoiQuyetToanVN + 1
                End If
                ' Tang so dong len 1, de tinh cho truong hop sau
                startRow = startRow + 1
            Next
            ' Tinh tong so nguoi nop thue khong duoc quyet toan tai doanh nghiep (Nguoi Nuoc ngoai Nam) tren To khi 04/TNCN
            startRow = startRow + 4 ' Bat dau dem so ca nhan khong thuc hien quyet toan tai co quan chi tra
            For i = startRow To .MaxRows
                .GetText .ColLetterToNumber("B"), i, vLastRow
                .GetText .ColLetterToNumber("C"), i, vTenDC
                .GetText .ColLetterToNumber("G"), i, vTNCT   ' Thu nhap chiu thue
                .GetText .ColLetterToNumber("M"), i, vTPN    ' Thue phai nop
                
                If UCase(vLastRow) = "BB" Then Exit For ' Check cho den het muc B phan chi tiet tong so nguoi Nuoc ngoai ko quyet toan tai co quan chi tra
                
                ' Neu chua den dong cuoi cung kiem tra Ten, Thu nhap chiu thue neu co du lieu thi moi duoc tinh la mot nguoi
                If (vTenDC <> "") And (vTNCT <> "") Then
                    sumNguoiQuyetToanNN = sumNguoiQuyetToanNN + 1
                End If
                ' Tang so dong len 1, de tinh cho truong hop sau
                startRow = startRow + 1
            Next
        
        End If
        
        ' Trong truong hop ca blCheck_S2 va blCheck_S2 khong co MST nao bi loi trong phu luc thi set lai gia tri cua Header la 1
        If (blCheck_S1 = True) And (blCheck_S2 = True) Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "0"
        End If
        
        ' Tong hop so lieu tu muc (B/ Phan chi tiet) va Phu luc 04-1/TNCN (neu co)
        .Sheet = 1
        mCurrentSheet = .Sheet
        .EventEnabled(EventAllEvents) = False
        ' Tong hop so lao dong Viet nam trong nam len to khai
        .SetText .ColLetterToNumber("M"), 10, sumNguoiKoQuyetToanVN + sumNguoiQuyetToanVN
        UpdateCell fps, .ColLetterToNumber("M"), 10, sumNguoiKoQuyetToanVN + sumNguoiQuyetToanVN
        ' Tong hop so lao dong Nuoc ngoai trong nam len to khai
        .SetText .ColLetterToNumber("M"), 11, sumNguoiKoQuyetToanNN + sumNguoiQuyetToanNN
        UpdateCell fps, .ColLetterToNumber("M"), 11, sumNguoiKoQuyetToanNN + sumNguoiQuyetToanNN
        ' Tong hop so lao dong Viet nam trong nam duoc co quan chi tra len to khai
        .SetText .ColLetterToNumber("M"), 13, sumNguoiQuyetToanVN
         UpdateCell fps, .ColLetterToNumber("M"), 13, sumNguoiQuyetToanVN
        ' Tong hop so lao dong Nuoc ngoai trong nam duoc co quan chi tra len to khai
        .SetText .ColLetterToNumber("M"), 14, sumNguoiQuyetToanNN
         UpdateCell fps, .ColLetterToNumber("M"), 14, sumNguoiQuyetToanNN
        ' Tong hop so lao dong Viet nam trong nam khong duoc co quan chi tra len to khai
        .SetText .ColLetterToNumber("M"), 26, sumNguoiKoQuyetToanVN
        UpdateCell fps, .ColLetterToNumber("M"), 26, sumNguoiKoQuyetToanVN
        ' Tong hop so lao dong Nuoc ngoai trong nam khong duoc co quan chi tra len to khai
        .SetText .ColLetterToNumber("M"), 27, sumNguoiKoQuyetToanNN
        UpdateCell fps, .ColLetterToNumber("M"), 27, sumNguoiKoQuyetToanNN
        .EventEnabled(EventAllEvents) = True
        .Sheet = iCurrentSheet
        mCurrentSheet = iCurrentSheet
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
Dim isheet As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
         isheet = mCurrentSheet
         mCurrentSheet = 1
        .Sheet = 1
        
        .Col = fps.ColLetterToNumber("P")
        .Row = 15
        .Lock = True
        UpdateCell fps, .Col, .Row, .value
        
        .Col = fps.ColLetterToNumber("P")
        .Row = 16
        .Lock = True
        UpdateCell fps, .Col, .Row, .value
        
        .Col = fps.ColLetterToNumber("P")
        .Row = 17
        .Lock = True
        UpdateCell fps, .Col, .Row, .value
        
        .Col = fps.ColLetterToNumber("P")
        .Row = 17
        .Lock = True
        UpdateCell fps, .Col, .Row, .value
        
        .Col = fps.ColLetterToNumber("P")
        .Row = 18
        .Lock = True
        UpdateCell fps, .Col, .Row, .value
        
        .Col = fps.ColLetterToNumber("P")
        .Row = 19
        .Lock = True
        UpdateCell fps, .Col, .Row, .value
        
        .Col = fps.ColLetterToNumber("P")
        .Row = 20
        .Lock = True
        UpdateCell fps, .Col, .Row, .value
        
        .Col = fps.ColLetterToNumber("P")
        .Row = 21
        .Lock = True
        UpdateCell fps, .Col, .Row, .value
        
        .Col = fps.ColLetterToNumber("P")
        .Row = 22
        .Lock = True
        UpdateCell fps, .Col, .Row, .value
        
        .Col = fps.ColLetterToNumber("P")
        .Row = 23
        .Lock = True
        UpdateCell fps, .Col, .Row, .value
        .EventEnabled(EventAllEvents) = True
        mCurrentSheet = isheet
        .ActiveSheet = mCurrentSheet
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
    ssSheet = mCurrentSheet
     .Sheet = 1
     mCurrentSheet = .Sheet
         For Each xmlCellNodeData In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
            ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
            .GetText lCol, lRow, varTemp
            UpdateCell fps, lCol, lRow, varTemp
          Next
      .Sheet = .ActiveSheet
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub



Public Sub FinishImport()

End Sub

