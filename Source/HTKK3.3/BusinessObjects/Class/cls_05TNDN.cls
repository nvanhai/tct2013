VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_05TNDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Khau tru tai nguon doi voi Hoa Hong Dai Ly" interface sheets
'this Class is belong to TAX_Business_v2 project which will be compline to DLL

Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Private ngayKyRow As Integer
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    With fps
   
        SetDateFormat fps, 1, 11, .ColLetterToNumber("F"), DDMMYYYY
        SetDateFormat fps, 1, 16, .ColLetterToNumber("I"), DDMMYYYY
        .Sheet = 1 'To khai
        .Row = 16
        .Col = .ColLetterToNumber("I")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    With fps
            .Sheet = 1
            .ActiveSheet = .Sheet
            .Row = 2
            .Col = .ColLetterToNumber("C")
            .CellType = CellTypeButton
            .TypeButtonTextColor = RGB(0, 0, 255)
            .TypeButtonColor = &H8000000F
            .TypeButtonAlign = TypeButtonAlignRight
            .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
    End With
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    With fps

    End With
End Sub

Private Sub fps_GotFocus()
    With fps
    
    End With
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Integer, iCol As Long, iRow As Long
    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 1 Then
                   fps.EventEnabled(EventAllEvents) = False
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 11
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 11
                         i = i + 1
                    Loop Until .Text = "aa"
                    .SetActiveCell iCol, iRow

                    .Row = iRow
                    .Col = iCol
                    fps.EventEnabled(EventAllEvents) = True
                    .SetActiveCell iCol, iRow
            End If
         End If
        .EventEnabled(EventAllEvents) = True
    End With
    
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim CellNguoiKy As String
    
    CellNguoiKy = GetAttribute(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length - 1), "CellID")
    With fps
        mCurrentSheet = .Sheet
        .EventEnabled(EventAllEvents) = False
      
        If .ActiveSheet = 1 Then
                If Col = .ColLetterToNumber("F") Then
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
                    UpdateCell fps, .Col, .Row, .Text
                End If
                If Col = .ColLetterToNumber("I") And Row = Right(CellNguoiKy, 2) Then
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
        End If
    .EventEnabled(EventAllEvents) = True
    End With
    
End Sub

Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
   With fps

   End With
End Sub

'Calling this function whenever we don't need this class anymore.
'No parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
    Dim vThueTNDNNop As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        
        CheckDynamicError 'Set Exception Error on cells of interface
            
        .Sheet = 1
        .GetText .ColLetterToNumber("I"), 7, vThueTNDNNop
        UpdateCell fps, .ColLetterToNumber("C"), 7, Convert(cReadNum(vThueTNDNNop), TCVN, UNICODE)
        
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
    Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    Dim strCheck As String, iCurrentSheet As Integer
    
    Dim vMST As Variant, vErrMsgMST As Variant, vLastRow As Variant, vValueKhautru As Variant
    Dim vNgayHD, vErrMsgNgayHD As Variant, vErrMsgkhautru As Variant
    Dim iFlagvMst As Integer, iFocusFlag As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String
    Dim i As Integer
    Dim blCheck_S1 As Boolean
    With fps
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        CheckErrorMST
        
        ' Lay thong bao ve sai ma so thue trong sheet header
        .GetText .ColLetterToNumber("E"), 15, vErrMsgMST
        ' Lay thong bao sai ve ngay thanh toan trong sheet header
        .GetText .ColLetterToNumber("E"), 17, vErrMsgNgayHD
        ' Lay thong bao sai ve ngay thanh toan trong sheet header
        .GetText .ColLetterToNumber("E"), 18, vErrMsgkhautru
        
        .Sheet = 1 'Check tinh xac thuc cua cac thong tin tren to khai 05/TNDN
        blCheck_S1 = True
        For i = 11 To .MaxRows
            .Row = i
            .Col = .ColLetterToNumber("D")
            .GetText .ColLetterToNumber("D"), .Row, vMST
            .GetText .ColLetterToNumber("B"), .Row, vLastRow
            If UCase(vLastRow) = "AA" Then Exit For ' Check cho den het cac dong co du lieu thi thoi
            'Check ma so thue tren to khai 05/TNDN
            If Trim(vMST) <> vbNullString Then
'               vMST = Replace(vMST, " ", "")
'               vMST = Replace(vMST, "-", "")
                vMST = Left(vMST, 10)
               strCheck = ""
               If Len(vMST) = 10 Or Len(vMST) = 13 Then
                    strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
               End If
               If Mid(strCheck, 2, 1) <> "0" Then
                    .Col = .ColLetterToNumber("D")
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
                     .Col = .ColLetterToNumber("D")
                     .CellNote = ""
                     .BackColor = mNonErrorColor
                End If
            Else
                .Col = .ColLetterToNumber("D")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
        Next
        ' Trong truong hop blCheck_S1 khong co MST nao bi loi thi set lai gia tri cua Header la 1
        If blCheck_S1 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "0"
        End If
        
        
        'adasdsadsad
        .Sheet = 1
        blCheck_S1 = True
        For i = 11 To .MaxRows
            .Row = i
            .Col = .ColLetterToNumber("F")
            .GetText .ColLetterToNumber("F"), .Row, vNgayHD
            .GetText .ColLetterToNumber("B"), .Row, vLastRow
            If UCase(vLastRow) = "AA" Then Exit For ' Check cho den het cac dong co du lieu thi thoi
            'Check ngay thanh toan tren to khai 05/TNDN
            If Trim(vNgayHD) <> "" Then
                 Dim strThang As Integer
                 Dim strNam As Integer
                 
                 strThang = TAX_Utilities_v2.Month
                 strNam = TAX_Utilities_v2.Year
                 
                 If CInt(Right(vNgayHD, 4)) > strNam Then
                     .Col = .ColLetterToNumber("F")
                     .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                     .CellNote = vErrMsgNgayHD 'static
                     .BackColor = mErrorColor
                      blCheck_S1 = False
                 ElseIf CInt(Right(vNgayHD, 4)) >= strNam And CInt(Right("00" & Mid(vNgayHD, 4, 2), 2)) > strThang Then
                      .Col = .ColLetterToNumber("F")
                     .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                     .CellNote = vErrMsgNgayHD 'static
                     .BackColor = mErrorColor
                      blCheck_S1 = False
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
        Next
        ' Trong truong hop blCheck_S1 khong co MST nao bi loi thi set lai gia tri cua Header la 1
        If blCheck_S1 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 17, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 17, "0"
        End If
        'asdasdsada
        
        .Sheet = 1
        blCheck_S1 = True
        For i = 11 To .MaxRows
            .Row = i
            .Col = .ColLetterToNumber("H")
            .GetText .ColLetterToNumber("H"), .Row, vValueKhautru
            .GetText .ColLetterToNumber("B"), .Row, vLastRow
            If UCase(vLastRow) = "AA" Then Exit For ' Check cho den het cac dong co du lieu thi thoi
            'Check ty le khau tru khong duoc lon hon 100
            If Trim(vValueKhautru) <> "" Then
                 If CDbl(vValueKhautru) > 100 Then
                     .Col = .ColLetterToNumber("H")
                     .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                     .CellNote = vErrMsgkhautru  'static
                     .BackColor = mErrorColor
                      blCheck_S1 = False
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
        
        If blCheck_S1 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 18, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 18, "0"
        End If
        'asdasdsada
        
        
        .Sheet = iCurrentSheet
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

End Sub
Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean

End Function

Public Sub SetData()

End Sub

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function
Public Sub FinishImport()

End Sub

