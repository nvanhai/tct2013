VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01NTNN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrTKThang_PS As String
Public loaiKHBS As String
Public StrSolanBosung As String
<<<<<<< .mine


=======

Public loaiKHBS As String

>>>>>>> .r183
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    With fps
        SetDateFormat fps, 1, 30, .ColLetterToNumber("AG"), DDMMYYYY
        SetDateFormat fps, 1, 55, .ColLetterToNumber("AD"), DDMMYYYY
        
        .Sheet = 1 'To khai 02/TNCN
        .Row = 30
        .Col = .ColLetterToNumber("AG")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
Dim ngay As Variant
Dim kythang As Variant
Dim kyngay As Variant

    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        ' vttoan: check to khai thang hoac lan phat sinh
        ngay = TAX_Utilities_New.Day
        .SetText .ColLetterToNumber("O"), 32, ngay
        .GetText .ColLetterToNumber("W"), 31, kythang
        .GetText .ColLetterToNumber("AS"), 31, kyngay
        
        If StrTKThang_PS = "TK_THANG" Then
            .Col = .ColLetterToNumber("AT")
            .Row = 5
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .SetText .ColLetterToNumber("P"), 5, kythang
            .Col = .ColLetterToNumber("BC")
            .Row = 5
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf StrTKThang_PS = "TK_LANPS" Then
            .Col = .ColLetterToNumber("BC")
            .Row = 5
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .SetText .ColLetterToNumber("P"), 5, kyngay
            .Col = .ColLetterToNumber("AT")
            .Row = 5
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        End If
      
        'end
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("C")
            .Row = 31
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 31
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 31
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("C")
            .Row = 31
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 31
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 31
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        .EventEnabled(EventAllEvents) = True
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
                   fps.EventEnabled(EventAllEvents) = True
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 55
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 55
                         i = i + 1
                    Loop Until .Text = "aa"
                    .SetActiveCell iCol, iRow

                    .Row = iRow
                    .Col = iCol
                     fps.EventEnabled(EventAllEvents) = False
                     .SetActiveCell iCol, iRow
            End If
         End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
        'check date
            If Col = .ColLetterToNumber("AG") And Row = 30 Then
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
            If Col = .ColLetterToNumber("AD") And Row = 55 Then
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
    Dim MST1 As Variant, MST2 As Variant, MST3 As Variant, MST4 As Variant
    Dim MST5 As Variant, MST6 As Variant, MST7 As Variant, MST8 As Variant
    Dim MST9 As Variant, MST10 As Variant, MST11 As Variant, MST12 As Variant, MST13 As Variant
    Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    Dim strCheck As String, iCurrentSheet As Integer
    Dim mstDN, strNoiDung As Variant
    Dim iFlagEmty, iFlagEmtyND As Variant
    Dim errMst As Variant
    Dim errEmty As Variant
    Dim errCt08 As Variant
    Dim ct07 As Variant
    Dim ct08 As Variant
    Dim ct11 As Variant
    Dim i As Variant
    Dim iWhiteRow As Boolean
    Dim strMst, strNgayHD, strSoTienTT, strNgayTT, strDoanhThu, strTyLeGT, strThueSuat, strDoanhThuTT, strTyLeTNDN, strSoThueMG As Variant
    With fps
    'vttoan
    ' bat cau truc mst
    .Sheet = .SheetCount
    .GetText .ColLetterToNumber("E"), 12, errMst
    .GetText .ColLetterToNumber("E"), 13, errEmty
    .GetText .ColLetterToNumber("E"), 14, errCt08
    .Sheet = 1
    .Col = .ColLetterToNumber("B")
    .Row = 55
    i = 1
    Do
        .GetText .ColLetterToNumber("C"), .Row, strNoiDung
        
        .GetText .ColLetterToNumber("R"), .Row, strNgayHD
        .GetText .ColLetterToNumber("X"), .Row, strSoTienTT
        .GetText .ColLetterToNumber("AD"), .Row, strNgayTT
        .GetText .ColLetterToNumber("AI"), .Row, strDoanhThu
        .GetText .ColLetterToNumber("AM"), .Row, strTyLeGT
        .GetText .ColLetterToNumber("AQ"), .Row, strThueSuat
        .GetText .ColLetterToNumber("AY"), .Row, strDoanhThuTT
        .GetText .ColLetterToNumber("BC"), .Row, strTyLeTNDN
        .GetText .ColLetterToNumber("BG"), .Row, strSoThueMG
        .GetText .ColLetterToNumber("L"), .Row, mstDN
        If Trim(mstDN) <> "" Or Trim(strNoiDung) <> "" Or Trim(strNgayHD) <> "" Or Val(strSoTienTT) <> 0 Or Trim(strNgayTT) <> "" _
        Or Val(strDoanhThu) <> 0 Or Val(strTyLeGT) <> 0 Or Val(strThueSuat) <> 0 Or Val(strDoanhThuTT) <> 0 Or Val(strTyLeTNDN) <> 0 Or Val(strSoThueMG) <> 0 Then
            iWhiteRow = False
        Else
            iWhiteRow = True
        End If
        
        
        If (strNoiDung = vbNullString Or Trim(strNoiDung) = "") And iWhiteRow = False Then
            .Col = .ColLetterToNumber("C")
            .CellNote = errEmty
            .BackColor = mErrorColor
            iFlagEmtyND = 1
        Else
            .Col = .ColLetterToNumber("C")
            .CellNote = ""
            .BackColor = mNonErrorColor
            iFlagEmtyND = 0
        End If
        
        
        
        If Trim(mstDN) <> vbNullString Then
            strCheck = ""
            If Len(mstDN) = 10 Or Len(mstDN) = 13 Then
                     strCheck = CheckTaxCode(Mid(mstDN, 1, 1), Mid(mstDN, 2, 1), Mid(mstDN, 3, 1), Mid(mstDN, 4, 1), Mid(mstDN, 5, 1), Mid(mstDN, 6, 1), Mid(mstDN, 7, 1), Mid(mstDN, 8, 1), Mid(mstDN, 9, 1), Mid(mstDN, 10, 1), Mid(mstDN, 11, 1), Mid(mstDN, 12, 1), Mid(mstDN, 13, 1))
                        If Len(mstDN) = 10 And strCheck <> "000" Then
                            .Col = .ColLetterToNumber("L")
                            .CellNote = errMst
                            .BackColor = mErrorColor
                            iFlagEmty = 1
                        ElseIf Len(mstDN) = 13 And strCheck <> "000" Then
                            .Col = .ColLetterToNumber("L")
                            .CellNote = errMst
                            .BackColor = mErrorColor
                            iFlagEmty = 1
                        Else
                            .Col = .ColLetterToNumber("L")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                   ElseIf Len(mstDN) <> 10 Or Len(mstDN) <> 13 Then
                        .Col = .ColLetterToNumber("L")
                        .CellNote = errMst
                        .BackColor = mErrorColor
                        iFlagEmty = 1
                   Else
                        .Col = .ColLetterToNumber("L")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                   End If
            Else
                .Col = .ColLetterToNumber("L")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
         ' bat dinh dang chi tieu [07],[08],[11]
         Dim ct09 As Variant
         Dim ct10 As Variant
         Dim ct12 As Variant
         Dim ct13 As Variant
         Dim ct14 As Variant
         Dim ct06 As Variant
          .GetText .ColLetterToNumber("AI"), .Row, ct06
'         .GetText .ColLetterToNumber("AU"), .Row, ct09
         .GetText .ColLetterToNumber("AY"), .Row, ct10
         .GetText .ColLetterToNumber("BG"), .Row, ct12
'         .GetText .ColLetterToNumber("BM"), .Row, ct13
         .GetText .ColLetterToNumber("AM"), .Row, ct07
         .GetText .ColLetterToNumber("AQ"), .Row, ct08
         .GetText .ColLetterToNumber("BC"), .Row, ct11
         
'            ' chi tieu 07
'            If Trim(ct07) <> "" Then
'                If Len(Trim(ct07)) <> 6 Or Mid(Trim(ct07), 3, 1) <> "," Or Mid(Trim(ct07), 6, 1) <> "%" Then
'                    .Col = .ColLetterToNumber("AM")
'                    .CellNote = errCt07
'                    .BackColor = mErrorColor
'                    iFlagEmty = 1
'                Else
'                    .Col = .ColLetterToNumber("AM")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                End If
'            Else
'                .Col = .ColLetterToNumber("AM")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'
'            End If
'            ' chi tieu 11
'            If Trim(ct11) <> "" Then
'                If Len(Trim(ct11)) <> 6 Or Mid(Trim(ct11), 3, 1) <> "," Or Mid(Trim(ct11), 6, 1) <> "%" Then
'                    .Col = .ColLetterToNumber("BC")
'                    .CellNote = errCt07
'                    .BackColor = mErrorColor
'                    iFlagEmty = 1
'                Else
'                    .Col = .ColLetterToNumber("BC")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                End If
'            Else
'                .Col = .ColLetterToNumber("BC")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'            End If
'            'chi tieu 08
'            If Trim(ct08) <> "" Then
'                If Len(Trim(ct08)) <> 3 Or Mid(Trim(ct08), 3, 1) <> "%" Then
'                    .Col = .ColLetterToNumber("AQ")
'                    .CellNote = errCt08
'                    .BackColor = mErrorColor
'                    iFlagEmty = 1
'                Else
'                    .Col = .ColLetterToNumber("AQ")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                End If
'            Else
'                .Col = .ColLetterToNumber("AQ")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'            End If
            If Trim(ct06) <> 0 And Trim(ct07) <> "" And Trim(ct08) <> "" Then
                ct09 = Round(Mid(ct07, 1, 5) / 100 * Mid(ct08, 1, 2) / 100 * Val(ct06), 0)
            Else
                ct09 = 0
            End If
            .SetText .ColLetterToNumber("AU"), .Row, ct09
            If Trim(ct10) <> 0 And Trim(ct11) <> "" Then
                ct13 = Round(Val(ct10) * Mid(ct11, 1, 5) / 100 - Val(ct12), 0)
            Else
                ct13 = 0
            End If
            .SetText .ColLetterToNumber("BM"), .Row, ct13
            ct14 = ct09 + ct13
            .SetText .ColLetterToNumber("BQ"), .Row, ct14
            ' end
            
    .Row = 55 + i
    i = i + 1
    .Col = .ColLetterToNumber("B")
    Loop Until .Text = "aa"
    
    .Sheet = .SheetCount
    .SetText .ColLetterToNumber("B"), 12, IIf(iFlagEmty = 1, "0", "1")
    .SetText .ColLetterToNumber("B"), 13, IIf(iFlagEmtyND = 1, "0", "1")
    
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

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

Public Sub SetActiveSheet()

End Sub

Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean

End Function

Public Sub SetData()

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
        If (GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "Active") <> "0") Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
                'If Not (lRow = 5 And (lCol = .ColLetterToNumber("AT") Or lCol = .ColLetterToNumber("BC"))) Then
                    ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
                'End If
             Next
        End If
      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing

End Sub

