VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_1_02TNCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
'Private Const S1Txt1Row = 9
'Private Const S1Txt1Col = "I"
'Private Const S1Txt1aRow = 10
'Private Const S1Txt1aCol = "I"
'Private Const S1Txt1bRow = 11
'Private Const S1Txt1bCol = "I"
'Private Const S1Txt1cRow = 12
'Private Const S1Txt1cCol = "I"
'Dim vt As String, va As String, vb As String, vc As String
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    With fps
        SetDateFormat fps, 1, 42, .ColLetterToNumber("N"), DDMMYYYY
        .Sheet = 1 'To khai 02/TNCN
        .Row = 42
        .Col = .ColLetterToNumber("N")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        SetDateFormat fps, 1, 31, .ColLetterToNumber("O"), DDMMYYYY
        SetDateFormat fps, 1, 35, .ColLetterToNumber("O"), DDMMYYYY
    End With
    
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim a2 As Variant, b2 As Variant, c2 As Variant
    Dim count1, count2 As Long
    With fps
'tinh so dong` them vao`
            mCurrentSheet = .ActiveSheet
            .Sheet = 1
            .Row = 31
            .Col = .ColLetterToNumber("B")
            Do While .Text <> "aa"
                count1 = count1 + 1
                .Row = .Row + 1
            Loop
            
            .Row = 35 + count1 - 1
            .Col = .ColLetterToNumber("B")
            Do While .Value <> "aa"
                count2 = count2 + 1
                .Row = .Row + 1
            Loop
            
            .Sheet = 1
            .Row = 40 + count1 + count2
            .Col = .ColLetterToNumber("N")
            If .Value = "" Then
                SetDateFormat fps, 1, 40 + count1 + count2, .ColLetterToNumber("N"), DDMMYYYY
                '.Value = "" + Conversion.CStr(Day(Now)) + Conversion.CStr(Month(Now)) + Conversion.CStr(Year(Now))
                .Text = Format(Date, "dd/mm/yyyy")
                UpdateCell fps, .Col, .Row, .Text
            End If
    End With
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
   Dim i As Integer
    Dim t As Variant, a As Variant, b As Variant, c As Variant
  
    Dim ArrRow As Variant, ArrCol As Variant
    Dim S As String
    
With fps


End With
End Sub

Private Sub fps_GotFocus()
    With fps
'        .Col = .ColLetterToNumber(S1Txt1Col)
'        .Row = S1Txt1Row
'        vt = GetAttribute(TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fps, .Col, .Row)), "Value")
'        .Row = S1Txt1aRow
'        va = GetAttribute(TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fps, .Col, .Row)), "Value")
'        .Row = S1Txt1bRow
'        vb = GetAttribute(TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fps, .Col, .Row)), "Value")
'        .Row = S1Txt1cRow
'        vc = GetAttribute(TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fps, .Col, .Row)), "Value")
    End With
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i, j As Integer, iCol As Long, iRow As Long
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
            If .ActiveSheet = 1 Then
                   fps.EventEnabled(EventAllEvents) = True
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 31
                    Do
                         .Text = str(i)
                         UpdateCell fps, .Col, .Row, .Text
                         .Row = i + 31
                         i = i + 1
                    Loop Until .Text = "aa"
                    .SetActiveCell iCol, iRow
        '****************************************
                    j = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 35 + i - 2
                    Do
                         .Text = str(j)
                         UpdateCell fps, .Col, .Row, .Text
                         .Row = j + 35 + i - 2
                         j = j + 1
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
    Dim varCT11 As Variant
    Dim varCT12 As Variant
    Dim count1, count2, sum61, sum62, sum71, sum72 As Long
    With fps
    .EventEnabled(EventAllEvents) = False
    mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
'tinh so dong` them vao`
            .Row = 31
            .Col = .ColLetterToNumber("B")
            Do While .Value <> "aa"
                count1 = count1 + 1
                .Row = .Row + 1
            Loop
            
            .Row = 35 + count1 - 1
            .Col = .ColLetterToNumber("B")
            Do While .Value <> "aa"
                count2 = count2 + 1
                .Row = .Row + 1
            Loop
           
'check date
            If Col = .ColLetterToNumber("N") And Row = 40 + count1 + count2 Then
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
            
             If Col = .ColLetterToNumber("O") Then
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
'check num >=0
            If Row < 22 And Row > 10 And (Col = 11 Or Col = 13) Then
            .Row = Row
            .Col = Col
                If .Value < 0 Then
                     DisplayMessage "0018", msOKOnly, miCriticalError
                    .SetActiveCell Col, Row
                End If
            End If
        End If
    .EventEnabled(EventAllEvents) = True
    End With
    'CellChange Col, Row
End Sub
Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
   With fps
'        .Col = .ColLetterToNumber(S1Txt1Col)
'        .Row = S1Txt1Row
'        vt = GetAttribute(TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fps, .Col, .Row)), "Value")
'        .Row = S1Txt1aRow
'        va = GetAttribute(TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fps, .Col, .Row)), "Value")
'        .Row = S1Txt1bRow
'        vb = GetAttribute(TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fps, .Col, .Row)), "Value")
'        .Row = S1Txt1cRow
'        vc = GetAttribute(TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fps, .Col, .Row)), "Value")
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
    Dim count, count1, count2 As Long
    Dim errorName, errorSct, errorDate, errorMst, error1, error2, vMST As Variant
    Dim colC, ColN, colO, colK As Variant
    
    Dim iFlagvName As Integer, iFlagvSct As Integer, iFlagvDate As Integer, iFlagvMst As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer
    Dim strFormula As String
    iFlagvName = 0
    iFlagvSct = 0
    iFlagvDate = 0
    iFlagvMst = 0
    iFocusFlag = 0
    Dim iWhiteRow As Integer
    With fps

    CheckErrorMST
    .Sheet = .SheetCount
    .GetText .ColLetterToNumber("E"), 15, errorName
    .GetText .ColLetterToNumber("E"), 17, errorSct
    .GetText .ColLetterToNumber("E"), 18, errorDate
    .GetText .ColLetterToNumber("E"), 19, errorMst
    .GetText .ColLetterToNumber("E"), 23, error1
    .GetText .ColLetterToNumber("E"), 24, error2
    
    .Sheet = 1
'check Null
    
    .Row = 31
    Do
        .GetText .ColLetterToNumber("C"), .Row, colC
        .GetText .ColLetterToNumber("N"), .Row, ColN
        .GetText .ColLetterToNumber("O"), .Row, colO
        .GetText .ColLetterToNumber("K"), .Row, colK
        
        If colC <> "" Or ColN <> "" Or colO <> "" Or colK <> 0 Then
        iWhiteRow = 0
        count = count + 1
'check name
            
            If colC = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("C")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = errorName  'static
                .BackColor = mErrorColor
                iFlagvName = 1
                'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
            Else
                .Col = .ColLetterToNumber("C")
                .BackColor = mNonErrorColor
                .CellNote = ""
         
            End If
     'so chung tu
          
            If ColN = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("N")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = errorSct  'static
                .BackColor = mErrorColor
                iFlagvSct = 1
                'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
            Else
                .Col = .ColLetterToNumber("N")
                .BackColor = mNonErrorColor
                .CellNote = ""
            End If
    'date
             If colO = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("O")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = errorDate  'static
                .BackColor = mErrorColor
                iFlagvDate = 1
                'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
            Else
                .Col = .ColLetterToNumber("O")
                .BackColor = mNonErrorColor
                .CellNote = ""
            End If
    ' Check MST
            .Col = .ColLetterToNumber("G")
            vMST = .Value
'            vMST = Replace(vMST, " ", "")
'            vMST = Replace(vMST, "-", "")
             vMST = Left(vMST, 10)
            strCheck = ""
            If (Len(vMST) = 10 Or Len(vMST) = 13) Then
                    strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                    If strCheck = "010" Then
                        .Col = .ColLetterToNumber("G")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errorMst  'static
                        .BackColor = mErrorColor
                        iFlagvMst = 1
                        'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
                    Else
                        .Col = .ColLetterToNumber("G")
                        .BackColor = mNonErrorColor
                        .CellNote = ""
                    End If
            ElseIf Len(vMST) = 0 Then
                        .BackColor = mNonErrorColor
                        .CellNote = ""
                  
            Else
                        .Col = .ColLetterToNumber("G")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errorMst  'static
                        .BackColor = mErrorColor
                        iFlagvMst = 1
                        'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
            End If
            
        End If
         count1 = count1 + 1
        .Row = count1 + 31
        .Col = .ColLetterToNumber("B")
    Loop Until .Text = "aa"
    
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 15, IIf(iFlagvName = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 17, IIf(iFlagvSct = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 18, IIf(iFlagvDate = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 19, IIf(iFlagvMst = 1, "0", "1")
    
     If iFocusFlag = 1 Then
            ' setcell for focus
            strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
            .Col = .ColLetterToNumber("C")
            .Row = 15
            .Formula = strFormula
            .Row = 17
            .Formula = strFormula
            .Row = 18
            .Formula = strFormula
            .Row = 19
            .Formula = strFormula
        End If
    
    
'check nguoi nuoc ngoai
    .Sheet = 1
    .Row = 31 + 3 + count1
    Do
        .GetText .ColLetterToNumber("C"), .Row, colC
        .GetText .ColLetterToNumber("N"), .Row, ColN
        .GetText .ColLetterToNumber("O"), .Row, colO
        .GetText .ColLetterToNumber("K"), .Row, colK
        
        If colC <> "" Or ColN <> "" Or colO <> "" Or colK <> 0 Then
            iWhiteRow = 0
            count = count + 1
            'check name
            
            If colC = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("C")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = errorName  'static
                .BackColor = mErrorColor
                iFlagvName = 1
                'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
            Else
                .Col = .ColLetterToNumber("C")
                .BackColor = mNonErrorColor
                .CellNote = ""
         
            End If
     'so chung tu
          
            If ColN = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("N")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = errorSct  'static
                .BackColor = mErrorColor
                iFlagvSct = 1
                'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
            Else
                .Col = .ColLetterToNumber("N")
                .BackColor = mNonErrorColor
                .CellNote = ""
            End If
    'date
             If colO = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("O")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = errorDate  'static
                .BackColor = mErrorColor
                iFlagvDate = 1
                'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
            Else
                .Col = .ColLetterToNumber("O")
                .BackColor = mNonErrorColor
                .CellNote = ""
            End If

    ' Check MST
            .Col = .ColLetterToNumber("G")
            vMST = .Value
'            vMST = Replace(vMST, " ", "")
'            vMST = Replace(vMST, "-", "")
             vMST = Left(vMST, 10)
            strCheck = ""
            If (Len(vMST) = 10 Or Len(vMST) = 13) Then
                    strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                    If strCheck = "010" Then
                        .Col = .ColLetterToNumber("G")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errorMst  'static
                        .BackColor = mErrorColor
                        iFlagvMst = 1
                        'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
                    Else
                        .Col = .ColLetterToNumber("G")
                        .BackColor = mNonErrorColor
                        .CellNote = ""
                    End If
            ElseIf Len(vMST) = 0 Then
                        .BackColor = mNonErrorColor
                        .CellNote = ""
                  
            Else
                        .Col = .ColLetterToNumber("G")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errorMst  'static
                        .BackColor = mErrorColor
                        iFlagvMst = 1
                        'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
            End If
         End If
         count2 = count2 + 1
        .Row = count1 + count2 + 34
        .Col = .ColLetterToNumber("B")
    Loop Until .Text = "aa"
    
    
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 15, IIf(iFlagvName = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 17, IIf(iFlagvSct = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 18, IIf(iFlagvDate = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 19, IIf(iFlagvMst = 1, "0", "1")
    
     If iFocusFlag = 1 Then
            ' setcell for focus
            strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
            .Col = .ColLetterToNumber("C")
            .Row = 15
            .Formula = strFormula
            .Row = 17
            .Formula = strFormula
            .Row = 18
            .Formula = strFormula
            .Row = 19
            .Formula = strFormula
        End If
    
    
'check count
            .Sheet = 1
            .Col = .ColLetterToNumber("K")
            .Row = 10
            .Value = count
            UpdateCell fps, .Col, .Row, .Text
'check nguoi ky
'        .Col = .ColLetterToNumber("N")
'        .Row = 38 + count1 + count2
'        If .Value <> "" Then
'            .BackColor = mNonErrorColor
'            .CellNote = ""
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 23, "1"
'            .Sheet = 1
'        Else
'            .BackColor = mErrorColor
'            .CellNote = error1
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 23, "0"
'            .Sheet = 1
'        End If
'check ngay ky
        .Col = .ColLetterToNumber("N")
        .Row = 40 + count1 + count2
        If .Value <> "" Then
            .BackColor = mNonErrorColor
            .CellNote = ""
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 23, "1"
            .Sheet = 1
        Else
            .BackColor = mErrorColor
            .CellNote = error1
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 23, "0"
            .Sheet = 1
        End If
    .Sheet = iCurrentSheet
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

End Sub

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function




