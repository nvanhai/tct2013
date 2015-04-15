VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01TNCN"
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
        SetDateFormat fps, 1, 31, .ColLetterToNumber("O"), DDMMYYYY
        
        .Sheet = 1 'To khai 01/TNCN
        .Row = 31
        .Col = .ColLetterToNumber("O")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    
    With fps
    End With
     
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
   Dim i As Integer
    Dim t As Variant, a As Variant, b As Variant, c As Variant
  
    Dim ArrRow As Variant, ArrCol As Variant
    Dim S As String
    
With fps
    'Because of unclearly-rised events, this function is here for turning it off
    'at the end of this sub, this event will be turned on
'    .EventEnabled(EventAllEvents) = False
'    .Sheet = .ActiveSheet
'
'    mCurrentSheet = .Sheet
'    GetCellSpan fps, Col, Row
'    .GetText .Col, S1Txt1Row, t
'    If (Row = S1Txt1Row And Col = .ColLetterToNumber(S1Txt1Col)) And vt <> CStr(t) Then
'        .Row = Row
'        .Col = Col
'        'UpdateCell fps, .Col, .Row, .Text & "chk"
'        'Xoa trang 1a
'        .Row = S1Txt1aRow
'        .Text = "0"
'        UpdateCell fps, .Col, .Row, .Value
'
'        'Xoa trang 1b
'        .Row = S1Txt1bRow
'        .Text = "0"
'        UpdateCell fps, .Col, .Row, .Text
'
'         'Xoa trang 1c
'        .Row = S1Txt1cRow
'        .Text = "0"
'        UpdateCell fps, .Col, .Row, .Text
'
'        .Row = S1Txt1Row
'        UpdateCell fps, .Col, .Row, .Value
'         .Row = 15
'        UpdateCell fps, .Col, .Row, .Value
'        .Row = 17
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Row = Row
'        .Col = Col
'    End If
'
'    If (Row = S1Txt1aRow And Col = .ColLetterToNumber(S1Txt1aCol)) Or _
'        (Row = S1Txt1bRow And Col = .ColLetterToNumber(S1Txt1bCol)) Or _
'        (Row = S1Txt1cRow And Col = .ColLetterToNumber(S1Txt1cCol)) Then
'        .GetText .Col, S1Txt1aRow, a
'        .GetText .Col, S1Txt1bRow, b
'        .GetText .Col, S1Txt1cRow, c
'        If va <> CStr(a) Or _
'            vb <> CStr(b) Or _
'            vc <> CStr(c) Then
'                .Row = S1Txt1Row
'                .Text = a + b + c
'                UpdateCell fps, .Col, .Row, .Value '.Text
'        End If
'
'        .Row = 15
'        UpdateCell fps, .Col, .Row, .Value
'        .Row = 17
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Row = Row
'        .Col = Col
'    End If
'    .EventEnabled(EventAllEvents) = True
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
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 9
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 9
                         i = i + 1
                    Loop Until .Text = "aa"
                    .SetActiveCell iCol, iRow

                    .Row = iRow
                    .Col = iCol
                     fps.EventEnabled(EventAllEvents) = False
                     .SetActiveCell iCol, iRow
            End If
            If .ActiveSheet = .SheetCount - 1 Then
                    fps.EventEnabled(EventAllEvents) = True
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 6
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 6
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
            If Col = .ColLetterToNumber("O") And Row = 31 Then
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
    Dim error2 As Variant
    With fps
    '.Visible = False
    '.ReDraw = False
    iCurrentSheet = .Sheet
    .Sheet = .SheetCount
    .GetText .ColLetterToNumber("E"), 15, error2
    CheckErrorMST
'check ngay ky

        .Sheet = 1
        .Col = .ColLetterToNumber("O")
        .Row = 31
        If .Value <> "" Then
            .BackColor = mNonErrorColor
            .CellNote = ""
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "1"
            .Sheet = 1
        Else
            .BackColor = mErrorColor
            .CellNote = error2
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "0"
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


