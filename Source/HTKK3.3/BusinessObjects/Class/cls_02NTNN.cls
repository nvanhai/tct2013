VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_02NTNN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrTKThang_PS As String
Public StrSolanBosung As String

Public loaiKHBS As String
Public FlagThaydoi As String

Private Const Row_MSTDN = 21
Private Const Col_MSTDN = "Y"

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
    With fps
'       nkhoan: format dinh dang ngay thang
        SetDateFormat fps, 1, 47, .ColLetterToNumber("AU"), DDMMYYYY
        SetDateFormat fps, 1, 22, .ColLetterToNumber("AH"), DDMMYYYY
'        SetDateFormat fps, 2, 23, .ColLetterToNumber("AS"), DDMMYYYY
'        SetDateFormat fps, 3, 24, .ColLetterToNumber("AR"), DDMMYYYY
            
        .Sheet = 1
'        .Row = 22
'        .Col = .ColLetterToNumber("AH")
'        .Text = Format(Date, "dd/mm/yyyy")
'        .TypeHAlign = TypeHAlignLeft
'
        
        .Row = 47
        .Col = .ColLetterToNumber("AU")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        SetDateFormat fps, .SheetCount - 1, 31, .ColLetterToNumber("BE"), DDMMYYYY
        SetDateFormat fps, .SheetCount - 1, 46, .ColLetterToNumber("BG"), DDMMYYYY
        .Sheet = .SheetCount - 1 'KHBS
        .Row = 46
        .Col = .ColLetterToNumber("BG")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
    End With
    
    setCQTQuanLyHoanThue fps

End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
        
    With fps
        .Sheet = 1
        'currentSheet2 = mCurrentSheet
        mCurrentSheet = .Sheet

        'set to khai chinh thuc hay bao sung
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("M")
            .Row = 23
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("N")
            .Row = 23
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("O")
            .Row = 23
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("M")
            .Row = 23
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("N")
            .Row = 23
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("O")
            .Row = 23
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
    End With
        
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    CellChange Col, Row
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    With fps
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            .Sheet = .ActiveSheet
            Select Case .Sheet
                Case 2
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 23
                    Do
                        .Text = str(i)
                        .Col = .ColLetterToNumber("B")
                        .Row = i + 23
                        i = i + 1
                    Loop Until .Text = "aa"

                Case 3
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 24
                    Do
                        .Sheet = 3
                        .Text = str(i)
                        .Col = .ColLetterToNumber("B")
                        .Row = i + 24
                        i = i + 1
                    Loop Until .Text = "aa"
            End Select
        End If
    End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, _
                          ByVal Row As Long, _
                          ByVal NewCol As Long, _
                          ByVal NewRow As Long, _
                          Cancel As Boolean)
    Dim varTemp  As Variant
    Dim i        As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        'check date
        Select Case .Sheet
            Case 1

                If Col = .ColLetterToNumber("AH") And Row = 22 Then
                    .GetText Col, Row, varTemp
                    If varTemp <> "" And varTemp <> "../../...." Then
                        If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                            .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                        Else
                            .SetActiveCell Col, Row
                        End If
                    Else
                        .SetText Col, Row, ""
                    End If

                    .Col = Col
                    .Row = Row
                    UpdateCell fps, Col, Row, .Text
                End If
                
                If Col = .ColLetterToNumber("AU") And Row = 47 Then
                    .GetText Col, Row, varTemp
                    If varTemp <> "" And varTemp <> "../../...." Then
                        If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                            .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                        Else
                            .SetActiveCell Col, Row
                        End If
                    Else
                        .SetText Col, Row, ""
                    End If
                    .Col = Col
                    .Row = Row
                    UpdateCell fps, Col, Row, .Text
                End If

'            Case 2
'
'                If Col = .ColLetterToNumber("AS") Then
'                    .GetText Col, Row, varTemp
'
'                    If varTemp <> "" And varTemp <> "../../...." Then
'                        If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
'                            .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                        Else
'
'                            .SetActiveCell Col, Row
'                        End If
'
'                    Else
'                        .SetText Col, Row, ""
'                    End If
'
'                    .Col = Col
'                    .Row = Row
'                    UpdateCell fps, Col, Row, .Text
'                End If
'
'            Case 3
'
'                If Col = .ColLetterToNumber("AR") Then
'                    .GetText Col, Row, varTemp
'
'                    If varTemp <> "" And varTemp <> "../../...." Then
'                        If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
'                            .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                        Else
'
'                            .SetActiveCell Col, Row
'                        End If
'
'                    Else
'                        .SetText Col, Row, ""
'                    End If
'
'                    .Col = Col
'                    .Row = Row
'                    UpdateCell fps, Col, Row, .Text
'                End If
        End Select
        .EventEnabled(EventAllEvents) = True
    End With
    
    'CellChange Col, Row
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
    ' check ma so thue
    If TAX_Utilities_v2.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0" Then
        CheckErrorMST2 23, 17, "O", 2
        'CheckErrorMST2 23, 18, "W", 2
        'CheckHanNop 23, "AS", 22, 2
    End If

    If TAX_Utilities_v2.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0" Then
        CheckErrorMST2 24, 19, "L", 3
        'CheckHanNop 24, "AR", 22, 3
    End If
               
End Sub



'<CSCM>
'--------------------------------------------------------------------------------
' Description:       check Thoi han hop dong khong duoc lon hon ngay hien tai
' Created by :       Khai Hoan
' Date-Time  :       9/16/2011-4:31:43 PM
' paramater  :       - iRow, iCol: toa do cua Cell Thoi han.
'                    - iSheet: sheet dag lam viec.
'--------------------------------------------------------------------------------
'</CSCM>
'Private Sub CheckHanNop(iRow As Integer, iCol As String, iRowErr As String, isheet As Integer)
'    Dim ngayNop, errHanNop  As Variant, ngayHienTai As Variant
'    Dim strFocusSheetName, strFocusRow, strFocusCol, strFormula As String
'    Dim iFocusFlag, iFlagvMst As Integer
'    With fps
'        .Sheet = .SheetCount
'        .GetText .ColLetterToNumber("E"), iRowErr, errHanNop
'        .Sheet = isheet
'        .Row = iRow
'        Do
'            ' check ngay nop khong dc lon hon ngay hien tai
'            .GetText .ColLetterToNumber(iCol), .Row, ngayNop
'            If Trim$(ngayNop) <> "" Then
'                ngayHienTai = IIf(Len(Day(Now())) = 1, "0" & Day(Now()), Day(Now())) & "/" & IIf(Len(Month(Now())) = 1, "0" & Month(Now()), Month(Now())) & "/" & Year(Now())
'                ngayHienTai = DateSerial(Int(Mid$(ngayHienTai, 7, 4)), Int(Mid$(ngayHienTai, 4, 2)), Int(Mid$(ngayHienTai, 1, 2)))
'                ngayNop = DateSerial(Int(Mid$(ngayNop, 7, 4)), Int(Mid$(ngayNop, 4, 2)), Int(Mid$(ngayNop, 1, 2)))
'                If ngayNop > ngayHienTai Then
'                    .Col = .ColLetterToNumber(iCol)
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = errHanNop   'static
'                    .BackColor = mErrorColor
'                    iFlagvMst = 1
'                    ' Forcus ve Cell co du lieu bi loi
'                    If iFocusFlag = 0 Then
'                        strFocusSheetName = "'" & .SheetName & "'"
'                        strFocusCol = str(.Col)
'                        strFocusRow = str(.Row)
'                        .Sheet = .SheetCount
'                        strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
'                        .Col = .ColLetterToNumber("C")
'                        .Row = iRowErr
'                        .Formula = strFormula
'                    End If
'                Else
'                    .Sheet = isheet
'                    .Col = .ColLetterToNumber(iCol)
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                End If
'            End If
'
'            .Sheet = isheet
'            iRow = iRow + 1
'            .Row = iRow
'            .Col = .ColLetterToNumber("B")
'        Loop Until .Text = "aa"
'        .Sheet = .SheetCount
'        .SetText .ColLetterToNumber("B"), iRowErr, IIf(iFlagvMst = 1, "0", "1")
'    End With
'End Sub


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

'<CSCM>
'--------------------------------------------------------------------------------
' Description:       check Ma so thue trong cac bang co dong dong
' Created by :       Khai Hoan
' Date-Time  :       9/16/2011-4:31:43 PM
' paramater  :       - iRow: dong thu 1 trong bang du lieu dong trong file Exel.
'                    - iCol: cot chua Cell du lieu Ma so thue.
'                    - iSheet: sheet dag lam viec.
'                    - iRow2: dong Messages thong bao loi cua Sheet Header trong file Exel
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub CheckErrorMST2(iRow As Integer, _
                          iRow2 As Integer, _
                          iCol As String, _
                          isheet As Integer)
    Dim strCheck, strFocusSheetName, strFocusRow, strFocusCol, strFormula As String

    Dim strNgayHD, strNgayTT, i, vErrMsgMST, vMST As Variant

    Dim iFocusFlag, iFlagvMst As Integer
    'Dim mNonErrorColor = vbWhite
    
    With fps
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), iRow2, vErrMsgMST
        iFocusFlag = 0
        .Sheet = isheet
        .Col = .ColLetterToNumber("B")
        .Row = iRow
        i = 1

        Do
            '----- check ma so thue  -------
            .GetText .ColLetterToNumber(iCol), .Row, vMST
            If Trim(vMST) <> vbNullString Then
'                vMST = Replace(vMST, " ", "")
'                vMST = Replace(vMST, "-", "")
                strCheck = ""
                If Len(vMST) = 14 And InStr(vMST, "-") > 0 Then
                   vMST = Left$(vMST, 10) & Right$(vMST, 3)
                End If
            
                If Len(vMST) = 10 Or Len(vMST) = 13 Then
                    strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                End If
                   
                If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                    .Col = .ColLetterToNumber(iCol)
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMST   'static
                    .BackColor = mErrorColor

                    iFlagvMst = 1
                        
                    ' Forcus ve Cell co du lieu bi loi
'                    If iFocusFlag = 0 Then
'                        strFocusSheetName = "'" & .SheetName & "'"
'                        strFocusCol = str(.Col)
'                        strFocusRow = str(.Row)
'                        .Sheet = .SheetCount
'                        strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
'                        .Col = .ColLetterToNumber("C")
'                        .Row = iRow2
'                        .Formula = strFormula
'                    End If

                Else
                    .Sheet = isheet
                    .Col = .ColLetterToNumber(iCol)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If

            Else
                .Sheet = isheet
                .Col = .ColLetterToNumber(iCol)
                .CellNote = ""
                .BackColor = mNonErrorColor
               
            End If

            .Sheet = isheet
            .Row = iRow + i
            i = i + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"

        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), iRow2, IIf(iFlagvMst = 1, "0", "1")
    End With

End Sub
Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

Public Sub SetActiveSheet()
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl        As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue     As String
    Dim strValue1    As String
    Dim strValue2    As String
    Dim strValue3    As String
    Dim strValue4    As Variant
    Dim strValue5    As Variant
    Dim iCol         As Long, iRow As Long

    
    Dim tong1        As Variant
    Dim tong2        As Variant
    Dim tong3        As Variant
    Dim tong4        As Variant
    Dim tong5        As Variant

    With fps
        'lay cac gia tri tong cua  sheet 2
        Set xmlCellNode = TAX_Utilities_v2.Data(1).nodeFromID(GetCellID(fps, .ColLetterToNumber("E"), 23))
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode.lastChild
        Set xmlCellNode = xmlCellNode.firstChild
        Set xmlCellNode = xmlCellNode.firstChild
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        .Sheet = 2
        .GetText iCol, iRow, tong1
        Set xmlCellNode = xmlCellNode.nextSibling
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        .GetText iCol, iRow, tong2
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        .GetText iCol, iRow, tong3
        
        ' lay cac dong tong cua sheet 3
        Set xmlCellNode = TAX_Utilities_v2.Data(2).nodeFromID(GetCellID(fps, .ColLetterToNumber("E"), 24))
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode.lastChild
        Set xmlCellNode = xmlCellNode.firstChild
        Set xmlCellNode = xmlCellNode.firstChild
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        .Sheet = 3
        .GetText iCol, iRow, tong4
        Set xmlCellNode = xmlCellNode.nextSibling
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        .GetText iCol, iRow, tong5
        
    End With
            
    For Each xmlNodeValid In TAX_Utilities_v2.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "01-1"

                If tong1 = 0 And tong2 = 0 And tong3 = 0 Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If
            
            Case "01-2"

                If tong4 = 0 And tong5 = 0 Then
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
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol         As Long, lRow As Long
    Dim IsUpdate     As Boolean
    Dim idtkhai      As Variant
    Dim totalCell1 As Long, countRowReset1 As Long
    
    With fps
        totalCell1 = TAX_Utilities_v2.Data(fps.SheetCount - 2).getElementsByTagName("Cell").length
        countRowReset1 = (totalCell1 - 8) / 6
        For Each xmlNodeReset In TAX_Utilities_v2.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            fps.Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            fps.Col = lCol
            fps.Row = lRow
            
            If mCurrentSheet = 1 Then
                If 29 < lRow And lRow < 44 And mCurrentSheet = 1 Then

                    Select Case fps.CellType

                        Case CellTypeCheckBox
                            fps.Text = vbNullString
                            UpdateCell fps, lCol, lRow, vbNullString

                        Case CellTypeComboBox
                            fps.Text = vbNullString
                            UpdateCell fps, lCol, lRow, vbNullString

                        Case CellTypeNumber
                            fps.value = 0
                            UpdateCell fps, lCol, lRow, vbNullString

                        Case Else
                            fps.value = vbNullString
                            UpdateCell fps, lCol, lRow, vbNullString
                    End Select

                End If
            ElseIf (lRow >= 20 + countRowReset1) And mCurrentSheet = (fps.SheetCount - 1) Then
            Else

                Select Case fps.CellType

                    Case CellTypeCheckBox
                        fps.Text = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString

                    Case CellTypeComboBox
                        fps.Text = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString

                    Case CellTypeNumber
                        fps.value = 0
                        UpdateCell fps, lCol, lRow, vbNullString

                    Case Else
                        fps.value = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                End Select

            End If
                
            TAX_Utilities_v2.AdjustData(mCurrentSheet - 1) = IIf(IsUpdate = True, IsUpdate, TAX_Utilities_v2.AdjustData(mCurrentSheet - 1))

        Next

    End With

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
        If (GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "Active") <> "0") Then
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
                'If Not (lRow = 5 And (lCol = .ColLetterToNumber("AT") Or lCol = .ColLetterToNumber("BC"))) Then
                    ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
                'End If
             Next
        End If
        
        .Sheet = 2
        mCurrentSheet = .Sheet
        If (GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(1), "Active") <> "0") Then
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(1).getElementsByTagName("Cell")
                'If Not (lRow = 5 And (lCol = .ColLetterToNumber("AT") Or lCol = .ColLetterToNumber("BC"))) Then
                    ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
                'End If
             Next
        End If
        
        
        .Sheet = 3
        mCurrentSheet = .Sheet
        If (GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(2), "Active") <> "0") Then
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(2).getElementsByTagName("Cell")
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

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       HTKK 3.1
' Description:       update mot so Cell cua to khai sau khi xoa hoac them phu luc
' Created by :       nkhoan
' Date-Time  :       10/24/2011-4:44:09 PM
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub updateSomeCell()
    Dim blCheck_S1 As Boolean
    Dim blCheck_S2 As Boolean

    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = 1
        blCheck_S1 = IIf(TAX_Utilities_v2.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)

        If blCheck_S1 = True Then
            ' tong hop chi tieu cot 10b PL 02-1/NTNN va cot 9b PL 02-2/NTNN
            .Col = .ColLetterToNumber("AF")
            .Row = 30
            .Formula = "BK30"
           
             ' tong hop chi tieu cot 11b PL 02-1/NTNN
            .Col = .ColLetterToNumber("AF")
            .Row = 31
            .Formula = "BK31"
        End If
        
        blCheck_S2 = IIf(TAX_Utilities_v2.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)

        If blCheck_S2 = True Then
            ' tong hop chi tieu cot 10b PL 02-1/NTNN va cot 9b PL 02-2/NTNN
             .Col = .ColLetterToNumber("AF")
             .Row = 30
             .Formula = "BK30"
             
               ' tong hop chi tieu cot 10b
            .Col = .ColLetterToNumber("AF")
            .Row = 32
            .Formula = "BK32"
        End If
      
        .EventEnabled(EventAllEvents) = True
    End With

End Sub
'doc vao mot so va tra ve chu cua so do
'Private Function doiSoThanhChu(ByVal so As String) As String
'    Dim result As String
'    result = ""
'    Select Case so
'            Case "0": result = ""
'            Case "1": result = GetAttribute(GetMessageCellById("0204"), "Msg")
'            Case "2": result = GetAttribute(GetMessageCellById("0205"), "Msg")
'            Case "3": result = GetAttribute(GetMessageCellById("0206"), "Msg")
'            Case "4": result = GetAttribute(GetMessageCellById("0207"), "Msg")
'            Case "5": result = GetAttribute(GetMessageCellById("0208"), "Msg")
'            Case "6": result = GetAttribute(GetMessageCellById("0209"), "Msg")
'            Case "7": result = GetAttribute(GetMessageCellById("0210"), "Msg")
'            Case "8": result = GetAttribute(GetMessageCellById("0211"), "Msg")
'            Case "9": result = GetAttribute(GetMessageCellById("0212"), "Msg")
'    End Select
'    doiSoThanhChu = result
'End Function
'ham doi tien tu so thanh chu
'Private Function doiTien(ByVal so As String) As String
'    Dim GTtien As String
'    Dim value As String
'    Dim i As Integer
'    'luu lai gia tri so ban dau de xu ly
'    Dim soTam2 As String
'    Dim soTam As String
'    soTam = so
'    soTam2 = "1"
'    For i = 0 To Len(so) - 1 Step 1
'
'            If Len(so) Mod 3 = 2 And Mid(so, 1, 1) = "1" Then
'                value = GetAttribute(GetMessageCellById("0213"), "Msg")
'            Else
'                value = doiSoThanhChu(Mid(so, 1, 1))
'            End If
'            doiTien = doiTien + value
'            If (Len(so) Mod 3 = 0) Then
'                soTam2 = so
'            End If
'            If (Len(so) Mod 3 = 0) And Mid(so, 1, 1) <> "0" Then
'                doiTien = doiTien + GetAttribute(GetMessageCellById("0214"), "Msg")
'            ElseIf Len(so) Mod 3 = 2 And Mid(so, 1, 1) <> "0" And Mid(so, 1, 1) <> "1" Then
'                doiTien = doiTien + GetAttribute(GetMessageCellById("0219"), "Msg")
'            ElseIf Len(so) Mod 3 = 2 And Mid(so, 1, 1) = "0" And Mid(so, 2, 1) <> "0" Then
'                doiTien = doiTien + GetAttribute(GetMessageCellById("0220"), "Msg")
'            ElseIf Len(so) Mod 3 = 1 And Len(so) <> 1 And Val(Mid(soTam2, 1, 3)) <> 0 Then
'                doiTien = doiTien + giaTriTien(so)
'            End If
'            so = Mid(so, 2, Len(so) - 1)
'    Next
'    doiTien = Trim(doiTien) + GetAttribute(GetMessageCellById("0221"), "Msg")
'    doiTien = UCase(Left(doiTien, 1)) + Mid(doiTien, 2)
'End Function
'xac dinh menh gia tien
'Private Function giaTriTien(ByVal tien As String) As String
'    If Len(tien) < 4 Then
'        giaTriTien = GetAttribute(GetMessageCellById("0214"), "Msg")
'    ElseIf Len(tien) >= 4 And Len(tien) < 7 Then
'        giaTriTien = GetAttribute(GetMessageCellById("0215"), "Msg")
'    ElseIf Len(tien) >= 7 And Len(tien) < 10 Then
'        giaTriTien = GetAttribute(GetMessageCellById("0216"), "Msg")
'    ElseIf Len(tien) >= 10 And Len(tien) < 13 Then
'        giaTriTien = GetAttribute(GetMessageCellById("0217"), "Msg")
'    ElseIf Len(tien) = "13" Then
'        giaTriTien = GetAttribute(GetMessageCellById("0218"), "Msg")
'    End If
'End Function

Public Function getValueCTDC(strFileLoadData As String) As String
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    'Dim strOriginDataFileName As String
    Dim result As String
    Dim iDom As Integer
    iDom = 0
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        If .Load(strFileLoadData) = True Then
            Set xmlNode = .nodeFromID("AF_40")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_02NTNN", "getValueCTDC", Err.number, Err.Description
End Function

Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("AF"), 40, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function

Function getMapCT() As String
    On Error GoTo ErrHandle
    
    Dim tempValue As Variant
    Dim strPath As String
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT_TT28.xml"
    
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "ID") = GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") Then
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
    SaveErrorLog "02_NTNN", "getMapCT", Err.number, Err.Description
End Function

Public Function getDieuChinhGiam(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue, strDieuChinhGiam, strDieuChinhTang, strDieuChinhTang4043 As String
    Dim tempArrOld()       As String
    Dim tempArrCurr()      As String
    Dim tempArrOldValue()  As String
    Dim tempArrCurrValue() As String
    Dim i, j As Integer
    Dim strMaCT         As String
    Dim arrMapCT()      As String ' Luu gia tri cua dong chi tieu
    Dim arrMapCTValue() As String ' Luu gia tri cua cac cell chi tieu
    strMaCT = getMapCT

    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("AF"), 41, varTemp
        strCurrentValue = "CT7a_" & Trim(varTemp)
        .GetText .ColLetterToNumber("AF"), 42, varTemp
        strCurrentValue = strCurrentValue & "~CT7b_" & Trim(varTemp)
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

            If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
                ' dieu chinh giam
                strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
            End If

            If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
                ' dieu chinh tang
                strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
            End If
        Next i
                
        If strDieuChinhGiam <> "" Then
            strDieuChinhGiam = Mid(strDieuChinhGiam, 2, Len(strDieuChinhGiam) - 1)
        End If

        If strDieuChinhTang <> "" Then
            strDieuChinhTang = Mid(strDieuChinhTang, 2, Len(strDieuChinhTang) - 1)
        End If

'        If strDieuChinhTang4043 <> "" Then
'            strDieuChinhTang4043 = Mid(strDieuChinhTang4043, 2, Len(strDieuChinhTang4043) - 1)
'        End If

        getDieuChinhGiam = strDieuChinhGiam & "###" & strDieuChinhTang & "###" & strDieuChinhTang4043
    End With

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
            Set xmlNode = .nodeFromID("AF_41")
            result = "CT7a_" & Val(GetAttribute(xmlNode, "Value"))

            Set xmlNode = .nodeFromID("AF_42")
            result = result & "~" & "CT7b_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
            

        End If
    End With
    getValueTK = result
    Exit Function
    
ErrorHandle:
    SaveErrorLog "cls_02NTNN", "getValueTK", Err.number, Err.Description
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
    For Each xmlNodeValid In TAX_Utilities_v2.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
                Case "KHBS"
                    blnNullValue = True
                    'Check value of last node in Data
                    Set xmlCellNode = TAX_Utilities_v2.Data(3).getElementsByTagName("Cell") _
                        (TAX_Utilities_v2.Data(3).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v2.Data(3).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v2.Data(3).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v2.Data(2).getElementsByTagName("Cell").length > 31 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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
    tuRowDL = 22
    denRowDL = 30
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "13_15_2~14_16_3~20_30_4"
    ' end
End Sub



'When cell's data is changed or a control at a cell is handled, This sub is called
'parameter Col: the column of cell that rise event
'parameter Row: the row of cell that rise event
'parameter f: the flag to mark some special option
'**: when this sub is on executing, all events of fpspread will be disabled until the end of the sub
Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
            
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim iCol         As Long, iRow As Long
            
            
    With fps
        'Because of unclearly-rised events, this function is here for turning it off
        'at the end of this sub, this event will be turned on
        .EventEnabled(EventAllEvents) = False
        .Sheet = .ActiveSheet
        mCurrentSheet = .ActiveSheet
        If .ActiveSheet = 1 Then
                mCurrentSheet = .Sheet
                If Col = .ColLetterToNumber("AF") And Row = 30 Then
                    .Formula = ""
                End If
                
                If Col = .ColLetterToNumber("AF") And Row = 31 Then
                    .Formula = ""
                End If
                
                If Col = .ColLetterToNumber("AF") And Row = 32 Then
                    .Formula = ""
                End If
        End If
        
        If .ActiveSheet = 2 Then
                mCurrentSheet = .Sheet
                
                Set xmlCellNode = TAX_Utilities_v2.Data(1).nodeFromID(GetCellID(fps, .ColLetterToNumber("E"), 23))
                Set xmlCellNode = xmlCellNode.parentNode
                Set xmlCellNode = xmlCellNode.parentNode
                Set xmlCellNode = xmlCellNode.parentNode.lastChild
                Set xmlCellNode = xmlCellNode.firstChild
                Set xmlCellNode = xmlCellNode.firstChild
                ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
                
                If Col = .ColLetterToNumber("BG") And Row >= 23 And Row <= iRow - 2 Then
                    .Sheet = 1
                    ' tong hop chi tieu cot 10b PL 02-1/NTNN va cot 9b PL 02-2/NTNN
                    .Col = .ColLetterToNumber("AF")
                    .Row = 30
                    .Formula = "BK30"
                End If
                
                If Col = .ColLetterToNumber("BY") And Row >= 23 And Row <= iRow - 2 Then
                    .Sheet = 1
                    ' tong hop chi tieu cot 11b PL 02-1/NTNN
                    .Col = .ColLetterToNumber("AF")
                    .Row = 31
                    .Formula = "BK31"
                End If
        End If
        
        If .ActiveSheet = 3 Then
                mCurrentSheet = .Sheet
                
                Set xmlCellNode = TAX_Utilities_v2.Data(2).nodeFromID(GetCellID(fps, .ColLetterToNumber("E"), 24))
                Set xmlCellNode = xmlCellNode.parentNode
                Set xmlCellNode = xmlCellNode.parentNode
                Set xmlCellNode = xmlCellNode.parentNode.lastChild
                Set xmlCellNode = xmlCellNode.firstChild
                Set xmlCellNode = xmlCellNode.firstChild
                ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
                
                If Col = .ColLetterToNumber("BD") And Row >= 24 And Row <= iRow - 2 Then
                    .Sheet = 1
                    ' tong hop chi tieu cot 10b PL 02-1/NTNN va cot 9b PL 02-2/NTNN
                    .Col = .ColLetterToNumber("AF")
                    .Row = 30
                    .Formula = "BK30"
                End If
                
                If Col = .ColLetterToNumber("BS") And Row >= 24 And Row <= iRow - 2 Then
                    .Sheet = 1
                    ' tong hop chi tieu cot 10b
                    .Col = .ColLetterToNumber("AF")
                    .Row = 32
                    .Formula = "BK32"
                End If
        End If
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
