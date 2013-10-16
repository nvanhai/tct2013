VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_02TAIN_DK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrSolanBosung As String

Private Const Row_MSTDN = 49
Private Const Col_MSTDN = "P"
Private Const Col_TKCT = "R"
Private Const Row_TKCT = 30
Private Const Col_TKBS = "W"
Private Const Row_TKBS = 30

Const SHEET_PHU_LUC_01 As Integer = 2
Const SHEET_PHU_LUC_SAN_LUONG As Integer = 3
Const sheetPl01_KHBS_GTGT As Integer = 4
Private Const ROW_START_02_1 = 27

' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String
Public chkDauTho As String
Public chkCondensate As String
Public chkKhiThienNhien As String
 
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()

    With fps
        .Sheet = 1
        .Row = 30
        .Col = .ColLetterToNumber("O")
        .Text = chkDauTho
        
        .Row = 30
        .Col = .ColLetterToNumber("P")
        .Text = chkCondensate
        
        .Row = 30
        .Col = .ColLetterToNumber("Q")
        .Text = chkKhiThienNhien
        
        .Row = 31
        .Col = .ColLetterToNumber("O")
        .Text = TAX_Utilities_New.Year
        
        SetDateFormat fps, 1, 92, fps.ColLetterToNumber("AR"), DDMMYYYY
        .Row = 92
        .Col = fps.ColLetterToNumber("AR")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        ' Phu luc 02-1
        SetDateFormat fps, SHEET_PHU_LUC_SAN_LUONG, ROW_START_02_1, fps.ColLetterToNumber("K"), DDMMYYYY
        SetDateFormat fps, SHEET_PHU_LUC_SAN_LUONG, ROW_START_02_1, fps.ColLetterToNumber("U"), DDMMYYYY
'        '--end set for sheet Phu luc 02
        
    End With

End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
Dim strColC As String
    Dim i As Long
    Dim intCtrl As Integer
    Dim strId As Variant
    Dim xmlNode As MSXML.IXMLDOMNode
    With fps
        .Sheet = 1
        mCurrentSheet = .Sheet

        'update MSTDN luôn bang MST tren phan TT chung
        UpdateMST fps, Col_MSTDN, Row_MSTDN

        'If Strloaitk = "TKCT" Then
        If Strloaitk <> "TKBS" Then
            .Col = .ColLetterToNumber(Col_TKCT)
            .Row = Row_TKCT
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
'            .Col = .ColLetterToNumber("F")
'            .Row = 59
'            .Text = ""
'            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber(Col_TKBS)
            .Row = Row_TKBS
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber(Col_TKCT)
            .Row = Row_TKCT
            .Text = ""
'            UpdateCell fps, .Col, .Row, .Text
'            .Col = .ColLetterToNumber("F")
'            .Row = 59
'            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber(Col_TKBS)
            .Row = Row_TKBS
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        mCurrentSheet = .Sheet
    End With
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, iCol As Long, iRow As Long
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tCol As Long, tRow As Long
Dim strId As Variant, Col7 As Variant, intIndexCombo As Integer

    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = SHEET_PHU_LUC_01 Or .ActiveSheet = SHEET_PHU_LUC_SAN_LUONG Then
                   fps.EventEnabled(EventAllEvents) = False
                    i = 0
                    .Col = .ColLetterToNumber("B")
                    .Row = 27
                    Do
                         i = i + 1
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = .Row + 1
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

Private Sub fps_LeaveCell(ByVal Col As Long, _
                          ByVal Row As Long, _
                          ByVal NewCol As Long, _
                          ByVal NewRow As Long, _
                          Cancel As Boolean)
    Dim varTemp     As Variant
    Dim CellNguoiKy As String
    Dim count       As Long

    'CellNguoiKy = GetAttribute(TAX_Utilities_New.Data(0).getElementsByTagName("Cell")(TAX_Utilities_New.Data(0).getElementsByTagName("Cell").length - 1), "CellID")
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet

        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("AR") And Row = 92 Then
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
        
        'format date for row static phu luc san luong 02-1
        If .ActiveSheet = SHEET_PHU_LUC_SAN_LUONG And Row >= 27 Then
            If Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("U") Then
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
                UpdateCell fps, .Col, .Row, .Text
            End If
        End If
        
        If .ActiveSheet = SHEET_PHU_LUC_01 Or .ActiveSheet = SHEET_PHU_LUC_SAN_LUONG Then
            'tinh so dong
            count = 0
            .Col = .ColLetterToNumber("B")
            .Row = 27

            Do
                count = count + 1
                .Row = .Row + 1
            Loop Until .value = "aa"
            
            If Col = .ColLetterToNumber("AD") And Row = 34 + count - 1 Then
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
    Dim strCheck          As String, iCurrentSheet As Integer, iWhiteRow As Integer
    Dim iFlagCheckMaST    As Integer, iFlagEmty As Integer, iFlagDateCompare As Integer
    Dim varTemp           As Variant, vMST As Variant, errMST As Variant, errEmty As Variant, errDateCompare As Variant, vTenNhaThau As Variant
    Dim vTyLePhanBo As Variant, vSoThuePhatSinh As Variant, vGhiChu As Variant, vDate1 As Variant, vDate2 As Variant
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer
    Dim arrDate1() As String, arrDate2() As String
    
    With fps
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 13, errMST
        .GetText .ColLetterToNumber("E"), 19, errEmty
        .GetText .ColLetterToNumber("E"), 22, errDateCompare
        iFlagCheckMaST = 0
        iFlagEmty = 0
        
        .Sheet = iCurrentSheet
    
        'Check MST static
        .Sheet = SHEET_PHU_LUC_01

        If .SheetVisible = True Then
            mCurrentSheet = .Sheet

            iFlagCheckMaST = 0
            iWhiteRow = 1
            .Row = 27

            Do
                .GetText .ColLetterToNumber("D"), .Row, vMST
                .GetText .ColLetterToNumber("L"), .Row, vTenNhaThau
                .GetText .ColLetterToNumber("U"), .Row, vTyLePhanBo
                .GetText .ColLetterToNumber("AB"), .Row, vSoThuePhatSinh
                .GetText .ColLetterToNumber("AL"), .Row, vGhiChu
                If Trim$(vMST) <> "" Or Trim$(vTenNhaThau) <> "" Or Trim$(vTyLePhanBo) <> "0" Or Trim$(vSoThuePhatSinh) <> "0" Or Trim$(vGhiChu) <> "" Then
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
                        .Col = .ColLetterToNumber("D")
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
                        .Col = .ColLetterToNumber("D")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If

                Else
                    If (Trim$(vMST) = vbNullString And iWhiteRow = 0) Then
                        .Col = .ColLetterToNumber("D")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        
                        iFlagEmty = 1
                    End If
                End If
                If (Trim$(vTenNhaThau) = vbNullString And iWhiteRow = 0) Then
                    .Col = .ColLetterToNumber("L")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = errEmty   'static
                    .BackColor = mErrorColor
                    
                    iFlagEmty = 1
                Else
                    .Col = .ColLetterToNumber("L")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                .Col = .ColLetterToNumber("B")
                .Row = .Row + 1
            Loop Until .Text = "aa"

        End If
        
        'validate sheet phu luc san luong
        .Sheet = SHEET_PHU_LUC_SAN_LUONG
        If .SheetVisible = True Then
            .Row = 27
            iFlagDateCompare = 0
            Do
                .GetText .ColLetterToNumber("K"), .Row, vDate1
                .GetText .ColLetterToNumber("U"), .Row, vDate2
                arrDate1 = Split(CStr(vDate1), "/")
                arrDate2 = Split(CStr(vDate2), "/")
                If (Trim$(vDate1) <> vbNullString And Trim$(vDate2) <> vbNullString) Then
                    If DateSerial(CInt(arrDate1(2)), CInt(arrDate1(1)), CInt(arrDate1(0))) >= DateSerial(CInt(arrDate2(2)), CInt(arrDate2(1)), CInt(arrDate2(0))) Then
                        iFlagDateCompare = 1
                        .Col = .ColLetterToNumber("U")
                        .BackColor = mErrorColor
                        .CellNote = errDateCompare
                    Else
                        .Col = .ColLetterToNumber("U")
                        .CellNote = ""  'static
                        .BackColor = mNonErrorColor
                    End If
                End If
                .Col = .ColLetterToNumber("B")
                .Row = .Row + 1
            Loop Until .Text = "aa"
        End If
        
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 19, IIf(iFlagEmty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 13, IIf(iFlagCheckMaST = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 22, IIf(iFlagDateCompare = 1, "0", "1")
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
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As Variant
    Dim strValue2 As Variant
    Dim strValue3 As Variant
    Dim strValue4 As Variant
    Dim strValue5 As Variant
    Dim strValue6 As Variant
    Dim strValue7 As Variant
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_New.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "01-1"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                fps.Sheet = SHEET_PHU_LUC_01
                fps.GetText fps.ColLetterToNumber("D"), 27, strValue4
                    If Val(strValue1) <> 0 Or (Not IsNullValue(strValue4)) Then
                            blnNullValue = False
                    End If

                    If blnNullValue Then
                        SetAttribute xmlNodeValid, "Active", "0"
                    End If
             Case "01-2"
                blnNullValue = True
                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
'                    (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
                fps.Sheet = SHEET_PHU_LUC_SAN_LUONG
                fps.GetText fps.ColLetterToNumber("D"), 27, strValue1
                fps.GetText fps.ColLetterToNumber("K"), 27, strValue2
                fps.GetText fps.ColLetterToNumber("P"), 27, strValue3
                fps.GetText fps.ColLetterToNumber("U"), 27, strValue4
                fps.GetText fps.ColLetterToNumber("Y"), 27, strValue5
                fps.GetText fps.ColLetterToNumber("AD"), 27, strValue6
                fps.GetText fps.ColLetterToNumber("AQ"), 27, strValue7
                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Or Val(strValue3) <> 0 Or (Not IsNullValue(strValue4)) Or Val(strValue5) <> 0 Or Val(strValue6) <> 0 Or (Not IsNullValue(strValue7)) Then
                            blnNullValue = False
                    End If

                    If blnNullValue Then
                        SetAttribute xmlNodeValid, "Active", "0"
                    End If
             Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_New.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_New.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell").length - 4)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
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

Public Sub SetData()
    Dim varTemp         As Variant
    Dim ssSheet         As Integer
    Dim lCol            As Long, lRow As Long
    Dim xmlCellNode     As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode

    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet
'        'clear checkbox
'        .Row = 28
'        .Col = .ColLetterToNumber("Y")
'        .value = 0
'        .Col = .ColLetterToNumber("AS")
'        .value = 0
        .Sheet = .ActiveSheet
        .EventEnabled(EventAllEvents) = True
    End With

    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

Private Sub UpdateSheets()
    Dim varTemp         As Variant
    Dim ssSheet         As Integer
    Dim lCol            As Long, lRow As Long
    Dim xmlCellNode     As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode

    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet

        If (GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "Active") <> "0") Then

            For Each xmlCellNodeData In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
               
            Next

        End If
        'updated phu luc 01
        .Sheet = SHEET_PHU_LUC_01
        mCurrentSheet = .Sheet
        For Each xmlCellNodeData In TAX_Utilities_New.Data(1).getElementsByTagName("Cell")
           ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
        Next
        'end pl 01
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
    tuRowDL = 13
    denRowDL = 18
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "2_3_2"
    ' end
End Sub

