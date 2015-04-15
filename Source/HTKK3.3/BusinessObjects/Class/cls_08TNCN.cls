VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_08TNCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrSolanBosung As String
Public strLoaiTKQT As String
Public strQuy As String

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
        SetDateFormat fps, 1, 66, .ColLetterToNumber("R"), DDMMYYYY
        .Sheet = 1 'To khai 08/TNCN
        .Row = 66
        .Col = .ColLetterToNumber("R")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
     
    End With
    
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim a2     As Variant, b2 As Variant, c2 As Variant
    Dim count1 As Long, lRow As Long, lCol As Long
    Dim tuThang, denThang As Variant
    Dim xmlNode As MSXML.IXMLDOMNode

    With fps
        'tinh so dong` them vao`
        mCurrentSheet = .ActiveSheet
        .Sheet = 1
        .Row = 68
           
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet

        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("C")
            .Row = 68
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 68
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("C")
            .Row = 68
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 68
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If

        'set thong tin header neu la to khai quy
'        If strQuy = "TK_QUY" Then
        .SetText .ColLetterToNumber("H"), 18, 1
        .SetText .ColLetterToNumber("R"), 18, ""
        .SetText .ColLetterToNumber("X"), 18, ""
'        Else
'            .SetText .ColLetterToNumber("H"), 18, 0
'            .SetText .ColLetterToNumber("R"), 18, ""
'            .SetText .ColLetterToNumber("X"), 18, ""
'            .SetText .ColLetterToNumber("K"), 4, TAX_Utilities_v2.FirstDay
'            .SetText .ColLetterToNumber("P"), 4, TAX_Utilities_v2.LastDay
'            'set thong tin header neu la to khai tuthang denthang
'            If strLoaiTKQT = "QT_TUNG_NAM" Then
'                .Col = .ColLetterToNumber("R")
'                .Row = 18
'                .Text = 1
''                .Col = .ColLetterToNumber("X")
''                .Text = ""
'            ElseIf strLoaiTKQT = "QT_NAM_DAU" Then
'                .Col = .ColLetterToNumber("X")
'                .Row = 18
'                .Text = 1
''                .Col = .ColLetterToNumber("R")
''                .Text = ""
'            End If
'
'        End If
          
        'lay gia tri va set hieu thoi gian tu thang den thang
'        .GetText .ColLetterToNumber("K"), 4, tuThang
'        .GetText .ColLetterToNumber("P"), 4, denThang
'
'        If tuThang <> vbNullString And denThang <> vbNullString Then
'            tuThang = "01/" & tuThang
'            denThang = "01/" & denThang
'            .Col = .ColLetterToNumber("AE")
'            .Row = 7
''            .Text = DateDiff("M", Format_ddmmyyyy(CStr(tuThang)), Format_ddmmyyyy(CStr(denThang))) + 1
'            .Text = DateDiff("M", Format("01/" & TAX_Utilities_v2.FirstDay, "dd/mm/yyyy"), Format("01/" & TAX_Utilities_v2.LastDay, "dd/mm/yyyy")) + 1
'        Else
'            .Col = .ColLetterToNumber("AE")
'            .Row = 7
'            .Text = 1
'        End If

        'set lai gia tri cho chi tieu [30]
        For Each xmlNode In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
            ParserCellID fps, GetAttribute(xmlNode, "CellID"), lCol, lRow
            If lCol = .ColLetterToNumber("R") And lRow = 58 Then
                .Col = lCol
                .Row = lRow
                .Text = GetAttribute(xmlNode, "Value")
            End If
        Next
        .EventEnabled(EventAllEvents) = True
    End With

    If Strloaitk <> "TKBS" Then
        setChitiet
    End If

End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim varCT11 As Variant
    Dim varCT12 As Variant
    Dim count1 As Long
    Dim CellNgayKy As String
    Dim strValue As String
    Dim varThangKT As Variant
    CellNgayKy = GetAttribute(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length - 3), "CellID")
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            ' check ngay ky
            If Col = .ColLetterToNumber("R") And Row = Right(CellNgayKy, 2) Then
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
            
'            CheckDynamicError
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
    Dim MST1 As Variant, MST2 As Variant, MST3 As Variant, MST4 As Variant
    Dim MST5 As Variant, MST6 As Variant, MST7 As Variant, MST8 As Variant
    Dim MST9 As Variant, MST10 As Variant, MST11 As Variant, MST12 As Variant, MST13 As Variant
    Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    Dim strCheck As String, iCurrentSheet As Integer
    Dim count, count1 As Long
    Dim errorName, errorDate, errorThang, errorMst, error1, vMST, errorPhuongXa  As Variant
    Dim colC, colH, colU As Variant
    
    Dim iFlagvName As Integer, iFlagvDate As Integer, iFlagvMst As Integer, iFlagvThang As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer
    Dim strFormula As String
    iFlagvName = 0
    iFlagvDate = 0
    iFlagvMst = 0
    iFlagvThang = 0
    iFocusFlag = 0
    Dim iWhiteRow As Integer
    Dim mErrCheckCT30, vCT30, vCT29, vCT31 As Variant, iFlagCheckCT30 As Integer
    Dim vCT25 As Variant, vCT26 As Variant
    Dim iFlagCt25 As Integer, iFlagCt26 As Integer
    Dim vErrorCt25 As Variant, vErrorCt26 As Variant
    Dim vErrorCt25_1 As Variant, vErrorCt26_1 As Variant
    
    With fps

        CheckErrorMST
        .Sheet = .SheetCount
     
        .GetText .ColLetterToNumber("E"), 19, error1
        .GetText .ColLetterToNumber("E"), 20, errorThang
        .GetText .ColLetterToNumber("E"), 24, mErrCheckCT30
        .GetText .ColLetterToNumber("E"), 21, vErrorCt25
        .GetText .ColLetterToNumber("E"), 22, vErrorCt26

        .GetText .ColLetterToNumber("E"), 25, vErrorCt25_1
        .GetText .ColLetterToNumber("E"), 26, vErrorCt26_1
        
        .Sheet = 1
        iFlagCheckCT30 = 0
        iFlagCt25 = 0
        iFlagCt26 = 0
        'get CT30,CT29, CT31
        .GetText .ColLetterToNumber("R"), 56, vCT29
        .GetText .ColLetterToNumber("R"), 57, vCT30
        .GetText .ColLetterToNumber("R"), 58, vCT31
        
        .GetText .ColLetterToNumber("R"), 52, vCT25
        .GetText .ColLetterToNumber("R"), 53, vCT26
        
                
        'check CT30 > CT31
        If vCT31 > vCT30 Then
            .Col = .ColLetterToNumber("R")
            .Row = 58
            .BackColor = mErrorColor
            .CellNote = mErrCheckCT30
            iFlagCheckCT30 = 1
        Else
            .Col = .ColLetterToNumber("R")
            .Row = 58
            .BackColor = mNonErrorColor
            .CellNote = ""
        End If
        ' kiem tra chi tieu 25, 26 la boi cua 9000000, 3600000
        If (TAX_Utilities_v2.Year = 2013 And TAX_Utilities_v2.ThreeMonths >= 3) Or (TAX_Utilities_v2.Year > 2013) Then
           ' ct25
            If Val(vCT25) Mod 9000000 <> 0 Or Val(vCT25) > 27000000 Then
                .Row = 52
                .Col = .ColLetterToNumber("R")
                .BackColor = mErrorColor
                .CellNote = vErrorCt25_1
                iFlagCt25 = 1
            Else
                .Row = 52
                .Col = .ColLetterToNumber("R")
                .BackColor = mNonErrorColor
                .CellNote = ""
            End If
           ' ct26
            If Val(vCT26) Mod 3600000 <> 0 Then
                .Row = 53
                .Col = .ColLetterToNumber("R")
                .BackColor = mErrorColor
                .CellNote = vErrorCt26_1
                iFlagCt26 = 1
            Else
                .Row = 53
                .Col = .ColLetterToNumber("R")
                .BackColor = mNonErrorColor
                .CellNote = ""
            End If
        Else
            'ct25
            If Val(vCT25) Mod 4000000 <> 0 Or Val(vCT25) > 12000000 Then
                .Row = 52
                .Col = .ColLetterToNumber("R")
                .BackColor = mErrorColor
                .CellNote = vErrorCt25
                iFlagCt25 = 1
            Else
                .Row = 52
                .Col = .ColLetterToNumber("R")
                .BackColor = mNonErrorColor
                .CellNote = ""
            End If
            'ct26
            If Val(vCT26) Mod 1600000 <> 0 Then
                 .Row = 53
                .Col = .ColLetterToNumber("R")
                .BackColor = mErrorColor
                .CellNote = vErrorCt26
                iFlagCt26 = 1
            Else
                .Row = 53
                .Col = .ColLetterToNumber("R")
                .BackColor = mNonErrorColor
                .CellNote = ""

            End If
        End If
        
        
        'check ngay ky
        .Col = .ColLetterToNumber("R")
        .Row = 66
        If .value <> "" Then
            .BackColor = mNonErrorColor
            .CellNote = ""
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 19, "1"
            .Sheet = 1
        Else
            .BackColor = mErrorColor
            .CellNote = error1
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 19, "0"
            .Sheet = 1
        End If
        'set lai header neu CT30 > CT29
        If iFlagCheckCT30 = 1 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 24, "0"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 24, "1"
        End If
'        .Row = 45
'        If .value Mod 1600000 = 0 Then
'            .BackColor = mNonErrorColor
'            .CellNote = ""
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 20, "1"
'            .Sheet = 1
'        Else
'            .BackColor = mErrorColor
'            .CellNote = errorThang
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 20, "0"
'            .Sheet = 1
'        End If
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 21, IIf(iFlagCt25 = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 22, IIf(iFlagCt26 = 1, "0", "1")

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

Private Sub setChitiet()
    
End Sub

Private Sub fillData(pGrid As fpSpread, xmlNodeList As MSXML.IXMLDOMNodeList)
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    Dim lRows As Long
    
    
    With pGrid
     .Sheet = mCurrentSheet
     .EventEnabled(EventChange) = False
      For Each xmlNodeCell In xmlNodeList
         ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
            If GetAttribute(xmlNodeCell, "FirstCell") = "1" Then
                lRows = GetDynRowCount(pGrid, xmlNodeCell.parentNode)
                InsertRow pGrid, lRow, lRows, True
            End If
       Next
      ' .EventEnabled(EventChange) = True
       
       For Each xmlNodeCell In xmlNodeList
            If GetAttribute(xmlNodeCell, "StatusID") <> vbNullString Then
                ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
                .Col = lCol
                .Row = lRow
                    Select Case .CellType
                        Case CellTypeNumber
                             .value = GetAttribute(xmlNodeCell, "Value")
                        Case CellTypeEdit
                             .Text = GetAttribute(xmlNodeCell, "Value")
                        Case CellTypePic
                             .Text = GetAttribute(xmlNodeCell, "Value")
                        Case Else
                             .value = GetAttribute(xmlNodeCell, "Value")
                    End Select
            End If
        Next
       .EventEnabled(EventChange) = True
   Set xmlNodeCell = Nothing
  End With
       
End Sub

Private Sub InsertRow(fpSpread1 As fpSpread, ByVal pRow As Long, lRows As Long, Optional blnFillingData As Boolean = False)
    On Error GoTo ErrorHandle
    
    Dim i As Long, lBgColor As Long
    Dim lRowCtrl As Long, lColCtrl As Long
    'Dim mCurrentSheet As Long
    
    With fpSpread1
        '.Visible = False
        .ReDraw = False
        '.Sheet = mCurrentSheet
        .MaxRows = .MaxRows + lRows
        .InsertRows pRow, lRows
        For lRowCtrl = 1 To lRows
        
            .CopyRowRange pRow - lRowCtrl, pRow - lRowCtrl, pRow + lRows - lRowCtrl
            .Row = pRow - lRowCtrl
            '.RowHeight(pRow - lRowCtrl) = 14
            If Not blnFillingData Then
                For i = 1 To fpSpread1.MaxCols
                    '***************************
                    'ThanhDX added
                    'Date: 26/12/2005
                    .Col = i
                    lBgColor = .BackColor
                    .Row = pRow + lRows - lRowCtrl
                    If Not .Lock Then
                        'Set BgColor to inserted cell
                        If lBgColor <> &HC0C0FF And lBgColor <> 12713215 Then 'vbRed
                            .BackColor = lBgColor
                        Else
                            .BackColor = vbWhite
                        End If
                    '***************************
                    ' ThanhDX added
                    ' Date: 29/04/06
                    Else
                        If Not TAX_Utilities_v2.Data(mCurrentSheet - 1).nodeFromID( _
                           GetCellID(fpSpread1, i, pRow - lRowCtrl)) Is Nothing Then
                            If .BackColor = &HC0C0FF Or .BackColor = 12713215 Then
                                .BackColor = vbWhite
                            End If
                        End If
                    '***************************
                    End If
                    '***************************
                    ' Reset empty value for new row on grid
                    If .Lock = False Then
                        Select Case .CellType
                            Case CellTypeNumber
                                .SetText i, .Row, 0
                            Case Else
                                .SetText i, .Row, vbNullString
                        End Select
                        .CellNote = vbNullString
                    '***************************
                    ' ThanhDX added
                    ' Date: 08/04/06
                    Else
                        If Not TAX_Utilities_v2.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, i, pRow - lRowCtrl)) Is Nothing Then
                            Select Case .CellType
                                Case CellTypeNumber
                                    .SetText i, .Row, 0
                                Case Else
                                    .SetText i, .Row, vbNullString
                            End Select
                            .CellNote = vbNullString
                        End If
                    '***************************
                    End If
                Next i
            End If
        Next lRowCtrl
        '.Visible = True
        .ReDraw = True
    End With
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "mdlFunctions", "InsertRow", Err.number, Err.Description
End Sub


Private Function GetDynRowCount(pGrid As fpSpread, xmlNodeCells As MSXML.IXMLDOMNode, Optional ByRef lReportRows As Long, Optional ByRef lMinRow As Long, Optional lMaxRow As Long)
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim lRow As Long, lCol As Long
    Dim lRow2 As Long, lCol2 As Long
    Dim lMaxRow2 As Long, lMinRow2 As Long
    
    lMinRow = 100000
    lMaxRow = 0
    lMinRow2 = 100000
    lMaxRow2 = 0
    
    If Not xmlNodeCells Is Nothing Then
        For Each xmlNodeCell In xmlNodeCells.childNodes
            'Get CellID
            ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID"), lCol, lRow
            
            'Get CellID2
            ParserCellID pGrid, GetAttribute(xmlNodeCell, "CellID2"), lCol2, lRow2
            
            'Get max row
            If lRow > lMaxRow Then
                lMaxRow = lRow
            End If
            
            'Get min row
            If lRow < lMinRow Then
                lMinRow = lRow
            End If
            
            'Get max row
            If lRow2 > lMaxRow2 Then
                lMaxRow2 = lRow2
            End If
            
            'Get min row
            If lRow2 < lMinRow2 Then
                lMinRow2 = lRow2
            End If
        Next
        
        GetDynRowCount = lMaxRow - lMinRow + 1
        lReportRows = lMaxRow2 - lMinRow2 + 1
    End If
End Function

Private Sub UpdateSheets()
    Dim varTemp As Variant
    Dim ssSheet As Integer
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    
    Dim blCheck_S2 As Boolean
    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet
        For Each xmlCellNodeData In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
           ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'               If (lRow = 18 And lCol = .ColLetterToNumber("R")) Or (lRow = 18 And lCol = .ColLetterToNumber("X")) Then
'                    .GetText lCol, lRow, varTemp
'                    If varTemp = "1" Or UCase(CStr(varTemp)) = "X" Then
'                        UpdateCell fps, lCol, lRow, "x"
'                    Else
'                        UpdateCell fps, lCol, lRow, ""
'                    End If
'               Else
            .GetText lCol, lRow, varTemp
            UpdateCell fps, lCol, lRow, varTemp
'              End If
        Next
    End With
End Sub


' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 22
    denRowDL = 27
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "2_3_2"
    ' end
End Sub

