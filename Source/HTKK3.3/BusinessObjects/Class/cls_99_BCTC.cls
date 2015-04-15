VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_99_BCTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    With fps
    
        SetDateFormat fps, 1, 212, .ColLetterToNumber("R"), DDMMYYYY
        .Sheet = 1 ' Bao cao ket qua hoat dong san xuat kinh doanh Bo 99 dung cho Cong ty chung khoan
        .Row = 212
        .Col = .ColLetterToNumber("R")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        ' Set multiline cho ten nguoi nop thue
        .Row = 5
        .Col = .ColLetterToNumber("G")
        .TypeEditMultiLine = True
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    setLastValueTarget
    setLastValueTarget1
    setLastValueTarget4
    setLastValueTarget5
End Sub
Public Sub fps_Change(ByVal Col As Long, ByVal Row As Long)

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
    With fps
    '.Visible = False
    '.ReDraw = False
    iCurrentSheet = .Sheet
    .Sheet = .SheetCount

    CheckErrorMST
    
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

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    With fps
        mCurrentSheet = .Sheet
        .EventEnabled(EventAllEvents) = False
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("R") And Row = 212 Then
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

Private Sub setLastValueTarget()
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    'Dim strOriginDataFileName As String
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim strValueTarget As String
    Dim iDom As Integer
    
    Dim iRow As Integer
    
    iDom = 0
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        
        strCurrentDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(iDom), "DataFile") & "_" & TAX_Utilities_v2.Year & ".xml"
        '***************************************
        
        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(GetValidityNode("21", , , Format((TAX_Utilities_v2.Year - 1), "####")).childNodes(iDom), "DataFile") & "_" & Format((TAX_Utilities_v2.Year - 1), "####") & ".xml"
        '***************************************
        
        If xmlDomCurrentData.Load(strCurrentDataFileName) = True Then
            GoTo Tidy_up
        End If
        If .Load(strDataFileName) = True Then
            ' Lay cac gia tri cua bang can doi ke toan
            ' Cac gia tri so nam nay cua ky truoc chuyen sang so nam truoc cua ky hien tai
            ' Va tu dong 40 den 208 se mang gia tri cua bang can doi ke toan
            For iRow = 40 To 208
                ' Lay cac dong ko co trong datafile de loai ra
                If iRow <> 72 Or iRow <> 101 Or iRow <> 105 Or iRow <> 115 Or iRow <> 138 Or iRow <> 148 Or iRow <> 170 Then
                    Set xmlNode = .nodeFromID("R_" & iRow) ' R_40 la bat dau mang gia tri cua bang can doi ke toan
                    strValueTarget = GetAttribute(xmlNode, "Value")
                    Set xmlNode = Nothing
                    
                    fps.Sheet = 1
                    fps.Col = fps.ColLetterToNumber("W")
                    fps.Row = iRow
                    If Trim(fps.Text) = "0" Or Trim(fps.Text) = "" Then
                        fps.Text = strValueTarget
                        mCurrentSheet = 1
                        UpdateCell fps, fps.Col, fps.Row, fps.value
                    End If
                End If
            Next
        ElseIf .parseError.reason <> vbNullString Then
            If InStr(1, .parseError.errorCode, "2146697210") <> 0 Then
                
            Else
                MsgBox .parseError.reason
                Exit Sub
            End If
        End If
        
    End With
    
Tidy_up:
    Set xmlDomLastData = Nothing
    Set xmlDomCurrentData = Nothing
    Set xmlNodeListCell = Nothing
    Set xmlNode = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_99_01_CDKT", "setLastValueTarget", Err.number, Err.Description
End Sub


Private Sub setLastValueTarget1()
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    'Dim strOriginDataFileName As String
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim strValueTarget As String
    Dim iDom As Integer
    
    Dim iRow As Integer
    
    iDom = 1
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        
        strCurrentDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(iDom), "DataFile") & "_" & TAX_Utilities_v2.Year & ".xml"
        '***************************************
        
        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(GetValidityNode("21", , , Format((TAX_Utilities_v2.Year - 1), "####")).childNodes(iDom), "DataFile") & "_" & Format((TAX_Utilities_v2.Year - 1), "####") & ".xml"
        '***************************************
        
        If xmlDomCurrentData.Load(strCurrentDataFileName) = True Then
            GoTo Tidy_up
        End If
        If .Load(strDataFileName) = True Then
            ' Lay cac gia tri cua bang can doi ke toan
            ' Cac gia tri so nam nay cua ky truoc chuyen sang so nam truoc cua ky hien tai
            ' Va tu dong 9 den 33 se mang gia tri cua bang san xuat kinh doanh
            For iRow = 9 To 33
                Set xmlNode = .nodeFromID("I_" & iRow) ' I_9 la bat dau mang gia tri cua bang san xuat kinh doanh
                strValueTarget = GetAttribute(xmlNode, "Value")
                Set xmlNode = Nothing
                
                fps.Sheet = 2
                fps.Col = fps.ColLetterToNumber("H")
                fps.Row = iRow
                If Trim(fps.Text) = "0" Or Trim(fps.Text) = "" Then
                    fps.Text = strValueTarget
                    mCurrentSheet = 2
                    UpdateCell fps, fps.Col, fps.Row, fps.value
                End If
            Next
        ElseIf .parseError.reason <> vbNullString Then
            If InStr(1, .parseError.errorCode, "2146697210") <> 0 Then
                
            Else
                MsgBox .parseError.reason
                Exit Sub
            End If
        End If
        
    End With
    
Tidy_up:
    Set xmlDomLastData = Nothing
    Set xmlDomCurrentData = Nothing
    Set xmlNodeListCell = Nothing
    Set xmlNode = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_99_02_SXKD", "setLastValueTarget", Err.number, Err.Description
End Sub

Private Sub setLastValueTarget4()
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    'Dim strOriginDataFileName As String
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim strValueTarget As String
    Dim iDom As Integer
    
    Dim iRow As Integer
    
    iDom = 4
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        
        strCurrentDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(iDom), "DataFile") & "_" & TAX_Utilities_v2.Year & ".xml"
        '***************************************
        
        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(GetValidityNode("21", , , Format((TAX_Utilities_v2.Year - 1), "####")).childNodes(iDom), "DataFile") & "_" & Format((TAX_Utilities_v2.Year - 1), "####") & ".xml"
        '***************************************
        
        If xmlDomCurrentData.Load(strCurrentDataFileName) = True Then
            GoTo Tidy_up
        End If
        If .Load(strDataFileName) = True Then
            ' Lay cac gia tri cua bang can doi ke toan
            ' Cac gia tri so nam nay cua ky truoc chuyen sang so nam truoc cua ky hien tai
            ' Va tu dong 40 den 70 se mang gia tri cua bang luu chuyen tien te truc tiep
            For iRow = 40 To 70
                ' Lay cac dong ko co trong datafile de loai ra
                If iRow <> 50 Or iRow <> 59 Then
                    Set xmlNode = .nodeFromID("R_" & iRow) ' R_40 la bat dau mang gia tri cua bang luu chuyen tien te truc tiep
                    strValueTarget = GetAttribute(xmlNode, "Value")
                    Set xmlNode = Nothing
                    
                    fps.Sheet = 5
                    fps.Col = fps.ColLetterToNumber("W")
                    fps.Row = iRow
                    If Trim(fps.Text) = "0" Or Trim(fps.Text) = "" Then
                        fps.Text = strValueTarget
                        mCurrentSheet = 5
                        UpdateCell fps, fps.Col, fps.Row, fps.value
                    End If
                End If
            Next
        ElseIf .parseError.reason <> vbNullString Then
            If InStr(1, .parseError.errorCode, "2146697210") <> 0 Then
                
            Else
                MsgBox .parseError.reason
                Exit Sub
            End If
        End If
        
    End With
    
Tidy_up:
    Set xmlDomLastData = Nothing
    Set xmlDomCurrentData = Nothing
    Set xmlNodeListCell = Nothing
    Set xmlNode = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_99_03_LCTTTT", "setLastValueTarget", Err.number, Err.Description
End Sub

Private Sub setLastValueTarget5()
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    'Dim strOriginDataFileName As String
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim strValueTarget As String
    Dim iDom As Integer
    
    Dim iRow As Integer
    
    iDom = 5
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        
        strCurrentDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(iDom), "DataFile") & "_" & TAX_Utilities_v2.Year & ".xml"
        '***************************************
        
        strDataFileName = TAX_Utilities_v2.DataFolder & GetAttribute(GetValidityNode("21", , , Format((TAX_Utilities_v2.Year - 1), "####")).childNodes(iDom), "DataFile") & "_" & Format((TAX_Utilities_v2.Year - 1), "####") & ".xml"
        '***************************************
        
        If xmlDomCurrentData.Load(strCurrentDataFileName) = True Then
            GoTo Tidy_up
        End If
        If .Load(strDataFileName) = True Then
            ' Lay cac gia tri cua bang can doi ke toan
            ' Cac gia tri so nam nay cua ky truoc chuyen sang so nam truoc cua ky hien tai
            ' Va tu dong 40 den 88 se mang gia tri cua bang luu chuyen tien te gian tiep
            For iRow = 40 To 88
                ' Lay cac dong ko co trong datafile de loai ra
                If iRow <> 69 Then
                    Set xmlNode = .nodeFromID("R_" & iRow) ' R_40 la bat dau mang gia tri cua bang luu chuyen tien te gian tiep
                    strValueTarget = GetAttribute(xmlNode, "Value")
                    Set xmlNode = Nothing
                    
                    fps.Sheet = 6
                    fps.Col = fps.ColLetterToNumber("W")
                    fps.Row = iRow
                    If Trim(fps.Text) = "0" Or Trim(fps.Text) = "" Then
                        fps.Text = strValueTarget
                        mCurrentSheet = 6
                        UpdateCell fps, fps.Col, fps.Row, fps.value
                    End If
                End If
            Next
        ElseIf .parseError.reason <> vbNullString Then
            If InStr(1, .parseError.errorCode, "2146697210") <> 0 Then
                
            Else
                MsgBox .parseError.reason
                Exit Sub
            End If
        End If
        
    End With
    
Tidy_up:
    Set xmlDomLastData = Nothing
    Set xmlDomCurrentData = Nothing
    Set xmlNodeListCell = Nothing
    Set xmlNode = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_99_04_LCTTGT", "setLastValueTarget", Err.number, Err.Description
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
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
        If (GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(1), "Active") <> "0") Then
        .Sheet = 2
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
        If (GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(2), "Active") <> "0") Then
        .Sheet = 3
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(2).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
        If (GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(3), "Active") <> "0") Then
        .Sheet = 4
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(3).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
        
      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub
