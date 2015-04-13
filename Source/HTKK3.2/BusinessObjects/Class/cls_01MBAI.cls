VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01MBAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Thu Nhap Doanh Nghiep" interface sheets
'this Class is belong to TAX_Business_v1 project which will be compline to DLL
'Coder: PhuocHK
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Private Const S1Txt1Row = 9
Private Const S1Txt1Col = "I"
Private Const S1Txt1aRow = 10
Private Const S1Txt1aCol = "I"
Private Const S1Txt1bRow = 11
Private Const S1Txt1bCol = "I"
Private Const S1Txt1cRow = 12
Private Const S1Txt1cCol = "I"
Dim vt As String, va As String, vb As String, vc As String
Dim strComboHien As String, strCombo As String

Dim mblKeyF5 As Boolean
Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim arrData() As String
Dim arrDanhSachKHBS() As String
Dim TenNVThue As Variant, SoChungChi As Variant
Public loaiKHBS As String
Public FlagThaydoi As String
Public Strloaitk As String
Public StrSolanBosung As String
Private Const Row_MSTDN = 14
Private Const Col_MSTDN = "H"

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
    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    With fps
        SetDateFormat fps, 1, 62, .ColLetterToNumber("R"), DDMMYYYY
        .Sheet = 1
        .Row = 62
        .Col = .ColLetterToNumber("R")
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
       ' set co quan thue quan ly
    setCQTQuanLyHoanThue fps
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
         'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("C")
            .Row = 35
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 35
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 35
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("C")
            .Row = 35
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 35
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 35
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        .EventEnabled(EventAllEvents) = True
    End With
    
End Sub


Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    Dim arrValue() As String
    Dim vonDK As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        If Col = .ColLetterToNumber("P") And Row = 45 Then
            .GetText Col, Row, vonDK
            arrValue = Split(convertBacTMB(vonDK), "~")
            ' set bac mon bai
            .Col = .ColLetterToNumber("U")
            .Row = Row
            .Text = arrValue(0)
            UpdateCell fps, .Col, .Row, .value
            ' set muc thu MB
            .Col = .ColLetterToNumber("Y")
            .Row = Row
            .Text = arrValue(1)
            UpdateCell fps, .Col, .Row, .value
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
            fps.EventEnabled(EventAllEvents) = False
                 i = 1
                 .Row = 49
                 .Col = .ColLetterToNumber("B")
                 Do
                     .Text = str(i)
                     
                     .Col = .ColLetterToNumber("U")
                     .Text = "4"
                     
                     .Col = .ColLetterToNumber("Y")
                     .value = "1000000"
                     
                     
                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                     i = i + 1
                 Loop Until .Text = "aa"
            End If
            fps.EventEnabled(EventAllEvents) = True
         End If
    End With
End Sub
Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("R") And Row = .MaxRows - 1 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                        .TypeHAlign = TypeHAlignLeft
                    Else
                        .SetActiveCell Col, Row
                        .SetFocus
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
    'CellChange Col, Row
End Sub
Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
  
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
            
        .EventEnabled(EventAllEvents) = False
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .Sheet = 1
        Dim i As Integer
         i = 1
         .Row = 49
         .Col = .ColLetterToNumber("B")
         Do
             .Text = str(i)
             .Col = .ColLetterToNumber("B")
             .Row = .Row + 1
             i = i + 1
         Loop Until .Text = "aa"
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
Dim iFlagEmty As Variant
Dim iFlagMucThue As Integer
Dim vErrMucThue As Variant
Dim vErrEmpty As Variant
Dim vonDK As Variant
Dim vMucThue As Variant
Dim arrValue() As String
Dim iWhiteRow As Integer
Dim ct2 As Variant
    With fps
     .Sheet = .SheetCount
     .GetText .ColLetterToNumber("E"), 17, vErrMucThue
     .GetText .ColLetterToNumber("E"), 15, vErrEmpty
     
     .Sheet = 1
     ' kiem tra cot 2 bat buoc nhap
     
     
     
     ' kiem tra muc thue
     .GetText .ColLetterToNumber("P"), 45, vonDK
     .GetText .ColLetterToNumber("Y"), 45, vMucThue
     .GetText .ColLetterToNumber("C"), 45, ct2
     iWhiteRow = 0
     If Trim$(ct2) <> "" Or Val(vonDK) <> 0 Or Val(vMucThue) <> 0 Then iWhiteRow = 1
     
     arrValue = Split(convertBacTMB(vonDK), "~")
     If vMucThue = arrValue(1) Or vMucThue = arrValue(1) / 2 Or iWhiteRow = 0 Then
        .Col = .ColLetterToNumber("Y")
        .Row = 45
        .CellNote = ""
        .BackColor = mNonErrorColor
     Else
        .Col = .ColLetterToNumber("Y")
        .Row = 45
        .CellNote = vErrMucThue
        .BackColor = mErrorColor
        iFlagMucThue = 1
     End If
     
     .Row = 49
     Do
        .GetText .ColLetterToNumber("P"), .Row, vonDK
        .GetText .ColLetterToNumber("Y"), .Row, vMucThue
        .GetText .ColLetterToNumber("C"), .Row, ct2
        
        iWhiteRow = 0
        If Trim$(ct2) <> "" Or Val(vonDK) <> 0 Then iWhiteRow = 1
        
        If Val(vMucThue) = 1000000 Or Val(vMucThue) = 500000 Or iWhiteRow = 0 Then
           .Col = .ColLetterToNumber("Y")
           .CellNote = ""
           .BackColor = mNonErrorColor
        Else
           .Col = .ColLetterToNumber("Y")
           .CellNote = vErrMucThue
           .BackColor = mErrorColor
           iFlagMucThue = 1
        End If
        
        ' kiem tra cot ten
        If Trim$(ct2) = "" Then
           .Col = .ColLetterToNumber("C")
           .CellNote = vErrEmpty
           .BackColor = mErrorColor
           iFlagEmty = 1
        Else
            .Col = .ColLetterToNumber("C")
           .CellNote = ""
           .BackColor = mNonErrorColor
        End If
        
        .Row = .Row + 1
        .Col = .ColLetterToNumber("B")
     Loop Until .Text = "aa"
     
     .Sheet = .SheetCount
     .SetText .ColLetterToNumber("B"), 15, IIf(iFlagEmty = 1, "0", "1")
     .SetText .ColLetterToNumber("B"), 17, IIf(iFlagMucThue = 1, "0", "1")
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
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    Dim strValue3 As String
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
             Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Then
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
        Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol         As Long, lRow As Long
    Dim IsUpdate     As Boolean
    Dim idtkhai      As Variant
    Dim totalCell1   As Long, countRowReset1 As Long
    Dim a            As Long
    Dim sheetTemp    As Integer
    
    With fps
        sheetTemp = mCurrentSheet
        mCurrentSheet = .ActiveSheet
        For Each xmlNodeReset In TAX_Utilities_v1.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            fps.Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            fps.Col = lCol
            fps.Row = lRow

                If .ActiveSheet = 1 Then
                    .Sheet = mCurrentSheet
                    If (lRow < 36 Or lRow > .MaxRows - 5) Then
                        GoTo nextStep
                    End If
                ElseIf .ActiveSheet = 2 Then
                    .Sheet = mCurrentSheet
                    If (lRow < 5 Or lRow > .MaxRows - 5) Then
                        GoTo nextStep
                    End If
                End If
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

                
                
'            TAX_Utilities_v1.AdjustData(mCurrentSheet - 1) = IIf(IsUpdate = True, IsUpdate, TAX_Utilities_v1.AdjustData(mCurrentSheet - 1))
nextStep:
        Next
        mCurrentSheet = sheetTemp
    End With
End Function
Public Sub SetData()

End Sub

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

Public Sub Prepare_KHBS()
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    Dim strPath As String
    Dim strComboName As String
    Dim strComboMCT As String
    Dim strCombo As String
    Dim blChon As Boolean
    Dim arrDanhsach() As String
    fps.Sheet = fps.SheetCount - 1
    fps.ActiveSheet = fps.Sheet
    
    blChon = False
End Sub


Public Sub UpdateChangeKHBS()
    
End Sub



Private Function seachCellNode(xmlCellNodelist As MSXML.IXMLDOMNodeList, cellid As String) As MSXML.IXMLDOMNode
Dim nodeReturn As MSXML.IXMLDOMNode
    For Each nodeReturn In xmlCellNodelist
        If GetAttribute(nodeReturn, "CellID") = cellid Then
            Set seachCellNode = nodeReturn
            Exit For
            Exit Function
        End If
    Next
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
            Set xmlNode = .nodeFromID("Y_44")
            result = "CT24_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01MBAI", "getValueTK", Err.number, Err.Description
End Function

Public Function getDieuChinhGiam(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue, strDieuChinhGiam, strDieuChinhTang As String
    Dim tempArrOld() As String
    Dim tempArrCurr() As String
    Dim tempArrOldValue() As String
    Dim tempArrCurrValue() As String
    Dim i, j As Integer
    Dim strMaCT As String
    Dim arrMapCT() As String ' Luu gia tri cua dong chi tieu
    Dim arrMapCTValue() As String ' Luu gia tri cua cac cell chi tieu
    strMaCT = getMapCT
    With fps
                .Sheet = 1
                .GetText .ColLetterToNumber("Y"), 44, varTemp
                strCurrentValue = "CT24_" & Trim(varTemp)
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
                getDieuChinhGiam = strDieuChinhGiam & "###" & strDieuChinhTang
      End With
End Function
Function getMapCT() As String
    On Error GoTo ErrHandle
    
    Dim tempValue As Variant
    Dim strPath As String
    Dim i As Integer
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT_TT28.xml"
    
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "ID") = GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "DataFile") Then
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
    SaveErrorLog "01_MBAI", "getMapCT", Err.number, Err.Description
End Function
' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("Y"), 44, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function

Public Function getValueCTDC(strFileLoadData As String) As String
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
            Set xmlNode = .nodeFromID("Y_44")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01MBAI", "getValueCTDC", Err.number, Err.Description
End Function
'******************************
'Set trang thai Flag KHBS active
Public Sub SetActiveKHBS() 'ByRef xmlNodeValid As MSXML.IXMLDOMNode)
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim blnNullValue As Boolean
    Dim strValue1 As String
    Dim strValue2 As String
    Dim strValue3 As String
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
                Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length > 31 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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


' 27-01-2011
Private Sub UpdateSheets()
    Dim varTemp As Variant
    Dim ssSheet As Integer
    Dim blCheckS2 As Boolean
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        
        .Sheet = 1
        mCurrentSheet = .Sheet
        For Each xmlCellNodeData In TAX_Utilities_v1.Data(.Sheet - 1).getElementsByTagName("Cell")
           ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
           If lCol = .ColLetterToNumber("P") And lRow = 20 Then
                .GetText lCol, lRow, varTemp
                If Trim$(varTemp) = "x" Or Trim$(varTemp) = "X" Or Trim$(varTemp) = "1" Then
                    UpdateCell fps, lCol, lRow, "x"
                Else
                    UpdateCell fps, lCol, lRow, ""
                End If
           Else
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
           End If
        Next

       ' Update lai sheet KHBS
        .Sheet = .SheetCount - 1
        mCurrentSheet = .Sheet
        blCheckS2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheckS2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(.Sheet - 1).getElementsByTagName("Cell")
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

' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 13
    denRowDL = 18
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "20_30_2"
    ' end
End Sub


Private Function convertBacTMB(vonDK As Variant) As String
    Dim result As String
    result = "0~0"
    If Val(CStr(vonDK)) / 1000 > 10000000 Then
        result = "1~3000000"
    ElseIf Val(CStr(vonDK)) / 1000 > 5000000 And Val(CStr(vonDK)) / 1000 <= 10000000 Then
        result = "2~2000000"
    ElseIf Val(CStr(vonDK)) / 1000 > 2000000 And Val(CStr(vonDK)) / 1000 <= 5000000 Then
        result = "3~1500000"
    ElseIf Val(CStr(vonDK)) / 1000 <= 2000000 Then
        result = "4~1000000"
    End If
    convertBacTMB = result
End Function


