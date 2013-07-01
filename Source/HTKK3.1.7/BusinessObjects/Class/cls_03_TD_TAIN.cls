VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_03_TD_TAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Thu Nhap Doanh Nghiep" interface sheets
'this Class is belong to TAX_Business project which will be compline to DLL
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
Dim arrTaiNguyen As New Scripting.Dictionary, mblKhoiDau As Boolean
Dim arrTaiNguyenGoc As New Scripting.Dictionary
Dim mblKeyF5 As Boolean
Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
Dim arrData() As String
Dim arrDanhSachKHBS() As String
Dim TenNVThue As Variant, SoChungChi As Variant
Public loaiKHBS As String
Public FlagThaydoi As String
Public Strloaitk As String
Public StrSolanBosung As String
Private Const Row_MSTDN = 11
Private Const Col_MSTDN = "N"

'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    With fps
        .Sheet = 1
        .Row = 63
        .Col = fps.ColLetterToNumber("O")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        .Sheet = 3 'To khai BS
        .Row = 23
        .Col = .ColLetterToNumber("BG")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
    End With
    
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
Dim i As Integer, intIndexCombo As Integer
Dim strLTN As Variant, Col7 As Variant, strLTNCu As Variant, strId As Variant
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode
Dim iCol As Long, iRow As Long, tRow As Long

Dim count1 As Long
    With fps
        mblKhoiDauToKhai = True
        .Sheet = .ActiveSheet
        mCurrentSheet = .ActiveSheet
        mblKhoiDau = False
        .Sheet = 1
        'If mCurrentSheet = 16 Then
         'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
         'set to khai bo xung hay chinh thuc
            If Strloaitk = "TKCT" Then
                .Col = .ColLetterToNumber("C")
                .Row = 27
                .Text = 1
                UpdateCell fps, .Col, .Row, .Text
                
                .Col = .ColLetterToNumber("F")
                .Row = 27
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                
                .Col = .ColLetterToNumber("I")
                .Row = 27
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
             ElseIf Strloaitk = "TKBS" Then
                .Col = .ColLetterToNumber("C")
                .Row = 27
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                
                 .Col = .ColLetterToNumber("F")
                .Row = 27
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                
                .Col = .ColLetterToNumber("I")
                .Row = 27
                .Text = Val(StrSolanBosung)
                UpdateCell fps, .Col, .Row, .Text
            End If
         'end
    End With
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
Dim strId As Variant, strLTN As Variant
Dim intIndexCombo As Integer
Dim Col7 As Variant
Dim iRow As Long, iCol As Long, GTSua As String
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tCol As Long, tRow As Long
Dim cCol As Long, cRow As Long
Dim cValue As Variant, temp As Variant, varTemp As Variant
Dim cValueG As Variant
Dim cTagF As String
Dim cValueF As Variant
Dim i As Integer
Dim xmlNodeH15 As MSXML.IXMLDOMNode


    fps.EventEnabled(EventAllEvents) = False
    With fps

    End With
   fps.EventEnabled(EventAllEvents) = True
End Sub

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim strId As Variant
Dim iCol As Long, iRow As Long
Dim intIndexCombo As Integer, i As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim itemp As Variant, temp As Variant
Dim cValue As Variant
With fps
 
    
End With
End Sub



Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, iCol As Long, iRow As Long
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tCol As Long, tRow As Long
Dim strId As Variant, Col7 As Variant, intIndexCombo As Integer
Dim NguoiKyCol As Long, NguoiKyRow As Long
Dim count1 As Integer
Dim STT As Variant

    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
         If .ActiveSheet = 1 Then
                If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
                    fps.EventEnabled(EventAllEvents) = False
                    .Col = .ColLetterToNumber("B")
                    .Row = 44
                    i = 1
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 44
                         i = i + 1
                    Loop Until .Text = "aa"
                     fps.EventEnabled(EventAllEvents) = True
                End If
           End If
    End With
    
End Sub
Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim CellNguoiKy As String
    
    CellNguoiKy = GetAttribute(TAX_Utilities_New.Data(0).getElementsByTagName("Cell")(TAX_Utilities_New.Data(0).getElementsByTagName("Cell").length - 4), "CellID")
    
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("AJ") And Row = Right(CellNguoiKy, 2) Then
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
    Dim ASheet As Integer, SSheet As Integer, i As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        
        CheckDynamicError 'Set Exception Error on cells of interface
        ' Update sheet KHBS
        UpdateSheets
        ' end
        .ActiveSheet = ASheet
        .Sheet = SSheet
        
        'set lai STT cac dong
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        .Col = .ColLetterToNumber("B")
        .Row = 44
        i = 1
        Do
             .Text = str(i)
             .Col = .ColLetterToNumber("B")
             .Row = i + 44
             i = i + 1
        Loop Until .Text = "aa"
 
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    Dim iCurrentSheet As Integer
    Dim vMST, vCTieu, vSTT, vTyLe, vThueGTGT, vMSTtemp As Variant
    Dim vTenNM, vSanLuongDien, vGiaTriTinh, vThue, vThueMG, vThuePS As Variant
    Dim iFlagMST As Integer, iFlagCtieu As Integer, iWhiteRow As Integer, iFocusFlag As Integer
    Dim errMST, errCTieu As Variant
    Dim blCheck_S2 As Boolean
    Dim strCheck As String
    
    With fps
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 15, errCTieu
        .GetText .ColLetterToNumber("E"), 16, errMST
        
        .Sheet = 1
        .Col = .ColLetterToNumber("B")
        .Row = 44
        Do
            iWhiteRow = 1
            ' reset cell note
            .Col = .ColLetterToNumber("C")
            .CellNote = ""
            .BackColor = mNonErrorColor
            ' end if
            
            .GetText .ColLetterToNumber("C"), .Row, vTenNM
            .GetText .ColLetterToNumber("F"), .Row, vSanLuongDien
            .GetText .ColLetterToNumber("H"), .Row, vGiaTriTinh
            .GetText .ColLetterToNumber("L"), .Row, vThue
            .GetText .ColLetterToNumber("O"), .Row, vThueMG
            .GetText .ColLetterToNumber("S"), .Row, vThuePS
            If Trim$(vTenNM) <> vbNullString Or Val(vSanLuongDien) <> 0 Or Val(vGiaTriTinh) <> 0 Or Val(vThue) <> 0 Or Val(vThueMG) <> 0 Or Val(vThuePS) <> 0 Then
                iWhiteRow = 0
            End If
            
            If Trim$(vTenNM) = "" And iWhiteRow = 0 Then
                   .Col = .ColLetterToNumber("C")
                   .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                   .CellNote = errCTieu
                   .BackColor = mErrorColor
                   iFlagCtieu = 1
            Else
                   .Col = .ColLetterToNumber("C")
                   .CellNote = ""
                   .BackColor = mNonErrorColor
            End If
            
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
         .Sheet = .SheetCount
         .SetText .ColLetterToNumber("B"), 15, IIf(iFlagCtieu = 1, "0", "1")
            
        .Sheet = 2
         blCheck_S2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
         If blCheck_S2 = True Then
         
                .Col = .ColLetterToNumber("B")
                .Row = 41
                Do
                      iWhiteRow = 1
                      .GetText .ColLetterToNumber("B"), .Row, vSTT
                      .GetText .ColLetterToNumber("C"), .Row, vCTieu
                      .GetText .ColLetterToNumber("L"), .Row, vMST
                      .GetText .ColLetterToNumber("R"), .Row, vTyLe
                      .GetText .ColLetterToNumber("U"), .Row, vThueGTGT
                      If Trim(vSTT) <> "" Or Trim(vCTieu) <> "" Or Trim(vMST) <> "" Or Val(Trim(vTyLe)) <> 0 Or Trim(vThueGTGT) <> "0" Then
                        iWhiteRow = 0
                      End If
                      ' Reset lai cac cellnote
                      ' Ctieu
                       .Col = .ColLetterToNumber("C")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                      ' MST
                       .Col = .ColLetterToNumber("L")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                      ' End reset
                      ' Kiem tra ctieu khong duoc trong
                      If Trim$(vCTieu) = "" And iWhiteRow = 0 Then
                            .Col = .ColLetterToNumber("C")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = errCTieu
                            .BackColor = mErrorColor
                            iFlagCtieu = 1
                     Else
                            .Col = .ColLetterToNumber("C")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                      End If
'                      ' Kiem tra MST khong duoc trang neu nhap cac du lieu khac
'                      If Trim(vMST) = "" And iWhiteRow = 0 Then
'                            .Col = .ColLetterToNumber("L")
'                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                            .CellNote = errCTieu
'                            .BackColor = mErrorColor
'                            iFlagCtieu = 1
'                      Else
                            ' Kiem tra cau truc MST
                        If Trim(vMST) <> vbNullString Then
                          vMSTtemp = vMST
                          vMST = Replace(vMST, " ", "")
                          vMST = Replace(vMST, "-", "")
                           'vMST = Left(vMST, 10)
                          strCheck = ""
                             If Len(vMST) = 10 Or Len(vMST) = 13 Then
                                  If Len(vMST) = Len(vMSTtemp) Or Mid$(vMSTtemp, 11, 1) = "-" Then
                                      strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                                  End If
                             End If
                             If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                                  .Col = .ColLetterToNumber("L")
                                  .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                  .CellNote = errMST
                                  .BackColor = mErrorColor
                                  iFlagMST = 1
                              End If
                          Else
                              .Col = .ColLetterToNumber("L")
                              .CellNote = ""
                              .BackColor = mNonErrorColor
                          End If
                      'End If
                      
                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                Loop Until .Text = "aa"
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 15, IIf(iFlagCtieu = 1, "0", "1")
                .SetText .ColLetterToNumber("B"), 16, IIf(iFlagMST = 1, "0", "1")
        End If
        
        
        
        
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
    
        CheckErrorMST

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
    For Each xmlNodeValid In TAX_Utilities_New.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
             Case "03_1"
                blnNullValue = True
                Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
                (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 5)
                strValue1 = GetAttribute(xmlCellNode, "Value")
                If (Not IsNullValue(strValue1)) Then
                     blnNullValue = False
                End If
                
                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If
             Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_New.Data(0).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(0).getElementsByTagName("Cell").length - 1)
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
    Dim totalCell  As Long
    Dim countRowReset As Long
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    totalCell = TAX_Utilities_New.Data(0).getElementsByTagName("Cell").length
    countRowReset = (totalCell - 11) / 6
    With fps
        For Each xmlNodeReset In TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            .Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            If (lRow < 37 Or lRow > 44 + countRowReset) And mCurrentSheet = 1 Then
            Else
                .Col = lCol
                .Row = lRow
                
                Select Case .CellType
                    Case CellTypeCheckBox
                        .Text = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                    Case CellTypeComboBox
                        .Text = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                    Case CellTypeNumber
                        .value = 0
                        UpdateCell fps, lCol, lRow, "0"
                    Case Else
                        .value = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                End Select
                'mAdjustData = IIf(IsUpdate = True, IsUpdate, mAdjustData)
                TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
            End If
        Next
    End With
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

Public Sub Prepare_KHBS()
   
    
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

Sub fps_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Long, TipWidth As Long, TipText As String, ShowTip As Boolean)
  
End Sub


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
            Set xmlNode = .nodeFromID("S_37")
            result = "CT28_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_03_TD_TAIN", "getValueTK", Err.number, Err.Description
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
                .GetText .ColLetterToNumber("S"), 37, varTemp
                strCurrentValue = "CT28_" & Trim(varTemp)
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
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT_TT28.xml"
    
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "ID") = GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") Then
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
    SaveErrorLog "cls_03_TD_TAIN", "getMapCT", Err.number, Err.Description
End Function



' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("S"), 37, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function

' Lay ve chi tieu 32 o to khai lan truoc
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
            Set xmlNode = .nodeFromID("S_37")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_03_TD_TAIN", "getValueCTDC", Err.number, Err.Description
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
    For Each xmlNodeValid In TAX_Utilities_New.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
                Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_New.Data(2).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(2).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_New.Data(2).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_New.Data(2).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_New.Data(2).getElementsByTagName("Cell").length > 20 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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
       ' Update lai sheet KHBS
        .Sheet = .SheetCount - 1
        mCurrentSheet = .Sheet
        blCheckS2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheckS2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(.Sheet - 1).getElementsByTagName("Cell")
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
