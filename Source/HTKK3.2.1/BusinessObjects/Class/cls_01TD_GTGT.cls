VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01TD_GTGT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Gia Tri Gia Tang" interface sheets
'this Class is belong to TAX_Business project which will be compline to DLL
'Coder: HTPhuong
Option Explicit

Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Private Const S1ChkRow = 23
Private Const S1ChkCol = "I"
Private Const S1Txt11Row = 24
Private Const S1Txt11Col = "L"
Private Const S4TxtF2row = 2
Private Const S4TxtF2Col = "F"

Private Const Row_MSTDN = 14
Private Const Col_MSTDN = "H"

Private Const THIRD_SHEET = 3
Private Const FOURTH_SHEET = 4
Private Const THIRD_SHEET_CELL1 = "D_20"
Private Const FOURTH_SHEET_CELL1 = "E_19"
Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
'Dim xNodeData0 As MSXML.IXMLDOMNode
Dim xNode As MSXML.IXMLDOMNode
Dim arrData() As String
Public Strloaitk As String
Public loaiKHBS As String
Public StrSolanBosung As String
Public FlagThaydoi As String





'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'no parameter
Public Sub Prepare1()
Dim strComboHien As Variant
With fps
    SetDateFormat fps, 1, 64, .ColLetterToNumber("R"), DDMMYYYY
    ' Sheet KHBS
    SetDateFormat fps, 5, 23, .ColLetterToNumber("BG"), DDMMYYYY
    .Sheet = 1 'To khai GTGT
        .Row = 64
        .Col = .ColLetterToNumber("R")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
    .Sheet = 2 'PL01_2
        .Row = 8
        .Col = .ColLetterToNumber("C")
        .Text = ""
        .CellTag = .CellTag & "~1"
        
        .Row = 13
        .Text = ""
        .CellTag = .CellTag & "~2"
       
        .Row = 18
        .Text = ""
        .CellTag = .CellTag & "~3"
           
    .Sheet = 5 'To khai TD GTGT
        .Row = 23
        .Col = .ColLetterToNumber("BG")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
   
    SetDateFormat fps, 2, 8, .ColLetterToNumber("E"), DDMMYYYY
    SetDateFormat fps, 2, 13, .ColLetterToNumber("E"), DDMMYYYY
    SetDateFormat fps, 2, 18, .ColLetterToNumber("E"), DDMMYYYY
    
    LockCellBySheet
End With

End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'no parameter
Public Sub Prepare2()
    Dim i As Long, DataText As Variant
    Dim arrCellTag() As String
    Dim ArrRow As Variant, ArrCol As Variant
    Dim arrRowChange As Variant, arrColChange As Variant
    Dim intCtrl As Integer
    Dim strId As Variant
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim tempCurrSheet As Integer
    With fps

    mblKhoiDauToKhai = True
        .EventEnabled(EventAllEvents) = True
        tempCurrSheet = mCurrentSheet
        .Sheet = 1 'To khai GTGT
            .ActiveSheet = .Sheet
            mCurrentSheet = 1
            'CellChange .ColLetterToNumber(S1ChkCol), S1ChkRow
            'LockCellBySheet
            'CellLock .ColLetterToNumber(S1ChkCol), S1ChkRow
            mblKhoiDauToKhai = False
            'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("C")
            .Row = 59
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 59
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 59
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("C")
            .Row = 59
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 59
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 59
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
            
'        .Sheet = 2 'PL2
'             .ActiveSheet = .Sheet
'                .Row = 1
'                .Col = .ColLetterToNumber("C")
'                .CellType = CellTypeButton
'                .TypeButtonTextColor = RGB(0, 0, 255)
'                .TypeButtonColor = &H8000000F
'                .TypeButtonAlign = TypeButtonAlignRight
'                .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")

        .ActiveSheet = 1
        mCurrentSheet = tempCurrSheet
         .EventEnabled(EventAllEvents) = True
    End With
    
  
    
End Sub

'inner class event: Change
Public Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
Dim str43Previous As String
Dim varTemp As Variant
Dim xmlNodeH15 As MSXML.IXMLDOMNode
Dim strId As Variant
Dim iCol As Long, iRow As Long
Dim cCol As Long, cRow As Long
Dim cValue As Variant, temp As Variant
Dim cValueG As Variant
Dim cTagF As String
Dim cValueF As Variant
Dim i As Integer



Dim dCT25 As Double

With fps
'    If .ActiveSheet = 1 Then
'        'Check value of 11
'        GetCellSpan fps, Col, Row
'        If Col = .ColLetterToNumber("L") And Row = 24 Then
'            .GetText Col, Row, varTemp
'            If GetLast43Previous(str43Previous) Then
'                If CStr(varTemp) <> str43Previous Then
'                    .Visible = True
'                    DisplayMessage "0060", msOKOnly, miWarning
'                    fps.SetFocus
'                End If
'            End If
'        End If
'    End If

         ' lay gia tri cua sheet 01-2, 01-4A, 01-4B
'    If .ActiveSheet = 3 Then
'        If Col = .ColLetterToNumber("N") And Row = 2 Then
'            .GetText Col, Row, temp
'            dCT25 = dCT25 + Val(temp)
'            .Sheet = 5
'            .GetText .ColLetterToNumber("D"), 20, temp
'            dCT25 = dCT25 + Val(temp)
'            .Sheet = 6
'            .GetText .ColLetterToNumber("E"), 22, temp
'            dCT25 = dCT25 + Val(temp)
'            .Sheet = 1
'            .SetText .ColLetterToNumber("L"), 28, dCT25
'            UpdateCell fps, .ColLetterToNumber("L"), 28, dCT25
'        End If
'    End If
'
'    If .ActiveSheet = 5 Then
'        If Col = .ColLetterToNumber("D") And Row = 20 Then
'            .GetText Col, Row, temp
'            dCT25 = dCT25 + Val(temp)
'            .Sheet = 3
'            .GetText .ColLetterToNumber("N"), 2, temp
'            dCT25 = dCT25 + Val(temp)
'            .Sheet = 6
'            .GetText .ColLetterToNumber("E"), 22, temp
'            dCT25 = dCT25 + Val(temp)
'            .Sheet = 1
'            .SetText .ColLetterToNumber("L"), 28, dCT25
'            UpdateCell fps, .ColLetterToNumber("L"), 28, dCT25
'        End If
'    End If
'
'    If .ActiveSheet = 6 Then
'        If Col = .ColLetterToNumber("E") And Row = 22 Then
'            .GetText Col, Row, temp
'            dCT25 = dCT25 + Val(temp)
'            .Sheet = 3
'            .GetText .ColLetterToNumber("N"), 2, temp
'            dCT25 = dCT25 + Val(temp)
'            .Sheet = 5
'            .GetText .ColLetterToNumber("D"), 20, temp
'            dCT25 = dCT25 + Val(temp)
'            .Sheet = 1
'            .SetText .ColLetterToNumber("L"), 28, dCT25
'            UpdateCell fps, .ColLetterToNumber("L"), 28, dCT25
'        End If
'    End If
    
'    ' set gia tri cho chi tieu 31
'    If .ActiveSheet = 2 Then
'        If Col = .ColLetterToNumber("N") And Row = 3 Then
''            .GetText Col, Row, temp
''            .Sheet = 1
''            .SetText .ColLetterToNumber("L"), 33, temp
''            UpdateCell fps, .ColLetterToNumber("L"), 33, temp
'                        ' chi tieu 33
'             .Sheet = 1
'            .Col = .ColLetterToNumber("L")
'            .Row = 34
'            .Formula = "'01_1'!N3"
'        End If
'    End If
'
'    ' set gia tri cho chi tieu 33
'    If .ActiveSheet = 2 Then
'        If Col = .ColLetterToNumber("N") And Row = 3 Then
'            .GetText Col, Row, temp
'            .Sheet = 1
'            .SetText .ColLetterToNumber("L"), 34, temp
'            UpdateCell fps, .ColLetterToNumber("L"), 34, temp
'        End If
'    End If
    
 mblKhoiDauToKhai = False
    
End With
    'CellChange Col, Row
    'CellChange1 Col, Row
    'LockCellBySheet
End Sub

'inner class event: ButtonClicked
Public Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim lResult As Integer
    Dim blTemp As Boolean
    
    fps.EventEnabled(EventAllEvents) = False
    
'    If Col = fps.ColLetterToNumber(S1ChkCol) And Row = S1ChkRow Then
'       If mblKhoiDauToKhai = False Then
'            fps.Sheet = fps.ActiveSheet
'            mCurrentSheet = fps.Sheet
'            fps.Row = Row
'            fps.Col = Col
'            blTemp = TAX_Utilities_New.AdjustData(fps.ActiveSheet - 1)
'            If fps.value = 1 Then
'                lResult = DisplayMessage("0034", msYesNo, miQuestion, , mrNo)
'                If lResult = 3 Then '<>6
'                    fps.value = 0
'                    UpdateCell fps, fps.Col, fps.Row, IIf(fps.value = 1, "x", vbNullString) 'fps.Value
'                    TAX_Utilities_New.AdjustData(fps.ActiveSheet - 1) = blTemp
'                    fps.EventEnabled(EventAllEvents) = True
'                    Exit Sub
'                Else
'                    UpdateCell fps, fps.Col, fps.Row, IIf(fps.value = 1, "x", vbNullString) 'fps.Value
'                    TAX_Utilities_New.AdjustData(fps.ActiveSheet - 1) = True
'                End If
'            ElseIf fps.value = 0 Then
'                UpdateCell fps, fps.Col, fps.Row, IIf(fps.value = 1, "x", vbNullString)
'                TAX_Utilities_New.AdjustData(fps.ActiveSheet - 1) = True
'                ' Khong check se lay lai gia tri cua phu luc len
'                SetValueOCheck
'
'            End If
'        End If
'
'        CellLock Col, Row
'    End If
    fps.EventEnabled(EventAllEvents) = True
    
    'CellChange Col, Row
    
End Sub

Private Sub fps_EditChange(ByVal Col As Long, ByVal Row As Long)
'    CellChange1 Col, Row
End Sub

Private Sub fps_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim i As Integer, iCol As Long, iRow As Long
     With fps
       If KeyCode = vbKeyF6 Then
            iCol = .ActiveCol
            iRow = .ActiveRow
            mblKeyF6 = True
            fps_Change iCol, iRow
            mblKeyF6 = False
        End If
     End With
End Sub
Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, iCol As Long, iRow As Long
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tCol As Long, tRow As Long
Dim strId As Variant, Col7 As Variant, intIndexCombo As Integer
Dim varTemp As Variant
Dim NguoiKyCol As Long, NguoiKyRow As Long
    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        fps.EventEnabled(EventAllEvents) = False
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 2 Then
                mCurrentSheet = 2
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 8
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 8
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "cc"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
                    'CellChange iCol, iRow
                    'CellChange2
                    'CellChange3
                    'LockCellBySheet
            End If
                        
        End If
  
      fps.EventEnabled(EventAllEvents) = True
    End With
    
End Sub


'inner class event: Leavecell
Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim tCol As Long, tRow As Long
    Dim cCol As Long, cRow As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim dCT25 As Double
    Dim temp As Variant
    
    With fps
    mCurrentSheet = .Sheet
    .EventEnabled(EventAllEvents) = False
        If .ActiveSheet = 1 Then
             If Col = .ColLetterToNumber("R") And Row = 64 Then
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
        If .ActiveSheet = 2 Then
                   If Col = .ColLetterToNumber("E") Then
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
        
        If .ActiveSheet = 5 Then
         Dim count1 As Integer
            .Row = 9
            .Col = .ColLetterToNumber("B")
            Do While .value <> "bb"
                count1 = count1 + 1
                .Row = .Row + 1
            Loop
                  If Col = .ColLetterToNumber("BG") And Row = 23 + count1 - 5 Then
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
                .Sheet = .SheetCount - 1
             
        End If
        .EventEnabled(EventAllEvents) = True
End With
    'CellChange Col, Row
    'LockCellBySheet
End Sub

'When cell's data is changed or a control at a cell is handled, This sub is called
'parameter Col: the column of cell that rise event
'parameter Row: the row of cell that rise event
'parameter f: the flag to mark some special option
'**: when this sub is on executing, all events of fpspread will be disabled until the end of the sub
Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
    Dim i As Integer
    Dim a As Double, b As Double, c As Double
    Dim ArrRow As Variant, ArrCol As Variant
    Dim arrCellTag() As String
    Dim cCol As Long, cRow As Long
    Dim varTemp As Variant
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim iCol  As Long, iRow As Long
    Dim strValue1  As Variant, strValue2 As Variant
With fps
    'Because of unclearly-rised events, this function is here for turning it off
    'at the end of this sub, this event will be turned on
    .EventEnabled(EventAllEvents) = False

    
    ' Kiem tra neu co phu luc moi set cong thuc tong hop so lieu tu tinh
    Dim tempCurrSheet As Integer
    Dim blCheck_S1 As Boolean
    Dim blCheck_S4A As Boolean
    blCheck_S1 = IIf(TAX_Utilities_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S4A = IIf(TAX_Utilities_New.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    
    .EventEnabled(EventAllEvents) = True
End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    Dim i As Long
    Dim vDoanhSo As Variant, vGTGT As Variant
    Dim vTarget15 As Variant, vTarget16 As Variant
    Dim vTarget25 As Variant, vTarget26 As Variant
    Dim vTarget17 As Variant, vTarget18 As Variant, vTarget19 As Variant, vTarget20 As Variant, vTarget21 As Variant, vTarget22 As Variant, vTarget23 As Variant, vTarget24 As Variant
    Dim vTarget38 As Variant, vTarget39 As Variant
    Dim vKyHieuHD As Variant, vSoHD As Variant, vNgayHD As Variant, vTennguoimua As Variant, vMST As Variant
    
    Dim vErrMsgKyHieuHD As Variant, vErrMsgSoHD As Variant, vErrMsgNgayHD As Variant, vErrMsgTennguoimua As Variant, vErrMsgMST As Variant
    Dim vErrMsg As Variant
    
    Dim iFlag5 As Integer, iFlag6 As Integer
    Dim iFlagvKyHieuHD As Integer, iFlagvSoHD As Integer, iFlagvNgayHD As Integer, iFlagvTennguoimua As Integer, iFlagvMst As Integer, iFlagvMst_PL6 As Integer
    Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    Dim iCurrentSheet As Integer, strCheck As String
    Dim iWhiteRow As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer
    Dim arrCellTag() As String
    Dim strCellTag As String
    Dim strFormula As String
    
    'PL 3
    Dim vCTieu, vSTT, vTyLe, vThueGTGT, vMSTtemp As Variant
    Dim errMST, errCTieu As Variant
    Dim iFlagMST As Integer, iFlagCtieu As Integer
    
    Dim flagTenCSPT As Integer
    Dim flagDiaChi As Integer
    Dim vErrorTenCSPT, vErrorDiaChi As Variant
    
    iFlag5 = 0
    iFlag6 = 0
    iFlagvKyHieuHD = 0
    iFlagvSoHD = 0
    iFlagvNgayHD = 0
    iFlagvTennguoimua = 0
    iFocusFlag = 0
    
    flagTenCSPT = 0
    flagDiaChi = 0
    
    Dim vErrMsgThuesuat As Variant, vThuesuat As Variant
    Dim iFlagvThuesuat As Integer
    

    Dim iflagSoChungTu As Integer
    Dim iflag_pl_3 As Integer
    Dim iflagNgayNop As Integer
    iflag_pl_3 = 0
    iflagSoChungTu = 0
    iflagNgayNop = 0
    Dim vErrorSoChungTu As Variant
    Dim vError_pl_3 As Variant
    Dim vErrorNgayNop As Variant
    ' end pl5
    Dim blCheck_S3 As Boolean, blCheck_S4 As Boolean
    
    
    Dim blCheck_S2 As Boolean
     With fps
            iCurrentSheet = .Sheet
            .Sheet = .SheetCount
            CheckErrorMST
            
            .Sheet = 1
        
             .Sheet = .SheetCount
             .GetText .ColLetterToNumber("E"), 21, vErrMsgKyHieuHD
             .GetText .ColLetterToNumber("E"), 22, vErrMsgSoHD
             .GetText .ColLetterToNumber("E"), 23, vErrMsgNgayHD
             .GetText .ColLetterToNumber("E"), 24, vErrMsgTennguoimua
             .GetText .ColLetterToNumber("E"), 25, vErrMsgMST
             .GetText .ColLetterToNumber("E"), 16, vErrMsgThuesuat
             .GetText .ColLetterToNumber("E"), 27, errCTieu
             .GetText .ColLetterToNumber("E"), 28, errMST
             
              
                
            .Sheet = 2
            blCheck_S3 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
            If blCheck_S3 = True Then
                 For i = 1 To .MaxRows
                     iWhiteRow = 1
                     .Row = i
                     .Col = .ColLetterToNumber("C")
                     arrCellTag = Split(.CellTag, "~")
                     If UBound(arrCellTag) > 0 Then
                     strCellTag = arrCellTag(1)
                     .GetText .ColLetterToNumber("C"), .Row, vKyHieuHD
                     .GetText .ColLetterToNumber("D"), .Row, vSoHD
                     .GetText .ColLetterToNumber("E"), .Row, vNgayHD
                     .GetText .ColLetterToNumber("F"), .Row, vTennguoimua
                     .GetText .ColLetterToNumber("I"), .Row, vDoanhSo
                     .GetText .ColLetterToNumber("K"), .Row, vGTGT
                     .GetText .ColLetterToNumber("G"), .Row, vMST
                     .GetText .ColLetterToNumber("J"), .Row, vThuesuat
             
                     If Trim(vKyHieuHD) <> "" Or Trim(vSoHD) <> "" Or Trim(vNgayHD) <> "" Or Trim(vTennguoimua) <> "" Or Trim(vDoanhSo) <> "0" Or Trim(vGTGT) <> "0" Then iWhiteRow = 0
                         Select Case strCellTag
                             Case "1", "2", "3"
                                 If Trim(vKyHieuHD) = "" And iWhiteRow = 0 Then
                                     .Col = .ColLetterToNumber("C")
                                     .CellNote = vErrMsgKyHieuHD  'static
                                     .BackColor = mAlertColor
                                     iFlagvKyHieuHD = 1
                                     
                                 Else
                                     .Col = .ColLetterToNumber("C")
                                     .CellNote = ""
                                     .BackColor = mNonErrorColor
                                 End If
                                 If Trim(vSoHD) = "" And iWhiteRow = 0 Then
                                     .Col = .ColLetterToNumber("D")
                                     .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                     .CellNote = vErrMsgSoHD  'static
                                     .BackColor = mErrorColor
                                     iFlagvSoHD = 1
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
                                 If Trim(vNgayHD) = "" And iWhiteRow = 0 Then
                                     .Col = .ColLetterToNumber("E")
                                     .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                     .CellNote = vErrMsgNgayHD  'static
                                     .BackColor = mErrorColor
                                     iFlagvNgayHD = 1
                                     
                                     'capture the first error cell
                                     If iFocusFlag = 0 Then
                                         strFocusSheetName = "'" & .SheetName & "'"
                                         strFocusCol = str(.Col)
                                         strFocusRow = str(.Row)
                                         iFocusFlag = 1
                                     End If
                                 ElseIf Trim(vNgayHD) <> "" And iWhiteRow = 0 Then
                                  If CInt(Right(vNgayHD, 4)) > CInt(TAX_Utilities_New.Year) Then
                                         .Col = .ColLetterToNumber("E")
                                         .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                         .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
                                         .BackColor = mErrorColor
                                         iFlagvNgayHD = 1
                                         If iFocusFlag = 0 Then
                                             strFocusSheetName = "'" & .SheetName & "'"
                                             strFocusCol = str(.Col)
                                             strFocusRow = str(.Row)
                                             iFocusFlag = 1
                                         End If
                                      ElseIf CInt(Right(vNgayHD, 4)) = CInt(TAX_Utilities_New.Year) And CInt(Right("00" & Mid(vNgayHD, 4, 2), 2)) > CInt(TAX_Utilities_New.Month) Then
                                         .Col = .ColLetterToNumber("E")
                                         .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                         .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
                                         .BackColor = mErrorColor
                                         iFlagvNgayHD = 1
                                         If iFocusFlag = 0 Then
                                             strFocusSheetName = "'" & .SheetName & "'"
                                             strFocusCol = str(.Col)
                                             strFocusRow = str(.Row)
                                             iFocusFlag = 1
                                         End If
                                      ElseIf (CheckValidNgayHD(CStr(vNgayHD)) = False) And strCellTag <> "2" Then
                                         .Col = .ColLetterToNumber("E")
                                         .CellNote = GetAttribute(GetMessageCellById("0124"), "Msg") 'static
                                         .BackColor = mAlertColor
                                          If iFocusFlag = 0 Then
                                             strFocusSheetName = "'" & .SheetName & "'"
                                             strFocusCol = str(.Col)
                                             strFocusRow = str(.Row)
                                             iFocusFlag = 1
                                           End If
                                      Else
                                         .Col = .ColLetterToNumber("E")
                                         .CellNote = ""
                                         .BackColor = mNonErrorColor
                                     End If
                                 Else
                                     .Col = .ColLetterToNumber("E")
                                     .CellNote = ""
                                     .BackColor = mNonErrorColor
                                 End If
                                 If Trim(vTennguoimua) = "" And iWhiteRow = 0 Then
                                     .Col = .ColLetterToNumber("F")
                                     .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                     .CellNote = vErrMsgTennguoimua  'static
                                     .BackColor = mErrorColor
                                     iFlagvTennguoimua = 1
                                     
                                     'capture the first error cell
                                     If iFocusFlag = 0 Then
                                         strFocusSheetName = "'" & .SheetName & "'"
                                         strFocusCol = str(.Col)
                                         strFocusRow = str(.Row)
                                         iFocusFlag = 1
                                     End If
                                 Else
                                     .Col = .ColLetterToNumber("F")
                                     .CellNote = ""
                                     .BackColor = mNonErrorColor
                                 End If
                      'Check MST
                                 If Trim(vMST) <> vbNullString Then
                                 vMST = Replace(vMST, " ", "")
                                 vMST = Replace(vMST, "-", "")
                                  'vMST = Left(vMST, 10)
                                 strCheck = ""
                                    If Len(vMST) = 10 Or Len(vMST) = 13 Then
                                         strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                                    End If
                                    If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                                         .Col = .ColLetterToNumber("G")
                                         .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                         .CellNote = vErrMsgMST   'static
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
                                 Else
                                     .Col = .ColLetterToNumber("G")
                                     .CellNote = ""
                                     .BackColor = mNonErrorColor
                                 End If
                                 
                                 If CDbl(vThuesuat) = 0 Or CDbl(vThuesuat) = 5 Or CDbl(vThuesuat) = 10 Then
                                     .Col = .ColLetterToNumber("J")
                                     .CellNote = ""
                                     .BackColor = mNonErrorColor
                                 Else
                                     .Col = .ColLetterToNumber("J")
                                     .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                     .CellNote = vErrMsgThuesuat   'static
                                     .BackColor = mErrorColor
                                     iFlagvThuesuat = 1
                                     'capture the first error cell
                                     If iFocusFlag = 0 Then
                                         strFocusSheetName = "'" & .SheetName & "'"
                                         strFocusCol = str(.Col)
                                         strFocusRow = str(.Row)
                                         iFocusFlag = 1
                                     End If
                                 End If
                                 
                                 ' kiem tra nguong cua thue GTGT
                                 If strCellTag = "1" Or strCellTag = "2" Or strCellTag = "3" Then
                                     If Trim(vDoanhSo) <> "0" And Trim(vGTGT) <> "0" Then
                                         If (Abs(Round((vDoanhSo * CDbl(vThuesuat) / 100), 0) - Round(vGTGT, 0)) > 100000) Or (Abs(Round((vDoanhSo * CDbl(vThuesuat) / 100), 0) - Round(vGTGT, 0)) > Round(Abs(vDoanhSo) * CDbl(vThuesuat) / 100 * 0.0001, 0)) Then
                                             .Col = .ColLetterToNumber("K")
                                             .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                             .CellNote = GetAttribute(GetMessageCellById("0118"), "Msg") 'static
                                             .BackColor = mAlertColor
                                         Else
                                             .Col = .ColLetterToNumber("K")
                                             .CellNote = ""
                                             .BackColor = mNonErrorColor
                                         End If
             
             
                                     End If
                                 End If
                                 
                         End Select
                 
                    End If
                 Next
             End If
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 21, IIf(iFlagvKyHieuHD = 1 And blCheck_S3 = True, "0", "1")
                .SetText .ColLetterToNumber("B"), 22, IIf(iFlagvSoHD = 1 And blCheck_S3 = True, "0", "1")
                .SetText .ColLetterToNumber("B"), 23, IIf(iFlagvNgayHD = 1 And blCheck_S3 = True, "0", "1")
                .SetText .ColLetterToNumber("B"), 24, IIf(iFlagvTennguoimua = 1 And blCheck_S3 = True, "0", "1")
                .SetText .ColLetterToNumber("B"), 25, IIf(iFlagvMst = 1 And blCheck_S3 = True, "0", "1")
                .SetText .ColLetterToNumber("B"), 16, IIf(iFlagvThuesuat = 1 And blCheck_S3 = True, "0", "1")
          ' Check sheet 3
          .Sheet = 3
            blCheck_S4 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
            If blCheck_S4 = True Then
                .Col = .ColLetterToNumber("B")
                .Row = 41
                Do
                      iWhiteRow = 1
                      ' clear cell note
                        .Col = .ColLetterToNumber("L")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                      ' end
                      
                      .GetText .ColLetterToNumber("B"), .Row, vSTT
                      .GetText .ColLetterToNumber("C"), .Row, vCTieu
                      .GetText .ColLetterToNumber("L"), .Row, vMST
                      .GetText .ColLetterToNumber("R"), .Row, vTyLe
                      .GetText .ColLetterToNumber("U"), .Row, vThueGTGT
                      If Trim(vSTT) <> "" Or Trim(vCTieu) <> "" Or Trim(vMST) <> "" Or Val(Trim(vTyLe)) <> 0 Or Trim(vThueGTGT) <> "0" Then
                        iWhiteRow = 0
                      End If
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
                      ' Kiem tra cau truc MST
                      If Trim(vMST) <> vbNullString Then
                        vMSTtemp = vMST
                        vMST = Replace(vMST, " ", "")
                        vMST = Replace(vMST, "-", "")
                         'vMST = Left(vMST, 10)
                        strCheck = ""
                        If Len(vMST) = 10 Or Len(vMST) = 13 Then
                                If Len(vMST) = Len(vMSTtemp) Or Mid(vMSTtemp, 11, 1) = "-" Then
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
                      
                      
                     .Col = .ColLetterToNumber("B")
                     .Row = .Row + 1
                Loop Until .Text = "aa"
                .Sheet = .SheetCount
                .SetText .ColLetterToNumber("B"), 27, IIf(iFlagCtieu = 1, "0", "1")
                .SetText .ColLetterToNumber("B"), 28, IIf(iFlagMST = 1, "0", "1")
                
            End If
                
            If iFocusFlag = 1 Then
                ' setcell for focus
                
                strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
                .Col = .ColLetterToNumber("C")
                .Row = 21
                .Formula = strFormula
                .Row = 22
                .Formula = strFormula
                .Row = 23
                .Formula = strFormula
                .Row = 24
                .Formula = strFormula
                .Row = 25
                .Formula = strFormula
                .Row = 16
                .Formula = strFormula
                
            End If
             
            .Sheet = iCurrentSheet
    End With
End Sub
'This function is used to support report
'some special data that we want to save to data files
Public Sub SaveExceptionData()
                  
End Sub

'Calling this function whenever we don't need this class anymore.
'It will do some special task with special Error/ special Data
'no parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim vTemp As Variant
    Dim iCol As Long, iRow As Long, i As Long
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        Set xmlCellNode = Nothing
        CheckDynamicError 'Set Exception Error on cells of interface
            
        UpdateSheet ' set gia tri ma hinh xuong XML
            
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
        
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
        .Col = .ColLetterToNumber("D")
        .Row = 3
        .CellNote = ""
        '.BackColor = mFormColor
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

'******************************
'Description: Set active properties in validity node
'Author: ThanhDX
'Date: 24/11/2005

'******************************
Public Sub SetActiveSheet() 'ByRef xmlNodeValid As MSXML.IXMLDOMNode)
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    Dim strValue3 As String
    Dim strValue4 As Variant
    Dim strValue5 As Variant
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_New.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
'            Case "01_2"
'                blnNullValue = True
'                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
'                    (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'                Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
'                    (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 2)
'                    strValue2 = GetAttribute(xmlCellNode, "Value")
'                'Set active voi dong 5
'                fps.Sheet = 2
'                fps.GetText fps.ColLetterToNumber("N"), 1, strValue4
'                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Or Val(strValue4) <> 0 Then
'                            blnNullValue = False
'                    End If
'
'                    If blnNullValue Then
'                        SetAttribute xmlNodeValid, "Active", "0"
'                    End If
               Case "01_2"
                blnNullValue = True

                For Each xmlCellNode In TAX_Utilities_New.Data(1).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
'                    If lCol = fps.ColLetterToNumber("D") Or _
'                       lCol = fps.ColLetterToNumber("E") Then
                        
                        strValue = GetAttribute(xmlCellNode, "Value")
                        If Not IsNullValue(strValue) Then
                            blnNullValue = False
                            Exit For
                        End If
'                    End If
                Next
                
                    If blnNullValue Then
                        SetAttribute xmlNodeValid, "Active", "0"
                    End If
                    
               Case "01_3"
                blnNullValue = True

                For Each xmlCellNode In TAX_Utilities_New.Data(2).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
'                    If lCol = fps.ColLetterToNumber("D") Or _
'                       lCol = fps.ColLetterToNumber("E") Then
                        
                        strValue = GetAttribute(xmlCellNode, "Value")
                        If Not IsNullValue(strValue) Then
                            blnNullValue = False
                            Exit For
                        End If
'                    End If
                Next
                
                    If blnNullValue Then
                        SetAttribute xmlNodeValid, "Active", "0"
                    End If
                    
               Case "01_4A"
                blnNullValue = True

                For Each xmlCellNode In TAX_Utilities_New.Data(3).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
'                    If lCol = fps.ColLetterToNumber("D") Or _
'                       lCol = fps.ColLetterToNumber("E") Then
                        strValue = GetAttribute(xmlCellNode, "Value")
                        If Not IsNullValue(strValue) Then
                            blnNullValue = False
                            Exit For
                        End If
'                    End If
                Next
                
                    If blnNullValue Then
                        SetAttribute xmlNodeValid, "Active", "0"
                    End If
                
             Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_New.Data(4).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(4).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_New.Data(4).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(4).getElementsByTagName("Cell").length - 4)
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

'******************************
'Description: Reset all error satus in sheet contain dynamic error.
'Author: ThanhDX
'Date: 31/12/2005

'******************************
Public Sub ResetErrorCells()
    Dim lRow As Long, lCtrl As Long
    Dim lBColumnColor As Long, lCColumnColor As Long
    Dim lDColumnColor As Long, lEColumnColor As Long
    Dim lFColumnColor As Long
    Dim intCurrentSheet As Integer
    
    With fps
        intCurrentSheet = .Sheet
        
        .Sheet = 2 'Sheet contain dynamic data
        For lCtrl = 1 To .MaxRows
            .Col = .ColLetterToNumber("C")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If

            .Col = .ColLetterToNumber("D")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If

            .Col = .ColLetterToNumber("E")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If

            .Col = .ColLetterToNumber("F")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If
            .Col = .ColLetterToNumber("G")
            .Row = lCtrl
            If Not .Lock Then
                .CellNote = ""
                .BackColor = vbWhite
            End If

        Next lCtrl
                
        .Sheet = intCurrentSheet
    End With
End Sub


'*******************************************************
'Description: ResetData procedure reset specified cells
'Author:ThanhDX
'Date:23/01/2006

'*******************************************************
Public Sub ResetData()
    
  
LockCellBySheet
End Sub

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()
      
End Sub

Public Sub Prepare_KHBS()
    
End Sub

Private Function GetValueKHBS(StrKytinhthue As String, StrNgayKHBS As String) As Double
 
End Function

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function


Private Function CheckValidNgayHD(vdate As String) As Boolean
  
   Dim dd As String, mm As String, yyyy As String, dDate As Date, dDateNow As Date
   Dim sDateNow As String
   On Error GoTo e
    dd = Right("00" & Left(vdate, InStr(vdate, "/") - 1), 2)
    mm = Right("00" & Mid(vdate, 4, 2), 2)
    yyyy = Right("0000" & vdate, 4)
    sDateNow = "01/" & TAX_Utilities_New.Month & "/" & TAX_Utilities_New.Year
    dDateNow = Format(CDate(sDateNow), "dd/mm/yyyy")
    dDate = Format(CDate(vdate), "dd/mm/yyyy")
        If CInt(yyyy) > CInt(TAX_Utilities_New.Year) Then
            CheckValidNgayHD = False
            Exit Function
        End If
        If DateAdd("m", 5, dDate) < dDateNow Then
        CheckValidNgayHD = False
            Exit Function
        End If
   CheckValidNgayHD = True
   Exit Function
   
e:
    CheckValidNgayHD = False
 End Function


Public Sub UpdateChangeKHBS()
   
End Sub

Public Sub FinishImport()
Dim i As Integer, iCol As Long, iRow As Long
With fps
fps.EventEnabled(EventAllEvents) = False
            If .ActiveSheet = 2 Then
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 8
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 8
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "cc"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
                    CellChange2
                    CellChange3
            End If
        fps.EventEnabled(EventAllEvents) = True
        End With
    'CellChange iCol, iRow
    'CellChange1 Col, Row
End Sub


Public Sub UpdateSheet()
    Dim varTemp As Variant
    Dim temp As Integer
    Dim blCheck_S2 As Boolean
    Dim blCheck_S8 As Boolean
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    With fps
        .Sheet = 1
        temp = mCurrentSheet
        mCurrentSheet = 1
        For Each xmlCellNodeData In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
            ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
            .GetText lCol, lRow, varTemp
            UpdateCell fps, lCol, lRow, varTemp
        Next

        .Sheet = 2
        mCurrentSheet = 2
        blCheck_S2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(1).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            Next
         End If
        .Sheet = 3
        mCurrentSheet = 3
        blCheck_S2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(2).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
        .Sheet = 4
        mCurrentSheet = 4
        blCheck_S2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(3).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
        ' Update lai sheet KHBS
        .Sheet = .SheetCount - 1
        mCurrentSheet = .Sheet
        blCheck_S2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(.Sheet - 1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
        
        mCurrentSheet = temp
    End With
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

' Ham set lai trang thai cua cac chi tieu khi check chi tieu 21
Sub CellLock(ByVal Col As Long, ByVal Row As Long)

End Sub


' Ham set cac cong thuc
Sub CellChange1(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
  
End Sub

' Ham set lai gia tri cua cong thuc khi tai bang ke 01_1
Sub CellChange2()
 
End Sub
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
                Set xmlCellNode = TAX_Utilities_New.Data(4).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(4).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_New.Data(4).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_New.Data(4).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_New.Data(4).getElementsByTagName("Cell").length > 20 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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
'When cell's data is changed or a control at a cell is handled, This sub is called
'parameter Col: the column of cell that rise event
'parameter Row: the row of cell that rise event
'parameter f: the flag to mark some special option
'**: when this sub is on executing, all events of fpspread will be disabled until the end of the sub
Public Sub LockCellBySheet()
    Dim i As Integer
    Dim a As Double, b As Double, c As Double
    Dim ArrRow As Variant, ArrCol As Variant
    Dim arrCellTag() As String
    Dim cCol As Long, cRow As Long
    Dim varTemp As Variant
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim iCol  As Long, iRow As Long
    Dim strValue1  As Variant, strValue2 As Variant
With fps
    'Because of unclearly-rised events, this function is here for turning it off
    'at the end of this sub, this event will be turned on
    .EventEnabled(EventAllEvents) = False

    
    ' Kiem tra neu co phu luc moi set cong thuc tong hop so lieu tu tinh
    Dim tempCurrSheet As Integer
    Dim blCheck_S1 As Boolean, blCheck_S4A As Boolean
    blCheck_S1 = IIf(TAX_Utilities_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S4A = IIf(TAX_Utilities_New.NodeValidity.childNodes(3).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    tempCurrSheet = mCurrentSheet
    .Sheet = 1
    If blCheck_S1 = True Or blCheck_S4A = True Then
        .Sheet = 1
        ' Chi tieu 26
        .Col = .ColLetterToNumber("S")
        .Row = 46
        .Lock = True
        .Formula = "AB46"
    Else
        .Sheet = 1
        ' Chi tieu 26
        .Col = .ColLetterToNumber("S")
        .Row = 46
        .Lock = False
        .Formula = ""
    End If
            
    .EventEnabled(EventAllEvents) = True
End With
End Sub



' Ham set lai gia tri cua cong thuc khi tai bang ke 01_2
Sub CellChange3()
 
End Sub


Sub SetValueOCheck()
    With fps
        ' Phu luc 01
        ' chi tieu 33
        .Sheet = 1
        .Col = .ColLetterToNumber("L")
        .Row = 34
        .Formula = "L53"
         ' chi tieu 26
        .Sheet = 1
        .Col = .ColLetterToNumber("J")
        .Row = 30
        .Formula = "E54"
         ' chi tieu 29
        .Sheet = 1
        .Col = .ColLetterToNumber("J")
        .Row = 32
        .Formula = "H53"
         ' chi tieu 30
        .Sheet = 1
        .Col = .ColLetterToNumber("J")
        .Row = 33
        .Formula = "I53"
         ' chi tieu 32
        .Sheet = 1
        .Col = .ColLetterToNumber("J")
        .Row = 34
        .Formula = "K53"
         ' chi tieu 31
        .Sheet = 1
        .Col = .ColLetterToNumber("L")
        .Row = 33
        .Formula = "J53"
        ' phu luc 01_2
        ' chi tieu 23
        .Sheet = 1
        .Col = .ColLetterToNumber("J")
        .Row = 27
        .Formula = "D53"
         ' chi tieu 24
        .Sheet = 1
        .Col = .ColLetterToNumber("L")
        .Row = 27
        .Formula = "E53"
         ' chi tieu 25
        .Sheet = 1
        .Col = .ColLetterToNumber("L")
        .Row = 28
        .Formula = "F53"
    End With
End Sub
Public Sub update_01_4A()
    Dim blCheck_S4A As Boolean
    Dim blCheck_S4B As Boolean
    Dim blCheck_S01_1 As Boolean
    Dim blCheck_S01_2 As Boolean
    
    blCheck_S4A = IIf(TAX_Utilities_New.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S4B = IIf(TAX_Utilities_New.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    
    blCheck_S01_1 = IIf(TAX_Utilities_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S01_2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    
    ' Tong hop len bang ke 01_4A
        With fps
'        .Sheet = .ActiveSheet
'        Select Case .Sheet
'            Case 2 ' Phu luc 01_1
                If blCheck_S4A = True And blCheck_S01_1 = True Then
                    .Sheet = 5
                    ' chi tieu tong doanh so ban ra tren phu luc 1
                    .Col = .ColLetterToNumber("D")
                    .Row = 16
                    .Formula = "C24"
                    ' chi tieu T?ng doanh thu hàng hóa, d?ch v? bán ra ch?u thu? GTGT
                    .Sheet = 5
                    .Col = .ColLetterToNumber("D")
                    .Row = 17
                    .Formula = "D24"
                End If
'            Case 3 ' phu luc 01_2
                If blCheck_S4A = True And blCheck_S01_2 = True Then
                    ' chi tieu tong dong 1 cot 10
                    'A1
                    .Sheet = 5
                    .Col = .ColLetterToNumber("D")
                    .Row = 7
                    .Formula = "C23"
                     ' chi tieu A2-04A
                    .Sheet = 5
                    .Col = .ColLetterToNumber("D")
                    .Row = 9
                    .Formula = "D23"
                    'chi tieu tong dong 3 cot 10
                    'A3
                    .Sheet = 5
                    .Col = .ColLetterToNumber("D")
                    .Row = 11
                    .Formula = "E23"
                    .Sheet = 5
                     ' chitieu B4
                    .Col = .ColLetterToNumber("D")
                    .Row = 19
                    .Formula = "E24"
                End If
                
                If blCheck_S4A = True Then
                    ' chi tieu 25
                    .Sheet = 1
                    .Col = .ColLetterToNumber("L")
                    .Row = 28
                    .Formula = "F53"
                End If
'         End Select
    End With
End Sub

Public Sub reset_01_4A()
    ' Tong hop len bang ke 01_4A
        With fps
        .Sheet = 5
        ' chi tieu tong doanh so ban ra tren phu luc 1
        .Col = .ColLetterToNumber("D")
        .Row = 16
        .Formula = ""
        ' chi tieu T?ng doanh thu hàng hóa, d?ch v? bán ra ch?u thu? GTGT
        .Sheet = 5
        .Col = .ColLetterToNumber("D")
        .Row = 17
        .Formula = ""
        ' chi tieu tong dong 1 cot 10
        'A1
        .Sheet = 5
        .Col = .ColLetterToNumber("D")
        .Row = 7
        .Formula = ""
         ' chi tieu A2-04A
        .Sheet = 5
        .Col = .ColLetterToNumber("D")
        .Row = 9
        .Formula = ""
        'chi tieu tong dong 3 cot 10
        'A3
        .Sheet = 5
        .Col = .ColLetterToNumber("D")
        .Row = 11
        .Formula = ""
        .Sheet = 5
         ' chitieu B4
        .Col = .ColLetterToNumber("D")
        .Row = 19
        .Formula = ""
        ' chi tieu 25
        .Sheet = 1
        .Col = .ColLetterToNumber("L")
        .Row = 28
        .Formula = ""
    End With
End Sub

' Lay ve danh sach cac chi tieu dieu chinh KHBS
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
    SaveErrorLog "cls_01TD_GTGT", "getMapCT", Err.number, Err.Description
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
            Set xmlNode = .nodeFromID("S_47")
            result = "CT27_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01TD_GTGT", "getValueTK", Err.number, Err.Description
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
                .GetText .ColLetterToNumber("S"), 47, varTemp
                strCurrentValue = "CT27_" & Trim(varTemp)
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

' Lay ve chi tieu D o to khai lan truoc
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
            Set xmlNode = .nodeFromID("S_47")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01TD_GTGT", "getValueCTDC", Err.number, Err.Description
End Function


' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("S"), 47, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function
