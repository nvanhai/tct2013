VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_03GTGT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
'Dim xNodeData0 As MSXML.IXMLDOMNode
Dim xNode As MSXML.IXMLDOMNode
Public Strloaitk As String
Dim arrData() As String
Public loaiKHBS As String
Public FlagThaydoi As String
Public StrSolanBosung As String

Private Const Row_MSTDN = 6
Private Const Col_MSTDN = "F"


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
        SetDateFormat fps, 1, 58, .ColLetterToNumber("O"), DDMMYYYY
        .Sheet = 1 'To khai GTGT
        .Row = 58
        .Col = .ColLetterToNumber("O")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        SetDateFormat fps, .SheetCount - 1, 23, .ColLetterToNumber("BG"), DDMMYYYY
        .Sheet = .SheetCount - 1 'KHBS
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
    Dim i As Long
    Dim intCtrl As Integer
    Dim strId As Variant
    Dim xmlNode As MSXML.IXMLDOMNode
    
    With fps
        'vttoan them trang thai to khai
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("C")
            .Row = 60
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 60
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 60
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("C")
            .Row = 60
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 60
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 60
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        .EventEnabled(EventAllEvents) = True
        'end
        .Sheet = 1
        mCurrentSheet = .Sheet
        
        
'        If TAX_Utilities_New.Month <> "01" Then
'            setMP10Target
'        Else
'            .Col = .ColLetterToNumber("M")
'            .Row = 10
'            .Lock = True
'            .value = 0
'            UpdateCell fps, .Col, .Row, .value
'            .Col = .ColLetterToNumber("P")
'            .Row = 10
'            .Lock = True
'            .value = 0
'            UpdateCell fps, .Col, .Row, .value
'
'
'            .SetActiveCell .ColLetterToNumber("M"), 9
'            .Col = .ColLetterToNumber("M")
'            .Row = 13
'            UpdateCell fps, .Col, .Row, .value
'            .Col = .ColLetterToNumber("P")
'            .Row = 13
'            UpdateCell fps, .Col, .Row, .value
'        End If
'        CellChange fps.ColLetterToNumber("M"), 13
'        CellChange fps.ColLetterToNumber("P"), 13
        
        If Not TAX_Utilities_New.DataKHBS Is Nothing Then
            
        
        Set xmlNode = TAX_Utilities_New.Data(0).nodeFromID(GetCellID(fps, fps.ColLetterToNumber("O"), 44))
        vThuephainop = GetAttribute(xmlNode, "Value")
       
        
        Set xmlNode = Nothing
        
          .Sheet = .SheetCount - 1 'KHBS
            Dim cCol As Long, cRow As Long
          '  .ActiveSheet = .Sheet
            Set xmlNode = TAX_Utilities_New.Data(.SheetCount - 2).childNodes(2).childNodes(2).lastChild.lastChild
            mCurrentSheet = .SheetCount - 1
            fps.Sheet = mCurrentSheet
            If GetAttribute(xmlNode, "Value") <> "" Then
                If GetAttribute(xmlNode, "Value") < 0 Then
                    ParserCellID fps, GetAttribute(xmlNode, "CellID"), cCol, cRow
                    .Col = .ColLetterToNumber("BE")
                    .Row = cRow + 3
                    .value = 0
                    UpdateCell fps, .Col, .Row, .value
                End If
            End If
            i = 1
            .Col = .ColLetterToNumber("B")
            .Row = 9
            Do
              .GetText .ColLetterToNumber("BC"), .Row, strId
              If strId <> "" Then
                .Col = .ColLetterToNumber("BE")
                .CellTag = .Row & "~" & Trim(Mid(strId, 100, 30))
              End If
              .Col = .ColLetterToNumber("B")
              .Row = .Row + 1
            Loop Until .Text = "bb"
        
            Set xmlNode = TAX_Utilities_New.Data(.SheetCount - 2).childNodes(2).firstChild
            SetAttribute xmlNode, "DateKHBS", TAX_Utilities_New.DateKHBS
            SetAttribute xmlNode, "loaiKHBS", loaiKHBS
            If loaiKHBS = "frmKHBS_BS" Then
                .SetText .ColLetterToNumber("BD"), 5, GetAttribute(GetMessageCellById("0115"), "Msg") & _
                GetAttribute(GetMessageCellById("0117"), "Msg") & Mid(TAX_Utilities_New.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_New.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_New.DateKHBS, 5, 4)
            Else
                .SetText .ColLetterToNumber("BD"), 5, GetAttribute(GetMessageCellById("0116"), "Msg") & _
                GetAttribute(GetMessageCellById("0117"), "Msg") & Mid(TAX_Utilities_New.DateKHBS, 1, 2) & "/" & Mid(TAX_Utilities_New.DateKHBS, 3, 2) & "/" & Mid(TAX_Utilities_New.DateKHBS, 5, 4)
            End If
            .EventEnabled(EventAllEvents) = False
            .Sheet = 1
            mCurrentSheet = 1
            .Col = .ColLetterToNumber("O")
            .Row = 56
            .Formula = "TRIM(Header!T5)"
            UpdateCell fps, .Col, .Row, .Text
            .Sheet = 1 'To khai GTGT
            .Row = 58
            .Col = .ColLetterToNumber("O")
            .Text = Format(Date, "dd/mm/yyyy")
            .TypeHAlign = TypeHAlignLeft
            .EventEnabled(EventAllEvents) = True
            ReDim arrData(TAX_Utilities_New.DataKHBS.getElementsByTagName("Cell").length)
            Dim j As Integer
            For Each xNode In TAX_Utilities_New.DataKHBS.getElementsByTagName("Cell")
                arrData(j) = GetAttribute(xNode, "CellID") & ":" & GetAttribute(xNode, "Value")
                j = j + 1
            Next
            Set xNode = Nothing
         End If
    End With
End Sub

Public Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
Dim varTemp As Variant
Dim strId As Variant
Dim iCol As Long, iRow As Long
Dim cCol As Long, cRow As Long
Dim cValue As Variant, temp As Variant
Dim cValueG As Variant
Dim cTagF As String
Dim cValueF As Variant
Dim i As Integer
Dim vM13, vM10 As Variant
Dim vP13, vP10 As Variant
Dim strGiaTriAm As String
Dim xmlNodeH15 As MSXML.IXMLDOMNode
With fps
    If .ActiveSheet = 1 Then
     mCurrentSheet = .ActiveSheet
     fps.EventEnabled(EventAllEvents) = False
           If Col = .ColLetterToNumber("M") And Row = 10 Then
                   .GetText Col, Row, vM10
                   If TAX_Utilities_New.Month <> "01" Then
                       If GetMP10Previous(strGiaTriAm) Then
                           If Val(strGiaTriAm) < 0 Then
                               .Visible = True
                               DisplayMessage "0099", msOKOnly, miWarning
                               fps.SetFocus
                           End If
                           If CStr(vM10) <> strGiaTriAm And Val(strGiaTriAm) >= 0 Then
                               .Visible = True
                               DisplayMessage "0085", msOKOnly, miWarning
                               fps.SetFocus
                           End If
                       End If
                   End If
                   .GetText .ColLetterToNumber("M"), 13, vM13
                   If Val(vM13) <= 0 Then
                       .Col = .ColLetterToNumber("M")
                       .Row = 14
                       .value = 0
                       .value = 0
                       UpdateCell fps, .Col, .Row, .value
                    Else
                       .Col = .ColLetterToNumber("M")
                       .Row = 14
                       UpdateCell fps, .Col, .Row, .value
                   End If
           End If
           
           If Col = .ColLetterToNumber("P") And Row = 10 Then
                   .GetText Col, Row, vP10
                   If TAX_Utilities_New.Month <> "01" Then
                       If GetMP10Previous(strGiaTriAm) Then
                           If Val(strGiaTriAm) < 0 Then
                               .Visible = True
                               DisplayMessage "0099", msOKOnly, miWarning
                               fps.SetFocus
                           End If
                           If CStr(vP10) <> strGiaTriAm And Val(strGiaTriAm) >= 0 Then
                               .Visible = True
                               DisplayMessage "0085", msOKOnly, miWarning
                               fps.SetFocus
                           End If
                       End If
                   End If
                   .GetText .ColLetterToNumber("P"), 13, vP13
                   If Val(vP13) <= 0 Then
                       .Col = .ColLetterToNumber("P")
                       .Row = 14
                       .value = 0
                       .value = 0
                       UpdateCell fps, .Col, .Row, .value
                    Else
                       .Col = .ColLetterToNumber("P")
                       .Row = 14
                       UpdateCell fps, .Col, .Row, .value
                   End If
           End If
        
           If Row = 13 And Col = .ColLetterToNumber("M") Then
               .GetText Col, Row, vM13
               If Val(vM13) <= 0 Then
                   .Col = .ColLetterToNumber("M")
                   .Row = 14
                   .value = 0
                   .Lock = True
                   UpdateCell fps, .Col, .Row, .value
               Else
                  .Col = .ColLetterToNumber("M")
                  .Row = 14
                  .Lock = True
                  UpdateCell fps, .Col, .Row, .value
               End If
           End If
           If Row = 13 And Col = .ColLetterToNumber("P") Then
               .GetText Col, Row, vP13
               If Val(vP13) <= 0 Then
                   .Col = .ColLetterToNumber("P")
                   .Row = 14
                   .value = 0
                   .Lock = True
                   UpdateCell fps, .Col, .Row, .value
               Else
                  .Col = .ColLetterToNumber("P")
                  .Row = 14
                  .Lock = True
                  UpdateCell fps, .Col, .Row, .value
               End If
           End If
       fps.EventEnabled(EventAllEvents) = True
    End If
    If .ActiveSheet = .SheetCount - 1 Then
         fps.EventEnabled(EventAllEvents) = False
        'Check value of 11
        GetCellSpan fps, Col, Row
        If blGiaTriCu And (Col = .ColLetterToNumber("BD") Or Col = .ColLetterToNumber("BE")) Then
                 DisplayMessage "0110", msOKOnly, miWarning
                .Row = RowGiaTriCu
                .Col = .ColLetterToNumber("BD")
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("BE")
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("BF")
                .Text = 0
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("BG")
                .Text = 0
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("BH")
                .Text = 0
                UpdateCell fps, .Col, .Row, .Text
                fps.SetFocus
                fps.EventEnabled(EventAllEvents) = True
                Exit Sub
         End If
         
         'get gia tri da ke khai
        .GetText .ColLetterToNumber("BC"), Row, strId
          If strId <> "" Then
                Dim arr() As String
                 For i = 0 To UBound(arrData)
                     arr() = Split(arrData(i), ":")
                     If arr(0) = Trim(Mid(strId, 100, 20)) Then
                         cValue = arr(1)
                         Exit For
                     End If
                 Next
               .Col = .ColLetterToNumber("BF")
               .Row = Row
               cTagF = .CellTag
               If cTagF <> "" Then
                         For i = 0 To UBound(arrData)
                               arr() = Split(arrData(i), ":")
                               If arr(0) = Right(cTagF, Len(cTagF) - InStr(cTagF, "~")) Then
                                   cValueF = arr(1)
                                   Exit For
                               End If
                           Next
                End If
               .CellTag = .Row & "~" & Trim(Mid(strId, 100, 20))
               .value = cValue
               UpdateCell fps, .Col, .Row, .value
               .Col = .ColLetterToNumber("BG")
               .Row = Row
               UpdateCell fps, .Col, .Row, .value
               .Col = .ColLetterToNumber("BH")
               .Row = Row
               UpdateCell fps, .Col, .Row, .value
                
               If loaiKHBS = "frmKHBS_TT" Then
                    .GetText .ColLetterToNumber("BG"), Row, cValueG
               Else
                    .GetText .ColLetterToNumber("BH"), Row, cValueG
               End If
              
              If cValueG <> "" Then
                      ParserCellID fps, Trim(Mid(strId, 100, 20)), cCol, cRow
                      .Sheet = 1
                      .SetText cCol, cRow, cValueG
                      If loaiKHBS = "frmKHBS_TT" Then
                        UpdateKHBSCell fps, cCol, cRow, cValueG
                      Else
                        UpdateKHBSCell fps, cCol, cRow, cValueG + cValueF
                      End If
              End If
              
' Set lai gia tri cu cua sheet 01
               .Sheet = .SheetCount - 1
               .Col = .ColLetterToNumber("BF")
               .Row = Row
               
               If cTagF <> "" And cTagF <> .CellTag Then
                  If Left(cTagF, InStr(cTagF, "~") - 1) = .Row Then
                   ' .GetText .ColLetterToNumber("F"), Row, temp
                    ParserCellID fps, Right(cTagF, Len(cTagF) - InStr(cTagF, "~")), iCol, iRow
                    UpdateKHBSCell fps, iCol, iRow, cValueF
                    .Sheet = 1
                    If loaiKHBS = "frmKHBS_TT" Then
                        .SetText iCol, iRow, cValueF
                    Else
                        .SetText iCol, iRow, 0
                    End If
                    .Sheet = .SheetCount - 1
                    If loaiKHBS = "frmKHBS_TT" Then
                        .GetText .ColLetterToNumber("BG"), Row, varTemp
                    Else
                        .GetText .ColLetterToNumber("BH"), Row, varTemp
                    End If
                    If varTemp <> "" Then
                        ParserCellID fps, Trim(Mid(strId, 100, 20)), iCol, iRow
                        If loaiKHBS = "frmKHBS_TT" Then
                            UpdateKHBSCell fps, cCol, cRow, varTemp
                        Else
                            UpdateKHBSCell fps, cCol, cRow, varTemp + cValueF
                        End If
                       ' UpdateKHBSCell fps, iCol, iRow, varTemp
                        .Sheet = 1
                        .SetText iCol, iRow, varTemp
                        .Sheet = .SheetCount - 1
                    End If
                   End If
               End If
'Xoa dong
               If mblKeyF6 = True Then
                ParserCellID fps, Right(cTagF, Len(cTagF) - InStr(cTagF, "~")), iCol, iRow
                UpdateKHBSCell fps, iCol, iRow, cValueF
                .Sheet = 1
                .SetText iCol, iRow, 0
                .Sheet = .SheetCount - 1
                .Col = .ColLetterToNumber("BF")
                .Row = Row
                .value = 0
                .Col = .ColLetterToNumber("BG")
                .Row = Row
                .value = 0
                .Col = .ColLetterToNumber("BH")
                .Row = Row
                .value = 0
               End If
               
            .Sheet = 1
            .GetText .ColLetterToNumber("O"), 44, varTemp

            
            .Sheet = .SheetCount - 1
            Set xmlNodeH15 = TAX_Utilities_New.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
                 ParserCellID fps, GetAttribute(xmlNodeH15, "CellID"), cCol, cRow
                .Col = cCol
                .Row = cRow
                If loaiKHBS = "frmKHBS_TT" Then
                    .value = CDbl(varTemp) - CDbl(vThuephainop)
                     UpdateCell fps, .Col, .Row, .value
                Else
                    .value = CDbl(varTemp)
                     UpdateCell fps, .Col, .Row, .value
                End If
                If .value < 0 Then
                    .Col = .ColLetterToNumber("BE")
                    .Row = cRow + 3
                    .value = 0
                    UpdateCell fps, .Col, .Row, 0
                Else
                    .Col = .ColLetterToNumber("BE")
                    .Row = cRow + 3
                    UpdateCell fps, .Col, .Row, .value
                End If
             Set xmlNodeH15 = Nothing
               
          End If
            fps.EventEnabled(EventAllEvents) = True
       End If
     mblKhoiDauToKhai = False
End With
'CellChange Col, Row
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
Dim tCol As Long, tRow As Long
    With fps
        fps.EventEnabled(EventAllEvents) = False
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = .SheetCount - 1 Then
                 fps_Change iCol, iRow
                 fps.EventEnabled(EventAllEvents) = False
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 9
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 9
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 3
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
                 fps.EventEnabled(EventAllEvents) = True
            End If
            
        End If
        fps.EventEnabled(EventAllEvents) = True
    End With

End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim tCol As Long, tRow As Long
    Dim cCol As Long, cRow As Long
    Dim xmlNode As MSXML.IXMLDOMNode
    With fps
        mCurrentSheet = .Sheet
        .EventEnabled(EventAllEvents) = False
        If .ActiveSheet = 1 Then
             If Col = .ColLetterToNumber("O") And Row = 58 Then
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
            
'            .GetText .ColLetterToNumber("L"), 16, varTemp
'            If TAX_Utilities_New.NodeValidity.childNodes(.SheetCount - 2).Attributes.getNamedItem("Active").nodeValue <> "0" Then
'                .Sheet = .SheetCount - 1
'                Set xmlNode = TAX_Utilities_New.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
'                ParserCellID fps, GetAttribute(xmlNode, "CellID"), cCol, cRow
'                .Col = cCol
'                .Row = cRow
'                .Value = CDbl(varTemp) - CDbl(vThuephainop)
'                'mCurrentSheet = .SheetCount - 1
'                UpdateCell fps, .Col, .Row, .Value
'                .Col = .ColLetterToNumber("BE")
'                .Row = .Row + 3
'                UpdateCell fps, .Col, .Row, .Value
'                .Sheet = 1
'                 mCurrentSheet = 1
'            End If
            
        End If
        
        'KHBS
       If .ActiveSheet = .SheetCount - 1 Then
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
 '   CellChange Col, Row
End Sub

'Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
'    Dim xmlNodeH15 As MSXML.IXMLDOMNode
'    Dim cCol As Long, cRow As Long
'    Dim varTemp As Variant
'With fps
'    .EventEnabled(EventAllEvents) = False
'    .Sheet = .ActiveSheet
'     GetCellSpan fps, Col, Row
'    Select Case .Sheet
'      Case .SheetCount - 1
'         mCurrentSheet = .Sheet
'            .Sheet = 1
'            .GetText .ColLetterToNumber("L"), 16, varTemp
'            .Sheet = .SheetCount - 1
'            Set xmlNodeH15 = TAX_Utilities_New.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
'                 ParserCellID fps, GetAttribute(xmlNodeH15, "CellID"), cCol, cRow
'                .Col = cCol
'                .Row = cRow
'                If loaiKHBS = "frmKHBS_TT" Then
'                    .Value = CDbl(varTemp) - CDbl(vThuephainop)
'                     UpdateCell fps, .Col, .Row, .Value
'                Else
'                    .Value = CDbl(varTemp)
'                     UpdateCell fps, .Col, .Row, .Value
'                End If
'                If .Value < 0 Then
'                    .Col = .ColLetterToNumber("BG")
'                    .Row = cRow + 3
'                    .Value = 0
'                    UpdateCell fps, .Col, .Row, 0
'                Else
'                    .Col = .ColLetterToNumber("BG")
'                    .Row = cRow + 3
'                    UpdateCell fps, .Col, .Row, .Value
'                End If
'          End Select
'     .Sheet = mCurrentSheet
'     .EventEnabled(EventAllEvents) = True
'End With
'End Sub

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim strId As Variant
Dim iCol As Long, iRow As Long
Dim intIndexCombo As Integer, i As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim itemp As Variant, temp As Variant
Dim cValue As Variant
With fps
   If .ActiveSheet = .SheetCount - 1 Then
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("BD") And .Text <> "" Then
            intIndexCombo = .TypeComboBoxCurSel
            .Col = .ColLetterToNumber("BC")
            .TypeComboBoxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
            'Chon lai combo cua colF
            .Col = .ColLetterToNumber("BE")
            .GetText .ColLetterToNumber("BC"), Row, strId
            .Text = Trim(Right(strId, 10))
            UpdateCell fps, .Col, .Row, .Text
        End If
         If Col = .ColLetterToNumber("BE") And .Text <> "" Then
            intIndexCombo = .TypeComboBoxCurSel
            .Col = .ColLetterToNumber("BC")
            .TypeComboBoxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
            'Chon lai combo cua colF
            .Col = .ColLetterToNumber("BD")
            .GetText .ColLetterToNumber("BC"), Row, strId
            .Text = Trim(Left(strId, 100))
            UpdateCell fps, .Col, .Row, .Text
        End If
   
   
        Set xmlCellNode = TAX_Utilities_New.Data(.SheetCount - 2).nodeFromID(GetCellID(fps, Col, Row))
        Set xmlCellNode = xmlCellNode.parentNode.parentNode.firstChild.firstChild
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        blGiaTriCu = False
        If Col = fps.ColLetterToNumber("BD") Or Col = fps.ColLetterToNumber("BE") Then
            fps.Col = .ColLetterToNumber("BC")
            fps.Row = Row
            temp = fps.Text
            Set xmlCellsNode = xmlCellNode.parentNode
            i = 1
            .Col = .ColLetterToNumber("B")
            ParserCellID fps, "BD_9", iCol, iRow
            .Row = iRow
            Do
                 .Col = .ColLetterToNumber("BC")
                 itemp = .Text
                 If .Row <> Row Then
                    If itemp = temp Then 'Neu co thi dat co blGiaTriCu = True
                        RowGiaTriCu = Row
                        ColGiaTriCu = .ColLetterToNumber("BD")
                        .Row = Row
                        .Col = .ColLetterToNumber("BD")
                        blGiaTriCu = True
                        Exit Sub
                    End If
                 End If
              
                 .Col = .ColLetterToNumber("B")
                 .Row = i + iRow
                 i = i + 1
            Loop Until .Text = "bb"
       End If
     
    End If
    
End With
    
   ' CellChange Col, Row
    TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
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
        
        UpdateSheet
            
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
                Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 1)
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
  '  CellChange fps.ColLetterToNumber(S1TxtL13Col), S1TxtL13Row
End Function

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()

End Sub

Private Function GetMP10Previous(ByRef strValue As String) As Boolean
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim strDataFileName As String
    Dim strM13Target, strP13Target As String
    Dim iDom As Integer
    Dim MY As Date
    
    iDom = 0
    GetMP10Previous = True
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        MY = DateSerial(TAX_Utilities_New.Year, TAX_Utilities_New.Month, 1)
        MY = DateAdd("M", -1, MY)
        
        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(GetValidityNode("04", Format(Month(MY), "0#"), , Format(Year(MY), "####")).childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        If .Load(strDataFileName) = True Then
            Set xmlNode = .nodeFromID("P_38")
            strM13Target = GetAttribute(xmlNode, "Value")
            Set xmlNode = .nodeFromID("S_38")
            strP13Target = GetAttribute(xmlNode, "Value")
            
            If fps.Col = fps.ColLetterToNumber("P") Then
                strValue = -(Val(strM13Target))
            Else
                strValue = -(Val(strP13Target))
            End If
            Set xmlNode = Nothing
        ElseIf .parseError.reason <> vbNullString Then
            If InStr(1, .parseError.errorCode, "2146697210") <> 0 Then
                'file data of last month does not exist
            Else
                MsgBox .parseError.reason
            End If
            GetMP10Previous = False
        End If
        
    End With
    
    Set xmlDomLastData = Nothing
    Set xmlNode = Nothing
    
    Exit Function
    
ErrorHandle:
    SaveErrorLog "cls004", "GetG6Previous", Err.number, Err.Description
End Function

Private Sub setMP10Target()
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNodeM13 As MSXML.IXMLDOMNode
    Dim xmlNodeP13 As MSXML.IXMLDOMNode
    Dim strOriginDataFileName As String
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim strTargetM13 As String
    Dim strTargetP13 As String
    Dim ltemp As Double
    Dim iDom As Integer
    Dim MY As Date
    
    iDom = 0
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
     
        MY = DateSerial(TAX_Utilities_New.Year, TAX_Utilities_New.Month, 1)
        
        MY = DateAdd("M", -1, MY)
        strOriginDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(1), "DataFile") & ".xml"
        strCurrentDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(iDom), "DataFile") & "_" & TAX_Utilities_New.Month & TAX_Utilities_New.Year & ".xml"
        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(GetValidityNode("04", Format(Month(MY), "0#"), , Format(Year(MY), "####")).childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        
        If xmlDomCurrentData.Load(strCurrentDataFileName) = True Then
            GoTo Tidy_up
        End If
        If .Load(strDataFileName) = True Then
            Set xmlNodeM13 = .nodeFromID("P_38")
            Set xmlNodeP13 = .nodeFromID("S_38")
            strTargetM13 = GetAttribute(xmlNodeM13, "Value")
            strTargetP13 = GetAttribute(xmlNodeP13, "Value")
            
            ltemp = Val(strTargetM13)
            
            Set xmlNodeM13 = Nothing
            Set xmlNodeP13 = Nothing
            fps.Sheet = 1
            fps.Col = fps.ColLetterToNumber("P")
            fps.Row = 38
            If Trim(fps.value) = "0" Or Trim(fps.value) = "" Then
                If ltemp < 0 Then
                    mCurrentSheet = 1
                    fps.value = -ltemp
                    UpdateCell fps, fps.Col, fps.Row, fps.value
                End If
            End If
            
            ltemp = Val(strTargetP13)
            fps.Col = fps.ColLetterToNumber("S")
            fps.Row = 38
            If Trim(fps.value) = "0" Or Trim(fps.value) = "" Then
                If ltemp < 0 Then
                    mCurrentSheet = 1
                    fps.value = -ltemp
                    UpdateCell fps, fps.Col, fps.Row, fps.value
                End If
            End If
            
            
        ElseIf .parseError.reason <> vbNullString Then
            If InStr(1, .parseError.errorCode, "2146697210") <> 0 Then
                'file data of last month does not exist
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
    Set xmlNodeM13 = Nothing
    Set xmlNodeP13 = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls004", "setG6Target", Err.number, Err.Description
End Sub

Public Sub Prepare_KHBS()
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    Dim strPath As String
    Dim strComboName As String
    Dim strComboMCT As String
    Dim strCombo As String
       
    mblKhoiDauToKhai = True
    fps.Sheet = fps.SheetCount - 1
    fps.ActiveSheet = fps.Sheet
    
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT.xml"
    
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
                strCombo = strCombo + CPab(GetAttribute(xmlNodeItem, "Caption"), 100) + CPab(GetAttribute(xmlNodeItem, "CellID"), 20) + GetAttribute(xmlNodeItem, "MCT") + Chr$(9)
                strComboName = strComboName + GetAttribute(xmlNodeItem, "Caption") + Chr$(9)
                strComboMCT = strComboMCT + GetAttribute(xmlNodeItem, "MCT") + Chr$(9)
    Next
     With fps
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        .Col = .ColLetterToNumber("BC")
        .Row = 9
        .TypeComboBoxList = strCombo
        .Row = 13
        .TypeComboBoxList = strCombo
        .Col = .ColLetterToNumber("BD")
        .Row = 9
        .TypeComboBoxList = strComboName
        .Row = 13
        .TypeComboBoxList = strComboName
        .Col = .ColLetterToNumber("BE")
        .Row = 9
        .TypeComboBoxList = strComboMCT
        .Row = 13
        .TypeComboBoxList = strComboMCT
    End With
    
End Sub

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

Public Sub UpdateChangeKHBS()
    Dim varTemp As Variant
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    With fps
        .Sheet = 1
        For Each xmlCellNodeData In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
            ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
            .GetText lCol, lRow, varTemp
            UpdateLastKHBSCell fps, lCol, lRow, varTemp
        Next
        .Sheet = mCurrentSheet
    End With
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
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
            Set xmlNode = .nodeFromID("M_39")
            result = "CT30_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("P_39")
            result = result & "~" & "CT31_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("S_39")
            result = result & "~" & "CT32_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_03GTGT", "getValueTK", Err.number, Err.Description
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
                .GetText .ColLetterToNumber("M"), 39, varTemp
                strCurrentValue = "CT30_" & Trim(varTemp)
                .GetText .ColLetterToNumber("P"), 39, varTemp
                strCurrentValue = strCurrentValue & "~CT31_" & Trim(varTemp)
                .GetText .ColLetterToNumber("S"), 39, varTemp
                strCurrentValue = strCurrentValue & "~CT32_" & Trim(varTemp)
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
                If Left(strMaCT, 3) = "" Then
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
                        ' dieu chinh giam
                        strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
                        ' dieu chinh tang
                        strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                Else
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
                        ' dieu chinh giam
                        strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
                        ' dieu chinh tang
                        strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                    End If
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
    SaveErrorLog "03_GTGT", "getMapCT", Err.number, Err.Description
End Function


Public Sub UpdateSheet()
    Dim varTemp As Variant
    Dim temp As Integer
    Dim blCheck_S2 As Boolean
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



' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("O"), 44, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function

' Lay ve chi tieu 34 o to khai lan truoc
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
            Set xmlNode = .nodeFromID("O_44")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_03GTGT", "getValueCTDC", Err.number, Err.Description
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
                    Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
                        (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length > 20 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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
    tuRowDL = 21
    denRowDL = 31
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "20_30_2"
    ' end
End Sub

