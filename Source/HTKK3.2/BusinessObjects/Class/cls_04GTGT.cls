VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_04GTGT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Private Const S1TxtL13Row = 13
Private Const S1TxtL13Col = "L"
Private Const S1TxtL14Row = 14
Private Const S1TxtL14Col = "L"
Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
'Dim xNodeData0 As MSXML.IXMLDOMNode
Dim xNode As MSXML.IXMLDOMNode
Dim arrData() As String
Public loaiKHBS As String
Public Strloaitk As String
Public StrSolanBosung As String
Public FlagThaydoi As String

Private Const Row_MSTDN = 14
Private Const Col_MSTDN = "H"

Public strMaNhom As String
Public strMaNhomAdd As String
Public lViTri As Long


' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

'Dinh nghia type luu du  Row
Public Type MY_ROW
    tenCT As String
    cot1 As Double ' Thue suat 0%
    cot2 As Double ' Thue suat 5%
    cot3 As Double ' Thue suat 10%
    Cot4 As Integer ' mas so
    Cot5 As Integer ' Vitri
End Type
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim strDataFileName As String
    Dim xmlNodeListItems As MSXML.IXMLDOMNodeList
    Dim strCboNhomNNKD As String
    Dim strCboNhomNNKDAn As String
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim i As Integer
    Dim fldList() As String
    strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_DM_Nhom_NNKD.xml")
    With fps
       If xmlDomData.Load(strDataFileName) Then
         Set xmlNodeListItems = xmlDomData.getElementsByTagName("Item")
         For Each xmlNode In xmlNodeListItems
             fldList = Split(GetAttribute(xmlNode, "Value"), "###")
             strCboNhomNNKD = strCboNhomNNKD + fldList(1) + Chr$(9)
             strCboNhomNNKDAn = strCboNhomNNKDAn + fldList(0) + Chr$(9)
         Next
     End If
          'set nhom nganh nghe KD
        .Sheet = 1
        .Row = 45
        .Col = .ColLetterToNumber("C")
        .TypeComboBoxList = strCboNhomNNKD
        ' set Ma nhom NNKD
        .Col = .ColLetterToNumber("AE")
        .TypeComboBoxList = strCboNhomNNKDAn
        
        ' Group 2
         .Row = 49
        .Col = .ColLetterToNumber("C")
        .TypeComboBoxList = strCboNhomNNKD
        ' set Ma nhom NNKD
        .Col = .ColLetterToNumber("AE")
        .TypeComboBoxList = strCboNhomNNKDAn
        
        ' Group 3
        .Row = 53
        .Col = .ColLetterToNumber("C")
        .TypeComboBoxList = strCboNhomNNKD
        ' set Ma nhom NNKD
        .Col = .ColLetterToNumber("AE")
        .TypeComboBoxList = strCboNhomNNKDAn
        
        ' group 4
        .Row = 57
        .Col = .ColLetterToNumber("C")
        .TypeComboBoxList = strCboNhomNNKD
        ' set Ma nhom NNKD
        .Col = .ColLetterToNumber("AE")
        .TypeComboBoxList = strCboNhomNNKDAn


        SetDateFormat fps, 1, 75, .ColLetterToNumber("T"), DDMMYYYY
        SetDateFormat fps, .SheetCount - 1, 23, .ColLetterToNumber("BG"), DDMMYYYY
        .Sheet = 1
        .Row = 75
        .Col = .ColLetterToNumber("T")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft

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
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        'update MSTDN luôn bang MST tren phan TT chung
        UpdateMST fps, Col_MSTDN, Row_MSTDN
        
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("C")
            .Row = 37
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 37
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 37
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("C")
            .Row = 37
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 37
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 37
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        ' Set so thu tu
        ' group 2
         fps.EventEnabled(EventAllEvents) = False
         ' Group 1
         i = 1
         .Col = .ColLetterToNumber("B")
         .Row = 45
         Do
              .Text = "1." & Trim(str(i))
              .Col = .ColLetterToNumber("B")
              .Row = i + 45
              i = i + 1
         Loop Until .Text = "aa"
         i = 1
        .Row = .Row + 3
         Do
              .Text = "2." & Trim(str(i))
              .Col = .ColLetterToNumber("B")
              .Row = 1 + .Row
              i = i + 1
         Loop Until .Text = "bb"
         
         ' Group 3
         i = 1
        .Row = .Row + 3
         Do
              .Text = "3." & Trim(str(i))
              .Col = .ColLetterToNumber("B")
              .Row = 1 + .Row
              i = i + 1
         Loop Until .Text = "cc"
         ' Group 4
         i = 1
        .Row = .Row + 3
         Do
              .Text = "4." & Trim(str(i))
              .Col = .ColLetterToNumber("B")
              .Row = 1 + .Row
              i = i + 1
         Loop Until .Text = "dd"
        fps.EventEnabled(EventAllEvents) = True
        
        mCurrentSheet = .Sheet
    End With
    
  '  CellChange fps.ColLetterToNumber(S1TxtL13Col), S1TxtL13Row
End Sub
Public Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
Dim strId As Variant
Dim iCol As Long, iRow As Long
Dim intIndexCombo As Integer, i As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim itemp As Variant, temp As Variant
Dim cValue As Variant

Dim tempRow As MY_ROW
Dim arrRow() As MY_ROW, arrRow2() As MY_ROW, arrRow3() As MY_ROW, arrRow4() As MY_ROW
Dim tempValue As Variant


    With fps
        If .ActiveSheet = 1 Then
                    'Huy chon Combo to khai khong hop le va ra thong bao
            blGiaTriCu = False
                    
            If blGiaTriCu And Col = .ColLetterToNumber("C") Then 'And Not mblKeyF5 Then
                If Not .Visible Then
                    .Visible = True
                    DisplayMessage "0242", msOKOnly, miWarning
                    .Visible = False
                Else
                    DisplayMessage "0242", msOKOnly, miWarning
                End If
                
                .Sheet = 1
                mCurrentSheet = .Sheet
                .Row = RowGiaTriCu
                .Col = ColGiaTriCu
                .Text = strGiaTriCu
                .Sheet = 1
                mCurrentSheet = .Sheet
                .Row = RowGiaTriCu
                .Col = ColGiaTriCu
                UpdateCell fps, .Col, .Row, .Text
                .Sheet = 1
                mCurrentSheet = .Sheet
                .Row = RowGiaTriCu
                .Col = .ColLetterToNumber("AE")
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                .Sheet = 1
                mCurrentSheet = .Sheet
                .Row = RowGiaTriCu
                .Col = .ColLetterToNumber("K")
                .Text = "0"
                UpdateCell fps, .Col, .Row, .Text
                .Sheet = 1
                mCurrentSheet = .Sheet
                .Row = RowGiaTriCu
                .Col = .ColLetterToNumber("Q")
                .Text = "0"
                UpdateCell fps, .Col, .Row, .Text
                .Sheet = 1
                mCurrentSheet = .Sheet
                .Row = RowGiaTriCu
                .Col = .ColLetterToNumber("W")
                .Text = "0"
                UpdateCell fps, .Col, .Row, .Text
                
                
                blGiaTriCu = False
'                mblLoi = True
            End If
'            .Col = Col
'            .row = row
'             If Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("Q") Or Col = .ColLetterToNumber("W") Then
'                ' Lay ve mang cac row hien tai cua group 1
'
'                .Col = .ColLetterToNumber("B")
'                .row = 45
'                i = 1
'                Do
'                    .GetText .ColLetterToNumber("C"), .row, tempValue
'                    tempRow.tenCT = CStr(tempValue)
'                    .GetText .ColLetterToNumber("K"), .row, tempValue
'                    tempRow.cot1 = Val(tempValue)
'                    .GetText .ColLetterToNumber("Q"), .row, tempValue
'                    tempRow.cot2 = Val(tempValue)
'                    .GetText .ColLetterToNumber("W"), .row, tempValue
'                    tempRow.cot3 = Val(tempValue)
'                    .GetText .ColLetterToNumber("AE"), .row, tempValue
'                    tempRow.Cot4 = Val(tempValue)
'                    tempRow.cot5 = i
'                    ReDim Preserve arrRow(i - 1)
'                    arrRow(i - 1) = tempRow
'                    i = i + 1
'                    .Col = .ColLetterToNumber("B")
'                    .row = .row + 1
'                Loop Until .Text = "aa"
'
'                ' Lay ve mang cac row hien tai cua group 2
'                .Col = .ColLetterToNumber("B")
'                .row = .row + 3
'                i = 1
'                Do
'                    tempRow.tenCT = arrRow(i - 1).tenCT
'                    .GetText .ColLetterToNumber("K"), .row, tempValue
'                    tempRow.cot1 = Val(tempValue)
'                    .GetText .ColLetterToNumber("Q"), .row, tempValue
'                    tempRow.cot2 = Val(tempValue)
'                    .GetText .ColLetterToNumber("W"), .row, tempValue
'                    tempRow.cot3 = Val(tempValue)
'                    .GetText .ColLetterToNumber("AE"), .row, tempValue
'                    tempRow.Cot4 = Val(tempValue)
'                    tempRow.cot5 = i
'                    ReDim Preserve arrRow2(i - 1)
'                    arrRow2(i - 1) = tempRow
'                    i = i + 1
'                    .Col = .ColLetterToNumber("B")
'                    .row = .row + 1
'                Loop Until .Text = "bb"
'
'                ' Lay ve mang cac row hien tai cua group 3
'                 .Col = .ColLetterToNumber("B")
'                .row = .row + 3
'                i = 1
'                Do
'                    .Col = .ColLetterToNumber("K")
'                    .Text = arrRow(i - 1).cot1 * arrRow2(i - 1).cot1
'                    'UpdateCell fps, .Col, .row, .Text
'
'                    .Col = .ColLetterToNumber("B")
'                    .row = .row + 1
'                Loop Until .Text = "cc"
'
'                ' Lay ve mang cac row hien tai cua group 4
'                .Col = .ColLetterToNumber("B")
'                .row = .row + 3
'                i = 1
'                Do
'                     tempRow.tenCT = arrRow(i - 1).tenCT
'                    .GetText .ColLetterToNumber("K"), .row, tempValue
'                    tempRow.cot1 = Val(tempValue)
'                    .GetText .ColLetterToNumber("Q"), .row, tempValue
'                    tempRow.cot2 = Val(tempValue)
'                    .GetText .ColLetterToNumber("W"), .row, tempValue
'                    tempRow.cot3 = Val(tempValue)
'                    .GetText .ColLetterToNumber("AE"), .row, tempValue
'                    tempRow.Cot4 = Val(tempValue)
'                    tempRow.cot5 = i
'                    ReDim Preserve arrRow4(i - 1)
'                    arrRow4(i - 1) = tempRow
'                    i = i + 1
'                    .Col = .ColLetterToNumber("B")
'                    .row = .row + 1
'                Loop Until .Text = "dd"
'              End If
         End If
     End With
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
Dim countRowDyn As Long, coundtRowDel As Long
Dim arrRow() As MY_ROW
Dim tempRow As MY_ROW
Dim tempValue As Variant, tempGroup As Variant, tempMaNhom As Variant
Dim tCol As Long, tRow As Long
    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
       ' Neu group II hoac III, IV thi khong them dong khi bam F5 hoac F6
       If Not isAddRow(iCol, iRow) Then
            Exit Sub
       End If
        
       If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
             If .ActiveSheet = 1 Then
                 fps_Change iCol, iRow
                 fps.EventEnabled(EventAllEvents) = False
                    ' Group 1
                    fps.Sheet = mCurrentSheet
        '            ' them dong o group 2
                    If KeyCode = vbKeyF5 Then
                        ' them dong o group 2
                         .Col = .ColLetterToNumber("B")
                         .Row = 45
                         Do
                             .Col = .ColLetterToNumber("B")
                             .Row = .Row + 1
                         Loop Until .Text = "aa"
                        ' Lay ve vi tri muon them dong
                        ' Group 2
                         .Col = .ColLetterToNumber("B")
                         .Row = .Row + 3
                         i = 1
                         Do
                              .GetText .ColLetterToNumber("AG"), .Row, tempGroup
                              .GetText .ColLetterToNumber("AE"), .Row, tempValue
                              If lViTri = i And Val(tempGroup) = 2 Then
                                    InsertNode .ColLetterToNumber("L"), .Row
                              End If
                              i = i + 1
                             .Col = .ColLetterToNumber("B")
                             .Row = .Row + 1
                         Loop Until .Text = "bb"
                         ' Group 3
                         .Col = .ColLetterToNumber("B")
                         .Row = .Row + 3
                         i = 1
                         Do
                              .GetText .ColLetterToNumber("AG"), .Row, tempGroup
                              .GetText .ColLetterToNumber("AE"), .Row, tempValue
                              If lViTri = i And Val(tempGroup) = 3 Then
                                    InsertNode .ColLetterToNumber("L"), .Row
                              End If
                              i = i + 1
                             .Col = .ColLetterToNumber("B")
                             .Row = .Row + 1
                         Loop Until .Text = "cc"
                         ' Group 4
                         .Col = .ColLetterToNumber("B")
                         .Row = .Row + 3
                         i = 1
                         Do
                              .GetText .ColLetterToNumber("AG"), .Row, tempGroup
                              .GetText .ColLetterToNumber("AE"), .Row, tempValue
                              If lViTri = i And Val(tempGroup) = 4 Then
                                    InsertNode .ColLetterToNumber("L"), .Row
                              End If
                              i = i + 1
                             .Col = .ColLetterToNumber("B")
                             .Row = .Row + 1
                         Loop Until .Text = "dd"
                         
                    End If
                    
                    ' Xoa cac row dong cua group 2,3,4
                    If KeyCode = vbKeyF6 Then
                        ' xoa dong o group 2
                         .Col = .ColLetterToNumber("B")
                         .Row = 45
                         Do
                             .Col = .ColLetterToNumber("B")
                             .Row = .Row + 1
                         Loop Until .Text = "aa"
                        ' Lay ve vi tri muon xoa
                        ' Group 2
                         .Col = .ColLetterToNumber("B")
                         .Row = .Row + 3
                         i = 1
                         Do
                              .GetText .ColLetterToNumber("AG"), .Row, tempGroup
                              .GetText .ColLetterToNumber("AE"), .Row, tempValue
                               If lViTri = i And Val(tempGroup) = 2 Then
                                    DeleteNode .ActiveSheet, .ColLetterToNumber("L"), .Row
                               End If
                               i = i + 1
                             .Col = .ColLetterToNumber("B")
                             .Row = .Row + 1
                         Loop Until .Text = "bb"
                         
                         ' xoa dong o group 3
                        ' Lay ve vi tri muon xoa
                         .Col = .ColLetterToNumber("B")
                         .Row = .Row + 3
                         i = 1
                         Do
                              .GetText .ColLetterToNumber("AG"), .Row, tempGroup
                              .GetText .ColLetterToNumber("AE"), .Row, tempValue
                              If lViTri = i And Val(tempGroup) = 3 Then
                                    DeleteNode .ActiveSheet, .ColLetterToNumber("L"), .Row
                              End If
                              i = i + 1
                             .Col = .ColLetterToNumber("B")
                             .Row = .Row + 1
                         Loop Until .Text = "cc"
                         ' xoa dong o group 4
                        ' Lay ve vi tri muon xoa
                         .Col = .ColLetterToNumber("B")
                         .Row = .Row + 3
                         i = 1
                         Do
                              .GetText .ColLetterToNumber("AG"), .Row, tempGroup
                              .GetText .ColLetterToNumber("AE"), .Row, tempValue
                              If lViTri = i And Val(tempGroup) = 4 Then
                                    DeleteNode .ActiveSheet, .ColLetterToNumber("L"), .Row
                              End If
                              i = i + 1
                             .Col = .ColLetterToNumber("B")
                             .Row = .Row + 1
                         Loop Until .Text = "dd"
                    End If

                    ' group 2
                    fps.EventEnabled(EventAllEvents) = False
                    ' Group 1
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 45
                    Do
                         .Text = "1." & Trim(str(i))
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 45
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 3
                    Do
                         .Text = "2." & Trim(str(i))
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"
                    coundtRowDel = i - 1
                    ' Group 3
                    i = 1
                   .Row = .Row + 3
                    Do
                         .Text = "3." & Trim(str(i))
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "cc"
                    ' Group 4
                    i = 1
                   .Row = .Row + 3
                    Do
                         .Text = "4." & Trim(str(i))
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "dd"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
                 fps.EventEnabled(EventAllEvents) = True
            End If
           
'            fps.Sheet = mCurrentSheet
''            ' them dong o group 2
'            If KeyCode = vbKeyF5 Then
'                InsertNode fps.ActiveCol, 49 + countRowDyn
'            End If
'            ' Xoa dong o group 2
'            If KeyCode = vbKeyF6 Then
'                DeleteNode 1, fps.ActiveCol, 50
'            End If
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
            If Col = .ColLetterToNumber("T") And Row = .MaxRows - 1 Then
                .GetText Col, Row, varTemp
                  If varTemp <> "" And varTemp <> "../../...." Then
                      If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                          .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                          .TypeHAlign = TypeHAlignLeft
                      Else
                               '.SetFocus
                               .SetActiveCell Col, Row
                      End If
                  Else
                   .SetText Col, Row, ""
                  End If
                 .Col = Col
                 .Row = Row
                 UpdateCell fps, .Col, .Row, .Text
            End If
            
'            .GetText .ColLetterToNumber("L"), 17, varTemp
'            If TAX_Utilities_New.NodeValidity.childNodes(.SheetCount - 2).Attributes.getNamedItem("Active").nodeValue <> "0" Then
'                .Sheet = .SheetCount - 1
'                Set xmlNode = TAX_Utilities_New.Data(.Sheet - 1).childNodes(2).childNodes(2).lastChild.lastChild
'                ParserCellID fps, GetAttribute(xmlNode, "CellID"), cCol, cRow
'                .Col = cCol
'                .Row = cRow
'                .Value = CDbl(varTemp) - CDbl(vThuephainop)
'                mCurrentSheet = .SheetCount - 1
'            '    UpdateCell fps, .Col, .Row, .Value
'                .Col = .ColLetterToNumber("BG")
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
    
    CellChange Col, Row
End Sub

Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
    Dim vL13, vG6 As Variant
    Dim strG6Previous As String
    Dim varTemp As Variant
    Dim temp As Variant
    Dim cCol As Long, cRow As Long
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = .ActiveSheet
        mCurrentSheet = .ActiveSheet
        GetCellSpan fps, Col, Row
        
        If .ActiveSheet = 1 Then
            If Row = 12 And Col = .ColLetterToNumber("P") Then
                    .GetText Col, Row, temp
                    If Val(temp) <= 0 Then
                        .Col = .ColLetterToNumber("P")
                        .Row = 13
                        .value = 0
                        .value = 0
                        UpdateCell fps, .Col, .Row, .value
                    Else
                       .Col = .ColLetterToNumber("P")
                       .Row = 13
                       UpdateCell fps, .Col, .Row, .value
                    End If
            End If
            If Row = 12 And Col = .ColLetterToNumber("M") Then
                    .GetText Col, Row, temp
                    If Val(temp) <= 0 Then
                        .Col = .ColLetterToNumber("M")
                        .Row = 13
                        .value = 0
                        .value = 0
                        UpdateCell fps, .Col, .Row, .value
                    Else
                       .Col = .ColLetterToNumber("M")
                       .Row = 13
                       UpdateCell fps, .Col, .Row, .value
                    End If
            End If
        
      End If
      
      'KHBS
        If .ActiveSheet = .SheetCount - 1 Then
            mCurrentSheet = .Sheet
           
        End If
      
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim strId As Variant
Dim iCol As Long, iRow As Long
Dim intIndexCombo As Integer, i As Integer
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim itemp As Variant, temp As Variant
Dim cValue As Variant

Dim tempRow As MY_ROW
Dim arrRow() As MY_ROW, arrRow2() As MY_ROW, arrRow3() As MY_ROW, arrRow4() As MY_ROW
Dim tempValue As Variant


    With fps
        If .ActiveSheet = 1 Then
                .Col = Col
                .Row = Row
                 If Col = .ColLetterToNumber("C") And .Text <> "" Then
                    intIndexCombo = .TypeComboBoxCurSel
                    .Col = .ColLetterToNumber("AE")
                    .TypeComboBoxCurSel = intIndexCombo
                    'Update combo C
                    UpdateCell fps, .Col, .Row, .Text
                End If
                
               ' Kiem tra trung ten
               Set xmlCellNode = TAX_Utilities_New.Data(1 - 1).nodeFromID(GetCellID(fps, Col, Row))
               Set xmlCellNode = xmlCellNode.parentNode.parentNode.firstChild.firstChild
               ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
               blGiaTriCu = False
        
'                If Col = fps.ColLetterToNumber("C") Then
'                    'kiem tra trung Ten NNKD trong to khai
'                    fps.Col = Col
'                    fps.Row = Row
'                    temp = fps.Text
'                    Set xmlCellsNode = xmlCellNode.parentNode
'                    i = 1
'                    .Col = .ColLetterToNumber("B")
'                    ParserCellID fps, GetAttribute(xmlCellsNode.childNodes(i - 1), "CellID"), iCol, iRow
'                    .Row = iRow
'
'                    Do
'                         .Col = .ColLetterToNumber("C")
'                         itemp = .Text
'                         If .Row <> Row Then
'                            If itemp = temp Then 'Neu co thi dat co blGiaTriCu = True
'                                RowGiaTriCu = Row
'                                ColGiaTriCu = .ColLetterToNumber("C")
'                                .Row = Row
'                                .Col = .ColLetterToNumber("C")
'                                blGiaTriCu = True
'                                Exit Sub
'                            End If
'                         End If
'                         .Col = .ColLetterToNumber("B")
'                         .Row = i + iRow
'                         i = i + 1
'                    Loop Until .Text = "aa"
'                    .Sheet = 1
'                    mCurrentSheet = 1
'                End If
                ' STT dong change
                
                ' Lay ve mang cac row hien tai cua group 1
                
                .Col = .ColLetterToNumber("B")
                .Row = 45
                i = 1
                Do
                    .GetText .ColLetterToNumber("C"), .Row, tempValue
                    tempRow.tenCT = CStr(tempValue)
                    .GetText .ColLetterToNumber("K"), .Row, tempValue
                    tempRow.cot1 = Val(tempValue)
                    .GetText .ColLetterToNumber("Q"), .Row, tempValue
                    tempRow.cot2 = Val(tempValue)
                    .GetText .ColLetterToNumber("W"), .Row, tempValue
                    tempRow.cot3 = Val(tempValue)
                    .GetText .ColLetterToNumber("AE"), .Row, tempValue
                    tempRow.Cot4 = Val(tempValue)
                    tempRow.Cot5 = i
                    ReDim Preserve arrRow(i - 1)
                    arrRow(i - 1) = tempRow
                    i = i + 1
                    .Col = .ColLetterToNumber("B")
                    .Row = .Row + 1
                Loop Until .Text = "aa"
                
                ' Lay ve mang cac row hien tai cua group 2
                ' bo check kiem tra trung ten
                blGiaTriCu = False
                
                If blGiaTriCu = False Then
                     .Col = .ColLetterToNumber("B")
                    .Row = .Row + 3
                    i = 1
                    Do
                        If i = Row - 44 Then
                            ' Ten
                            .Col = .ColLetterToNumber("C")
                            .TypeComboBoxCurSel = intIndexCombo
                            ' Ma nhom
                            .Col = .ColLetterToNumber("AE")
                            .TypeComboBoxCurSel = intIndexCombo
                            UpdateCell fps, .Col, .Row, .Text
                         End If
                        tempRow.tenCT = arrRow(i - 1).tenCT
                        .GetText .ColLetterToNumber("K"), .Row, tempValue
                        tempRow.cot1 = Val(tempValue)
                        .GetText .ColLetterToNumber("Q"), .Row, tempValue
                        tempRow.cot2 = Val(tempValue)
                        .GetText .ColLetterToNumber("W"), .Row, tempValue
                        tempRow.cot3 = Val(tempValue)
                        .GetText .ColLetterToNumber("AE"), .Row, tempValue
                        tempRow.Cot4 = Val(tempValue)
                        tempRow.Cot5 = i
                        ReDim Preserve arrRow2(i - 1)
                        arrRow2(i - 1) = tempRow
                        i = i + 1
                        .Col = .ColLetterToNumber("B")
                        .Row = .Row + 1
                    Loop Until .Text = "bb"
                    
                    ' Lay ve mang cac row hien tai cua group 3
                     .Col = .ColLetterToNumber("B")
                    .Row = .Row + 3
                    i = 1
                    Do
                        If i = Row - 44 Then
                            ' Ten
                            .Col = .ColLetterToNumber("C")
                            .TypeComboBoxCurSel = intIndexCombo
                            ' Ma nhom
                            .Col = .ColLetterToNumber("AE")
                            .TypeComboBoxCurSel = intIndexCombo
                            UpdateCell fps, .Col, .Row, .Text
                         End If
                        tempRow.tenCT = arrRow(i - 1).tenCT
                        .GetText .ColLetterToNumber("K"), .Row, tempValue
                        tempRow.cot1 = Val(tempValue)
                        .GetText .ColLetterToNumber("Q"), .Row, tempValue
                        tempRow.cot2 = Val(tempValue)
                        .GetText .ColLetterToNumber("W"), .Row, tempValue
                        tempRow.cot3 = Val(tempValue)
                        .GetText .ColLetterToNumber("AE"), .Row, tempValue
                        tempRow.Cot4 = Val(tempValue)
                        tempRow.Cot5 = i
                        ReDim Preserve arrRow3(i - 1)
                        arrRow3(i - 1) = tempRow
                        i = i + 1
                        .Col = .ColLetterToNumber("B")
                        .Row = .Row + 1
                    Loop Until .Text = "cc"
    
                    ' Lay ve mang cac row hien tai cua group 4
                     .Col = .ColLetterToNumber("B")
                    .Row = .Row + 3
                    i = 1
                    Do
                        If i = Row - 44 Then
                            ' Ten
                            .Col = .ColLetterToNumber("C")
                            .TypeComboBoxCurSel = intIndexCombo
                            ' Ma nhom
                            .Col = .ColLetterToNumber("AE")
                            .TypeComboBoxCurSel = intIndexCombo
                            UpdateCell fps, .Col, .Row, .Text
                         End If
                        tempRow.tenCT = arrRow(i - 1).tenCT
                        .GetText .ColLetterToNumber("K"), .Row, tempValue
                        tempRow.cot1 = Val(tempValue)
                        .GetText .ColLetterToNumber("Q"), .Row, tempValue
                        tempRow.cot2 = Val(tempValue)
                        .GetText .ColLetterToNumber("W"), .Row, tempValue
                        tempRow.cot3 = Val(tempValue)
                        .GetText .ColLetterToNumber("AE"), .Row, tempValue
                        tempRow.Cot4 = Val(tempValue)
                        tempRow.Cot5 = i
                        ReDim Preserve arrRow4(i - 1)
                        arrRow4(i - 1) = tempRow
                        i = i + 1
                        .Col = .ColLetterToNumber("B")
                        .Row = .Row + 1
                    Loop Until .Text = "dd"
                    
                End If
          End If
     End With
    'CellChange Col, Row
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
    
    Dim i As Integer
    With fps
    '.Visible = False
    '.ReDraw = False
    iCurrentSheet = .Sheet
    .Sheet = .SheetCount

    CheckErrorMST
    
    ' Set lai so TT
    ' group 2
    
       fps.EventEnabled(EventAllEvents) = False
       ' Group 1
       .Sheet = 1
       i = 1
       .Col = .ColLetterToNumber("B")
       .Row = 45
       Do
            .Text = "1." & Trim(str(i))
            .Col = .ColLetterToNumber("B")
            .Row = i + 45
            i = i + 1
       Loop Until .Text = "aa"
       i = 1
      .Row = .Row + 3
       Do
            .Text = "2." & Trim(str(i))
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
       Loop Until .Text = "bb"
       ' Group 3
       i = 1
      .Row = .Row + 3
       Do
            .Text = "3." & Trim(str(i))
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
       Loop Until .Text = "cc"
       ' Group 4
       i = 1
      .Row = .Row + 3
       Do
            .Text = "4." & Trim(str(i))
            .Col = .ColLetterToNumber("B")
            .Row = 1 + .Row
            i = i + 1
       Loop Until .Text = "dd"
    fps.EventEnabled(EventAllEvents) = True
    
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
    Dim totalCell  As Long, totalCell1  As Long
    Dim countRowReset As Long, countRowReset1 As Long
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    totalCell = TAX_Utilities_New.Data(0).getElementsByTagName("Cell").length
    totalCell1 = TAX_Utilities_New.Data(fps.SheetCount - 2).getElementsByTagName("Cell").length
    countRowReset = (totalCell - 26) / 6
    countRowReset1 = (totalCell1 - 8) / 6
    With fps
        For Each xmlNodeReset In TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            .Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            If ((lRow < 43 And lRow <> 25) Or lRow > 62 + countRowReset) And mCurrentSheet = 1 Then
            ElseIf (lRow >= 20 + countRowReset1) And mCurrentSheet = (fps.SheetCount - 1) Then
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
    Dim i As Integer
    With fps
       ' Set so thu tu
        ' group 2
         fps.EventEnabled(EventAllEvents) = False
         ' Group 1
         If mCurrentSheet = 1 Then
             .Sheet = 1
             i = 1
             .Col = .ColLetterToNumber("B")
             .Row = 45
             Do
                  .Text = "1." & Trim(str(i))
                  .Col = .ColLetterToNumber("B")
                  .Row = i + 45
                  i = i + 1
             Loop Until .Text = "aa"
             i = 1
            .Row = .Row + 3
             Do
                  .Text = "2." & Trim(str(i))
                  .Col = .ColLetterToNumber("B")
                  .Row = 1 + .Row
                  i = i + 1
             Loop Until .Text = "bb"
             
             ' Group 3
             i = 1
            .Row = .Row + 3
             Do
                  .Text = "3." & Trim(str(i))
                  .Col = .ColLetterToNumber("B")
                  .Row = 1 + .Row
                  i = i + 1
             Loop Until .Text = "cc"
             ' Group 4
             i = 1
            .Row = .Row + 3
             Do
                  .Text = "4." & Trim(str(i))
                  .Col = .ColLetterToNumber("B")
                  .Row = 1 + .Row
                  i = i + 1
             Loop Until .Text = "dd"
        End If
        fps.EventEnabled(EventAllEvents) = True
        
        mCurrentSheet = .Sheet
    End With
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

' Insert them 1 row
Private Sub InsertNode(ByVal pCol As Long, ByVal pRow As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCell As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    Dim lLRowBound As Long, lURowBound As Long
    Dim lRow2s As Long, lRows As Long
    
    ' Get cellspan for merge cell on interface templates
    GetCellSpan fps, pCol, pRow
    
    Set xmlNodeCells = TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fps, pCol, pRow)).parentNode
    
    lRows = GetDynRowCount(fps, xmlNodeCells, lRow2s, lLRowBound, lURowBound)
    
    'If Not xmlNodeCells.nextSibling Is Nothing Then GoTo EXIT_SUB
    If GetAttribute(xmlNodeCells.parentNode, "Dynamic") <> "1" Then GoTo EXIT_SUB
    If Val(GetAttribute(xmlNodeCells.parentNode, "MaxRows")) = xmlNodeCells.parentNode.childNodes.length Then GoTo EXIT_SUB
    
    ' insert new row on grid
    InsertRow fps, lURowBound + 1, lRows
    'fpSpread1.SetFocus
    
    ' increase value of row in xmlDocument
    IncreaseRowInDOM fps, TAX_Utilities_New.Data(mCurrentSheet - 1), lURowBound + 1, lRows, lRow2s
    'IncreaseRowInDOM lURowBound + 1, lRows, lRow2s

    Set xmlNodeNewCells = xmlNodeCells.cloneNode(True)
    For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
        ' Set new ID for node (CellID)
        ParserCellID fps, GetAttribute(xmlNodeNewCell, "CellID"), lCol, lRow
        SetAttribute xmlNodeNewCell, "CellID", GetCellID(fps, lCol, lRow + lRows)
                
        ' Set first cell = 1
        SetAttribute xmlNodeNewCell, "FirstCell", "1"
        
        ' Reset empty value for new node
        fps.Col = lCol
        fps.Row = lRow
        Select Case fps.CellType
            Case CellTypeNumber
                SetAttribute xmlNodeNewCell, "Value", "0"
            Case Else
                SetAttribute xmlNodeNewCell, "Value", vbNullString
        End Select
        
        ' Set new ID for node (CellID2)
        ParserCellID fps, GetAttribute(xmlNodeNewCell, "CellID2"), lCol, lRow
        SetAttribute xmlNodeNewCell, "CellID2", GetCellID(fps, lCol, lRow + lRow2s)
    Next
    
    ' Insert new node to DOM object
    If Not xmlNodeCells.nextSibling Is Nothing Then
        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, xmlNodeCells.nextSibling
    Else
        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, Null
    End If
'   **********************************
'    added
'   Date: 12/04/06
    'mAdjustData = True
    TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
'   **********************************
EXIT_SUB:
    Set xmlNodeNewCell = Nothing
    Set xmlNodeNewCells = Nothing
    Set xmlNodeCells = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "04_GTGT", "InsertNode", Err.number, Err.Description
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
                        If Not TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID( _
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
                        If Not TAX_Utilities_New.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, i, pRow - lRowCtrl)) Is Nothing Then
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
    SaveErrorLog "04_GTGT", "InsertRow", Err.number, Err.Description
End Sub

Private Sub IncreaseRowInDOM(fpSpread1 As fpSpread, xmlDomData As MSXML.DOMDocument, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, i As Long
        
    If xmlDomData Is Nothing Then Exit Sub
    Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
    
    For i = xmlNodeListCell.length - 1 To 0 Step -1
        ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
        If lRow >= pRow Then
            ' Increase value of row attribute + 1 (CellID)
            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fpSpread1, lCol, lRow + lRows)
            
            ' Increase value of row attribute + 1 (CellID2)
            ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fpSpread1, lCol, lRow + lRow2s)
        End If
    Next
        
    Set xmlNodeListCell = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "04_GTGT", "IncreaseRowInDOM", Err.number, Err.Description
End Sub


Private Sub DeleteNode(ByVal intSheet As Integer, ByVal pCol As Long, ByVal pRow As Long, Optional ByVal blnForce As Boolean = True)
    On Error GoTo ErrorHandle
    Dim lLRowBound As Long, lURowBound As Long
    Dim lRow2s As Long, lRows As Long
    Dim xmlNodeCells As MSXML.IXMLDOMNode, xmlNodeTemp As MSXML.IXMLDOMNode
    Dim xmlNodeTemp2 As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    
    GetCellSpan fps, pCol, pRow
    
    Set xmlNodeCells = TAX_Utilities_New.Data(intSheet - 1).nodeFromID(GetCellID(fps, pCol, pRow)).parentNode
    
    lRows = GetDynRowCount(fps, xmlNodeCells, lRow2s, lLRowBound, lURowBound)

    If GetAttribute(xmlNodeCells.parentNode, "Dynamic") <> "1" Then GoTo EXIT_SUB
    
'*********************************************************
' added
'Date: 01/03/06
    'Check whether user want to delete
'    If lRows > 1 And blnForce And Not IsEmptyValue(xmlNodeCells) Then
'        If DisplayMessage("0075", msYesNo, miQuestion, , mrYes) = mrNo Then
'            Exit Sub
'        End If
'    End If
'*********************************************************
        
    If xmlNodeCells.parentNode.childNodes.length <= 1 Then
        ClearRows xmlNodeCells
        TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
        GoTo EXIT_SUB
    Else
        If GetAttribute(xmlNodeCells.firstChild, "FirstCell") = "0" Then
            'Set FirstCell attr to "0" for next Cells node
            Set xmlNodeCells = xmlNodeCells.nextSibling
            SetAttribute xmlNodeCells.firstChild, "FirstCell", "0"
            Set xmlNodeCells = xmlNodeCells.previousSibling
        End If
    End If
    
    'If xmlNodeCells.xml = xmlNodeCells.parentNode.childNodes(0).xml Then GoTo EXIT_SUB
    
    'Jump active cell to prevous section
    'fpSpread1.SetActiveCell fpSpread1.ActiveCol, fpSpread1.ActiveRow - lRows
    
    ' Delete curent row on Form
    DeleteRow intSheet, lLRowBound, lRows
    'fpSpread1.SetFocus
    
    xmlNodeCells.parentNode.removeChild xmlNodeCells
    
    ' Decrease value of row in xmlDocument
    DecreaseRowInDOM intSheet, lLRowBound + 1, lRows, lRow2s
    
    fps.Col = fps.ActiveCol
    fps.Row = fps.ActiveRow
    If fps.Lock = True Then
        Do
            fps.Row = fps.Row - 1
        Loop Until (Not fps.Lock) Or (fps.Row = 1)
        fps.SetActiveCell fps.Col, fps.Row
    End If
'   ************************************
'    added
'   Date: 12/04/06
    TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
'   ************************************
EXIT_SUB:
    Set xmlNodeCells = Nothing
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "04_GTGT", "DeleteNode", Err.number, Err.Description
End Sub


Private Sub ClearRows(xmlCellsNode As MSXML.IXMLDOMNode) '(ByVal lRow As Long, ByVal lRows As Long)
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode

    With fps
        For Each xmlCellNode In xmlCellsNode.childNodes
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
            .Sheet = mCurrentSheet
            .Col = lCol
            .Row = lRow
            If .Lock = False Or (.Lock = True And .Formula = vbNullString) Then
                Select Case .CellType
                    Case CellTypeNumber
                        .Text = "0"
                        SetAttribute xmlCellNode, "Value", "0"
                    Case Else
                        .Text = ""
                        SetAttribute xmlCellNode, "Value", ""
                End Select
           End If
                
            
            If .CellNote <> vbNullString Then
                .CellNote = vbNullString
                .BackColor = vbWhite
            End If
            '*******************************
        Next
    End With
End Sub


Private Sub DeleteRow(ByVal intSheet As Integer, ByVal pRow As Long, ByVal lRows As Long)
    On Error GoTo ErrorHandle
    
    With fps
        .EventEnabled(EventChange) = False
        .ReDraw = False
        '.Visible = False
        .Sheet = intSheet
        .DeleteRows pRow, lRows
        .MaxRows = .MaxRows - lRows
        '.Visible = True
        .ReDraw = True
        .EventEnabled(EventChange) = True
    End With
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "04_GTGT", "DeleteRow", Err.number, Err.Description
End Sub


Private Sub DecreaseRowInDOM(ByVal intSheet As Integer, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, i As Long
    
    If TAX_Utilities_New.Data(intSheet - 1) Is Nothing Then Exit Sub
    Set xmlNodeListCell = TAX_Utilities_New.Data(intSheet - 1).getElementsByTagName("Cell")
    
    For i = 0 To xmlNodeListCell.length - 1
        ParserCellID fps, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
        If lRow >= pRow Then
            ' Decrease value of row attribute - 1 "CellID"
            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fps, lCol, lRow - lRows)
            
            ' Decrease value of row attribute - 1 "CellID2"
            ParserCellID fps, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fps, lCol, lRow - lRow2s)
        End If
    Next
    
    Set xmlNodeListCell = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "04_GTGT", "DecreaseRowInDOM", Err.number, Err.Description
End Sub


' Kiem tra xem neu la group II, III, IV ma bam F5, F6 se khong them dong
Public Function isAddRow(ByVal iCol As Integer, ByVal iRow As Integer) As Boolean
    Dim varTemp  As Variant
    With fps
        .Sheet = 1
        ' Group 1
        .Col = .ColLetterToNumber("AF")
        .Row = iRow
        .GetText .Col, .Row, varTemp
        If Trim$(CStr(varTemp)) = "F" Then
            isAddRow = False
        Else
            isAddRow = True
        End If
    End With
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

' Dieu chinh to khai BS
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
            Set xmlNode = .nodeFromID("L_25")
            result = "CT30_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("R_25")
            result = result & "~" & "CT31_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("X_25")
            result = result & "~" & "CT32_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_04GTGT", "getValueTK", Err.number, Err.Description
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
                .GetText .ColLetterToNumber("L"), 25, varTemp
                strCurrentValue = "CT30_" & Trim(varTemp)
                .GetText .ColLetterToNumber("R"), 25, varTemp
                strCurrentValue = strCurrentValue & "~CT31_" & Trim(varTemp)
                .GetText .ColLetterToNumber("X"), 25, varTemp
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
    SaveErrorLog "cls_04GTGT", "getMapCT", Err.number, Err.Description
End Function

' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("E"), 25, varTemp
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
            Set xmlNode = .nodeFromID("E_25")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_04GTGT", "getValueCTDC", Err.number, Err.Description
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
    tuRowDL = 13
    denRowDL = 18
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "20_30_2"
    ' end
End Sub

