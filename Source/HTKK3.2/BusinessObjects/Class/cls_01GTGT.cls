VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01GTGT"
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

Private Const Row_MSTDN = 7
Private Const Col_MSTDN = "D"

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
Public strLoaiNNKD As String

' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String




'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'no parameter
Public Sub Prepare1()
Dim strComboHien As Variant
Dim ThangKK As Integer
Dim NamKK As Integer
Dim Thongtu As String
Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
Dim strDataFileName As String
Dim xmlNodeListItems As MSXML.IXMLDOMNodeList
Dim strCboLyDo As String
Dim strCboLyDoAn As String
Dim xmlNode As MSXML.IXMLDOMNode
Dim i As Integer
Dim fldList() As String



With fps
'  'danh sach  list cua combobox
'    strComboHien = "0%" + Chr$(9) + "5%" + Chr$(9) + "10%" + Chr$(9)
    '--
    SetDateFormat fps, 1, 53, .ColLetterToNumber("K"), DDMMYYYY
    
    SetDateFormat fps, 7, 6, .ColLetterToNumber("D"), DDMMYYYY
    
    SetDateFormat fps, 9, 23, .ColLetterToNumber("BG"), DDMMYYYY
    .Sheet = 1 'To khai GTGT
        .Row = 53
        .Col = .ColLetterToNumber("K")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
    .Sheet = 2 'PL01_1
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



        .Row = 23
        .Text = ""
        .CellTag = .CellTag & "~4"
        
        
        
        .Row = 28
        .Text = ""
        .CellTag = .CellTag & "~5"
        

    
    .Sheet = 3 ''PL01_1
        .Row = 8
        .Col = .ColLetterToNumber("C")
        .Text = ""
        .CellTag = .CellTag & "~1"
 
 
'        .Col = .ColLetterToNumber("J")
'        .TypeComboBoxList = strComboHien
        
        .Row = 13
        .Text = ""
        .CellTag = .CellTag & "~2"

'        .Col = .ColLetterToNumber("J")
'        .TypeComboBoxList = strComboHien
        
        .Row = 18
        .Text = ""
        .CellTag = .CellTag & "~3"
        
'        .Col = .ColLetterToNumber("J")
'        .TypeComboBoxList = strComboHien
        
        .Row = 23
        .Text = ""
        .CellTag = .CellTag & "~4"

'        .Col = .ColLetterToNumber("J")
'        .TypeComboBoxList = strComboHien
        
        .Row = 28
        .Text = ""
        .CellTag = .CellTag & "~5"
   
'        .Col = .ColLetterToNumber("J")
'        .TypeComboBoxList = strComboHien
    .Sheet = 9 'To khai GTGT
        .Row = 23
        .Col = .ColLetterToNumber("BG")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
   
    SetDateFormat fps, 2, 8, .ColLetterToNumber("E"), DDMMYYYY
    SetDateFormat fps, 2, 13, .ColLetterToNumber("E"), DDMMYYYY
    SetDateFormat fps, 2, 18, .ColLetterToNumber("E"), DDMMYYYY
    SetDateFormat fps, 2, 23, .ColLetterToNumber("E"), DDMMYYYY
    SetDateFormat fps, 2, 28, .ColLetterToNumber("E"), DDMMYYYY
    
    SetDateFormat fps, 3, 8, .ColLetterToNumber("E"), DDMMYYYY
    SetDateFormat fps, 3, 13, .ColLetterToNumber("E"), DDMMYYYY
    SetDateFormat fps, 3, 18, .ColLetterToNumber("E"), DDMMYYYY
    SetDateFormat fps, 3, 23, .ColLetterToNumber("E"), DDMMYYYY
    SetDateFormat fps, 3, 28, .ColLetterToNumber("E"), DDMMYYYY
    
    
    SetDateFormat fps, 7, 16, .ColLetterToNumber("O"), DDMMYYYY
    
    ' Set gia tri check gia han cho to khai 01/GTGT
    .Sheet = 1
    ThangKK = TAX_Utilities_New.Month
    NamKK = TAX_Utilities_New.Year
    If NamKK = 2012 And (ThangKK = 4 Or ThangKK = 5 Or ThangKK = 6) Then
        Thongtu = GetAttribute(GetMessageCellById("0262"), "Msg")
        .SetText .ColLetterToNumber("C"), 18, Thongtu
        
        ' Ky dc gia han moi set gia tri ly do gia han
        strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_DM_Lydo_GH.xml")
        If xmlDomData.Load(strDataFileName) Then
             Set xmlNodeListItems = xmlDomData.getElementsByTagName("Item")
             For Each xmlNode In xmlNodeListItems
                 fldList = Split(GetAttribute(xmlNode, "Value"), "###")
                 strCboLyDo = strCboLyDo + fldList(1) + Chr$(9)
                 strCboLyDoAn = strCboLyDoAn + fldList(0) + Chr$(9)
             Next
         End If
              'set ly do gia han
            .Row = 20
            .Col = .ColLetterToNumber("D")
            .TypeComboBoxList = strCboLyDo
            ' set Ma ly do gia han
            .Col = .ColLetterToNumber("P")
            .TypeComboBoxList = strCboLyDoAn
'    ElseIf NamKK = 2013 And (ThangKK = 1 Or ThangKK = 2 Or ThangKK = 3) Then
'         Thongtu = GetAttribute(GetMessageCellById("0263"), "Msg")
'        .SetText .ColLetterToNumber("C"), 18, Thongtu
'        .Row = 20
'        .RowHidden = True
    Else
'         .SetText .ColLetterToNumber("C"), 18, ""
'        .Row = 18
'        .Col = .ColLetterToNumber("B")
''        .CellType = CellTypeEdit
'
'        .Text = ""
'        .Lock = True
'        .Row = 18
'        .RowHidden = True
'        .Row = 20
'        .RowHidden = True
        Thongtu = GetAttribute(GetMessageCellById("0263"), "Msg")
        .SetText .ColLetterToNumber("C"), 18, Thongtu
        ' Ky dc gia han moi set gia tri ly do gia han
        strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_DM_Lydo_GH_TT16.xml")
        If xmlDomData.Load(strDataFileName) Then
             Set xmlNodeListItems = xmlDomData.getElementsByTagName("Item")
             For Each xmlNode In xmlNodeListItems
                 fldList = Split(GetAttribute(xmlNode, "Value"), "###")
                 strCboLyDo = strCboLyDo + fldList(1) + Chr$(9)
                 strCboLyDoAn = strCboLyDoAn + fldList(0) + Chr$(9)
             Next
         End If
              'set ly do gia han
        .Row = 20
        .Col = .ColLetterToNumber("D")
        .TypeComboBoxList = strCboLyDo
        ' set Ma ly do gia han
        .Col = .ColLetterToNumber("P")
        .TypeComboBoxList = strCboLyDoAn
    End If
    
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
            LockCellBySheet
            CellLock .ColLetterToNumber(S1ChkCol), S1ChkRow
            mblKhoiDauToKhai = False
            'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("K")
            .Row = 55
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("L")
            .Row = 55
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("O")
            .Row = 55
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("K")
            .Row = 55
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("L")
            .Row = 55
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("O")
            .Row = 55
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        'ma NNKD
        .Col = .ColLetterToNumber("L")
        .Row = 14
        .Text = strLoaiNNKD
        UpdateCell fps, .Col, .Row, .Text
        
        
        
        ' Neu check gia han moi cho phep chon GT
        .Col = .ColLetterToNumber("B")
        .Row = 18
        If .Text = "1" Or UCase$(.Text) = "X" Then
            .Col = .ColLetterToNumber("D")
            .Row = 20
            .Lock = False
        Else
            .Col = .ColLetterToNumber("D")
            .Row = 20
            .Lock = True
        End If
        ' End set gia han
            
        .Sheet = 2 'PL1
            .ActiveSheet = .Sheet
                .Row = 1
                .Col = .ColLetterToNumber("C")
                .CellType = CellTypeButton
                .TypeButtonTextColor = RGB(0, 0, 255)
                .TypeButtonColor = &H8000000F
                .TypeButtonAlign = TypeButtonAlignRight
                .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
        .Sheet = 3 'PL2
             .ActiveSheet = .Sheet
                .Row = 1
                .Col = .ColLetterToNumber("C")
                .CellType = CellTypeButton
                .TypeButtonTextColor = RGB(0, 0, 255)
                .TypeButtonColor = &H8000000F
                .TypeButtonAlign = TypeButtonAlignRight
                .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
        .Sheet = 4 'PL3
'            .ActiveSheet = .Sheet
'            Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
'            Set xmlCellNode = TAX_Utilities_New.Data(4 - 1).nodeFromID(GetCellID(fps, fps.ColLetterToNumber("F"), 2))
'            mCurrentSheet = 4
'            fps.Sheet = mCurrentSheet
'            If GetAttribute(xmlCellNode, "Value") = "" Then
'                .Col = fps.ColLetterToNumber(S4TxtF2Col)
'                .Row = S4TxtF2row
'                .Text = TAX_Utilities_New.Year - 1
'                .TypeHAlign = TypeHAlignLeft
'                UpdateCell fps, .Col, .Row, .Text
'            End If
            
'         .Sheet = 9 'KHBS
'            Dim cCol As Long, cRow As Long
'          '  .ActiveSheet = .Sheet
'            mCurrentSheet = 9
'            fps.Sheet = mCurrentSheet
'            .Row = 23
'            .Col = .ColLetterToNumber("BG")
'            .Text = Format(Date, "dd/mm/yyyy")
        .ActiveSheet = 1
        mCurrentSheet = tempCurrSheet
         .EventEnabled(EventAllEvents) = True
    End With
    
    
    'Get last [43] target and assign to current [11] target
    setLast43Target
    
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
    If .ActiveSheet = 1 Then
        'Check value of 11
        GetCellSpan fps, Col, Row
        If Col = .ColLetterToNumber("L") And Row = 24 Then
            .GetText Col, Row, varTemp
            If GetLast43Previous(str43Previous) Then
                If CStr(varTemp) <> str43Previous Then
                    .Visible = True
                    DisplayMessage "0060", msOKOnly, miWarning
                    fps.SetFocus
                End If
            End If
        End If
    End If

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
    LockCellBySheet
End Sub

'inner class event: ButtonClicked
Public Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim lResult As Integer
    Dim blTemp As Boolean
    
    fps.EventEnabled(EventAllEvents) = False
    
    If Col = fps.ColLetterToNumber(S1ChkCol) And Row = S1ChkRow Then
       If mblKhoiDauToKhai = False Then
            fps.Sheet = fps.ActiveSheet
            mCurrentSheet = fps.Sheet
            fps.Row = Row
            fps.Col = Col
            blTemp = TAX_Utilities_New.AdjustData(fps.ActiveSheet - 1)
            If fps.value = 1 Then
                lResult = DisplayMessage("0034", msYesNo, miQuestion, , mrNo)
                If lResult = 3 Then '<>6
                    fps.value = 0
                    UpdateCell fps, fps.Col, fps.Row, IIf(fps.value = 1, "x", vbNullString) 'fps.Value
                    TAX_Utilities_New.AdjustData(fps.ActiveSheet - 1) = blTemp
                    fps.EventEnabled(EventAllEvents) = True
                    Exit Sub
                Else
                    UpdateCell fps, fps.Col, fps.Row, IIf(fps.value = 1, "x", vbNullString) 'fps.Value
                    TAX_Utilities_New.AdjustData(fps.ActiveSheet - 1) = True
                End If
            ElseIf fps.value = 0 Then
                UpdateCell fps, fps.Col, fps.Row, IIf(fps.value = 1, "x", vbNullString)
                TAX_Utilities_New.AdjustData(fps.ActiveSheet - 1) = True
                ' Khong check se lay lai gia tri cua phu luc len
                SetValueOCheck
                
            End If
        End If
        
        CellLock Col, Row
    End If
    
    ' set gia lock cell ly do gia han
    If fps.ActiveSheet = 1 Then
            ' Neu check gia han moi cho phep chon GT
            If Col = fps.ColLetterToNumber("B") And Row = 18 Then
                fps.Col = fps.ColLetterToNumber("B")
                fps.Row = 18
                If fps.Text = "1" Or UCase$(fps.Text) = "X" Then
                    fps.Col = fps.ColLetterToNumber("D")
                    fps.Row = 20
                    fps.Lock = False
                    If Strloaitk = "TKBS" Then
                        TAX_Utilities_New.CheckToKhaiGH = True
                        TinhHanNop_KHBS fps
                        fps.Sheet = 1
                    End If
                Else
                    fps.Col = fps.ColLetterToNumber("D")
                    fps.Row = 20
                    fps.Lock = True
                    fps.SetText fps.ColLetterToNumber("D"), 20, ""
                    fps.SetText fps.ColLetterToNumber("P"), 20, "00"
                    UpdateCell fps, fps.ColLetterToNumber("P"), 20, "00"
                    UpdateCell fps, fps.ColLetterToNumber("D"), 20, ""
                    If Strloaitk = "TKBS" Then
                        TAX_Utilities_New.CheckToKhaiGH = False
                        TinhHanNop_KHBS fps
                        fps.Sheet = 1
                    End If
                End If
            End If
    End If
    
    fps.EventEnabled(EventAllEvents) = True
    
    'CellChange Col, Row
    
End Sub

Private Sub fps_EditChange(ByVal Col As Long, ByVal Row As Long)
    CellChange1 Col, Row
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
Private Sub fps_KeyPress(KeyAscii As Integer)
     With fps
        If .ActiveSheet = 6 Then
            If .Col = .ColLetterToNumber("F") And .Row = 2 Then
                If (KeyAscii < 48 And KeyAscii <> 3 And KeyAscii <> 8 And KeyAscii <> 22 And KeyAscii <> 13) Or KeyAscii > 57 Then
                    KeyAscii = 0
                End If
            End If
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
                         .Col = .ColLetterToNumber("J")
                         .Text = "0%"
                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("J")
                         .Text = "5%"
                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "cc"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("J")
                         .Text = "10%"
                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "dd"
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                     
                    .SetActiveCell iCol, iRow
                    CellChange2
                    CellChange3
                    LockCellBySheet
            End If
            If .ActiveSheet = 3 Then
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 8
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("J")
'                         .CellType = CellTypePercent
'                         .TypePercentDecPlaces = 0
'                         .TypePercentMax = 100
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 8
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("J")
'                         .CellType = CellTypePercent
'                         .TypePercentDecPlaces = 0
'                         .TypePercentMax = 100
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("J")
'                         .CellType = CellTypePercent
'                         .TypePercentDecPlaces = 0
'                         .TypePercentMax = 100
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "cc"
                    
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "dd"
                    
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "ee"
                    
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
                    'CellChange iCol, iRow
                    CellChange2
                    CellChange3
                    LockCellBySheet
            End If
            
            If .ActiveSheet = 4 Then
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 8
                    Do
                         .Text = str(i)
                         .GetText .ColLetterToNumber("D"), .Row, varTemp
                         If varTemp = vbNullString Or Trim(varTemp) = "" Then
                            .SetText .ColLetterToNumber("D"), .Row, TAX_Utilities_New.Convert("ChiÕc", TCVN, UNICODE)
                         End If
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 8
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 4
                     Do
                          .Text = str(i)
                                 .GetText .ColLetterToNumber("D"), .Row, varTemp
                                 If varTemp = vbNullString Or Trim(varTemp) = "" Then
                                    .SetText .ColLetterToNumber("D"), .Row, TAX_Utilities_New.Convert("ChiÕc", TCVN, UNICODE)
                                 End If
                          .Col = .ColLetterToNumber("B")
                          .Row = 1 + .Row
                          i = i + 1
                     Loop Until .Text = "bb"
                     i = 1
                    .Row = .Row + 4
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
            End If
            If .ActiveSheet = 7 Then
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
                    .SetActiveCell iCol, iRow
            End If
            If .ActiveSheet = 8 Then
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 8
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 8
                         i = i + 1
                    Loop Until .Text = "aa"
                    .SetActiveCell iCol, iRow

                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
            End If
            
'            If .ActiveSheet = .SheetCount - 1 Then
'                 fps_Change iCol, iRow
'                 fps.EventEnabled(EventAllEvents) = False
'                    i = 1
'                    .Col = .ColLetterToNumber("B")
'                    .Row = 9
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("B")
'                         .Row = i + 9
'                         i = i + 1
'                    Loop Until .Text = "aa"
'                    i = 1
'                   .Row = .Row + 3
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("B")
'                         .Row = 1 + .Row
'                         i = i + 1
'                    Loop Until .Text = "bb"
'                    .SetActiveCell iCol, iRow
'                    .Row = iRow
'                    .Col = iCol
'                    .SetActiveCell iCol, iRow
'                 fps.EventEnabled(EventAllEvents) = True
'            End If
            
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
             If Col = .ColLetterToNumber("K") And Row = 50 Then
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
        If .ActiveSheet = 2 Or .ActiveSheet = 3 Then
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
        
        If .ActiveSheet = 4 Then
'                   If Col = .ColLetterToNumber("C") Then
'                       .GetText Col, Row, varTemp
'                       If varTemp <> "" And varTemp <> "../...." Then
'                           If Format_mmyyyy(CStr(varTemp)) <> "" Then
'                               .SetText Col, Row, Format_mmyyyy(CStr(varTemp))
'                           Else
'                               .SetActiveCell Col, Row
'                           End If
'                       Else
'                        .SetText Col, Row, ""
'                       End If
'                       .Col = Col
'                       .Row = Row
'                      UpdateCell fps, .Col, .Row, .Text
'                   End If
'                   If Col = .ColLetterToNumber("D") Then
'                       .GetText Col, Row, varTemp
'                       If varTemp <> "" And varTemp <> "../../...." Then
'                           If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
'                               .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                           Else
'                               .SetActiveCell Col, Row
'                           End If
'                       Else
'                        .SetText Col, Row, ""
'                       End If
'                       .Col = Col
'                       .Row = Row
'                      UpdateCell fps, .Col, .Row, .Text
'                   End If
        End If
        
'        If .ActiveSheet = 5 Then
'            .Col = Col
'            .Row = Row
'            UpdateCell fps, .Col, .Row, .Text
'        End If
'
        If .ActiveSheet = 7 Then
                   If Col = .ColLetterToNumber("D") Then
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
                      UpdateCell fps, .Col, .Row, .Text
                   End If
        End If
        If .ActiveSheet = 9 Then
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
    LockCellBySheet
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
    Dim ct21 As Variant
    Dim blCheck_S1 As Boolean
    Dim blCheck_S2 As Boolean
    Dim blCheck_S5 As Boolean
    Dim blCheck_S4A As Boolean
    Dim blCheck_S4B As Boolean
    blCheck_S1 = IIf(TAX_Utilities_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S5 = IIf(TAX_Utilities_New.NodeValidity.childNodes(6).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S4A = IIf(TAX_Utilities_New.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S4B = IIf(TAX_Utilities_New.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    tempCurrSheet = mCurrentSheet
    .Sheet = 1
    ' Chi tieu 21
    .Col = .ColLetterToNumber("I")
    .Row = 23
    ct21 = .Text
'       If ct21 = "0" Then
'            If blCheck_S2 = True And ct21 = "0" Then
'                .Sheet = 1
'                ' Chi tieu 23
'                .Col = .ColLetterToNumber("J")
'                .Row = 27
'                '.Lock = True
'                .Formula = "D53"
'                .BackColor = vbWhite
'                ' Chi tieu 24
'                .Col = .ColLetterToNumber("L")
'                .Row = 27
'                '.Lock = True
'                .Formula = "E53"
'                .BackColor = vbWhite
        
                ' shet bang ke 01-4A
                ' chi tieu tong dong 2 cot 10
'                If blCheck_S4a = True And ct21 = "0" Then
'                    .Sheet = 5
'                    .Col = .ColLetterToNumber("D")
'                    .Row = 9
'                    .Formula = "D23"
'                 '   .Lock = True
'                    ' chi tieu tong dong 1 cot 10
'                    .Col = .ColLetterToNumber("D")
'                    .Row = 7
'                    .Formula = "C23"
'                  '  .Lock = True
'                    ' chi tieu tong dong 3 cot 10
'                    .Col = .ColLetterToNumber("D")
'                    .Row = 11
'                    .Formula = "E23"
'                   ' .Lock = True
'                    ' chitieu B4
'                    .Col = .ColLetterToNumber("D")
'                    .Row = 19
'                    .Formula = "E24"
'                    '.Lock = True
'                End If
                
                '        ' set du lieu cho sheet 6
                '        If blCheck_S4a = True Then
                '            .Sheet = 6
                '            ' chi tieu A1
                '            .Col = .ColLetterToNumber("E")
                '            .Row = 7
                '            .Formula = "D24"
                '            ' chi tieu A2
                '            .Col = .ColLetterToNumber("E")
                '            .Row = 9
                '            .Formula = "E24"
                '            ' chi tieu A3
                '            .Col = .ColLetterToNumber("E")
                '            .Row = 11
                '            .Formula = "F24"
                '            ' chi tieu B4
                '            .Col = .ColLetterToNumber("E")
                '            .Row = 19
                '            .Formula = "F25"
                '
                '        End If
''                        .Sheet = 1
'                        ' chi tieu 23
'                        .Col = .ColLetterToNumber("J")
'                        .Row = 27
'                     '   .Lock = False
'                        .Formula = ""
'                        .BackColor = vbWhite
'                        ' chi tieu 24
'                        .Col = .ColLetterToNumber("L")
'                        .Row = 27
'                      '  .Lock = False
'                        .Formula = ""
'                        .BackColor = vbWhite
'                        ' shet bang ke 01-4A
'                        If blCheck_S4a = True And ct21 = "0" Then
'                            .Sheet = 5
'                            .Col = .ColLetterToNumber("D")
'                            .Row = 9
'                            .Formula = ""
'                       '     .Lock = False
'                            ' chi tieu tong dong 1 cot 10
'                            .Col = .ColLetterToNumber("D")
'                            .Row = 7
'                            .Formula = ""
'                        '    .Lock = False
'                            ' chi tieu tong dong 3 cot 10
'                            .Col = .ColLetterToNumber("D")
'                            .Row = 11
'                            .Formula = ""
'                         '   .Lock = False
'                            ' chitieu B4
'                            .Col = .ColLetterToNumber("D")
'                            .Row = 19
'                            .Formula = "E24"
'                          '  .Lock = False
'                        End If
                
                '        ' set du lieu cho sheet 6
                '        If blCheck_S4b = True Then
                '            .Sheet = 6
                '            ' chi tieu A1
                '            .Col = .ColLetterToNumber("E")
                '            .Row = 7
                '            .Formula = ""
                '            ' chi tieu A2
                '            .Col = .ColLetterToNumber("E")
                '            .Row = 9
                '            .Formula = ""
                '            ' chi tieu A3
                '            .Col = .ColLetterToNumber("E")
                '            .Row = 11
                '            .Formula = ""
                '             ' chi tieu B4
                '            .Col = .ColLetterToNumber("E")
                '            .Row = 19
                '            .Formula = ""
                '
                '        End If
                '
                
'                    End If
'                    ' kiem tra sheet phu luc 01-1
'                    If blCheck_S1 = True And ct21 = "0" Then
'                        .Sheet = 1
'                     '   If .ActiveSheet = 2 Then
'                            ' chi tieu 29
'                        .Col = .ColLetterToNumber("J")
'                        .Row = 32
'                        '.Lock = True
'                        .Formula = "H53"
'                        ' chi tieu 30
'                        .Col = .ColLetterToNumber("J")
'                        .Row = 33
'                        '.Lock = True
'                        .Formula = "I53"
'                        ' chi tieu 32
'                        .Col = .ColLetterToNumber("J")
'                        .Row = 34
'                        '.Lock = True
'                        .Formula = "K53"
'                        '.Lock = True
'                        'chi tieu 31
'                        .Col = .ColLetterToNumber("L")
'                        .Row = 33
'                        .Formula = "J53"
'                        '.Lock = True
'                        ' chi tieu 33
''                        .Col = .ColLetterToNumber("L")
''                        .Row = 34
''                        .Formula = "L53"
'
'                        ' chi tieu 26
'                        .Col = .ColLetterToNumber("J")
'                        .Row = 30
'                        .Formula = "E54"
'                        '.Lock = True
                        ' set du lieu tren sheet 5
'                        If blCheck_S4a = True And ct21 = "0" Then
'                            .Sheet = 5
'                            ' chi tieu tong doanh so ban ra tren phu luc 1
'                            .Col = .ColLetterToNumber("D")
'                            .Row = 16
'                            .Formula = "C24"
'                         '   .Lock = True
'                            ' chi tieu T?ng doanh thu hàng hóa, d?ch v? bán ra ch?u thu? GTGT
'                            .Col = .ColLetterToNumber("D")
'                            .Row = 17
'                            .Formula = "D24"
'                          '  .Lock = True
'                        End If
                
                '        ' set du lieu tren sheet 6
                '        If blCheck_S4b = True Then
                '            .Sheet = 6
                '             ' chi tieu B1
                '            .Col = .ColLetterToNumber("E")
                '            .Row = 16
                '            .Formula = "D25"
                '             ' chi tieu B2
                '            .Col = .ColLetterToNumber("E")
                '            .Row = 17
                '            .Formula = "E25"
                '
                '        End If
                '
                      '   End If
'                    Else
'                        .Sheet = 1
'                        ' chi tieu 29
'                        .Col = .ColLetterToNumber("J")
'                        .Row = 32
'                        '.Lock = False
'                        .Formula = ""
'                         ' chi tieu 30
'                        .Col = .ColLetterToNumber("J")
'                        .Row = 33
'                        '.Lock = False
'                        .Formula = ""
'                        ' chi tieu 32
'                        .Col = .ColLetterToNumber("J")
'                        .Row = 34
'                        '.Lock = False
'                        .Formula = ""
'                        ' chi tieu 31
'                        .Col = .ColLetterToNumber("L")
'                        .Row = 33
'                        .Formula = ""
'                        ' chi tieu 33
'                        .Col = .ColLetterToNumber("L")
'                        .Row = 34
'                        .Formula = ""
'
'                        ' chi tieu 26
'                        .Col = .ColLetterToNumber("J")
'                        .Row = 30
'                        .Formula = ""
                
                        ' set du lieu tren sheet 5
'                        If blCheck_S4a = True And ct21 = "0" Then
'                            .Sheet = 5
'                            ' chi tieu tong doanh so ban ra tren phu luc 1
'                            .Col = .ColLetterToNumber("D")
'                            .Row = 16
'                            .Formula = ""
'                         '   .Lock = False
'                            ' chi tieu T?ng doanh thu hàng hóa, d?ch v? bán ra ch?u thu? GTGT
'                            .Col = .ColLetterToNumber("D")
'                            .Row = 17
'                            .Formula = ""
'                          '  .Lock = False
'                        End If
                '        ' set du lieu tren sheet 6
                '        If blCheck_S4b = True Then
                '            .Sheet = 6
                '             ' chi tieu B1
                '            .Col = .ColLetterToNumber("E")
                '            .Row = 16
                '            .Formula = ""
                '             ' chi tieu B2
                '            .Col = .ColLetterToNumber("E")
                '            .Row = 17
                '            .Formula = ""
                '
                '        End If
                
                
'                    End If
                '
                ''    ' phu luc 01-4a
                ''    If blCheck_S4a = True Then
                ''        If blCheck_S4b = True Then
                ''            .Sheet = 6
                ''            ' chi tieu B6
                ''            .Col = .ColLetterToNumber("E")
                ''            .Row = 21
                ''            .Formula = "D26"
                ''
                ''        End If
                ''    Else
                ''        If blCheck_S4b = True Then
                ''            .Sheet = 6
                ''            ' chi tieu B6
                ''            .Col = .ColLetterToNumber("E")
                ''            .Row = 21
                ''            .Formula = ""
                ''
                ''        End If
                ''
                ''    End If
                '
                    ' kiem tra sheet phu luc 01-5
'                    If blCheck_S5 = True And ct21 = "0" Then
'                        .Sheet = 1
'                        ' chi tieu 39
'                        .Col = .ColLetterToNumber("L")
'                        .Row = 40
'                        .Formula = "D54"
'                        '.Lock = True
'                    Else
'                        .Sheet = 1
'                        ' chi tieu 39
'                        .Col = .ColLetterToNumber("L")
'                        .Row = 40
'                        .Formula = ""
'                        '.Lock = False
'
'                    End If
                    
'                    ' kiem tra phu luc 01-2,01-4A,01-4B
'                    If blCheck_S2 = True Or blCheck_S4a = True Or blCheck_S4b = True And ct21 = "0" Then
'                        .Sheet = 1
'                        ' chi tieu 25
'                        .Col = .ColLetterToNumber("L")
'                        .Row = 28
'                        .Formula = "F53"
'                        .BackColor = vbWhite
'                        '.Lock = True
'                    Else
'                        .Sheet = 1
'                        ' chi tieu 25
'                        .Col = .ColLetterToNumber("L")
'                        .Row = 28
'                        .Formula = ""
'                        .BackColor = vbWhite
'                        '.Lock = False
'                    End If
                    ' Rang buoc [40]-[42]
'                    .Sheet = 1
'                    ' chi tieu 40
'                    .Col = .ColLetterToNumber("L")
'                    .Row = 44
'                    If Val(.Text) > 0 Then
'                        .Sheet = 1
'                        ' chi tieu 42
'                        .Col = .ColLetterToNumber("L")
'                        .Row = 46
'                        '.Lock = True
'                    Else
'                        .Sheet = 1
'                        ' chi tieu 42
'                        .Col = .ColLetterToNumber("L")
'                        .Row = 46
'                        '.Lock = False
'                    End If
    'End If
   ' mCurrentSheet = tempCurrSheet
    
        .Sheet = .ActiveSheet
    Select Case .Sheet
        Case 1 ' To khai GTGT
            mCurrentSheet = .Sheet
            GetCellSpan fps, Col, Row
            If Row = S1ChkRow And Col = .ColLetterToNumber(S1ChkCol) Then
                '14 J11;15 L11;16 J12;17 L12;26 J20;29 J22;30 J23;32 J24; 31 L23; 33 L24
                '14, 15, 16, 17, 26, 29, 30,31 32 và 33
   'htphuong edit 15/12/2007
'                ArrRow = Array(11, 11, 12, 12, 20, 22, 23, 24, 23, 24, 23, 24)
'                ArrCol = Array("J", "L", "J", "L", "J", "J", "J", "J", "L", "L", "L", "L")
   'htphuong edit 03/03/2008
                'ArrRow = Array(27, 27, 28, 30, 31, 31, 32, 33, 33, 34, 34, 35, 35)
                'ArrCol = Array("J", "L", "L", "J", "J", "L", "J", "J", "L", "J", "L", "J", "L")
'                ArrRow = Array(27, 27, 28, 30, 31, 31, 32, 33, 33, 34, 34, 35, 35)
'                ArrCol = Array("J", "L", "L", "J", "J", "L", "J", "J", "L", "J", "L", "J", "L")
'               .Col = .ColLetterToNumber("I")
'                .Row = 23
'                'UpdateCell fps, .Col, .Row, .Text & "chk"
'
'                If .Text = "1" And Strloaitk <> "TKBS" Then
'                    For i = 0 To UBound(ArrRow)
'                        .Row = ArrRow(i)
'                        .Col = .ColLetterToNumber(ArrCol(i))
'                        .Text = "0"
'                        UpdateCell fps, .Col, .Row, .Text
'                        .Lock = True
'                    Next
'                ElseIf .Text = "0" Then
'                    For i = 0 To UBound(ArrRow)
'                        .Row = ArrRow(i)
'                        .Col = .ColLetterToNumber(ArrCol(i))
'                        If (ArrCol(i) = "J" Or ArrCol(i) = "L") And (ArrRow(i) = 31 Or ArrRow(i) = 35) Then
'                            .Lock = True
'                        Else
'                            .Lock = False
'                        End If
'                    Next
'                End If
                

             'Check rang buôc CT 40 va 42
                .Row = 40
                .Col = .ColLetterToNumber("L")
                If Val(.Text) > 0 Then
                    .Row = 46
                    .Col = .ColLetterToNumber("L")
                    .Lock = True
                Else
                    .Row = 46
                    .Col = .ColLetterToNumber("L")
                    .Lock = False
                End If
'
            End If
        
        Case 2 'PL1
'            mCurrentSheet = .Sheet
'            Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
'                          (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 1)
'            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'            .Col = iCol
'            .Row = iRow
'            strValue1 = .value
'            UpdateCell fps, .Col, .Row, .value
'            Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
'                          (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 2)
'            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'            .Col = iCol
'            .Row = iRow
'            strValue2 = .value
'            UpdateCell fps, .Col, .Row, .value
'            If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                   fps.EventEnabled(EventAllEvents) = False
'                    .Sheet = 1
'                    mCurrentSheet = .Sheet
'                    fps.EventEnabled(EventAllEvents) = True
'              End If
        Case 3 'PL2
'            mCurrentSheet = .Sheet
'
'            Set xmlCellNode = TAX_Utilities_New.Data(2).getElementsByTagName("Cell") _
'                          (TAX_Utilities_New.Data(2).getElementsByTagName("Cell").length - 2)
'            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'            .Col = iCol
'            .Row = iRow
'            strValue1 = .value
'            UpdateCell fps, .Col, .Row, .value
'            Set xmlCellNode = TAX_Utilities_New.Data(2).getElementsByTagName("Cell") _
'                          (TAX_Utilities_New.Data(2).getElementsByTagName("Cell").length - 1)
'            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'            .Col = iCol
'            .Row = iRow
'            strValue2 = .value
'            UpdateCell fps, .Col, .Row, .value
'            If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Then
'                    fps.EventEnabled(EventAllEvents) = False
'                        .Sheet = 1
'                        mCurrentSheet = .Sheet
''                        'Chi tieu [12]
'                        .GetText .ColLetterToNumber("I"), 38, varTemp
'                        .Col = .ColLetterToNumber("J")
'                        .Row = 10
'                        .Value = varTemp
'                        UpdateCell fps, .Col, .Row, .Value
'                        'Chi tieu [13]
'                        .GetText .ColLetterToNumber("J"), 38, varTemp
'                        .Col = .ColLetterToNumber("L")
'                        .Row = 10
'                        .Value = varTemp
'                        UpdateCell fps, .Col, .Row, .Value
'                      .Sheet = .ActiveSheet
'                      mCurrentSheet = .Sheet
'                      fps.EventEnabled(EventAllEvents) = True
'             End If
        Case 4 'PL3
 '           mCurrentSheet = .Sheet
            'baclt added
            'ValidateDate fps, 4, 2, .ColLetterToNumber("F"), yyyy
    End Select
    
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
    
    ' pl5
    Dim vSoChungTu As Variant
    Dim vNgayNop As Variant
    Dim vNoiNop As Variant
    Dim vCoQuanQL As Variant
    Dim vSoTienDN As Variant
    'pl3
    Dim vLoaiXe As Variant
    Dim vDVtinh As Variant
    Dim vTongsoxe As Variant
    Dim vSoLuongxe As Variant
    Dim vGiaxe As Variant
    
    '/
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
    
    
    
    Dim blCheck_S2 As Boolean
    Dim blCheck_S3 As Boolean
    Dim blCheck_S4 As Boolean
    Dim blCheck_S5 As Boolean
     Dim blCheck_S7 As Boolean
    Dim blCheck_S8 As Boolean
    
    ' Kiem tra check gia han
    Dim vCheckGH, vLyDoGH, vErrorCheckGH As Variant
    Dim iflagCheckGH As Integer
    
    
    With fps
    
    '.Visible = False
    '.ReDraw = False
    iCurrentSheet = .Sheet
    .Sheet = .SheetCount
    CheckErrorMST
    
    .Sheet = 1

     .Sheet = .SheetCount
     .GetText .ColLetterToNumber("E"), 27, vErrMsgKyHieuHD
     .GetText .ColLetterToNumber("E"), 28, vErrMsgSoHD
     .GetText .ColLetterToNumber("E"), 29, vErrMsgNgayHD
     .GetText .ColLetterToNumber("E"), 30, vErrMsgTennguoimua
     .GetText .ColLetterToNumber("E"), 31, vErrMsgMST
     .GetText .ColLetterToNumber("E"), 37, vErrMsgThuesuat
     
     .GetText .ColLetterToNumber("E"), 37, vErrMsgThuesuat
     .GetText .ColLetterToNumber("E"), 57, vErrorSoChungTu
     .GetText .ColLetterToNumber("E"), 58, vError_pl_3
     
     .GetText .ColLetterToNumber("E"), 55, vErrorNgayNop
     
     .GetText .ColLetterToNumber("E"), 62, vErrorTenCSPT
     .GetText .ColLetterToNumber("E"), 63, vErrorDiaChi
     
     .GetText .ColLetterToNumber("E"), 64, vErrorCheckGH
     
     
     ' Kiem tra check gia han
     .Sheet = 1
     .GetText .ColLetterToNumber("B"), 18, vCheckGH
     .GetText .ColLetterToNumber("D"), 20, vLyDoGH
     'If TAX_Utilities_New.Year = 2012 And (TAX_Utilities_New.Month = 4 Or TAX_Utilities_New.Month = 5 Or TAX_Utilities_New.Month = 6) Then
        If Trim(vCheckGH) = "1" Or Trim$(UCase(vCheckGH)) = "X" Then
            If Trim$(vLyDoGH) = "" Then
                .Col = .ColLetterToNumber("D")
                .Row = 20
                .CellNote = vErrorCheckGH  'static
                .BackColor = mErrorColor
                iflagCheckGH = 1
            Else
                .Col = .ColLetterToNumber("D")
                .Row = 20
                .CellNote = ""  'static
                .BackColor = mNonErrorColor
                iflagCheckGH = 0
            End If
        Else
                .Col = .ColLetterToNumber("D")
                .Row = 20
                .CellNote = ""  'static
                .BackColor = mNonErrorColor
                iflagCheckGH = 0
        End If
        
    .Sheet = .SheetCount
    .SetText .ColLetterToNumber("B"), 64, IIf(iflagCheckGH = 1, "0", "1")
     'End If
     ' end kiem tra check gia han
    
    ' if sheet invisible then set value="1"
    .Sheet = 2
    blCheck_S2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    If blCheck_S2 = True Then
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
            
    
            If Trim(vKyHieuHD) <> "" Or Trim(vSoHD) <> "" Or Trim(vNgayHD) <> "" Or Trim(vTennguoimua) <> "" Or Trim(vDoanhSo) <> "0" Or Trim(vGTGT) <> "0" Then iWhiteRow = 0
                Select Case strCellTag
                    Case "1", "2", "3", "4", "5"
                        If Trim(vKyHieuHD) = "" And iWhiteRow = 0 Then
                            .Col = .ColLetterToNumber("C")
                            .CellNote = vErrMsgKyHieuHD  'static
                            .BackColor = mAlertColor
                            iFlagvKyHieuHD = 1
                            
                            'capture the first error cell
    '                        If iFocusFlag = 0 Then
    '                            strFocusSheetName = "'" & .SheetName & "'"
    '                            strFocusCol = str(.Col)
    '                            strFocusRow = str(.Row)
    '                            iFocusFlag = 1
    '                        End If
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
                             ElseIf (CheckValidNgayHD(CStr(vNgayHD)) = False) Then
                                .Col = .ColLetterToNumber("E")
                                .CellNote = GetAttribute(GetMessageCellById("0124"), "Msg") 'static
                                .BackColor = mAlertColor
                                iFlagvNgayHD = 0
    '                             If iFocusFlag = 0 Then
    '                                strFocusSheetName = "'" & .SheetName & "'"
    '                                strFocusCol = str(.Col)
    '                                strFocusRow = str(.Row)
    '                                iFocusFlag = 1
    '                              End If
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
          ' Check MST
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
                        
                        If strCellTag = "3" Then
                            If Trim(vDoanhSo) <> "0" And Trim(vGTGT) <> "0" Then
                                If (Abs(Round((vDoanhSo * 0.05), 0) - Round(vGTGT, 0)) > 100000) Or (Abs(Round((vDoanhSo * 0.05), 0) - Round(vGTGT, 0)) > Round(Abs(vDoanhSo) * 0.05 * 0.0001, 0)) Then
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
                        If strCellTag = "4" Then
                          If Trim(vDoanhSo) <> "0" Or Trim(vGTGT) <> "0" Then
                                If (Abs(Round((vDoanhSo * 0.1), 0) - Round(vGTGT, 0)) > 100000) Or (Abs(Round((vDoanhSo * 0.1), 0) - Round(vGTGT, 0)) > Round(Abs(vDoanhSo) * 0.1 * 0.0001, 0)) Then
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
            .SetText .ColLetterToNumber("B"), 27, IIf(iFlagvKyHieuHD = 1 And blCheck_S2 = True, "0", "1")
            .SetText .ColLetterToNumber("B"), 28, IIf(iFlagvSoHD = 1 And blCheck_S2 = True, "0", "1")
            .SetText .ColLetterToNumber("B"), 29, IIf(iFlagvNgayHD = 1 And blCheck_S2 = True, "0", "1")
            .SetText .ColLetterToNumber("B"), 30, IIf(iFlagvTennguoimua = 1 And blCheck_S2 = True, "0", "1")
            .SetText .ColLetterToNumber("B"), 31, IIf(iFlagvMst = 1 And blCheck_S2 = True, "0", "1")
        
            If iFocusFlag = 1 Then
                ' setcell for focus
                strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
                .Col = .ColLetterToNumber("C")
                .Row = 27
                .Formula = strFormula
                .Row = 28
                .Formula = strFormula
                .Row = 29
                .Formula = strFormula
                .Row = 30
                .Formula = strFormula
                .Row = 31
                .Formula = strFormula
            End If
         
         'Check MST
         
         
         .Sheet = .SheetCount
    '     .GetText .ColLetterToNumber("B"), 19, vTarget17
    '     .GetText .ColLetterToNumber("B"), 20, vTarget18
         .GetText .ColLetterToNumber("B"), 21, vTarget21
         .GetText .ColLetterToNumber("B"), 22, vTarget22
         .GetText .ColLetterToNumber("B"), 23, vTarget23
         .GetText .ColLetterToNumber("B"), 24, vTarget24
         .GetText .ColLetterToNumber("B"), 25, vTarget25
         .GetText .ColLetterToNumber("B"), 26, vTarget26
         
         If blCheck_S2 Then
            .SetText .ColLetterToNumber("B"), 21, IIf(vTarget21 = "1", "1", "0")
            .SetText .ColLetterToNumber("B"), 22, IIf(vTarget22 = "1", "1", "0")
            .SetText .ColLetterToNumber("B"), 23, IIf(vTarget23 = "1", "1", "0")
            .SetText .ColLetterToNumber("B"), 24, IIf(vTarget24 = "1", "1", "0")
            .SetText .ColLetterToNumber("B"), 25, IIf(vTarget25 = "1", "1", "0")
            .SetText .ColLetterToNumber("B"), 26, IIf(vTarget26 = "1", "1", "0")
          Else
           .SetText .ColLetterToNumber("B"), 21, "1"
           .SetText .ColLetterToNumber("B"), 22, "1"
           .SetText .ColLetterToNumber("B"), 23, "1"
           .SetText .ColLetterToNumber("B"), 24, "1"
           .SetText .ColLetterToNumber("B"), 25, "1"
           .SetText .ColLetterToNumber("B"), 26, "1"
          End If
  
    
    .Sheet = 3
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
                    Case "1", "2", "3", "4", "5"
                        If Trim(vKyHieuHD) = "" And iWhiteRow = 0 Then
                            .Col = .ColLetterToNumber("C")
                            .CellNote = vErrMsgKyHieuHD  'static
                            .BackColor = mAlertColor
                            iFlagvKyHieuHD = 1
                            
                            'capture the first error cell
    '                        If iFocusFlag = 0 Then
    '                            strFocusSheetName = "'" & .SheetName & "'"
    '                            strFocusCol = str(.Col)
    '                            strFocusRow = str(.Row)
    '                            iFocusFlag = 1
    '                        End If
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
    '                        If CheckValidNgayHD(CStr(vNgayHD)) = False Then
    '                           .Col = .ColLetterToNumber("E")
    '                           .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    '                           .CellNote = GetAttribute(GetMessageCellById("0112"), "Msg") 'static
    '                           .BackColor = mErrorColor
    '                           iFlagvNgayHD = 1
    '                            'capture the first error cell
    '                            If iFocusFlag = 0 Then
    '                                strFocusSheetName = "'" & .SheetName & "'"
    '                                strFocusCol = str(.Col)
    '                                strFocusRow = str(.Row)
    '                                iFocusFlag = 1
    '                            End If
    '                         Else
    '                            .Col = .ColLetterToNumber("E")
    '                            .CellNote = ""
    '                            .BackColor = mNonErrorColor
    '                         End If
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
                        If strCellTag = "1" Or strCellTag = "2" Or strCellTag = "3" Or strCellTag = "4" Or strCellTag = "5" Then
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
        .SetText .ColLetterToNumber("B"), 32, IIf(iFlagvKyHieuHD = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 33, IIf(iFlagvSoHD = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 34, IIf(iFlagvNgayHD = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 35, IIf(iFlagvTennguoimua = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 36, IIf(iFlagvMst = 1 And blCheck_S3 = True, "0", "1")
        .SetText .ColLetterToNumber("B"), 37, IIf(iFlagvThuesuat = 1 And blCheck_S3 = True, "0", "1")
        If iFocusFlag = 1 Then
            ' setcell for focus
            
            strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
            .Col = .ColLetterToNumber("C")
            .Row = 32
            .Formula = strFormula
            .Row = 33
            .Formula = strFormula
            .Row = 34
            .Formula = strFormula
            .Row = 35
            .Formula = strFormula
            .Row = 36
            .Formula = strFormula
            .Row = 37
            .Formula = strFormula
            
        End If
     
     .Sheet = .SheetCount
     .GetText .ColLetterToNumber("B"), 19, vTarget19
     .GetText .ColLetterToNumber("B"), 20, vTarget20
    
     
     
     If blCheck_S3 Then
        .SetText .ColLetterToNumber("B"), 19, IIf(vTarget19 = "1", "1", "0")
        .SetText .ColLetterToNumber("B"), 20, IIf(vTarget20 = "1", "1", "0")
      Else
       .SetText .ColLetterToNumber("B"), 19, "1"
       .SetText .ColLetterToNumber("B"), 20, "1"
      End If
    
    .GetText .ColLetterToNumber("B"), 38, vTarget38
    .GetText .ColLetterToNumber("B"), 39, vTarget39
   ' .GetText .ColLetterToNumber("E"), 40, vErrMsg
  
    .Sheet = 4
     blCheck_S4 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
     If blCheck_S4 = True Then
            .Col = .ColLetterToNumber("B")
            .Row = 8
            i = 1
            Do
                  iWhiteRow = 1
                  .GetText .ColLetterToNumber("C"), .Row, vLoaiXe
                  .GetText .ColLetterToNumber("D"), .Row, vDVtinh
                  .GetText .ColLetterToNumber("E"), .Row, vTongsoxe
                  .GetText .ColLetterToNumber("F"), .Row, vSoLuongxe
                  .GetText .ColLetterToNumber("G"), .Row, vGiaxe
                  If Trim(vLoaiXe) <> "" Or Trim(vDVtinh) <> "" Or Trim(vTongsoxe) <> "0" Or Trim(vSoLuongxe) <> "0" Or Trim(vGiaxe) <> "0" Then
                    iWhiteRow = 0
                  End If
                  If (Val(Trim(vTongsoxe)) < Val(Trim(vSoLuongxe))) And (vTongsoxe <> vbNullString) And (vSoLuongxe <> vbNullString) Then
                        ' Tong so luong xe ban ra
                        .Col = .ColLetterToNumber("E")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vError_pl_3   'static
                        .BackColor = mErrorColor
                        ' So luong xe ban cho nguoi tieu dung
                        .Col = .ColLetterToNumber("F")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vError_pl_3   'static
                        .BackColor = mErrorColor
                        
                        iflag_pl_3 = 1
                        'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
                  Else
                        .Col = .ColLetterToNumber("E")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = ""   'static
                        .BackColor = mNonErrorColor
                        ' so luong xe ban tieu dung
                        .Col = .ColLetterToNumber("F")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = ""   'static
                        .BackColor = mNonErrorColor
                  End If
                 .Col = .ColLetterToNumber("B")
                 .Row = i + 8
                 i = i + 1
                 If .Text = "aa" Then
                    i = i + 3
                 End If
            Loop Until .Text = "bb"
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 58, IIf(iflag_pl_3 = "1", "0", "1")
    End If
    
    .Sheet = 5
    blCheck_S5 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    
    .Sheet = 7
    blCheck_S7 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    If blCheck_S7 = True Then
    .Col = .ColLetterToNumber("B")
    .Row = 6
    i = 1
    Do
          iWhiteRow = 1
          .GetText .ColLetterToNumber("C"), .Row, vSoChungTu
          .GetText .ColLetterToNumber("D"), .Row, vNgayNop
          .GetText .ColLetterToNumber("E"), .Row, vNoiNop
          .GetText .ColLetterToNumber("F"), .Row, vCoQuanQL
          .GetText .ColLetterToNumber("G"), .Row, vSoTienDN
          If Trim(vSoChungTu) <> "" Or Trim(vNgayNop) <> "" Or Trim(vNoiNop) <> "" Or Trim(vCoQuanQL) <> "" Or Trim(vSoTienDN) <> "0" Then
            iWhiteRow = 0
          End If
          
          
          If (Trim(vSoChungTu) = "" Or vSoChungTu = vbNullString) And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("C")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrorSoChungTu   'static
                .BackColor = mErrorColor
                iflagSoChungTu = 1
                'capture the first error cell
                If iFocusFlag = 0 Then
                    strFocusSheetName = "'" & .SheetName & "'"
                    strFocusCol = str(.Col)
                    strFocusRow = str(.Row)
                    iFocusFlag = 1
                End If
          Else
                .Col = .ColLetterToNumber("C")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""   'static
                .BackColor = mNonErrorColor
          End If
          ' Kiem tra noi nop
           If (Trim(vNoiNop) = "" Or vNoiNop = vbNullString) And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("E")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrorSoChungTu   'static
                .BackColor = mErrorColor
                iflagSoChungTu = 1
                'capture the first error cell
                If iFocusFlag = 0 Then
                    strFocusSheetName = "'" & .SheetName & "'"
                    strFocusCol = str(.Col)
                    strFocusRow = str(.Row)
                    iFocusFlag = 1
                End If
          Else
                .Col = .ColLetterToNumber("E")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""   'static
                .BackColor = mNonErrorColor
          End If
          ' Kiem tra co quan thue
          If (Trim(vCoQuanQL) = "" Or vCoQuanQL = vbNullString) And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("F")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrorSoChungTu   'static
                .BackColor = mErrorColor
                iflagSoChungTu = 1
                'capture the first error cell
                If iFocusFlag = 0 Then
                    strFocusSheetName = "'" & .SheetName & "'"
                    strFocusCol = str(.Col)
                    strFocusRow = str(.Row)
                    iFocusFlag = 1
                End If
          Else
                .Col = .ColLetterToNumber("F")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""   'static
                .BackColor = mNonErrorColor
          End If
          ' Kiem tra so thue
          If (Trim(vSoTienDN) = "0" Or vSoTienDN = vbNullString) And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("G")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrorSoChungTu   'static
                .BackColor = mErrorColor
                iflagSoChungTu = 1
                'capture the first error cell
                If iFocusFlag = 0 Then
                    strFocusSheetName = "'" & .SheetName & "'"
                    strFocusCol = str(.Col)
                    strFocusRow = str(.Row)
                    iFocusFlag = 1
                End If
          Else
                .Col = .ColLetterToNumber("G")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = ""   'static
                .BackColor = mNonErrorColor
          End If
          
          
          If (Trim(vNgayNop) = "" Or vNgayNop = vbNullString) And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("D")
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = vErrorSoChungTu   'static
                .BackColor = mErrorColor
                iflagSoChungTu = 1
                'capture the first error cell
                If iFocusFlag = 0 Then
                    strFocusSheetName = "'" & .SheetName & "'"
                    strFocusCol = str(.Col)
                    strFocusRow = str(.Row)
                    iFocusFlag = 1
                End If
          Else
                If iWhiteRow = 0 Then
                    If DateDiff("D", Date, Format(vNgayNop, "dd/mm/yyyy")) > 0 Then
                           .Col = .ColLetterToNumber("D")
                           .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                           .CellNote = vErrorNgayNop   'static
                           .BackColor = mErrorColor
                           iflagNgayNop = 1
                           'capture the first error cell
                           If iFocusFlag = 0 Then
                               strFocusSheetName = "'" & .SheetName & "'"
                               strFocusCol = str(.Col)
                               strFocusRow = str(.Row)
                               iFocusFlag = 1
                           End If
                       Else
                            .Col = .ColLetterToNumber("D")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = ""   'static
                            .BackColor = mNonErrorColor
                      End If
                Else
                            .Col = .ColLetterToNumber("D")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = ""   'static
                            .BackColor = mNonErrorColor
                End If
          End If
                 
         ' .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg") 'static
                 
         .Col = .ColLetterToNumber("B")
         .Row = i + 6
         i = i + 1
    Loop Until .Text = "aa"
      .Sheet = .SheetCount
    .SetText .ColLetterToNumber("B"), 57, IIf(iflagSoChungTu = "1", "0", "1")
    .SetText .ColLetterToNumber("B"), 55, IIf(iflagNgayNop = 1, "0", "1")
 End If
    .Sheet = 8
    iFlagvMst_PL6 = 1
    blCheck_S8 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
'    For i = 8 To .MaxRows - 13
    .Col = .ColLetterToNumber("B")
    .Row = 8
    i = 1
    Do

        iWhiteRow = 1
        '.Row = i
        .Col = .ColLetterToNumber("C")
        .GetText .ColLetterToNumber("C"), .Row, vKyHieuHD
        .GetText .ColLetterToNumber("D"), .Row, vSoHD
        .GetText .ColLetterToNumber("E"), .Row, vMST
        .GetText .ColLetterToNumber("F"), .Row, vTennguoimua
        .GetText .ColLetterToNumber("G"), .Row, vDoanhSo
        .GetText .ColLetterToNumber("H"), .Row, vGTGT
        If Trim(vKyHieuHD) <> "" Or Trim(vSoHD) <> "" Or Trim(vMST) <> "" Or Trim(vTennguoimua) <> "0" Or Trim(vDoanhSo) <> "0" Then iWhiteRow = 0
      ' Check MST
                If Trim(vMST) <> vbNullString Then
                   vMST = Replace(vMST, " ", "")
                   vMST = Replace(vMST, "-", "")
                    'vMST = Left(vMST, 10)

                   strCheck = ""
                   If Len(vMST) = 10 Or Len(vMST) = 13 Then
                        strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                   End If
                   
                   If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                        .Col = .ColLetterToNumber("E")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgMST   'static
                        .BackColor = mErrorColor
                        iFlagvMst_PL6 = 0
                        
                        'capture the first error cell
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
                
             ' kiem tra bat buoc nhap truong ten va dia chi
              If Trim(vKyHieuHD) = "" And iWhiteRow = 0 Then
                    .Col = .ColLetterToNumber("C")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrorTenCSPT   'static
                    .BackColor = mErrorColor
                    flagTenCSPT = 1
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
              Else
                     .Col = .ColLetterToNumber("C")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = ""   'static
                    .BackColor = mNonErrorColor
              End If
                    
              If Trim(vSoHD) = "" And iWhiteRow = 0 Then
                    .Col = .ColLetterToNumber("D")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrorDiaChi   'static
                    .BackColor = mErrorColor
                    flagDiaChi = 1
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
              Else
                     .Col = .ColLetterToNumber("D")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = ""   'static
                    .BackColor = mNonErrorColor
              End If
          .Col = .ColLetterToNumber("B")
         .Row = i + 8
         i = i + 1
    Loop Until .Text = "aa"
                   
                    
'    Next
    .Sheet = .SheetCount
     If blCheck_S8 Then
        .SetText .ColLetterToNumber("B"), 48, IIf(iFlagvMst_PL6 = "1", "1", "0")
        .SetText .ColLetterToNumber("B"), 62, IIf(flagTenCSPT = "1", "0", "1")
        .SetText .ColLetterToNumber("B"), 63, IIf(flagDiaChi = "1", "0", "1")
      Else
        .SetText .ColLetterToNumber("B"), 48, "1"
        .SetText .ColLetterToNumber("B"), 62, "1"
        .SetText .ColLetterToNumber("B"), 63, "1"
      End If
    .Sheet = iCurrentSheet
    '.ReDraw = True
    '.Visible = True
    End With
End Sub
'This function is used to support report
'some special data that we want to save to data files
Public Sub SaveExceptionData()
    Dim i As Long, DataText As Variant
    Dim ArrRow As Variant, ArrCol As Variant
    Dim arrCellTag() As String
    With fps
        .Sheet = 1
        mCurrentSheet = .Sheet
        ArrRow = Array(10, 10, 16, 19, 19, 21, 21, 28, 28, 30, 31, 33, 20, 22, 23, 24, 23, 24)
        ArrCol = Array("J", "L", "L", "J", "L", "J", "L", "J", "L", "L", "L", "L", "J", "J", "J", "J", "L", "L")
        For i = 0 To UBound(ArrRow)
            .Row = ArrRow(i)
            .Col = .ColLetterToNumber(ArrCol(i))
            UpdateCell fps, .Col, .Row, .value
        Next
    End With
                  
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
        
        'SaveExceptionData 'Save Exception Data to DOM
         
'          mCurrentSheet = 2
'         .Sheet = mCurrentSheet
'        Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
'        (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 1)
'         ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'         .GetText iCol, iRow, vTemp
'         UpdateCell fps, iCol, iRow, vTemp
'
'         Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
'        (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 2)
'          ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'         .GetText iCol, iRow, vTemp
'         UpdateCell fps, iCol, iRow, vTemp
'             i = 1
'             .Col = .ColLetterToNumber("B")
'             .Row = 8
'             Do
'                  .Text = str(i)
'                  .Col = .ColLetterToNumber("B")
'                  .Row = i + 8
'                  i = i + 1
'             Loop Until .Text = "aa"
'             i = 1
'            .Row = .Row + 4
'             Do
'                  .Text = str(i)
'                  .Col = .ColLetterToNumber("J")
'                  .Text = "0%"
'                  UpdateCell fps, .Col, .Row, .Text
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "bb"
'             i = 1
'            .Row = .Row + 4
'             Do
'                  .Text = str(i)
'                  .Col = .ColLetterToNumber("J")
'                  .Text = "5%"
'                  UpdateCell fps, .Col, .Row, .Text
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "cc"
'             i = 1
'            .Row = .Row + 4
'             Do
'                  .Text = str(i)
'                  .Col = .ColLetterToNumber("J")
'                  .Text = "10%"
'                  UpdateCell fps, .Col, .Row, .Text
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "dd"
'             i = 1
'            .Row = .Row + 4
'             Do
'                  .Text = str(i)
'                  .Col = .ColLetterToNumber("J")
'                  .Text = "10%"
'                  UpdateCell fps, .Col, .Row, .Text
'                  .Col = .ColLetterToNumber("B")
'                  .Row = 1 + .Row
'                  i = i + 1
'             Loop Until .Text = "ee"
'
'
'         mCurrentSheet = 3
'         .Sheet = mCurrentSheet
'        Set xmlCellNode = TAX_Utilities_New.Data(2).getElementsByTagName("Cell") _
'        (TAX_Utilities_New.Data(2).getElementsByTagName("Cell").length - 1)
'         ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'         .GetText iCol, iRow, vTemp
'         UpdateCell fps, iCol, iRow, vTemp
'
'         Set xmlCellNode = TAX_Utilities_New.Data(2).getElementsByTagName("Cell") _
'        (TAX_Utilities_New.Data(2).getElementsByTagName("Cell").length - 2)
'          ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
'         .GetText iCol, iRow, vTemp
'         UpdateCell fps, iCol, iRow, vTemp
'                      i = 1
'                    .Col = .ColLetterToNumber("B")
'                    .Row = 8
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("J")
'                         .CellType = CellTypePercent
'                         .Col = .ColLetterToNumber("B")
'                         .Row = i + 8
'                         i = i + 1
'                    Loop Until .Text = "aa"
'                    i = 1
'                   .Row = .Row + 4
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("J")
'                         .CellType = CellTypePercent
'                         .Col = .ColLetterToNumber("B")
'                         .Row = 1 + .Row
'                         i = i + 1
'                    Loop Until .Text = "bb"
'                    i = 1
'                   .Row = .Row + 4
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("J")
'                         .CellType = CellTypePercent
'                         .Col = .ColLetterToNumber("B")
'                         .Row = 1 + .Row
'                         i = i + 1
'                    Loop Until .Text = "cc"
         
        Set xmlCellNode = Nothing
         
'         mCurrentSheet = 6
'        .Sheet = mCurrentSheet
'        Dim vF2 As Variant
'        .GetText .ColLetterToNumber("F"), 2, vF2
'        If vF2 <> "" Then
'            If CDbl(vF2) < CDbl(TAX_Utilities_New.Year) Then
'                ' update gia tri dung (khong loi) trong sheet header
'                .Sheet = .SheetCount
'                .SetText .ColLetterToNumber("B"), 47, "1"
'                .SetText .ColLetterToNumber("B"), 48, "1"
'                'Chuyen chi tieu C4 ve trang thai dung
'                mCurrentSheet = 4
'                .Sheet = 4
'                .Col = .ColLetterToNumber("F")
'                .Row = 2
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'            Else
'                ' update gia tri sai (co loi) trong sheet header
'                .Sheet = .SheetCount
'                .SetText .ColLetterToNumber("B"), 47, "0"
'                ' Lay thong bao loi
'                Dim vErrMsg1 As Variant
'                .GetText .ColLetterToNumber("E"), 47, vErrMsg1
'                ' Chuyen chi tieu ve trang thai loi
'                mCurrentSheet = 4
'                .Sheet = 4
'                .Col = .ColLetterToNumber("F")
'                .Row = 2
'                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                .CellNote = vErrMsg1 'static
'                .BackColor = mErrorColor
'            End If
'        Else
'                ' update gia tri sai (co loi) trong sheet header
'                .Sheet = .SheetCount
'                .SetText .ColLetterToNumber("B"), 48, "0"
'                ' Lay thong bao loi
'                Dim vErrMsg2 As Variant
'                .GetText .ColLetterToNumber("E"), 48, vErrMsg2
'                ' Chuyen chi tieu ve trang thai loi
'                mCurrentSheet = 4
'                .Sheet = 4
'                .Col = .ColLetterToNumber("F")
'                .Row = 2
'                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                .CellNote = vErrMsg2 'static
'                .BackColor = mErrorColor
'        End If
'
        CheckDynamicError 'Set Exception Error on cells of interface
            
        UpdateSheet ' set gia tri ma hinh xuong XML
            
        .ActiveSheet = ASheet
        .Sheet = SSheet
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
            .Col = .ColLetterToNumber("BB")
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
      .Col = .ColLetterToNumber("BE")
      .Row = Row
      If .Text = "23" Then
         .Col = .ColLetterToNumber("BG")
         .Row = .Row
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
      Else
        .Col = .ColLetterToNumber("BG")
        .Row = .Row
        .TypeNumberMin = Val("0")
        .TypeNumberMax = Val("99999999999999")
      End If
    ElseIf .ActiveSheet = 1 Then
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("D") And Row = 20 Then
            intIndexCombo = .TypeComboBoxCurSel
            .Col = .ColLetterToNumber("P")
            .TypeComboBoxCurSel = intIndexCombo
            'Update combo C
            UpdateCell fps, .Col, .Row, .Text
        End If
    End If
    
End With
    
    'CellChange Col, Row
    LockCellBySheet
    TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
End Sub

Private Sub setLast43Target()
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    'Dim strOriginDataFileName As String
    Dim strDataFileName As String, strCurrentDataFileName As String
    Dim str43Target As String
    Dim iDom As Integer
    Dim MY As Date
    
    iDom = 0
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        MY = DateSerial(TAX_Utilities_New.Year, TAX_Utilities_New.Month, 1)
        
        MY = DateAdd("M", -1, MY)
        ' TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(1), "DataFile") & ".xml"
        strCurrentDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(iDom), "DataFile") & "_" & TAX_Utilities_New.Month & TAX_Utilities_New.Year & ".xml"
        '***************************************
        'ThanhDX added
        'Date: 14/04/06
        
        'strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(GetValidityNode("01", Format(Month(MY), "0#"), , Format(Year(MY), "####")).childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        '***************************************
        
        
        If xmlDomCurrentData.Load(strCurrentDataFileName) = True Then
            GoTo Tidy_up
        End If
        If .Load(strDataFileName) = True Then
            Set xmlNode = .nodeFromID("L_47")  'L_47: col & row of 43 target
            str43Target = GetAttribute(xmlNode, "Value")
            Set xmlNode = Nothing
            If Trim(str43Target) <> "" And Trim(str43Target) <> "0" Then
                fps.Sheet = 1
                fps.Col = fps.ColLetterToNumber(S1Txt11Col)
                fps.Row = S1Txt11Row
                If Trim(fps.Text) = "0" Or Trim(fps.Text) = "" Then
                    fps.Text = str43Target
                    mCurrentSheet = 1
'                    UpdateCell fps, fps.Col, fps.Row, fps.Value
'                    'update cell L30
'                    fps.Col = fps.ColLetterToNumber("L")
'                    fps.Row = 30
'                    UpdateCell fps, fps.Col, fps.Row, fps.Value
'                    'update cell L31
'                    fps.Col = fps.ColLetterToNumber("L")
'                    fps.Row = 31
'                    UpdateCell fps, fps.Col, fps.Row, fps.Value
'                    'update cell L33
'                    fps.Col = fps.ColLetterToNumber("L")
'                    fps.Row = 33
'                    UpdateCell fps, fps.Col, fps.Row, fps.Value
                    
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
    Set xmlNode = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls001", "setLast43Target", Err.number, Err.Description
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
            Case "01_1"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 2)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 3)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
                'Set active voi dong 5
                fps.Sheet = 2
                fps.GetText fps.ColLetterToNumber("N"), 1, strValue4
                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Or Val(strValue4) <> 0 Then
                            blnNullValue = False
                    End If

                    If blnNullValue Then
                        SetAttribute xmlNodeValid, "Active", "0"
                    End If

             Case "01_2"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_New.Data(2).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(2).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_New.Data(2).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(2).getElementsByTagName("Cell").length - 2)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 'Set active voi dong 5
                fps.Sheet = 3
                fps.GetText fps.ColLetterToNumber("N"), 1, strValue5

                    If (Not IsNullValue(strValue1)) Or (Not IsNullValue(strValue2)) Or Val(strValue5) <> 0 Then
                            blnNullValue = False
                    End If

                    If blnNullValue Then
                        SetAttribute xmlNodeValid, "Active", "0"
                    End If
                    
            Case "01_3"
               blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_New.Data(3).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in E column and F column
                    If lCol = fps.ColLetterToNumber("E") Or _
                       lCol = fps.ColLetterToNumber("F") Then
                        strValue = GetAttribute(xmlCellNode, "Value")
                        If Not IsNullValue(strValue) Then
                            blnNullValue = False
                            Exit For
                        End If
                    End If
                Next
            
                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If
                
            Case "01_4A"
                'Check the fourth sheet

                blnNullValue = True
                For Each xmlCellNode In TAX_Utilities_New.Data(4).getElementsByTagName("Cell")
                    If GetAttribute(xmlCellNode, "CellID") = THIRD_SHEET_CELL1 Then _
                        Exit For
                Next

                strValue = GetAttribute(xmlCellNode, "Value")
                
                Set xmlCellNode = TAX_Utilities_New.Data(4).getElementsByTagName("Cell").Item(0)
                strValue2 = GetAttribute(xmlCellNode, "Value")
                
                If Not (IsNullValue(strValue) And IsNullValue(strValue2)) Then
                    blnNullValue = False
                End If

                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If
             
             Case "01_4B"
                'Check the fivth sheet
                blnNullValue = True
                For Each xmlCellNode In TAX_Utilities_New.Data(5).getElementsByTagName("Cell")
                    If GetAttribute(xmlCellNode, "CellID") = FOURTH_SHEET_CELL1 Then _
                        Exit For
                Next
            
                strValue = GetAttribute(xmlCellNode, "Value")
                If Not IsNullValue(strValue) Then
                    blnNullValue = False
                End If
            
                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If
             
             Case "01_5"
              blnNullValue = True
                'Check the last sheet
               For Each xmlCellNode In TAX_Utilities_New.Data(6).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value G F column
                    If lCol = fps.ColLetterToNumber("G") Then
                        strValue = GetAttribute(xmlCellNode, "Value")
                        If Not IsNullValue(strValue) Then
                            blnNullValue = False
                            Exit For
                        End If
                    End If
                Next
            
                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If
             Case "01_6"
                'Check the last sheet
               For Each xmlCellNode In TAX_Utilities_New.Data(7).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value I column
                    If lCol = fps.ColLetterToNumber("I") Then
                        strValue = GetAttribute(xmlCellNode, "Value")
                        If Not IsNullValue(strValue) Then
                            blnNullValue = False
                            Exit For
                        End If
                    End If
                Next
            
                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If
             Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_New.Data(8).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(8).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_New.Data(8).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(8).getElementsByTagName("Cell").length - 4)
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
        
        .Sheet = 3 'Sheet contain dynamic data

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

Private Function GetLast43Previous(ByRef strValue As String) As Boolean
     
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim strDataFileName As String
    Dim str43Target As String
    Dim iDom As Integer
    Dim MY As Date
    
    iDom = 0
    GetLast43Previous = True
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        MY = DateSerial(TAX_Utilities_New.Year, TAX_Utilities_New.Month, 1)
        MY = DateAdd("M", -1, MY)
        
        '***************************************
        'ThanhDX added
        'Date: 14/04/06
        
        'strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        strDataFileName = TAX_Utilities_New.DataFolder & GetAttribute(GetValidityNode("01", Format(Month(MY), "0#"), , Format(Year(MY), "####")).childNodes(iDom), "DataFile") & "_" & Format(Month(MY), "0#") & Format(Year(MY), "####") & ".xml"
        '***************************************
                
        If .Load(strDataFileName) = True Then
            Set xmlNode = .nodeFromID("L_33")  'L_33: col & row of 43 target
            strValue = GetAttribute(xmlNode, "Value")
            Set xmlNode = Nothing
        ElseIf .parseError.reason <> vbNullString Then
            If InStr(1, .parseError.errorCode, "2146697210") <> 0 Then
                'file data of last month does not exist
            Else
                MsgBox .parseError.reason
            End If
            GetLast43Previous = False
        End If
        
    End With
    
    Set xmlDomLastData = Nothing
    Set xmlNode = Nothing
    
    Exit Function
    
ErrorHandle:
    SaveErrorLog "cls001", "GetLast43Previous", Err.number, Err.Description
End Function

'*******************************************************
'Description: ResetData procedure reset specified cells
'Author:ThanhDX
'Date:23/01/2006

'*******************************************************
Public Sub ResetData()
    Dim intCurSheet As Integer, intCtrl As Long
    Dim lCol As Long, lRow As Long
    Dim arrStrCellUnLock() As String
    Dim strCell As String
    
    'Array of reset cells
    'arrStrCellUnLock = Split("J_11,J_12,L_11,L_12,J_20,J_22,J_23,L_23,J_24,L_24", ",")
    
    'htphuong edit 15/12/2007
    'arrStrCellUnLock = Split("J_10,L_10,J_20,J_22,J_23,L_23,J_24,L_24", ",")
    
    'htphuong edit 09/03/2007
'    arrStrCellUnLock = Split("J_11,L_11,J_12,L_12,J_20,J_22,J_23,L_23,J_24,L_24", ",")
'
'    With fps
'        intCurSheet = .Sheet
'        .Sheet = 1
'        mCurrentSheet = .Sheet
'        For intCtrl = 0 To UBound(arrStrCellUnLock)
'            strCell = arrStrCellUnLock(intCtrl)
'            ParserCellID fps, strCell, lCol, lRow
'            .Col = lCol
'            .Row = lRow
'            .Lock = False
'        Next
'        .Sheet = intCurSheet
'        mCurrentSheet = .Sheet
'    End With
'dhdang sua lock cell theo Sheet active
LockCellBySheet
End Sub

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()
      
        If fps.ActiveSheet = 6 Then
            mCurrentSheet = 6
            fps.Sheet = mCurrentSheet
                With fps
                    .Col = fps.ColLetterToNumber("F")
                    .Row = 2
                    .Text = TAX_Utilities_New.Year - 1
                    .TypeHAlign = TypeHAlignLeft
                    UpdateCell fps, .Col, .Row, .Text
                End With
        End If
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

Private Function GetValueKHBS(StrKytinhthue As String, StrNgayKHBS As String) As Double
    Dim lngIndex As Long
    Dim fso As New FileSystemObject
    Dim fle As File
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlDocument As New MSXML.DOMDocument
    Dim cValue As Long
    Dim StrFileName As String
    For Each fle In fso.GetFolder(GetAbsolutePath(TAX_Utilities_New.DataFolder)).Files
       If Right$(fle.Name, 4) = ".xml" Then
            If Mid$(fle.Name, 1, Len(fle.Name) - 4) = "KHBS_01_GTGT_" & StrKytinhthue & "_" & StrNgayKHBS Then
                StrFileName = TAX_Utilities_New.DataFolder & fle.Name
                 xmlDocument.Load StrFileName
                  Set xmlNode = xmlDocument.childNodes(2).childNodes(2).lastChild.lastChild
                        cValue = CDbl(GetAttribute(xmlNode, "Value"))
                        If cValue < 0 Then
                            GetValueKHBS = GetValueKHBS + cValue
                            Exit For
                        End If
                  
            End If
       End If
    Next
      Set xmlNode = Nothing
      Set xmlDocument = Nothing
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
    Dim varTemp As Variant
    Dim var28 As Variant
    Dim var17 As Variant
    Dim var7 As Variant
    Dim var30 As Variant
    Dim var31 As Variant
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
        
        If loaiKHBS = "frmKHBS_BS" Then
            var28 = GetAttribute(TAX_Utilities_New.DataKHBS.nodeFromID("L_28"), "Value")
            var17 = GetAttribute(TAX_Utilities_New.DataKHBS.nodeFromID("L_17"), "Value")
            var7 = GetAttribute(TAX_Utilities_New.DataKHBS.nodeFromID("L_7"), "Value")

            If var28 - var17 - var7 > 0 Then
                SetAttribute TAX_Utilities_New.DataKHBS.nodeFromID("L_30"), "Value", CStr(var28 - var17 - var7)
                SetAttribute TAX_Utilities_New.DataKHBS.nodeFromID("L_31"), "Value", 0
                SetAttribute TAX_Utilities_New.DataKHBS.nodeFromID("L_33"), "Value", 0
                
            Else
                SetAttribute TAX_Utilities_New.DataKHBS.nodeFromID("L_30"), "Value", 0
                SetAttribute TAX_Utilities_New.DataKHBS.nodeFromID("L_31"), "Value", CStr(var28 - var17 - var7)
                SetAttribute TAX_Utilities_New.DataKHBS.nodeFromID("L_33"), "Value", 0
            End If
        End If
        .Sheet = mCurrentSheet
    End With
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
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
'                         .Col = .ColLetterToNumber("K")
'                         .Text = "0"
'                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 8
                         i = i + 1
                    Loop Until .Text = "aa"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("J")
'                         .TypeHAlign = TypeHAlignCenter
'                         .Text = "0%"
                         'UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("K")
                         .Text = "0"
                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "bb"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("J")
'                         .TypeHAlign = TypeHAlignCenter
'                         .Text = "5%"
'                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "cc"
                    i = 1
                   .Row = .Row + 4
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("J")
'                         .TypeHAlign = TypeHAlignCenter
'                         .Text = "10%"
'                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")

                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "dd"
                    i = 1
                    .Row = .Row + 4
                    Do
                         .Text = str(i)
'                         .Col = .ColLetterToNumber("J")
'                         .TypeHAlign = TypeHAlignCenter
'                         .Text = "10%"
'                         UpdateCell fps, .Col, .Row, .Text
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "ee"
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol

                    .SetActiveCell iCol, iRow
                    CellChange2
                    CellChange3
            End If
            If .ActiveSheet = 3 Then
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
                     i = 1
                    .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "dd"
                     i = 1
                    .Row = .Row + 4
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = 1 + .Row
                         i = i + 1
                    Loop Until .Text = "ee"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    .SetActiveCell iCol, iRow
                    CellChange2
                    CellChange3
            End If
'            If .ActiveSheet = 4 Then
'                    i = 1
'                    .Col = .ColLetterToNumber("B")
'                    .Row = 20
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("B")
'                         .Row = i + 20
'                         i = i + 1
'                    Loop Until .Text = "aa"
'                    i = 1
'                   .Row = .Row + 4
'                    Do
'                         .Text = str(i)
'                         .Col = .ColLetterToNumber("B")
'                         .Row = 1 + .Row
'                         i = i + 1
'                    Loop Until .Text = "bb"
'            End If
        fps.EventEnabled(EventAllEvents) = True
        End With
    'CellChange iCol, iRow
    'CellChange1 Col, Row
End Sub

Public Function getDieuChinhGiam(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue, strDieuChinhGiam, strDieuChinhTang, strDieuChinhTang4043 As String
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
                .GetText .ColLetterToNumber("L"), 24, varTemp
                strCurrentValue = "CT22_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 27, varTemp
'                strCurrentValue = strCurrentValue & "~CT24_" & Trim(varTemp)
                .GetText .ColLetterToNumber("L"), 28, varTemp
                strCurrentValue = strCurrentValue & "~CT25_" & Trim(varTemp)
                .GetText .ColLetterToNumber("L"), 33, varTemp
                strCurrentValue = strCurrentValue & "~CT31_" & Trim(varTemp)
                .GetText .ColLetterToNumber("L"), 34, varTemp
                strCurrentValue = strCurrentValue & "~CT33_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 43, varTemp
'                strCurrentValue = strCurrentValue & "~CT40b_" & Trim(varTemp)
                .GetText .ColLetterToNumber("L"), 46, varTemp
                strCurrentValue = strCurrentValue & "~CT42_" & Trim(varTemp)
                .GetText .ColLetterToNumber("L"), 44, varTemp
                strCurrentValue = strCurrentValue & "~CT40_" & Trim(varTemp)
                .GetText .ColLetterToNumber("L"), 47, varTemp
                strCurrentValue = strCurrentValue & "~CT43_" & Trim(varTemp)
                ' tinh chenh lech giua 2 lan dieu chinh
                tempArrOld = Split(strOldValue, "~")
                tempArrCurr = Split(strCurrentValue, "~")
                arrMapCT = Split(strMaCT, "~")
                strDieuChinhGiam = ""
                strDieuChinhTang = ""
                strDieuChinhTang4043 = ""
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
                    If (Left(strMaCT, 3) = "40_" Or Val(Left(strMaCT, 2)) = 43) Then
                        'If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) <> 0 Then
                            ' dieu chinh giam
                            strDieuChinhTang4043 = strDieuChinhTang4043 & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                        'End If
                    End If
                    If Left(strMaCT, 3) = "22_" Or Left(strMaCT, 3) = "25_" Or Left(strMaCT, 3) = "42_" Then
                        If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 And Left(strMaCT, 3) <> "40_" And Val(Left(strMaCT, 2)) <> 43 Then
                            ' dieu chinh giam
                            strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                        End If
                        If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 And Left(strMaCT, 3) <> "40_" And Val(Left(strMaCT, 2)) <> 43 Then
                            ' dieu chinh tang
                            strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                        End If
                    Else
                        If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 And Left(strMaCT, 3) <> "40_" And Val(Left(strMaCT, 2)) <> 43 Then
                                ' dieu chinh giam
                                strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
                            End If
                            If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 And Left(strMaCT, 3) <> "40_" And Val(Left(strMaCT, 2)) <> 43 Then
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
                If strDieuChinhTang4043 <> "" Then
                    strDieuChinhTang4043 = Mid(strDieuChinhTang4043, 2, Len(strDieuChinhTang4043) - 1)
                End If
                getDieuChinhGiam = strDieuChinhGiam & "###" & strDieuChinhTang & "###" & strDieuChinhTang4043
      End With
End Function
'Public Function getDieuChinhGiam(strOldValue As String) As String
'    Dim varTemp As Variant
'    Dim strCurrentValue, strDieuChinhGiam, strDieuChinhTang As String
'    Dim tempArrOld() As String
'    Dim tempArrCurr() As String
'    Dim tempArrOldValue() As String
'    Dim tempArrCurrValue() As String
'    Dim i, j As Integer
'    Dim strMaCT As String
'    Dim arrMapCT() As String ' Luu gia tri cua dong chi tieu
'    Dim arrMapCTValue() As String ' Luu gia tri cua cac cell chi tieu
'    strMaCT = getMapCT
'    With fps
'                .Sheet = 1
'                .GetText .ColLetterToNumber("L"), 24, varTemp
'                strCurrentValue = "CT22_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 27, varTemp
'                strCurrentValue = strCurrentValue & "~CT24_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 29, varTemp
'                strCurrentValue = strCurrentValue & "~CT25_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 33, varTemp
'                strCurrentValue = strCurrentValue & "~CT31_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 34, varTemp
'                strCurrentValue = strCurrentValue & "~CT33_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 38, varTemp
'                strCurrentValue = strCurrentValue & "~CT37_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 39, varTemp
'                strCurrentValue = strCurrentValue & "~CT38_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 43, varTemp
'                strCurrentValue = strCurrentValue & "~CT40b_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 46, varTemp
'                strCurrentValue = strCurrentValue & "~CT42_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 44, varTemp
'                strCurrentValue = strCurrentValue & "~CT40_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("L"), 47, varTemp
'                strCurrentValue = strCurrentValue & "~CT43_" & Trim(varTemp)
'                ' tinh chenh lech giua 2 lan dieu chinh
'                tempArrOld = Split(strOldValue, "~")
'                tempArrCurr = Split(strCurrentValue, "~")
'                arrMapCT = Split(strMaCT, "~")
'                strDieuChinhGiam = ""
'                strDieuChinhTang = ""
'                For i = 0 To UBound(tempArrOld)
'                    tempArrOldValue = Split(tempArrOld(i), "_")
'                    tempArrCurrValue = Split(tempArrCurr(i), "_")
'                    ' lay cac gia tri cua tung cot trong sheet KHBS
'                    For j = 0 To UBound(arrMapCT)
'                        arrMapCTValue = Split(arrMapCT(j), "_")
'                        ' phan tu so 0 la ma chi tieu
'                        If arrMapCTValue(0) = Mid(tempArrOldValue(0), 3) Then
'                            strMaCT = arrMapCT(j)
'                            Exit For
'                        End If
'                    Next j
'
'                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) > 0 Then
'                        ' dieu chinh giam
'                        strDieuChinhGiam = strDieuChinhGiam & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
'                    End If
'                    If Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1)) < 0 Then
'                        ' dieu chinh tang
'                        strDieuChinhTang = strDieuChinhTang & "~" & strMaCT & "_" & Val(tempArrOldValue(1)) & "_" & Val(tempArrCurrValue(1)) & "_" & Val(tempArrCurrValue(1)) - Val(tempArrOldValue(1))
'                    End If
'                Next i
'                If strDieuChinhGiam <> "" Then
'                    strDieuChinhGiam = Mid(strDieuChinhGiam, 2, Len(strDieuChinhGiam) - 1)
'                End If
'                If strDieuChinhTang <> "" Then
'                    strDieuChinhTang = Mid(strDieuChinhTang, 2, Len(strDieuChinhTang) - 1)
'                End If
'                getDieuChinhGiam = strDieuChinhGiam & "###" & strDieuChinhTang
'      End With
'End Function

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
            Set xmlNode = .nodeFromID("L_24")
            result = "CT22_" & Val(GetAttribute(xmlNode, "Value"))
'            Set xmlNode = .nodeFromID("L_27")
'            result = result & "~" & "CT24_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("L_28")
            result = result & "~" & "CT25_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("L_33")
            result = result & "~" & "CT31_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("L_34")
            result = result & "~" & "CT33_" & Val(GetAttribute(xmlNode, "Value"))
'            Set xmlNode = .nodeFromID("L_43")
'            result = result & "~" & "CT40b_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("L_46")
            result = result & "~" & "CT42_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("L_44")
            result = result & "~" & "CT40_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("L_47")
            result = result & "~" & "CT43_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
            

        End If
    End With
    getValueTK = result
    Exit Function
    
ErrorHandle:
    SaveErrorLog "cls_02GTGT", "getValueTK", Err.number, Err.Description
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
    SaveErrorLog "01_GTGT", "getMapCT", Err.number, Err.Description
End Function

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
            If lCol = .ColLetterToNumber("I") And lRow = 23 Then
            ElseIf lCol = .ColLetterToNumber("B") And lRow = 18 Then
            Else
                
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            End If
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
        .Sheet = 5
        mCurrentSheet = 5
        blCheck_S2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(4).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
        .Sheet = 6
        mCurrentSheet = 6
        blCheck_S2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(5).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
        .Sheet = 7
        mCurrentSheet = 7
        blCheck_S2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(6).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
        .Sheet = 8
        mCurrentSheet = 8
        blCheck_S2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(7).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
        
        .Sheet = 9
        mCurrentSheet = 9
        blCheck_S2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(8).getElementsByTagName("Cell")
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


' set cho du ham
' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        getChiTieuTongDC = "0"
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

 
    getValueCTDC = "0"
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01BTNDN", "getValueCTDC", Err.number, Err.Description
End Function


' Ham set lai trang thai cua cac chi tieu khi check chi tieu 21
Sub CellLock(ByVal Col As Long, ByVal Row As Long)

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
       .EventEnabled(EventAllEvents) = False
       .Sheet = .ActiveSheet
       Select Case .Sheet
           Case 1 ' To khai GTGT
               mCurrentSheet = .Sheet
               GetCellSpan fps, Col, Row
               If Row = S1ChkRow And Col = .ColLetterToNumber(S1ChkCol) Then
                   ArrRow = Array(27, 27, 28, 30, 31, 31, 32, 33, 33, 34, 34, 35, 35)
                   ArrCol = Array("J", "L", "L", "J", "J", "L", "J", "J", "L", "J", "L", "J", "L")
                   .Col = .ColLetterToNumber("I")
                   .Row = 23
                   If .Text = "1" Then
                       For i = 0 To UBound(ArrRow)
                           .Row = ArrRow(i)
                           .Col = .ColLetterToNumber(ArrCol(i))
                           .Text = "0"
                           UpdateCell fps, .Col, .Row, .Text
                           .Lock = True
                       Next
                   ElseIf .Text = "0" Then
                       For i = 0 To UBound(ArrRow)
                           .Row = ArrRow(i)
                           .Col = .ColLetterToNumber(ArrCol(i))
                           If (ArrCol(i) = "J" Or ArrCol(i) = "L") And (ArrRow(i) = 31 Or ArrRow(i) = 35) Then
                               .Lock = True
                           Else
                               .Lock = False
                           End If
                       Next
                   End If
                   
            End If
            
       End Select
       
       .EventEnabled(EventAllEvents) = True
    End With
End Sub


' Ham set cac cong thuc
Sub CellChange1(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
    Dim varTemp As Variant
    Dim cellTagArr() As String
    Dim ct21 As Variant
    
    Dim blCheck_S4A As Boolean
    Dim blCheck_S4B As Boolean
    
    blCheck_S4A = IIf(TAX_Utilities_New.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S4B = IIf(TAX_Utilities_New.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    
    With fps
        
        .Sheet = 1
        ' Chi tieu 21
        .Col = .ColLetterToNumber("I")
        .Row = 23
        ct21 = .Text
        .Sheet = .ActiveSheet
        Select Case .Sheet
            ' Tong hop du lieu tu phu luc 01_1 len to khai
            Case 2 ' phu luc 01_1
                .Col = .ColLetterToNumber("C")
                .Row = Row
                varTemp = .CellTag
                cellTagArr = Split(varTemp, "~")
                If varTemp <> "" Then
                    If (Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("I")) And cellTagArr(1) = "4" And ct21 = "0" Then
                        ' chi tieu 33
                            .Sheet = 1
                            .Col = .ColLetterToNumber("L")
                            .Row = 34
                            .Formula = "L53"
                            .Sheet = .ActiveSheet
                    End If
                    If Col = .ColLetterToNumber("I") And cellTagArr(1) = "1" And ct21 = "0" Then
                             ' chi tieu 26
                            .Sheet = 1
                            .Col = .ColLetterToNumber("J")
                            .Row = 30
                            .Formula = "E54"
                            .Sheet = .ActiveSheet
                    End If
                    If Col = .ColLetterToNumber("I") And cellTagArr(1) = "2" And ct21 = "0" Then
                             ' chi tieu 29
                            .Sheet = 1
                            .Col = .ColLetterToNumber("J")
                            .Row = 32
                            .Formula = "H53"
                            .Sheet = .ActiveSheet
                    End If
                    If Col = .ColLetterToNumber("I") And cellTagArr(1) = "3" And ct21 = "0" Then
                             ' chi tieu 30
                            .Sheet = 1
                            .Col = .ColLetterToNumber("J")
                            .Row = 33
                            .Formula = "I53"
                            .Sheet = .ActiveSheet
                    End If
                    If Col = .ColLetterToNumber("I") And cellTagArr(1) = "4" And ct21 = "0" Then
                             ' chi tieu 32
                            .Sheet = 1
                            .Col = .ColLetterToNumber("J")
                            .Row = 34
                            .Formula = "K53"
                            .Sheet = .ActiveSheet
                    End If
                    If (Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("I")) And cellTagArr(1) = "3" And ct21 = "0" Then
                             ' chi tieu 31
                            .Sheet = 1
                            .Col = .ColLetterToNumber("L")
                            .Row = 33
                            .Formula = "J53"
                            .Sheet = .ActiveSheet
                    End If
                    
                        ' kiem tra neu co phu luc 01_4A thi moi set cong thuc
                    If blCheck_S4A = True Then
                        If Col = .ColLetterToNumber("I") And (cellTagArr(1) = "1" Or cellTagArr(1) = "2" Or cellTagArr(1) = "3" Or cellTagArr(1) = "4") Then
                                .Sheet = 5
                                ' chi tieu tong doanh so ban ra tren phu luc 1
                                .Col = .ColLetterToNumber("D")
                                .Row = 16
                                .Formula = "C24"
                                .Sheet = .ActiveSheet
                        End If
                        If Col = .ColLetterToNumber("I") And (cellTagArr(1) = "2" Or cellTagArr(1) = "3" Or cellTagArr(1) = "4") Then
                                ' chi tieu T?ng doanh thu hàng hóa, d?ch v? bán ra ch?u thu? GTGT
                                .Sheet = 5
                                .Col = .ColLetterToNumber("D")
                                .Row = 17
                                .Formula = "D24"
                                .Sheet = .ActiveSheet
                        End If
                    End If
                End If
            Case 3 ' phu luc 01_2
                .Col = .ColLetterToNumber("C")
                .Row = Row
                varTemp = .CellTag
                cellTagArr = Split(varTemp, "~")
                If varTemp <> "" Then
                    If Col = .ColLetterToNumber("I") And (cellTagArr(1) = "1" Or cellTagArr(1) = "2" Or cellTagArr(1) = "3") And ct21 = "0" Then
                            ' chi tieu 23
                            .Sheet = 1
                            .Col = .ColLetterToNumber("J")
                            .Row = 27
                            .Formula = "D53"
                            .Sheet = .ActiveSheet
                    End If
                    If (Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("I") Or Col = .ColLetterToNumber("J")) And (cellTagArr(1) = "1" Or cellTagArr(1) = "2" Or cellTagArr(1) = "3") And ct21 = "0" Then
                             ' chi tieu 24
                            .Sheet = 1
                            .Col = .ColLetterToNumber("L")
                            .Row = 27
                            .Formula = "E53"
                            .Sheet = .ActiveSheet
                    End If
                    
                    If (Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("I") Or Col = .ColLetterToNumber("J")) And (cellTagArr(1) = "1") Then
                             ' chi tieu 25
                            .Sheet = 1
                            .Col = .ColLetterToNumber("L")
                            .Row = 28
                            .Formula = "F53"
                            .Sheet = .ActiveSheet
                    End If
                    
                    
                    ' kiem tra neu co phu luc 01_4A thi moi set cong thuc
                    If blCheck_S4A = True Then
                        If (Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("I") Or Col = .ColLetterToNumber("J")) And (cellTagArr(1) = "1") Then
                                ' chi tieu tong dong 1 cot 10
                                'A1
                                .Sheet = 5
                                .Col = .ColLetterToNumber("D")
                                .Row = 7
                                .Formula = "C23"
                                .Sheet = .ActiveSheet
                        End If
                        
                        If (Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("I") Or Col = .ColLetterToNumber("J")) And (cellTagArr(1) = "2") Then
                                 ' chi tieu A2-04A
                                .Sheet = 5
                                .Col = .ColLetterToNumber("D")
                                .Row = 9
                                .Formula = "D23"
                                .Sheet = .ActiveSheet
                        End If
                        
                        If (Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("I") Or Col = .ColLetterToNumber("J")) And (cellTagArr(1) = "3") Then
                                'chi tieu tong dong 3 cot 10
                                'A3
                                .Sheet = 5
                                .Col = .ColLetterToNumber("D")
                                .Row = 11
                                .Formula = "E23"
                                .Sheet = .ActiveSheet
                        End If
                        
                        If (Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("I") Or Col = .ColLetterToNumber("J")) And (cellTagArr(1) = "3") Then
                                .Sheet = 5
                                 ' chitieu B4
                                .Col = .ColLetterToNumber("D")
                                .Row = 19
                                .Formula = "E24"
                                .Sheet = .ActiveSheet
                        End If
                   End If
                End If
            Case 5 ' phu luc 01_4A
                        ' bo cac ct tren phu luc 01_4A
                        reset_01_4A
                        If Col = .ColLetterToNumber("D") And Row = 19 Then
                                ' chi tieu 25
                                .Sheet = 1
                                .Col = .ColLetterToNumber("L")
                                .Row = 28
                                .Formula = "F53"
                                .Sheet = .ActiveSheet
                        End If
            Case 6 ' phu luc 01_4B
                        If Col = .ColLetterToNumber("E") And Row = 21 Then
                                ' chi tieu 25
                                .Sheet = 1
                                .Col = .ColLetterToNumber("L")
                                .Row = 28
                                .Formula = "F53"
                                .Sheet = .ActiveSheet
                        End If
            Case 7 ' phu luc 01_5
                    If Col = .ColLetterToNumber("G") Then
                            ' chi tieu 39
                            .Sheet = 1
                            .Col = .ColLetterToNumber("L")
                            .Row = 40
                            .Formula = "D54"
                            .Sheet = .ActiveSheet
                    End If
        End Select
    End With
End Sub

' Ham set lai gia tri cua cong thuc khi tai bang ke 01_1
Sub CellChange2()
    Dim varTemp As Variant
    Dim cellTagArr() As String
    With fps
        .Sheet = .ActiveSheet
        Select Case .Sheet
            Case 2 ' phu luc 01_1
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
        Case 3 ' phu luc 01_2
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
            
        End Select
    End With
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
    Dim strValue4 As String
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_New.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
                Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_New.Data(8).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(8).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_New.Data(8).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(8).getElementsByTagName("Cell").length - 4)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_New.Data(8).getElementsByTagName("Cell")(1)
                   strValue3 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_New.Data(8).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_New.Data(8).getElementsByTagName("Cell").length > 25 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Or (Not IsNullValue(strValue4)) Then
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
    Dim ct21 As Variant
    Dim blCheck_S1 As Boolean
    Dim blCheck_S2 As Boolean
    Dim blCheck_S5 As Boolean
    Dim blCheck_S4A As Boolean
    Dim blCheck_S4B As Boolean
    blCheck_S1 = IIf(TAX_Utilities_New.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S2 = IIf(TAX_Utilities_New.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S5 = IIf(TAX_Utilities_New.NodeValidity.childNodes(6).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S4A = IIf(TAX_Utilities_New.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S4B = IIf(TAX_Utilities_New.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    tempCurrSheet = mCurrentSheet
    .Sheet = 1
    ' Chi tieu 21
    .Col = .ColLetterToNumber("I")
    .Row = 23
    ct21 = .Text
            If blCheck_S1 = True And ct21 = "0" Then
                .Sheet = 1
               ' chi tieu 26
               .Col = .ColLetterToNumber("J")
               .Row = 30
               .Lock = True
                .Sheet = 1
                ' chi tieu 29
                .Col = .ColLetterToNumber("J")
                .Row = 32
                .Lock = True
                 ' chi tieu 30
                .Col = .ColLetterToNumber("J")
                .Row = 33
                .Lock = True
                ' chi tieu 32
                .Col = .ColLetterToNumber("J")
                .Row = 34
                .Lock = True
                ' chi tieu 31
                .Col = .ColLetterToNumber("L")
                .Row = 33
                .Lock = True
                ' chi tieu 33
                .Col = .ColLetterToNumber("L")
                .Row = 34
                 .Lock = True
            ElseIf (blCheck_S1 = False Or blCheck_S1 = True) And ct21 = "1" Then
                .Sheet = 1
               ' chi tieu 26
               .Col = .ColLetterToNumber("J")
               .Row = 30
               .Lock = True
                .Sheet = 1
                ' chi tieu 29
                .Col = .ColLetterToNumber("J")
                .Row = 32
                .Lock = True
                 ' chi tieu 30
                .Col = .ColLetterToNumber("J")
                .Row = 33
                .Lock = True
                ' chi tieu 32
                .Col = .ColLetterToNumber("J")
                .Row = 34
                .Lock = True
                ' chi tieu 31
                .Col = .ColLetterToNumber("L")
                .Row = 33
                .Lock = True
                ' chi tieu 33
                .Col = .ColLetterToNumber("L")
                .Row = 34
                 .Lock = True
            Else
                .Sheet = 1
               ' chi tieu 26
               .Col = .ColLetterToNumber("J")
               .Row = 30
               .Lock = False
                .Sheet = 1
                ' chi tieu 29
                .Col = .ColLetterToNumber("J")
                .Row = 32
                .Lock = False
                 ' chi tieu 30
                .Col = .ColLetterToNumber("J")
                .Row = 33
                .Lock = False
                ' chi tieu 32
                .Col = .ColLetterToNumber("J")
                .Row = 34
                .Lock = False
                ' chi tieu 31
                .Col = .ColLetterToNumber("L")
                .Row = 33
                .Lock = False
                ' chi tieu 33
                .Col = .ColLetterToNumber("L")
                .Row = 34
                 .Lock = False
            End If
            
            If blCheck_S2 = True And ct21 = "0" Then
                .Sheet = 1
                ' Chi tieu 23
                .Col = .ColLetterToNumber("J")
                .Row = 27
                .Lock = True
                ' Chi tieu 24
                .Col = .ColLetterToNumber("L")
                .Row = 27
                .Lock = True
            ElseIf (blCheck_S2 = False Or blCheck_S2 = True) And ct21 = "1" Then
                .Sheet = 1
                ' Chi tieu 23
                .Col = .ColLetterToNumber("J")
                .Row = 27
                .Lock = True
                ' Chi tieu 24
                .Col = .ColLetterToNumber("L")
                .Row = 27
                .Lock = True
            Else
                .Sheet = 1
                ' Chi tieu 23
                .Col = .ColLetterToNumber("J")
                .Row = 27
                .Lock = False
                ' Chi tieu 24
                .Col = .ColLetterToNumber("L")
                .Row = 27
                .Lock = False
            End If
            
            If blCheck_S5 = True Then
                .Sheet = 1
                ' Chi tieu 39
                .Col = .ColLetterToNumber("L")
                .Row = 40
                .Lock = True
            Else
                .Sheet = 1
                ' Chi tieu 39
                .Col = .ColLetterToNumber("L")
                .Row = 40
                .Lock = False
            End If
    .EventEnabled(EventAllEvents) = True
End With
End Sub



' Ham set lai gia tri cua cong thuc khi tai bang ke 01_2
Sub CellChange3()
    Dim blCheck_S4A As Boolean
    Dim blCheck_S4B As Boolean
    
    blCheck_S4A = IIf(TAX_Utilities_New.NodeValidity.childNodes(4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    blCheck_S4B = IIf(TAX_Utilities_New.NodeValidity.childNodes(5).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
    
    ' Tong hop len bang ke 01_4A
        With fps
        .Sheet = .ActiveSheet
        Select Case .Sheet
            Case 2 ' Phu luc 01_1
                If blCheck_S4A = True Then
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
            Case 3 ' phu luc 01_2
                If blCheck_S4A = True Then
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
         End Select
    End With
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

' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 22
    denRowDL = 32
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "9_11_2~9_11_3~11_13_4~10_12_5~10_12_6~10_11_7~9_11_8~20_30_9"
    ' end
End Sub
