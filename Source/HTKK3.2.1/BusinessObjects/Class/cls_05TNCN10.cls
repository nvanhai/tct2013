VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_05TNCN10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "To khai quyet toan thu nhap ca nhan mau 04-TNCN" interface sheets
'this Class is belong to TAX_Business project which will be compline to DLL

Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrSolanBosung As String

Private arrMST5A() As String
Private arrMST5B() As String

Public flgLoadToKhai As Boolean
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    Dim vCell As Variant
    With fps
        SetDateFormat fps, 1, 43, .ColLetterToNumber("I"), DDMMYYYY
        
        .Sheet = 1 'To khai 04/TNCN
        .Row = 43
        .Col = .ColLetterToNumber("I")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        resetFormula
        
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
     With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        
        If Strloaitk = "TKCT" And StrSolanBosung = "" Then
            .Col = .ColLetterToNumber("C")
            .Row = 40
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 40
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 40
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKCT" And Val(StrSolanBosung) > 0 Then
            .Col = .ColLetterToNumber("C")
            .Row = 40
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 40
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 40
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        .Sheet = 2
        .ActiveSheet = .Sheet
        .Row = 2
        .Col = .ColLetterToNumber("D")
        .CellType = CellTypeButton
        .TypeButtonTextColor = RGB(0, 0, 255)
        .TypeButtonColor = &H8000000F
        .TypeButtonAlign = TypeButtonAlignRight
        .Font = "Arial"
        .FontSize = 8
        .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
        
        
        .Sheet = 3
        .ActiveSheet = .Sheet
        .Row = 2
        .Col = .ColLetterToNumber("C")
        .CellType = CellTypeButton
        .TypeButtonTextColor = RGB(0, 0, 255)
        .TypeButtonColor = &H8000000F
        .TypeButtonAlign = TypeButtonAlignRight
        .Font = "Arial"
        .FontSize = 8
        .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Function GetValidatedCellIndex(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long) As Integer
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long
    
    Set xmlCellNode = TAX_Utilities_New.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
    Set xmlCellsNode = xmlCellNode.parentNode
    
    'Get Index of anchor cell
    For lCtrl = 1 To xmlCellsNode.childNodes.length
        If GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID") = GetAttribute(xmlCellNode, "CellID") Then _
            Exit For
    Next
    GetValidatedCellIndex = lCtrl
    
End Function

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    'UpdateSheets
   
End Sub


Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Integer, iCol As Long, iRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
    Dim tCol As Long, tRow As Long
    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 2 Then
                    fps.EventEnabled(EventAllEvents) = False
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 22
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 22
                         i = i + 1
                    Loop Until .Text = "aa"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    
                    .SetActiveCell iCol, iRow
            End If
            If .ActiveSheet = 3 Then
                    fps.EventEnabled(EventAllEvents) = False
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 21
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 21
                         i = i + 1
                    Loop Until .Text = "aa"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                     
                    .SetActiveCell iCol, iRow
            End If
        End If
  
      fps.EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim varHoVaTen, varMST, varCMT  As Variant
    Dim varThuNhap, varBaoHiem, varSoNguoiGiamTru, varSoThangGiamTru, varThuNhapGiamThue, varSoThueDuocGiam, varSoThueKhauTru As Variant
    Dim strCheck As String
    Dim flgErr As Boolean
    
    
    flgErr = True
    
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            'check date
            If Col = .ColLetterToNumber("I") And Row = 43 Then
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
        ElseIf .ActiveSheet = 2 Then
            ' Cot Ho va ten
            If Col = .ColLetterToNumber("D") Then
                .GetText .ColLetterToNumber("D"), Row, varHoVaTen
                ' Cu them mot lao dong thi tinh la mot nguoi
                If Trim(varHoVaTen) <> vbNullString Then
                    ' Them mot nguoi vao cot N
                    .Col = .ColLetterToNumber("N")
                    .Text = "1"
                Else ' Neu ko thi ko tinh nguoi do
                    flgErr = False
                    .Col = .ColLetterToNumber("N")
                    .Text = "0"
                End If
            End If
            
            ' Lay so thue da khau tru cua lao dong do
            If Col = .ColLetterToNumber("M") Then
                .GetText .ColLetterToNumber("M"), Row, varSoThueKhauTru
                If Val(varSoThueKhauTru) > 0 Then
                    ' Va tinh la lao dong co thue khau tru
                    .Col = .ColLetterToNumber("O")
                    .Text = "1"
                    ' Tinh tong thu nhap cua rieng lao dong co thue khau tru nay
                    .GetText .ColLetterToNumber("G"), Row, varThuNhap
                    .Col = .ColLetterToNumber("P")
                    .Text = varThuNhap
                    ' Tinh tong so nguoi phu thuoc duoc giam tru
                    .GetText .ColLetterToNumber("I"), Row, varSoNguoiGiamTru
                    .Col = .ColLetterToNumber("Q")
                    .Text = varSoNguoiGiamTru
                Else
                    ' Va tinh la lao dong ko co thue khau tru
                    .Col = .ColLetterToNumber("O")
                    .Text = "0"
                    ' Khong tinh tong thu nhap cua lao dong nay
                    .Col = .ColLetterToNumber("P")
                    .Text = "0"
                    ' Khong tinh so nguoi phu thuoc duoc giam tru
                    .Col = .ColLetterToNumber("Q")
                    .Text = "0"
                End If
            End If
        ElseIf .ActiveSheet = 3 Then
            If Col = .ColLetterToNumber("C") Then
                .GetText .ColLetterToNumber("C"), Row, varTemp
                ' Cu them mot lao dong thi tinh la mot nguoi
                If Trim(varTemp) <> vbNullString Then
                    .Col = .ColLetterToNumber("J")
                    .Text = "1"
                Else ' Neu ko thi ko tinh nguoi do
                    .Col = .ColLetterToNumber("J")
                    .Text = "0"
                End If
            End If
            ' Lay so thue da khau tru cua lao dong do
            If Col = .ColLetterToNumber("I") Then
                .GetText .ColLetterToNumber("I"), Row, varTemp
                If Val(varTemp) > 0 Then
                    ' Va tinh la lao dong co thue khau tru
                    .Col = .ColLetterToNumber("K")
                    .Text = "1"
                    ' Tinh tong thu nhap cua rieng lao dong co thue khau tru nay
                    .GetText .ColLetterToNumber("F"), Row, varTemp
                    .Col = .ColLetterToNumber("L")
                    .Text = varTemp
                Else
                    ' Va tinh la lao dong ko co thue khau tru
                    .Col = .ColLetterToNumber("K")
                    .Text = "0"
                    ' Khong tinh tong thu nhap cua lao dong nay
                    .Col = .ColLetterToNumber("L")
                    .Text = "0"
                End If
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
    
End Sub

'Calling this function whenever we don't need this class anymore.
'No parameter
Public Sub finish()
    Dim ASheet As Integer, SSheet As Integer
            
    If flgLoadToKhai = False Then defaultFormula
       
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        
        SetData
        
        ' Lay cac mang MST trong cac bang ke va phu luc 05A, 05B
        prepareMST
        
        kiemTraDuLieuImport
        
        
        
        UpdateSheets
        CheckDynamicError 'Set Exception Error on cells of interface
            
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
        
    End With
    
    resetFormula
    
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
        
    With fps
        CheckErrorMST
        ' Neu ton tai file data bang ke 05A thi kiem tra tren bang ke 5A
        If IIf(TAX_Utilities_New.NodeValidity.childNodes(.SheetCount - 4).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            checkSheet1
        End If
        If IIf(TAX_Utilities_New.NodeValidity.childNodes(.SheetCount - 3).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            checkSheet2
        End If
    End With
End Sub

Private Sub checkSheet1()
    Dim vErrMsgHoVaTen, vErrMsgMST, vErrMsgMST1, vErrMsgMST2, vErrMsgCMT, vErrMsgCMT1, vErrMsgMSTCMT As Variant
    Dim vErrMsgBaoHiem, vErrMsgThangGiamTru, vErrMsgThuNhapGiamThue, vErrMsgSoThueDuocGiam, vErrMsgSoThueKhauTru As Variant
    Dim vLastRow As Variant
    Dim varTemp As Variant
    Dim i, j As Integer
    Dim iFocusFlag As Integer
    Dim strFocusSheetName, strFocusCol, strFocusRow As String
    
    With fps
        .Sheet = .SheetCount
                
        .GetText .ColLetterToNumber("E"), 21, vErrMsgHoVaTen
        .GetText .ColLetterToNumber("E"), 22, vErrMsgMST
        .GetText .ColLetterToNumber("E"), 23, vErrMsgMST1
        .GetText .ColLetterToNumber("E"), 24, vErrMsgMST2
        .GetText .ColLetterToNumber("E"), 25, vErrMsgCMT
        .GetText .ColLetterToNumber("E"), 26, vErrMsgCMT1
        .GetText .ColLetterToNumber("E"), 27, vErrMsgMSTCMT
        
        .GetText .ColLetterToNumber("E"), 28, vErrMsgBaoHiem
        .GetText .ColLetterToNumber("E"), 29, vErrMsgThangGiamTru
        .GetText .ColLetterToNumber("E"), 30, vErrMsgThuNhapGiamThue
        .GetText .ColLetterToNumber("E"), 31, vErrMsgSoThueDuocGiam
        .GetText .ColLetterToNumber("E"), 32, vErrMsgSoThueKhauTru
        
        .Sheet = 2 'Check tinh xac thuc cua cac thong tin tren bang ke 05A/TNCN
                        
        ' Cot check loi la AD24 duoc danh dau xem bang ke co con bi sai hay khong. Neu sai thi moi check loi tiep
'        .GetText .ColLetterToNumber("AD"), 16, varTemp
'        If Val(varTemp) = 0 Then Exit Sub
        
        For j = 22 To .MaxRows
            .Row = j
            .GetText .ColLetterToNumber("B"), .Row, vLastRow
            If UCase(vLastRow) = "AA" Then Exit For ' Check cho den het cac dong co du lieu thi thoi
            
            .GetText .ColLetterToNumber("AD"), .Row, varTemp
            If Val(varTemp) = 1 Then
                .GetText .ColLetterToNumber("R"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("D")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgHoVaTen
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("S"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("E")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMST
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("T"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("E")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMST1
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("U"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("E")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMST2
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("V"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("F")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgCMT
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("W"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("F")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgCMT1
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("X"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("E")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMSTCMT
                    .BackColor = mErrorColor

                    .Col = .ColLetterToNumber("F")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMSTCMT
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("Y"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("H")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgBaoHiem
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("Z"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("J")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgThangGiamTru
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("AA"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("K")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgThuNhapGiamThue
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("AB"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("L")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgSoThueDuocGiam
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("AC"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("M")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgSoThueKhauTru
                    .BackColor = mErrorColor
                End If
            Else
                .Col = .ColLetterToNumber("D")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("E")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("F")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("G")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("G")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("H")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("I")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("J")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("K")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("L")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("M")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
        Next
        
    End With
End Sub

Private Sub checkSheet2()
    Dim vErrMsgHoVaTen, vErrMsgMST, vErrMsgMST1, vErrMsgMST2, vErrMsgCMT, vErrMsgCMT1, vErrMsgMSTCMT As Variant
    Dim vErrMsgThuNhapGiamThue, vErrMsgSoThueDuocGiam, vErrMsgSoThueKhauTru As Variant
    Dim vLastRow As Variant
    Dim varTemp As Variant
    Dim i, j As Integer
    Dim iFocusFlag As Integer
    Dim strFocusSheetName, strFocusCol, strFocusRow As String
    
    With fps
        .Sheet = .SheetCount
                
        .GetText .ColLetterToNumber("E"), 38, vErrMsgHoVaTen
        .GetText .ColLetterToNumber("E"), 39, vErrMsgMST
        .GetText .ColLetterToNumber("E"), 40, vErrMsgMST1
        .GetText .ColLetterToNumber("E"), 41, vErrMsgMST2
        .GetText .ColLetterToNumber("E"), 42, vErrMsgCMT
        .GetText .ColLetterToNumber("E"), 43, vErrMsgCMT1
        .GetText .ColLetterToNumber("E"), 44, vErrMsgMSTCMT
                
        .GetText .ColLetterToNumber("E"), 45, vErrMsgThuNhapGiamThue
        .GetText .ColLetterToNumber("E"), 46, vErrMsgSoThueDuocGiam
        .GetText .ColLetterToNumber("E"), 47, vErrMsgSoThueKhauTru
        
        .Sheet = 3 'Check tinh xac thuc cua cac thong tin tren bang ke 05A/TNCN
                        
        ' Cot check loi la AD24 duoc danh dau xem bang ke co con bi sai hay khong. Neu sai thi moi check loi tiep
'        .GetText .ColLetterToNumber("W"), 16, varTemp
'        If Val(varTemp) = 0 Then Exit Sub
        
        For j = 21 To .MaxRows
            .Row = j
            .GetText .ColLetterToNumber("B"), .Row, vLastRow
            If UCase(vLastRow) = "AA" Then Exit For ' Check cho den het cac dong co du lieu thi thoi
            
            .GetText .ColLetterToNumber("W"), .Row, varTemp
            If Val(varTemp) = 1 Then
                .GetText .ColLetterToNumber("M"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("C")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgHoVaTen
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("N"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("D")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMST
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("O"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("D")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMST1
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("P"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("D")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMST2
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("Q"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("E")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgCMT
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("R"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("E")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgCMT1
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("S"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("D")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMSTCMT
                    .BackColor = mErrorColor

                    .Col = .ColLetterToNumber("E")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMSTCMT
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("T"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("G")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgThuNhapGiamThue
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("U"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("H")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgSoThueDuocGiam
                    .BackColor = mErrorColor
                End If
                
                .GetText .ColLetterToNumber("V"), .Row, varTemp
                If Val(varTemp) = 1 Then
                    .Col = .ColLetterToNumber("I")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgSoThueKhauTru
                    .BackColor = mErrorColor
                End If
            Else
                .Col = .ColLetterToNumber("C")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("D")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("E")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("G")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("H")
                .CellNote = ""
                .BackColor = mNonErrorColor
                .Col = .ColLetterToNumber("I")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
        Next
        
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
        .Row = 3
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
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_New.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "05A"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_New.Data(1).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in E column and F column
                    If lCol = fps.ColLetterToNumber("D") Or _
                       lCol = fps.ColLetterToNumber("E") Then
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

             Case "05B"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_New.Data(2).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in E column and F column
                    If lCol = fps.ColLetterToNumber("C") Or _
                       lCol = fps.ColLetterToNumber("D") Then
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
        End Select
    Next
    
    Set xmlCellNode = Nothing
    Set xmlNodeValid = Nothing
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
        
        If (GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "Active") <> "0") Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
        If (GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(1), "Active") <> "0") Then
        .Sheet = 2
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_New.Data(1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
        If (GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(2), "Active") <> "0") Then
        .Sheet = 3
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_New.Data(2).getElementsByTagName("Cell")
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

Public Sub FinishImport()

End Sub

Private Sub prepareMST()
    Dim countMST As Long, i As Integer
    Dim varMST As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 2 'Bang ke 5A
        .Row = 22
        countMST = 0
        Do
            countMST = countMST + 1
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        If countMST = 0 Then
            ReDim arrMST5A(1) As String
            arrMST5A(0) = ""
        Else
            .Row = 22
            ReDim arrMST5A(22 + countMST - 1) As String
            For i = 0 To 21
                arrMST5A(i) = ""
            Next
            Do
                .GetText .ColLetterToNumber("E"), .Row, varMST
                If Trim(varMST) <> vbNullString Then
                    arrMST5A(.Row) = varMST
                Else
                    arrMST5A(.Row) = ""
                End If
                .Row = .Row + 1
                .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
        End If
        
        .Sheet = 3 'Bang ke 5B
        .Row = 21
        countMST = 0
        Do
            countMST = countMST + 1
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        If countMST = 0 Then
            ReDim arrMST5B(1) As String
            arrMST5B(0) = ""
        Else
            .Row = 21
            ReDim arrMST5B(21 + countMST - 1) As String
            For i = 0 To 20
                arrMST5B(i) = ""
            Next
            Do
                .GetText .ColLetterToNumber("D"), .Row, varMST
                If Trim(varMST) <> vbNullString Then
                    arrMST5B(.Row) = varMST
                Else
                    arrMST5B(.Row) = ""
                End If
                .Row = .Row + 1
                .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
        End If
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Public Sub kiemTraDuLieuImport()
    kiemtra5A
    kiemtra5B
End Sub

Private Sub kiemtra5A()
    Dim varTemp As Variant
    Dim varHoVaTen, varMST, varCMT  As Variant
    Dim varThuNhap, varBaoHiem, varSoNguoiGiamTru, varSoThangGiamTru, varThuNhapGiamThue, varSoThueDuocGiam, varSoThueKhauTru As Variant
    Dim strCheck As String
    Dim flgErr As Boolean
    Dim varMSTCQCT As Variant
    
    Dim startRowMST As Long
    Dim rowCountMST As Long
    Dim totalRowMST As Long
    
    flgErr = True
    totalRowMST = UBound(arrMST5A)
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 2 'Bang ke 5A
        ' MST co quan chi tra
        .GetText .ColLetterToNumber("E"), 9, varMSTCQCT
                
        .Row = 22
        Do
            ' Cot Ho va ten
            .GetText .ColLetterToNumber("D"), .Row, varHoVaTen
            ' Cu them mot lao dong thi tinh la mot nguoi
            If Trim(varHoVaTen) <> vbNullString Then
                ' Them mot nguoi vao cot N
                .Col = .ColLetterToNumber("N")
                .Text = "1"
                ' Va danh dau cot Ho va ten nay dung de sau nay ko bat loi
                .Col = .ColLetterToNumber("R")
                .Text = "0"
            Else ' Neu ko thi ko tinh nguoi do
                flgErr = False
                .Col = .ColLetterToNumber("N")
                .Text = "0"
                ' Va danh dau cot Ho va ten nay sai de sau nay bat loi
                .Col = .ColLetterToNumber("R")
                .Text = "1"
            End If
            
            ' Cot MST
            
            .GetText .ColLetterToNumber("E"), .Row, varMST
            ' Ma la null
            If Trim(varMST) = vbNullString Then
                flgErr = False
                ' Danh dau cot MST nay sai de sau nay bat loi
                .Col = .ColLetterToNumber("X")
                .Text = "1"
                ' Cac cot lien quan den MST con lai la dung
                .Col = .ColLetterToNumber("S")
                .Text = "0"
                .Col = .ColLetterToNumber("T")
                .Text = "0"
                .Col = .ColLetterToNumber("U")
                .Text = "0"
            Else
                ' Ma > 10 so la sai
                If Len(Trim(varMST)) <> 10 Then
                    flgErr = False
                    ' Danh dau cot MST nay sai de sau nay bat loi
                    .Col = .ColLetterToNumber("S")
                    .Text = "1"
                    ' Cac cot lien quan den MST con lai la dung
                    .Col = .ColLetterToNumber("X")
                    .Text = "0"
                    .Col = .ColLetterToNumber("T")
                    .Text = "0"
                    .Col = .ColLetterToNumber("U")
                    .Text = "0"
                End If
                ' Ma 10 so
                If Len(Trim(varMST)) = 10 Then
                    strCheck = CheckTaxCode(Mid(Trim(varMST), 1, 1), Mid(Trim(varMST), 2, 1), Mid(Trim(varMST), 3, 1), Mid(Trim(varMST), 4, 1), Mid(Trim(varMST), 5, 1), Mid(Trim(varMST), 6, 1), Mid(Trim(varMST), 7, 1), Mid(Trim(varMST), 8, 1), Mid(Trim(varMST), 9, 1), Mid(Trim(varMST), 10, 1), Mid(Trim(varMST), 11, 1), Mid(Trim(varMST), 12, 1), Mid(Trim(varMST), 13, 1))
                    If Mid(strCheck, 2, 1) <> "0" Then
                        flgErr = False
                        ' Danh dau cot MST nay sai de sau nay bat loi
                        .Col = .ColLetterToNumber("S")
                        .Text = "1"
                        ' Cac cot lien quan den MST con lai la dung
                        .Col = .ColLetterToNumber("X")
                        .Text = "0"
                        .Col = .ColLetterToNumber("T")
                        .Text = "0"
                        .Col = .ColLetterToNumber("U")
                        .Text = "0"
                    Else
                        .Col = .ColLetterToNumber("S")
                        .Text = "0"
                        ' Kiem tra trung voi MST cqct
                        If Trim(varMST) = Trim(varMSTCQCT) Then
                           flgErr = False
                           .Col = .ColLetterToNumber("T")
                           .Text = "1"
                           .Col = .ColLetterToNumber("S")
                           .Text = "0"
                           ' Cac cot lien quan den MST con lai la dung
                           .Col = .ColLetterToNumber("X")
                           .Text = "0"
                           .Col = .ColLetterToNumber("U")
                           .Text = "0"
                        Else
                            .Col = .ColLetterToNumber("T")
                            .Text = "0"
                            ' Neu ko co loi nao kiem tra xem co trung voi MST nao tren bang ke khong
                            startRowMST = .Row + 1
                            If totalRowMST > 1 Then
                               For rowCountMST = startRowMST To totalRowMST
                                   If Trim(varMST) = Trim(arrMST5A(rowCountMST)) Then
                                        flgErr = False
                                        .Col = .ColLetterToNumber("U")
                                        .Text = "1"
                                        ' Danh dau cot MST nay dung
                                        .Col = .ColLetterToNumber("S")
                                        .Text = "0"
                                        ' Cac cot lien quan den MST con lai la dung
                                        .Col = .ColLetterToNumber("X")
                                        .Text = "0"
                                        Exit For
                                   Else
                                        .Col = .ColLetterToNumber("U")
                                        .Text = "0"
                                   End If
                               Next
                            End If
                         End If
                    End If
                End If
            End If
                        
            ' Cot CMT
            
            .GetText .ColLetterToNumber("E"), .Row, varMST
            .GetText .ColLetterToNumber("F"), .Row, varCMT
            If Trim(varMST) = vbNullString And Trim(varCMT) = vbNullString Then
                flgErr = False
                ' Danh dau ca cot MST va CMT deu khong co
                .Col = .ColLetterToNumber("X")
                .Text = "1"
                ' Danh dau cac cot con lai la dung
                .Col = .ColLetterToNumber("S")
                .Text = "0"
                .Col = .ColLetterToNumber("T")
                .Text = "0"
                .Col = .ColLetterToNumber("U")
                .Text = "0"
                .Col = .ColLetterToNumber("V")
                .Text = "0"
                .Col = .ColLetterToNumber("W")
                .Text = "0"
            ElseIf Trim(varCMT) <> vbNullString Then
                If InStr(1, Trim(varCMT), " ", vbTextCompare) <> 0 Then
                    flgErr = False
                    ' Danh dau cot CMT nay sai de sau nay bat loi
                    .Col = .ColLetterToNumber("W")
                    .Text = "1"
                    ' Danh dau cac cot con lai la dung
                    .Col = .ColLetterToNumber("V")
                    .Text = "0"
                    .Col = .ColLetterToNumber("X")
                    .Text = "0"
                Else
                    If Trim(varMST) = vbNullString Then flgErr = True
                    ' Danh dau cac cot con lai la dung
                    .Col = .ColLetterToNumber("W")
                    .Text = "0"
                    .Col = .ColLetterToNumber("V")
                    .Text = "0"
                    .Col = .ColLetterToNumber("X")
                    .Text = "0"
                End If
            End If
            
            ' Cot Bao hiem bat buoc dong
            
            .GetText .ColLetterToNumber("G"), .Row, varThuNhap
            .GetText .ColLetterToNumber("H"), .Row, varBaoHiem
            If Val(varBaoHiem) > Val(varThuNhap) Then
                flgErr = False
                ' Danh dau cot Bao hiem nay sai sau nay bat loi
                .Col = .ColLetterToNumber("Y")
                .Text = "1"
            Else
                ' Danh dau cac cot con lai la dung
                .Col = .ColLetterToNumber("Y")
                .Text = "0"
            End If
            
            
            ' Cot So thang giam tru

            .GetText .ColLetterToNumber("I"), .Row, varSoNguoiGiamTru
            .GetText .ColLetterToNumber("J"), .Row, varSoThangGiamTru
            ' Nam 2009
            If Val(TAX_Utilities_New.Year) = 2009 And Val(varSoThangGiamTru) > Val(varSoNguoiGiamTru) * 6 Then
                flgErr = False
                ' Danh dau cot Bao hiem nay sai sau nay bat loi
                .Col = .ColLetterToNumber("Z")
                .Text = "1"
            ' Tu nam 2010 tro di
            ElseIf Val(TAX_Utilities_New.Year) > 2009 And Val(varSoThangGiamTru) > Val(varSoNguoiGiamTru) * 12 Then
                flgErr = False
                ' Danh dau cot Bao hiem nay sai sau nay bat loi
                .Col = .ColLetterToNumber("Z")
                .Text = "1"
            Else
                ' Danh dau cac cot con lai la dung
                .Col = .ColLetterToNumber("Z")
                .Text = "0"
            End If

            
            ' Cot thu nhap giam thue

            .GetText .ColLetterToNumber("G"), .Row, varThuNhap
            .GetText .ColLetterToNumber("K"), .Row, varThuNhapGiamThue
            If Val(varThuNhapGiamThue) > Val(varThuNhap) Then
                flgErr = False
                ' Danh dau cot thu nhap giam thue nay sai sau nay bat loi
                .Col = .ColLetterToNumber("AA")
                .Text = "1"
            Else
                ' Danh dau cac cot con lai la dung
                .Col = .ColLetterToNumber("AA")
                .Text = "0"
            End If

            
            ' Cot Thue thu nhap ca nhan giam thue

            .GetText .ColLetterToNumber("L"), .Row, varSoThueDuocGiam
            .GetText .ColLetterToNumber("K"), .Row, varThuNhapGiamThue
            If (Val(varSoThueDuocGiam) > 0) And Val(varSoThueDuocGiam) >= Val(varThuNhapGiamThue) Then
                flgErr = False
                ' Danh dau cot so thue thu nhap duoc giam nay sai sau nay bat loi
                .Col = .ColLetterToNumber("AB")
                .Text = "1"
            Else
                ' Danh dau cac cot con lai la dung
                .Col = .ColLetterToNumber("AB")
                .Text = "0"
            End If

            
            ' Cot Thue da khau tru

            .GetText .ColLetterToNumber("M"), .Row, varSoThueKhauTru
            .GetText .ColLetterToNumber("G"), .Row, varThuNhap
            If Val(varSoThueKhauTru) > 0 And Val(varSoThueKhauTru) >= Val(varThuNhap) Then
                flgErr = False
                ' Danh dau cot so thue khau tru nay sai sau nay bat loi
                .Col = .ColLetterToNumber("AC")
                .Text = "1"
            Else
                ' Danh dau cac cot con lai la dung
                .Col = .ColLetterToNumber("AC")
                .Text = "0"
            End If

            
            ' Neu co bat cu mot chi tieu nao do bi sai thi danh dau vao de con bat loi
            If flgErr = False Then
                ' Danh dau cot check loi nay sai sau nay bat loi
                .Col = .ColLetterToNumber("AD")
                .Text = "1"
            Else
                .Col = .ColLetterToNumber("AD")
                .Text = "0"
            End If
            
            ' Lay so thue da khau tru cua lao dong do

            .GetText .ColLetterToNumber("M"), .Row, varSoThueKhauTru
            If Val(varSoThueKhauTru) > 0 Then
                ' Va tinh la lao dong co thue khau tru
                .Col = .ColLetterToNumber("O")
                .Text = "1"
                ' Tinh tong thu nhap cua rieng lao dong co thue khau tru nay
                .GetText .ColLetterToNumber("G"), .Row, varThuNhap
                .Col = .ColLetterToNumber("P")
                .Text = varThuNhap
                ' Tinh tong so nguoi phu thuoc duoc giam tru
                .GetText .ColLetterToNumber("I"), .Row, varSoNguoiGiamTru
                .Col = .ColLetterToNumber("Q")
                .Text = varSoNguoiGiamTru
            Else
                ' Va tinh la lao dong ko co thue khau tru
                .Col = .ColLetterToNumber("O")
                .Text = "0"
                ' Khong tinh tong thu nhap cua lao dong nay
                .Col = .ColLetterToNumber("P")
                .Text = "0"
                ' Khong tinh so nguoi phu thuoc duoc giam tru
                .Col = .ColLetterToNumber("Q")
                .Text = "0"
            End If

            ' Trong truong hop cot Ho ten, MST, CMT, ThuNhap khong co gia tri thi ko kiem tra
            If (Trim(varHoVaTen) = vbNullString And (Trim(varMST) = vbNullString Or Trim(varCMT) = vbNullString) And Val(varThuNhap) = 0) Then
                ' Va danh dau cac cot sau de sau nay ko bat loi
                ' Ho va ten
                .Col = .ColLetterToNumber("R")
                .Text = "0"
                ' MST
                .Col = .ColLetterToNumber("S")
                .Text = "0"
                .Col = .ColLetterToNumber("T")
                .Text = "0"
                .Col = .ColLetterToNumber("U")
                .Text = "0"
                ' CMT
                .Col = .ColLetterToNumber("V")
                .Text = "0"
                .Col = .ColLetterToNumber("W")
                .Text = "0"
                ' MST & CMT
                .Col = .ColLetterToNumber("X")
                .Text = "0"
                ' Cot check loi
                .Col = .ColLetterToNumber("AD")
                .Text = "0"
            End If
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub kiemtra5B()
    Dim varTemp As Variant
    Dim varHoVaTen, varMST, varCMT  As Variant
    Dim varThuNhap, varBaoHiem, varSoNguoiGiamTru, varSoThangGiamTru, varThuNhapGiamThue, varSoThueDuocGiam, varSoThueKhauTru As Variant
    Dim strCheck As String
    Dim flgErr As Boolean
    Dim varMSTCQCT As Variant
    
    Dim startRowMST As Long
    Dim rowCountMST As Long
    Dim totalRowMST As Long
    
    flgErr = True
    totalRowMST = UBound(arrMST5B)
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 3 'Bang ke 5B
        ' MST co quan chi tra
        .GetText .ColLetterToNumber("D"), 8, varMSTCQCT
                
        .Row = 21
        Do
            ' Cot Ho va ten
            .GetText .ColLetterToNumber("C"), .Row, varHoVaTen
            ' Cu them mot lao dong thi tinh la mot nguoi
            If Trim(varHoVaTen) <> vbNullString Then
                ' Them mot nguoi vao cot
                .Col = .ColLetterToNumber("J")
                .Text = "1"
                ' Va danh dau cot Ho va ten nay dung de sau nay ko bat loi
                .Col = .ColLetterToNumber("M")
                .Text = "0"
            Else ' Neu ko thi ko tinh nguoi do
                flgErr = False
                .Col = .ColLetterToNumber("J")
                .Text = "0"
                ' Va danh dau cot Ho va ten nay sai de sau nay bat loi
                .Col = .ColLetterToNumber("M")
                .Text = "1"
            End If
            
            ' Cot MST
            
            .GetText .ColLetterToNumber("D"), .Row, varMST
            ' Ma la null
            If Trim(varMST) = vbNullString Then
                flgErr = False
                ' Danh dau cot MST nay sai de sau nay bat loi
                .Col = .ColLetterToNumber("S")
                .Text = "1"
                ' Cac cot lien quan den MST con lai la dung
                .Col = .ColLetterToNumber("N")
                .Text = "0"
                .Col = .ColLetterToNumber("O")
                .Text = "0"
                .Col = .ColLetterToNumber("P")
                .Text = "0"
            Else
                ' Ma > 10 so la sai
                If Len(Trim(varMST)) <> 10 Then
                    flgErr = False
                    ' Danh dau cot MST nay sai de sau nay bat loi
                    .Col = .ColLetterToNumber("N")
                    .Text = "1"
                    ' Cac cot lien quan den MST con lai la dung
                    .Col = .ColLetterToNumber("S")
                    .Text = "0"
                    .Col = .ColLetterToNumber("O")
                    .Text = "0"
                    .Col = .ColLetterToNumber("P")
                    .Text = "0"
                End If
                ' Ma 10 so
                If Len(Trim(varMST)) = 10 Then
                    strCheck = CheckTaxCode(Mid(Trim(varMST), 1, 1), Mid(Trim(varMST), 2, 1), Mid(Trim(varMST), 3, 1), Mid(Trim(varMST), 4, 1), Mid(Trim(varMST), 5, 1), Mid(Trim(varMST), 6, 1), Mid(Trim(varMST), 7, 1), Mid(Trim(varMST), 8, 1), Mid(Trim(varMST), 9, 1), Mid(Trim(varMST), 10, 1), Mid(Trim(varMST), 11, 1), Mid(Trim(varMST), 12, 1), Mid(Trim(varMST), 13, 1))
                    If Mid(strCheck, 2, 1) <> "0" Then
                        flgErr = False
                        ' Danh dau cot MST nay sai de sau nay bat loi
                        .Col = .ColLetterToNumber("N")
                        .Text = "1"
                        ' Cac cot lien quan den MST con lai la dung
                        .Col = .ColLetterToNumber("S")
                        .Text = "0"
                        .Col = .ColLetterToNumber("O")
                        .Text = "0"
                        .Col = .ColLetterToNumber("P")
                        .Text = "0"
                    Else
                        .Col = .ColLetterToNumber("N")
                        .Text = "0"
                        ' Kiem tra trung voi MST cqct
                        If Trim(varMST) = Trim(varMSTCQCT) Then
                           flgErr = False
                           .Col = .ColLetterToNumber("O")
                           .Text = "1"
                           .Col = .ColLetterToNumber("N")
                           .Text = "0"
                           ' Cac cot lien quan den MST con lai la dung
                           .Col = .ColLetterToNumber("P")
                           .Text = "0"
                           .Col = .ColLetterToNumber("S")
                           .Text = "0"
                        Else
                            .Col = .ColLetterToNumber("O")
                            .Text = "0"
                            ' Neu ko co loi nao kiem tra xem co trung voi MST nao tren bang ke khong
                            startRowMST = .Row + 1
                            If totalRowMST > 1 Then
                               For rowCountMST = startRowMST To totalRowMST
                                   If Trim(varMST) = Trim(arrMST5B(rowCountMST)) Then
                                        flgErr = False
                                        .Col = .ColLetterToNumber("P")
                                        .Text = "1"
                                        ' Danh dau cot MST nay dung
                                        .Col = .ColLetterToNumber("N")
                                        .Text = "0"
                                        ' Cac cot lien quan den MST con lai la dung
                                        .Col = .ColLetterToNumber("S")
                                        .Text = "0"
                                        Exit For
                                   Else
                                        .Col = .ColLetterToNumber("P")
                                        .Text = "0"
                                   End If
                               Next
                            End If
                         End If
                    End If
                End If
            End If
                        
            ' Cot CMT
            
            .GetText .ColLetterToNumber("D"), .Row, varMST
            .GetText .ColLetterToNumber("E"), .Row, varCMT
            If Trim(varMST) = vbNullString And Trim(varCMT) = vbNullString Then
                flgErr = False
                ' Danh dau ca cot MST va CMT deu khong co
                .Col = .ColLetterToNumber("S")
                .Text = "1"
                ' Danh dau cac cot con lai la dung
                .Col = .ColLetterToNumber("N")
                .Text = "0"
                .Col = .ColLetterToNumber("O")
                .Text = "0"
                .Col = .ColLetterToNumber("P")
                .Text = "0"
                .Col = .ColLetterToNumber("Q")
                .Text = "0"
                .Col = .ColLetterToNumber("R")
                .Text = "0"
            ElseIf Trim(varCMT) <> vbNullString Then
                If InStr(1, Trim(varCMT), " ", vbTextCompare) <> 0 Then
                    flgErr = False
                    ' Danh dau cot CMT nay sai de sau nay bat loi
                    .Col = .ColLetterToNumber("R")
                    .Text = "1"
                    ' Danh dau cac cot con lai la dung
                    .Col = .ColLetterToNumber("Q")
                    .Text = "0"
                    .Col = .ColLetterToNumber("S")
                    .Text = "0"
                Else
                    If Trim(varMST) = vbNullString Then flgErr = True
                    ' Danh dau cac cot con lai la dung
                    .Col = .ColLetterToNumber("R")
                    .Text = "0"
                    .Col = .ColLetterToNumber("Q")
                    .Text = "0"
                    .Col = .ColLetterToNumber("S")
                    .Text = "0"
                End If
            End If
                        
            ' Cot thu nhap giam thue

            .GetText .ColLetterToNumber("F"), .Row, varThuNhap
            .GetText .ColLetterToNumber("G"), .Row, varThuNhapGiamThue
            If Val(varThuNhapGiamThue) > Val(varThuNhap) Then
                flgErr = False
                ' Danh dau cot thu nhap giam thue nay sai sau nay bat loi
                .Col = .ColLetterToNumber("T")
                .Text = "1"
            Else
                ' Danh dau cac cot con lai la dung
                .Col = .ColLetterToNumber("T")
                .Text = "0"
            End If

            
            ' Cot Thue thu nhap ca nhan giam thue

            .GetText .ColLetterToNumber("H"), .Row, varSoThueDuocGiam
            .GetText .ColLetterToNumber("G"), .Row, varThuNhapGiamThue
            If (Val(varSoThueDuocGiam) > 0) And Val(varSoThueDuocGiam) >= Val(varThuNhapGiamThue) Then
                flgErr = False
                ' Danh dau cot so thue thu nhap duoc giam nay sai sau nay bat loi
                .Col = .ColLetterToNumber("U")
                .Text = "1"
            Else
                ' Danh dau cac cot con lai la dung
                .Col = .ColLetterToNumber("U")
                .Text = "0"
            End If

            
            ' Cot Thue da khau tru

            .GetText .ColLetterToNumber("I"), .Row, varSoThueKhauTru
            .GetText .ColLetterToNumber("F"), .Row, varThuNhap
            If Val(varSoThueKhauTru) > 0 And Val(varSoThueKhauTru) >= Val(varThuNhap) Then
                flgErr = False
                ' Danh dau cot so thue khau tru nay sai sau nay bat loi
                .Col = .ColLetterToNumber("I")
                .Text = "1"
            Else
                ' Danh dau cac cot con lai la dung
                .Col = .ColLetterToNumber("I")
                .Text = "0"
            End If

            
            ' Neu co bat cu mot chi tieu nao do bi sai thi danh dau vao de con bat loi
            If flgErr = False Then
                ' Danh dau cot check loi nay sai sau nay bat loi
                .Col = .ColLetterToNumber("W")
                .Text = "1"
            Else
                .Col = .ColLetterToNumber("W")
                .Text = "0"
            End If
            
            ' Lay so thue da khau tru cua lao dong do

            .GetText .ColLetterToNumber("I"), .Row, varSoThueKhauTru
            If Val(varSoThueKhauTru) > 0 Then
                ' Va tinh la lao dong co thue khau tru
                .Col = .ColLetterToNumber("K")
                .Text = "1"
                ' Tinh tong thu nhap cua rieng lao dong co thue khau tru nay
                .GetText .ColLetterToNumber("F"), .Row, varThuNhap
                .Col = .ColLetterToNumber("L")
                .Text = varThuNhap
            Else
                ' Va tinh la lao dong ko co thue khau tru
                .Col = .ColLetterToNumber("K")
                .Text = "0"
                ' Khong tinh tong thu nhap cua lao dong nay
                .Col = .ColLetterToNumber("L")
                .Text = "0"
            End If

            ' Trong truong hop cot Ho ten, MST, CMT, ThuNhap khong co gia tri thi ko kiem tra
            If (Trim(varHoVaTen) = vbNullString And (Trim(varMST) = vbNullString Or Trim(varCMT) = vbNullString) And Val(varThuNhap) = 0) Then
                ' Va danh dau cac cot sau de sau nay ko bat loi
                ' Ho va ten
                .Col = .ColLetterToNumber("M")
                .Text = "0"
                ' MST
                .Col = .ColLetterToNumber("N")
                .Text = "0"
                .Col = .ColLetterToNumber("O")
                .Text = "0"
                .Col = .ColLetterToNumber("P")
                .Text = "0"
                ' CMT
                .Col = .ColLetterToNumber("Q")
                .Text = "0"
                .Col = .ColLetterToNumber("R")
                .Text = "0"
                ' MST & CMT
                .Col = .ColLetterToNumber("S")
                .Text = "0"
                ' Cot check loi
                .Col = .ColLetterToNumber("W")
                .Text = "0"
            End If
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
Private Sub resetFormula()
    Dim i As Integer
    With fps
        .Sheet = 1
        .Col = .ColLetterToNumber("I")
        For i = 24 To 29
            .Row = i
            .Formula = ""
        Next
        For i = 34 To 38
            .Row = i
            .Formula = ""
        Next
    End With
End Sub

Private Sub defaultFormula()
    Dim i As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        .Col = .ColLetterToNumber("I")
        .Row = 24
        .Formula = "'05AKK-TNCN'!N16"
        .Row = 25
        .Formula = "'05AKK-TNCN'!G16"
        .Row = 26
        .Formula = "'05AKK-TNCN'!O16"
        .Row = 27
        .Formula = "'05AKK-TNCN'!P16"
        .Row = 28
        .Formula = "'05AKK-TNCN'!Q16"
        .Row = 29
        .Formula = "'05AKK-TNCN'!M16"
        
        .Row = 34
        .Formula = "'05BKK-TNCN'!J16"
        .Row = 35
        .Formula = "'05BKK-TNCN'!K16"
        .Row = 36
        .Formula = "'05BKK-TNCN'!F16"
        .Row = 37
        .Formula = "'05BKK-TNCN'!L16"
        .Row = 38
        .Formula = "'05BKK-TNCN'!I16"
        .EventEnabled(EventAllEvents) = True
    End With
End Sub
