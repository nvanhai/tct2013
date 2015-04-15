VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01BCTL_DK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Thu Nhap Doanh Nghiep" interface sheets
'this Class is belong to TAX_Business project which will be compline to DLL
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1


Dim vt As String, va As String, vb As String, vc As String
Dim strComboHien As String, strCombo As String
Dim arrTaiNguyen As New Scripting.Dictionary, mblKhoiDau As Boolean
Dim arrTaiNguyenGoc As New Scripting.Dictionary
Dim blKiemTraDieuChinhCu As Boolean
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
Dim strDVTCu As Variant, strMucThueCu As Variant
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
Private Const Row_MSTDN = 18
Private Const Col_MSTDN = "I"

' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

Public isImportXML As Boolean

'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    With fps
        SetDateFormat fps, 1, 68, .ColLetterToNumber("R"), DDMMYYYY
        .Sheet = 1
        .Row = 68
        .Col = .ColLetterToNumber("R")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft

    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
        Dim checkUq, checkDLT As Variant
    With fps
        mblKhoiDauToKhai = True
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        'update MSTDN luôn bang MST tren phan TT chung
        UpdateMST fps, Col_MSTDN, Row_MSTDN
        
        If Strloaitk = "TKCT" Or Strloaitk = "" Then
            .Col = .ColLetterToNumber("C")
            .Row = 36
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 36
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("C")
            .Row = 36
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 36
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        
         If isCheckTTDLT = True Then
            .Row = 17
            .Col = .ColLetterToNumber("I")
             checkUq = .value
             
             .Row = 17
            .Col = .ColLetterToNumber("O")
             checkDLT = .value
             
             If checkUq = 0 And checkDLT = 0 Then
                .Row = 17
                .Col = .ColLetterToNumber("O")
                 .value = 1
                 UpdateCell fps, .Col, .Row, "x"
             End If
             
        End If
        'end
    End With
End Sub




        

Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
     With fps
     ' xu ly check 1 trong 2 o uy quyen hoac DLT
        If Col = .ColLetterToNumber("I") And Row = 17 And .ActiveSheet = 1 Then
            fps.EventEnabled(EventAllEvents) = False
            .Sheet = 1
            .Row = 17
            .Col = .ColLetterToNumber("I")
            If (.value = "1") Then
                .Row = 17
                .Col = .ColLetterToNumber("O")
                .value = 0
                UpdateCell fps, .Col, .Row, ""
            Else
                .Row = 17
                .Col = .ColLetterToNumber("O")
                .value = 1
                UpdateCell fps, .Col, .Row, "x"
            End If
            fps.EventEnabled(EventAllEvents) = True
        End If
        
        If Col = .ColLetterToNumber("O") And Row = 17 And .ActiveSheet = 1 Then
            fps.EventEnabled(EventAllEvents) = False
            .Sheet = 1
            .Row = 17
            .Col = .ColLetterToNumber("O")
            If (.value = "1") Then
                .Row = 17
                .Col = .ColLetterToNumber("I")
                .value = 0
                UpdateCell fps, .Col, .Row, ""
            Else
                .Row = 17
                .Col = .ColLetterToNumber("I")
                .value = 1
                UpdateCell fps, .Col, .Row, "x"
            End If
            fps.EventEnabled(EventAllEvents) = True
        End If
     End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    Dim CellNguoiKy As String
    Dim tongNTNN As Variant, tienDoiThanhChu As String
    
    CellNguoiKy = GetAttribute(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length - 1), "CellID")
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("R") And Row = 68 Then
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
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    Dim iCurrentSheet As Integer
    Dim flagDVUyQuyen As Integer
    Dim flagDVUyQuyenEmpty As Integer
    Dim vDVUyQuyen, vDaiLyThue As Variant
    
    Dim vThongTinDL As Variant
    Dim flagThongTinDL As Integer
    Dim vErrorEmpty As Variant
    Dim vErrorDVUQ As Variant
    
    With fps
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 23, vErrorEmpty
        .GetText .ColLetterToNumber("E"), 22, vErrorDVUQ
        
        .GetText .ColLetterToNumber("E"), 25, vThongTinDL
        
        ' kiem tra to khai
        .Sheet = 1
        .GetText .ColLetterToNumber("I"), 17, vDVUyQuyen
        .GetText .ColLetterToNumber("O"), 17, vDaiLyThue
        ' kiem tra chi duoc chon  uy quyen hoac dai ly thue
        If isCheckTTDLT = True Then
            If Trim$(vDVUyQuyen) = "1" And Trim$(vDaiLyThue) = "1" Then
                .Col = .ColLetterToNumber("I")
                .Row = 17
                .CellNote = vErrorDVUQ
                .BackColor = mErrorColor
                
                .Col = .ColLetterToNumber("O")
                .Row = 17
                .CellNote = vErrorDVUQ
                .BackColor = mErrorColor
                
                flagDVUyQuyen = 1
            ElseIf Trim$(vDVUyQuyen) = "0" And Trim$(vDaiLyThue) = "0" Then
                .Col = .ColLetterToNumber("I")
                .Row = 17
                .CellNote = vErrorEmpty
                .BackColor = mErrorColor
                
                .Col = .ColLetterToNumber("O")
                .Row = 17
                .CellNote = vErrorEmpty
                .BackColor = mErrorColor
                flagDVUyQuyenEmpty = 1
            Else
                .Col = .ColLetterToNumber("I")
                .Row = 17
                .CellNote = ""
                .BackColor = mNonErrorColor
                
                .Col = .ColLetterToNumber("O")
                .Row = 17
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
        Else
                ' neu khong co don vi uy quyen hoac dai ly thue thi khong duoc check chon
                If Trim$(vDVUyQuyen) = "0" And Trim$(vDaiLyThue) = "0" Then
                    .Col = .ColLetterToNumber("I")
                    .Row = 17
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                    
                    .Col = .ColLetterToNumber("O")
                    .Row = 17
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                Else
                    .Col = .ColLetterToNumber("I")
                    .Row = 17
                    .CellNote = vThongTinDL
                    .BackColor = mErrorColor
                    
                    .Col = .ColLetterToNumber("O")
                    .Row = 17
                    .CellNote = vThongTinDL
                    .BackColor = mErrorColor
                    flagThongTinDL = 1
                End If
        End If
         .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 22, IIf(flagDVUyQuyen = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 23, IIf(flagDVUyQuyenEmpty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 25, IIf(flagThongTinDL = 1, "0", "1")
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
    With fps
        .EventEnabled(EventAllEvents) = False
        If isImportXML = True Then
            .Sheet = 1
            .Row = 17
            .Col = .ColLetterToNumber("I")
            .value = "0"
    
            .Col = .ColLetterToNumber("O")
            .value = "0"
        End If
        .EventEnabled(EventAllEvents) = True
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
    ' Set tip to display and set tip's content
      If Col = fps.ColLetterToNumber("G") Then
        ShowTip = True
        TipText = TipText
      End If
End Sub
Public Function getValueTK(strFileLoadData As String) As String
    Dim result As String
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01TBVMT", "getValueTK", Err.number, Err.Description
End Function

Public Function getDieuChinhGiam(strOldValue As String) As String
    getDieuChinhGiam = ""
End Function
Function getMapCT() As String
    getMapCT = ""
    Exit Function
ErrHandle:
    SaveErrorLog "01_BCTL_DK", "getMapCT", Err.number, Err.Description
End Function

' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        getChiTieuTongDC = ""
    End With
End Function
' Lay ve chi tieu 32 o to khai lan truoc
Public Function getValueCTDC(strFileLoadData As String) As String
    getValueCTDC = ""
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01BCTL_DK", "getValueCTDC", Err.number, Err.Description
End Function
'******************************
'Set trang thai Flag KHBS active
Public Sub SetActiveKHBS() 'ByRef xmlNodeValid As MSXML.IXMLDOMNode)

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
         For Each xmlCellNodeData In TAX_Utilities_v2.Data(.Sheet - 1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               If (lRow = 17 And lCol = .ColLetterToNumber("I")) Or (lRow = 17 And lCol = .ColLetterToNumber("O")) Then
                    .GetText lCol, lRow, varTemp

                    If varTemp = "1" Or UCase(CStr(varTemp)) = "X" Then
                        UpdateCell fps, lCol, lRow, "x"
                    Else
                        UpdateCell fps, lCol, lRow, ""
                    End If
               Else
                    .Col = lCol
                    .Row = lRow
                    If .CellType = CellTypeNumber Then
                         varTemp = .value
                    Else
                         varTemp = .Text
                    End If
                    UpdateCell fps, lCol, lRow, varTemp
               End If
         Next
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
    tuRowDL = 16
    denRowDL = 21
    'TuRow_DenRow_sheet
End Sub


' Lay ve ky tu ngan cach so thap phan va so hang nghin
Private Sub test()
    Dim LocaleDecimal As String
    Dim LocaleComma As String
    LocaleDecimal = Mid$(CStr(11 / 10), 2, 1)
    LocaleComma = Chr$(90 - Asc(LocaleDecimal))
End Sub
