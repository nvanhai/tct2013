VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01TTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Dim mblKeyF6 As Boolean
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

Public Sub Prepare1()
    With fps
        SetDateFormat fps, 1, 61, .ColLetterToNumber("O"), DDMMYYYY
        SetDateFormat fps, 1, 13, .ColLetterToNumber("R"), DDMMYYYY
        .Sheet = 1
        .Row = 61
        .Col = .ColLetterToNumber("O")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
      
      End With
    
End Sub

Public Sub Prepare2()
    Dim a2 As Variant, b2 As Variant, c2 As Variant
    Dim count1 As Long
    With fps
            .EventEnabled(EventAllEvents) = False
            .Sheet = 1
            mCurrentSheet = .Sheet
        
            If Strloaitk = "TKCT" Then
                .Col = .ColLetterToNumber("C")
                .Row = 29
                .Text = 1
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("I")
                .Row = 29
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                
            ElseIf Strloaitk = "TKBS" Then
                        
                .Col = .ColLetterToNumber("C")
                .Row = 29
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("I")
                .Row = 29
                .Text = StrSolanBosung
                UpdateCell fps, .Col, .Row, .Text
            End If
            
'            .Row = 54 + count1
'            .Row = 61
'            .Col = .ColLetterToNumber("O")
'            If .value = "" Then
'                SetDateFormat fps, 1, 61, .ColLetterToNumber("O"), DDMMYYYY
'                .Text = Format(Date, "dd/mm/yyyy")
'                .TypeHAlign = TypeHAlignLeft
'                UpdateCell fps, .Col, .Row, .Text
'            End If
          'set thong tin header neu la to khai quy
          If strQuy = "TK_QUY" Then
            .SetText .ColLetterToNumber("G"), 29, 1
            .SetText .ColLetterToNumber("O"), 29, "0"
            .SetText .ColLetterToNumber("R"), 29, "0"
          Else
            .SetText .ColLetterToNumber("G"), 29, 0
            'lay tu thang den thang
            .Col = .ColLetterToNumber("U")
            .Row = 29
            'lay so thang quyet toan
            .Text = DateDiff("M", Format("01/" & TAX_Utilities_v1.FirstDay, "dd/mm/yyyy"), Format("01/" & TAX_Utilities_v1.LastDay, "dd/mm/yyyy")) + 1
          'set thong tin header neu la to khai tuthang denthang
             If strLoaiTKQT = "QT_TUNG_NAM" Then
               .Col = .ColLetterToNumber("O")
               .Row = 29
               .Text = 1
               .Col = .ColLetterToNumber("R")
               .Text = "0"
               .SetText .ColLetterToNumber("P"), 3, TAX_Utilities_v1.FirstDay
               .SetText .ColLetterToNumber("T"), 3, TAX_Utilities_v1.LastDay
            ElseIf strLoaiTKQT = "QT_NAM_DAU" Then
               .Col = .ColLetterToNumber("R")
               .Row = 29
               .Text = 1
               .Col = .ColLetterToNumber("O")
               .Text = "0"
               .SetText .ColLetterToNumber("P"), 3, TAX_Utilities_v1.FirstDay
               .SetText .ColLetterToNumber("T"), 3, TAX_Utilities_v1.LastDay
             End If
          End If
        .EventEnabled(EventAllEvents) = True
    End With
    If Strloaitk <> "TKBS" Then
        
    End If
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
'Dim i, j As Integer, iCol As Long, iRow As Long
'Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
'Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
'Dim tCol As Long, tRow As Long
'Dim strId As Variant, Col7 As Variant, intIndexCombo As Integer
'Dim NguoiKyCol As Long, NguoiKyRow As Long
'Dim strtext As Variant
'Dim lValue As Variant, lrow43 As Variant, lrow44 As Variant
'Dim strSumValue As Long
'Dim CountCaNhanKD As Integer
'Dim CellNguoiKy As String
'    With fps
'        iCol = .ActiveCol
'        iRow = .ActiveRow
'        GetCellSpan fps, iCol, iRow
'        CountCaNhanKD = 0
'            If .ActiveSheet = 1 Then
'                   fps.EventEnabled(EventAllEvents) = False
'                    i = 1
'                    .Col = .ColLetterToNumber("B")
'                    .Row = 51
'                    Do
'                         .Text = str(i)
'                         UpdateCell fps, .Col, .Row, .Text
'                         .GetText .ColLetterToNumber("B"), .Row, strtext
'                         If strtext <> "" Then
'                                CountCaNhanKD = CountCaNhanKD + 1
'                         End If
'                          If i = 1 Then
'                                .Col = .ColLetterToNumber("L")
'                                .Row = 50
'                                UpdateCell fps, .Col, .Row, .value
'                                lValue = .value
'                                .Col = .ColLetterToNumber("R")
'                                .Row = 43
'                                UpdateCell fps, .Col, .Row, .value
'                                lrow43 = .value
'                                .Col = .ColLetterToNumber("R")
'                                .Row = 44
'                                 UpdateCell fps, .Col, .Row, .value
'                                lrow44 = .value
'                                .Col = .ColLetterToNumber("N")
'                                .Row = 50
'                                .value = lrow43 * lValue / 100
'                                 UpdateCell fps, .Col, .Row, .value
'                                .Col = .ColLetterToNumber("T")
'                                .Row = 50
'                                UpdateCell fps, .Col, .Row, .value
'                                .Col = .ColLetterToNumber("R")
'                                .Row = 50
'                                UpdateCell fps, .Col, .Row, .value
'                                .Col = .ColLetterToNumber("X")
'                                .Row = 50
'                                UpdateCell fps, .Col, .Row, .value
'                                .Col = .ColLetterToNumber("Z")
'                                .Row = 50
'                                UpdateCell fps, .Col, .Row, .value
'
'                           Else
'
'                                .Col = .ColLetterToNumber("L")
'                                .Row = 50 + i - 1
'                                UpdateCell fps, .Col, .Row, .value
'                                lValue = .value
'                                .Col = .ColLetterToNumber("R")
'                                .Row = 43
'                                UpdateCell fps, .Col, .Row, .value
'                                lrow43 = .value
'                                .Col = .ColLetterToNumber("R")
'                                .Row = 44
'                                 UpdateCell fps, .Col, .Row, .value
'                                lrow44 = .value
'                                .Col = .ColLetterToNumber("N")
'                                .Row = 50 + i - 1
'                                .value = lrow43 * lValue / 100
'                                 UpdateCell fps, .Col, .Row, .value
'                                .Col = .ColLetterToNumber("T")
'                                .Row = 50 + i - 1
'                                UpdateCell fps, .Col, .Row, .value
'                                .Col = .ColLetterToNumber("R")
'                                .Row = 50 + i - 1
'                                UpdateCell fps, .Col, .Row, .value
'                                .Col = .ColLetterToNumber("X")
'                                .Row = 50 + i - 1
'                                UpdateCell fps, .Col, .Row, .value
'                                .Col = .ColLetterToNumber("Z")
'                                .Row = 50 + i - 1
'                                UpdateCell fps, .Col, .Row, .value
'
'                           End If
'                          .Col = .ColLetterToNumber("B")
'                         .Row = i + 50
'                         i = i + 1
'                    Loop Until .Text = "aa"
'                         .Col = .ColLetterToNumber("N")
'                           .Row = 50 + i
'                           UpdateCell fps, .Col, .Row, .value
'                           .Col = .ColLetterToNumber("P")
'                           .Row = 50 + i
'                           UpdateCell fps, .Col, .Row, .value
'                           .Col = .ColLetterToNumber("R")
'                           .Row = 50 + i
'                           UpdateCell fps, .Col, .Row, .value
'                           .Col = .ColLetterToNumber("T")
'                           .Row = 50 + i
'                           UpdateCell fps, .Col, .Row, .value
'                           .Col = .ColLetterToNumber("V")
'                           .Row = 50 + i
'                           UpdateCell fps, .Col, .Row, .value
'                           .Col = .ColLetterToNumber("X")
'                           .Row = 50 + i
'                           UpdateCell fps, .Col, .Row, .value
'                           .Col = .ColLetterToNumber("Z")
'                           .Row = 50 + i
'                           UpdateCell fps, .Col, .Row, .value
'          End If
'
'                 .SetActiveCell iCol, iRow
'                    '************************
'                    .Row = iRow
'                    .Col = iCol
'                    fps.EventEnabled(EventAllEvents) = True
'                    .SetActiveCell iCol, iRow
'
'        .EventEnabled(EventAllEvents) = True
'    End With
'End Sub
'
'Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim i, j As Integer, iCol As Long, iRow As Long
'Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
'Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
'Dim tCol As Long, tRow As Long
'Dim strId As Variant, Col7 As Variant, intIndexCombo As Integer
'Dim NguoiKyCol As Long, NguoiKyRow As Long
'Dim strtext As Variant
'Dim lValue As Variant
'Dim strSumValue As Long
'Dim CountCaNhanKD As Integer
'Dim CellNguoiKy As String
'    With fps
'        iCol = .ActiveCol
'        iRow = .ActiveRow
'        GetCellSpan fps, iCol, iRow
'        CountCaNhanKD = 0
'      If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
'            If .ActiveSheet = 1 Then
'                   fps.EventEnabled(EventAllEvents) = False
'                    i = 1
'                    .Col = .ColLetterToNumber("B")
'                    .Row = 51
'                    Do
'                         .Text = str(i)
'                         UpdateCell fps, .Col, .Row, .Text
'                         .Row = i + 50
'                         '.Col = .ColLetterToNumber("B")
'                         i = i + 1
'                    Loop Until .Text = "aa"
'
'             End If
''
'                 .SetActiveCell iCol, iRow
'                    '************************
'                    .Row = iRow
'                    .Col = iCol
'                    fps.EventEnabled(EventAllEvents) = True
'                    .SetActiveCell iCol, iRow
'    End If
'
'        .EventEnabled(EventAllEvents) = True
'    End With
'
'End Sub
'  Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Dim varTemp As Variant
'    Dim varCT11 As Variant
'    Dim varCT12 As Variant
'    Dim count1 As Long
'    Dim CellNguoiKy As String
'    Dim strValue As String
'    Dim varThangKT As Variant
'
'    CellNguoiKy = GetAttribute(TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")(TAX_Utilities_v1.Data(0).getElementsByTagName("Cell").length - 4), "CellID")
'    With fps
'        .EventEnabled(EventAllEvents) = False
'        mCurrentSheet = .Sheet
'        If .ActiveSheet = 1 Then
'            ' check ngay ky
'            If Col = .ColLetterToNumber("R") And Row = Right(CellNguoiKy, 2) Then
'                .GetText Col, Row, varTemp
'                If varTemp <> "" And varTemp <> "../../...." Then
'                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
'                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
'                    Else
'                        .SetFocus
'                        .SetActiveCell Col, Row
'                    End If
'                Else
'                    .SetText Col, Row, ""
'                End If
'                .Col = Col
'                .Row = Row
'                UpdateCell fps, Col, Row, .Text
'            End If
'
'            CheckDynamicError
'        End If
'        .EventEnabled(EventAllEvents) = True
'    End With
'    'CellChange Col, Row
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
    Dim colC, colH, colU, TiLeTNCT As Variant
    Dim vErrMsgMST As Variant
   
    Dim iFlagvName As Integer, iFlagvDate As Integer, iFlagvMst As Integer, iFlagvThang As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer
    Dim mErrTongTNCN As Variant, iFlagTongTNCN As Integer, vTongTNCN As Variant
    Dim mErrTongCT36 As Variant, iflagTongCT36 As Integer, vTongCT36   As Variant
    Dim merrTongCT37 As Variant, iflagTongCT37 As Integer, vTongCT37 As Variant
    Dim vCT38 As Variant, vCT40 As Variant
    Dim mErrEmty As Variant, iFlagEmty As Integer, iFlagCheckData As Boolean
    Dim vCT27, vCT28, vNgayKy, iFlagNgayKy As Variant
    Dim mErrMSTTonTai, vMSTCT5 As Variant, iFlagMSTTonTai, rowMST As Integer
    Dim countRow As Long, rowDes As Long, iFlagMSTTrung As Integer, mErrMSTTrung As Variant
        
    Dim vNgayVB As Variant
    Dim vErrorNgayVB As Variant
    Dim iFlagNgayVB As Integer
        
    Dim strFormula As String
    Const colorCell = 16709097
    iFlagvName = 0
    iFlagvDate = 0
    iFlagvMst = 0
    iFlagvThang = 0
    iFocusFlag = 0
    iFlagMSTTonTai = 0
    
    Dim iWhiteRow As Integer
    With fps

        CheckErrorMST
        .Sheet = .SheetCount
        
        .GetText .ColLetterToNumber("E"), 24, mErrEmty
        .GetText .ColLetterToNumber("E"), 30, vErrMsgMST
        .GetText .ColLetterToNumber("E"), 33, mErrTongTNCN
        .GetText .ColLetterToNumber("E"), 34, mErrTongCT36
        .GetText .ColLetterToNumber("E"), 35, merrTongCT37
        .GetText .ColLetterToNumber("E"), 36, vErrorNgayVB
        

        .Sheet = 1
     
        vTongTNCN = 0
        vTongCT36 = 0
        vTongCT37 = 0
        
        'check ngay nhap khong dc de trong
        .GetText .ColLetterToNumber("O"), .MaxRows - 6, vNgayKy
        If vNgayKy = vbNullString Or vNgayKy = "../../...." Then
            .Col = .ColLetterToNumber("O")
            .Row = .MaxRows - 6
            .BackColor = mErrorColor
            .CellNote = error1
            iFlagNgayKy = 1
        Else
            .Col = .ColLetterToNumber("O")
            .Row = .MaxRows - 6
            .BackColor = mNonErrorColor
            .CellNote = ""
        End If
        
        ' kiem tra ngay van ban
        .GetText .ColLetterToNumber("R"), 13, vNgayVB
        If Trim$(vNgayVB) <> "" Then
            If DateDiff("D", Date, Format(vNgayVB, "dd/mm/yyyy")) > 0 Then
                 .Col = .ColLetterToNumber("R")
                 .Row = 13
                 .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                 .CellNote = vErrorNgayVB   'static
                 .BackColor = mErrorColor
                 iFlagNgayVB = 1
                 'capture the first error cell
                 If iFocusFlag = 0 Then
                     strFocusSheetName = "'" & .SheetName & "'"
                     strFocusCol = str(.Col)
                     strFocusRow = str(.Row)
                     iFocusFlag = 1
                 End If
             Else
                  .Col = .ColLetterToNumber("R")
                  .Row = 13
                  .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                  .CellNote = ""   'static
                  .BackColor = mNonErrorColor
            End If
        Else
            .Col = .ColLetterToNumber("R")
            .Row = 13
            .BackColor = mNonErrorColor
            .CellNote = ""
        End If
        
        .Row = 40
        Do
            .GetText .ColLetterToNumber("C"), .Row, colC
            .GetText .ColLetterToNumber("E"), .Row, vMST
            .GetText .ColLetterToNumber("I"), .Row, TiLeTNCT
            .GetText .ColLetterToNumber("P"), .Row, vCT38
            .GetText .ColLetterToNumber("T"), .Row, vCT40
            
            If colC <> vbNullString Or vMST <> vbNullString Or TiLeTNCT <> 0 Or vCT38 <> 0 Or vCT40 <> 0 Then
                  
                   If Trim(vMST) <> vbNullString Then
    '                vMST = Replace(vMST, " ", "")
    '                vMST = Replace(vMST, "-", "")
                     vMST = Left(vMST, 10)
                     strCheck = ""
                     If Len(vMST) = 10 Or Len(vMST) = 13 Then
                         strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                     End If
                     If Mid(strCheck, 2, 1) <> "0" Then
                        .Col = .ColLetterToNumber("E")
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
                     Else
                        .Col = .ColLetterToNumber("E")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                   Else
                    .Col = .ColLetterToNumber("E")
                    .CellNote = mErrEmty
                    .BackColor = mErrorColor
                    iFlagEmty = 1
                End If
                'check du lieu k dc de trong
                If colC = vbNullString Then
                    .Col = .ColLetterToNumber("C")
                    .BackColor = mErrorColor
                    .CellNote = mErrEmty
                    iFlagEmty = 1
                Else
                    .Col = .ColLetterToNumber("C")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If
                
                If vTongTNCN = vbNullString Then
                    .Col = .ColLetterToNumber("I")
                    .BackColor = mErrorColor
                    .CellNote = mErrEmty
                    iFlagEmty = 1
                Else
                    .Col = .ColLetterToNumber("I")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                End If
                
'                If vCT38 = vbNullString Then
'                    .Col = .ColLetterToNumber("P")
'                    .BackColor = mErrorColor
'                    .CellNote = mErrEmty
'                    iFlagEmty = 1
'                Else
'                    .Col = .ColLetterToNumber("P")
'                    .BackColor = mNonErrorColor
'                    .CellNote = ""
'                End If
'
'                If vCT40 = vbNullString Then
'                    .Col = .ColLetterToNumber("T")
'                    .BackColor = mErrorColor
'                    .CellNote = mErrEmty
'                    iFlagEmty = 1
'                Else
'                    .Col = .ColLetterToNumber("T")
'                    .BackColor = mNonErrorColor
'                    .CellNote = ""
'                End If
                
                'danh dau neu dong co chua du lieu
                iFlagCheckData = True
        End If
        
             count1 = count1 + 1
            .Row = count1 + 40
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
    
        'neu dong co chua du lieu thi kiem tra cac rang buoc trong cac chi tieu
        If iFlagCheckData Then
            'lay CT 27
            .GetText .ColLetterToNumber("I"), 23, vCT27
            'lay CT28
            .GetText .ColLetterToNumber("I"), 24, vCT28
            'lay tong ty le chia TNCT
            .GetText .ColLetterToNumber("I"), .MaxRows - 25, vTongTNCN
            'lay tong CT 36
            .GetText .ColLetterToNumber("N"), .MaxRows - 25, vTongCT36
            'lay tong CT37
            .GetText .ColLetterToNumber("O"), .MaxRows - 25, vTongCT37
            'check tong TNCT phai = 100
            If Round(CDbl(vTongTNCN), 2) <> 100 Then
                .Col = .ColLetterToNumber("I")
                .Row = .MaxRows - 25
                .BackColor = mErrorColor
                .CellNote = mErrTongTNCN
                iFlagTongTNCN = 1
            Else
                .Col = .ColLetterToNumber("I")
                .Row = .MaxRows - 25
                .BackColor = colorCell
                .CellNote = ""
            End If
            'check Tong CT 36 phai = CT27
            If vTongCT36 <> vCT27 Then
                .Col = .ColLetterToNumber("N")
                .Row = .MaxRows - 25
                .BackColor = mErrorColor
                .CellNote = mErrTongCT36
                iflagTongCT36 = 1
            Else
                .Col = .ColLetterToNumber("N")
                .Row = .MaxRows - 25
                .BackColor = colorCell
                .CellNote = ""
            End If
            'check Tong CT37 = CT28
            If vTongCT37 <> vCT28 Then
                .Col = .ColLetterToNumber("O")
                .Row = .MaxRows - 25
                .BackColor = mErrorColor
                .CellNote = merrTongCT37
                iflagTongCT37 = 1
            Else
                .Col = .ColLetterToNumber("O")
                .Row = .MaxRows - 25
                .BackColor = colorCell
                .CellNote = ""
            End If
        End If
        
       
    
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 33, IIf(iFlagTongTNCN = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 34, IIf(iflagTongCT36 = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 35, IIf(iflagTongCT37 = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 30, IIf(iFlagvMst = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 24, IIf(iFlagEmty = 1, "0", "1")
    
        .Sheet = iCurrentSheet
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
Public Sub SetData()
    With fps
        'set lai STT cho CT 25
        .SetText .ColLetterToNumber("B"), 51, 1
    End With
End Sub

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function
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
                        If Not TAX_Utilities_v1.Data(mCurrentSheet - 1).nodeFromID( _
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
                        If Not TAX_Utilities_v1.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, i, pRow - lRowCtrl)) Is Nothing Then
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

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim iCol As Long, iRow As Long
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tCol As Long, tRow As Long
Dim i As Integer
    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 1 Then
                    fps.EventEnabled(EventAllEvents) = False
                    
                    .Col = .ColLetterToNumber("B")
                    i = 1
                    .Row = 40
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = .Row + 1
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
        With fps
    'check date
        If Col = .ColLetterToNumber("O") And Row = .MaxRows - 2 Then
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
                UpdateCell fps, Col, Row, .Text
            End If
        
            If Col = .ColLetterToNumber("R") And Row = 13 Then
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
                UpdateCell fps, Col, Row, .Text
            End If
        
        End With
End Sub
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
        For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
           ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'               If (lRow = 19 And lCol = .ColLetterToNumber("R")) Or (lRow = 19 And lCol = .ColLetterToNumber("O")) Then
'                    .GetText lCol, lRow, varTemp
'                    If varTemp = "1" Or UCase(CStr(varTemp)) = "X" Then
'                        UpdateCell fps, lCol, lRow, "x"
'                    Else
'                        UpdateCell fps, lCol, lRow, ""
'                    End If
'               Else
            .GetText lCol, lRow, varTemp
            UpdateCell fps, lCol, lRow, varTemp
'               End If
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


