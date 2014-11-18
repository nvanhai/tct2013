VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_02TAIN_DK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrSolanBosung As String

Public FlagThaydoi As String

Const SHEET_PHU_LUC_01 As Integer = 3
Const SHEET_PHU_LUC_SAN_LUONG As Integer = 2
Const sheetPl01_KHBS_GTGT As Integer = 4
Private Const ROW_START_02_1 = 29
Private Const Row_MSTDN = 51
Private Const Col_MSTDN = "P"
Private Const Col_TKCT = "R"
Private Const Row_TKCT = 30
Private Const Col_TKBS = "W"
Private Const Row_TKBS = 30



' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String
Public strLoaiToKhaiDK As String
Public isImportXML As Boolean

Private arrStrXMLFileNames() As String
 
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    Dim ct29KyTruoc As String
    Dim arrKytruoc() As String

    With fps
        .EventEnabled(EventAllEvents) = False
        SetDateFormat fps, 1, 94, fps.ColLetterToNumber("AR"), DDMMYYYY
        SetDateFormat fps, SHEET_PHU_LUC_SAN_LUONG, 29, .ColLetterToNumber("K"), DDMMYYYY
        SetDateFormat fps, SHEET_PHU_LUC_SAN_LUONG, 29, .ColLetterToNumber("U"), DDMMYYYY
        SetDateFormat fps, sheetPl01_KHBS_GTGT, 46, .ColLetterToNumber("BG"), DDMMYYYY
        SetDateFormat fps, sheetPl01_KHBS_GTGT, 31, .ColLetterToNumber("BE"), DDMMYYYY
        .Sheet = 1
        
        .Row = 94
        .Col = fps.ColLetterToNumber("AR")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        If isCheckTTDLT = True Then
            .Row = 49
            .Col = .ColLetterToNumber("P")
            .Lock = False
            
            .Row = 49
            .Col = .ColLetterToNumber("AA")
            .Lock = False
            .value = 1
            
        Else
            .Row = 49
            .Col = .ColLetterToNumber("P")
            .Lock = True
            
            .Row = 49
            .Col = .ColLetterToNumber("AA")
            .Lock = True
        End If
        
        
        ct29KyTruoc = getCT29
        arrKytruoc = Split(ct29KyTruoc, "~")
        ' chi tieu 29
        .Row = 57
        .Col = .ColLetterToNumber("AV")
        .value = arrKytruoc(0)
        ' chi tieu 33
        .Row = 67
        .Col = .ColLetterToNumber("AV")
        .value = arrKytruoc(1)
        
        
        .Sheet = 2
        .Row = 29
        .Col = .ColLetterToNumber("D")
        If strLoaiToKhaiDK = "DT" Then
            .Text = GetAttribute(GetMessageCellById("0324"), "Msg")
        Else
            .Text = GetAttribute(GetMessageCellById("0325"), "Msg")
        End If
        
        '--end set for sheet Phu luc 02
        .EventEnabled(EventAllEvents) = True
    End With
    
    setCQTQuanLyHoanThue fps
    
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
Dim strColC As String
    Dim i As Long
    Dim intCtrl As Integer
    Dim strId As Variant
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim checkUq, checkDLT As Variant
    
    With fps
        .Sheet = 1
        mCurrentSheet = .Sheet
        .EventEnabled(EventAllEvents) = False
        'update MSTDN luôn bang MST tren phan TT chung
        UpdateMST fps, Col_MSTDN, Row_MSTDN
        .Sheet = 1
        .Row = 30
        .Col = .ColLetterToNumber("O")
        If strLoaiToKhaiDK = "DT" Then
            .Text = "1"
        Else
            .Text = "0"
        End If
        '.Text = chkDauTho
        
        .Row = 30
        .Col = .ColLetterToNumber("P")
        If strLoaiToKhaiDK = "CD" Then
            .Text = "1"
        Else
            .Text = "0"
        End If
       ' .Text = chkCondensate
        
        .Row = 30
        .Col = .ColLetterToNumber("Q")
        If strLoaiToKhaiDK = "KTN" Then
            .Text = "1"
        Else
            .Text = "0"
        End If
        '.Text = chkKhiThienNhien

        'If Strloaitk = "TKCT" Then
        If Strloaitk = "TKCT" Or Strloaitk = "" Then
            .Col = .ColLetterToNumber(Col_TKCT)
            .Row = Row_TKCT
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber(Col_TKBS)
            .Row = Row_TKBS
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber(Col_TKCT)
            .Row = Row_TKCT
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber(Col_TKBS)
            .Row = Row_TKBS
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
         If isCheckTTDLT = True Then
            .Row = 49
            .Col = .ColLetterToNumber("P")
             checkUq = .value
             
             .Row = 49
            .Col = .ColLetterToNumber("AA")
             checkDLT = .value
             
             If checkUq = 0 And checkDLT = 0 Then
                .Row = 49
                .Col = .ColLetterToNumber("AA")
                 .value = 1
                 UpdateCell fps, .Col, .Row, .value
             End If
             
        End If
        
        ' set tu thang den thang
        .Col = .ColLetterToNumber("AD")
        .Row = 30
        .Text = TAX_Utilities_v1.FirstDay
        UpdateCell fps, .Col, .Row, .Text
        
        .Col = .ColLetterToNumber("AE")
        .Row = 30
        .Text = TAX_Utilities_v1.LastDay
        UpdateCell fps, .Col, .Row, .Text
        
        mCurrentSheet = .Sheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
 With fps
       
        
        
        ' xu ly check 1 trong 2 o uy quyen hoac DLT
        If Col = .ColLetterToNumber("P") And Row = 49 And .ActiveSheet = 1 Then
            fps.EventEnabled(EventAllEvents) = False
            .Sheet = 1
            .Row = 49
            .Col = .ColLetterToNumber("P")
            If (.value = "1") Then
                .Row = 49
                .Col = .ColLetterToNumber("AA")
                .value = 0
                UpdateCell fps, .Col, .Row, ""
            Else
                .Row = 49
                .Col = .ColLetterToNumber("AA")
                .value = 1
                UpdateCell fps, .Col, .Row, "x"
            End If
            fps.EventEnabled(EventAllEvents) = True
        End If
        
        If Col = .ColLetterToNumber("AA") And Row = 49 And .ActiveSheet = 1 Then
            fps.EventEnabled(EventAllEvents) = False
            .Sheet = 1
            .Row = 49
            .Col = .ColLetterToNumber("AA")
            If (.value = "1") Then
                .Row = 49
                .Col = .ColLetterToNumber("P")
                .value = 0
                UpdateCell fps, .Col, .Row, ""
            Else
                .Row = 49
                .Col = .ColLetterToNumber("P")
                .value = 1
                UpdateCell fps, .Col, .Row, "x"
            End If
            fps.EventEnabled(EventAllEvents) = True
        End If
        
        
        
        
        
    End With
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    CellChange Col, Row
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, iCol As Long, iRow As Long

    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = SHEET_PHU_LUC_01 Or .ActiveSheet = SHEET_PHU_LUC_SAN_LUONG Then
                   fps.EventEnabled(EventAllEvents) = False
                    i = 0
                    .Col = .ColLetterToNumber("B")
                    If .ActiveSheet = SHEET_PHU_LUC_01 Then
                          .Row = 27
                    Else
                           .Row = 29
                    End If
                    Do
                         i = i + 1
                         .Text = str$(i)
                         .Col = .ColLetterToNumber("B")
                         If .ActiveSheet = SHEET_PHU_LUC_SAN_LUONG Then
                            .Col = .ColLetterToNumber("D")
                             If strLoaiToKhaiDK = "DT" Then
                                .Text = GetAttribute(GetMessageCellById("0324"), "Msg")
                             Else
                                .Text = GetAttribute(GetMessageCellById("0325"), "Msg")
                             End If
                             .Col = .ColLetterToNumber("B")
                         End If
                         .Row = .Row + 1
                    Loop Until .Text = "aa"
                    .SetActiveCell iCol, iRow

                    .Row = iRow
                    .Col = iCol
                     fps.EventEnabled(EventAllEvents) = True
                     .SetActiveCell iCol, iRow
       
            End If
         End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet

        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("AR") And Row = 94 Then
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
        
        'format date for row static phu luc san luong 02-1
        If .ActiveSheet = SHEET_PHU_LUC_SAN_LUONG And Row >= 29 Then
            If Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("U") Then
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
        
        

        .EventEnabled(EventAllEvents) = True
    End With
    
    CellChange Col, Row
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
    Dim strCheck          As String, iCurrentSheet As Integer, iWhiteRow As Integer
    Dim iFlagCheckMaST    As Integer, iFlagEmty As Integer
    Dim vMST As Variant, errMST As Variant, vTenNhaThau As Variant, errEmty As Variant
    Dim vTyLePhanBo As Variant, vSoThuePhatSinh As Variant, vGhiChu As Variant
    Dim vTongTyLe As Variant
    Dim vErrorTongTyLe As Variant
    Dim flagTongTyLe As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer
    Dim vErrorEmpty As Variant
    Dim vErrorDVUQ As Variant
    Dim flagDVUyQuyen As Integer
    Dim flagDVUyQuyenEmpty As Integer
    Dim vDVUyQuyen, vDaiLyThue As Variant
    
    Dim iFlagCheckData As Boolean
    Dim vCt09 As Variant
    Dim flagCt09 As Integer
    Dim vErrorCt09 As Variant
    
    Dim iFlagDateCompare As Integer
    Dim vDate1, vDate2 As Variant
    Dim arrDate1() As String, arrDate2() As String
    Dim errDateCompare As Variant
    With fps
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 13, errMST
        .GetText .ColLetterToNumber("E"), 19, errEmty
        .GetText .ColLetterToNumber("E"), 23, vErrorEmpty
        .GetText .ColLetterToNumber("E"), 25, vErrorDVUQ
        
        .GetText .ColLetterToNumber("E"), 28, vErrorTongTyLe
        .GetText .ColLetterToNumber("E"), 29, vErrorCt09

        iFlagCheckMaST = 0
        iFlagEmty = 0
        ' kiem tra to khai
        .Sheet = 1
        .GetText .ColLetterToNumber("P"), 49, vDVUyQuyen
        .GetText .ColLetterToNumber("Z"), 49, vDaiLyThue
        ' kiem tra chi duoc chon  uy quyen hoac dai ly thue
        If isCheckTTDLT = True Then
            If Trim$(vDVUyQuyen) = "1" And Trim$(vDaiLyThue) = "1" Then
                .Col = .ColLetterToNumber("P")
                .Row = 49
                .CellNote = vErrorDVUQ
                .BackColor = mErrorColor
                
                .Col = .ColLetterToNumber("AA")
                .Row = 49
                .CellNote = vErrorDVUQ
                .BackColor = mErrorColor
                
                flagDVUyQuyen = 1
            ElseIf Trim$(vDVUyQuyen) = "0" And Trim$(vDaiLyThue) = "0" Then
                .Col = .ColLetterToNumber("P")
                .Row = 49
                .CellNote = vErrorEmpty
                .BackColor = mErrorColor
                
                .Col = .ColLetterToNumber("AA")
                .Row = 49
                .CellNote = vErrorEmpty
                .BackColor = mErrorColor
                flagDVUyQuyenEmpty = 1
            Else
                .Col = .ColLetterToNumber("P")
                .Row = 49
                .CellNote = ""
                .BackColor = mNonErrorColor
                
                .Col = .ColLetterToNumber("AA")
                .Row = 49
                .CellNote = ""
                .BackColor = mNonErrorColor
            
            End If
        Else
                .Col = .ColLetterToNumber("P")
                .Row = 49
                .CellNote = ""
                .BackColor = mNonErrorColor
                
                .Col = .ColLetterToNumber("AA")
                .Row = 49
                .CellNote = ""
                .BackColor = mNonErrorColor
                        
        End If
        
        .Sheet = iCurrentSheet
        .Sheet = SHEET_PHU_LUC_01
        'Check emty data
        'Check MST static
        If .SheetVisible = True Then
            mCurrentSheet = .Sheet

            iFlagCheckMaST = 0
            'iWhiteRow = 1
            .Row = 27
            iWhiteRow = 1
            .GetText .ColLetterToNumber("K"), 19, vCt09
            Do
                .GetText .ColLetterToNumber("D"), .Row, vMST
                .GetText .ColLetterToNumber("L"), .Row, vTenNhaThau
                .GetText .ColLetterToNumber("U"), .Row, vTyLePhanBo
                .GetText .ColLetterToNumber("AB"), .Row, vSoThuePhatSinh
                .GetText .ColLetterToNumber("AL"), .Row, vGhiChu
                If Trim$(vMST) <> "" Or Trim$(vTenNhaThau) <> "" Or Trim$(vTyLePhanBo) <> "0" Or Trim$(vSoThuePhatSinh) <> "0" Or Trim$(vGhiChu) <> "" Then
                    iWhiteRow = 0
                    iFlagCheckData = True
                End If
                If Trim$(vMST) <> vbNullString Then
                    vMST = Replace(vMST, " ", "")
                    vMST = Replace(vMST, "-", "")
                    'vMST = Left(vMST, 10)

                    strCheck = ""

                    If Len(vMST) = 10 Or Len(vMST) = 13 Then
                        strCheck = CheckTaxCode(Mid$(vMST, 1, 1), Mid$(vMST, 2, 1), Mid$(vMST, 3, 1), Mid$(vMST, 4, 1), Mid$(vMST, 5, 1), Mid$(vMST, 6, 1), Mid$(vMST, 7, 1), Mid$(vMST, 8, 1), Mid$(vMST, 9, 1), Mid$(vMST, 10, 1), Mid$(vMST, 11, 1), Mid$(vMST, 12, 1), Mid$(vMST, 13, 1))
                    End If
                       
                    If Mid$(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                        .Col = .ColLetterToNumber("D")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errMST   'static
                        .BackColor = mErrorColor

                        iFlagCheckMaST = 1
                            
                        'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str$(.Col)
                            strFocusRow = str$(.Row)

                            iFocusFlag = 1
                        End If

                    Else
                        .Col = .ColLetterToNumber("D")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If

                Else
                    If (Trim$(vMST) = vbNullString And iWhiteRow = 0) Then
                        'is emty
                        .Col = .ColLetterToNumber("D")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = errEmty   'static
                        .BackColor = mErrorColor
                        
                        iFlagEmty = 1
                    End If
                End If
                If (Trim$(vTenNhaThau) = vbNullString And iWhiteRow = 0) Then
                    .Col = .ColLetterToNumber("L")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = errEmty   'static
                    .BackColor = mErrorColor
                    
                    iFlagEmty = 1
                Else
                    .Col = .ColLetterToNumber("L")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                .Col = .ColLetterToNumber("B")
                .Row = .Row + 1
            Loop Until .Text = "aa"
                ' kiem tra tong ty le va tong so thue
            .GetText .ColLetterToNumber("U"), 22, vTongTyLe
             'neu dong co chua du lieu thi kiem tra cac rang buoc trong cac chi tieu
             Const colorCell = 16709097
            If iFlagCheckData Then
               'lay tong ty le chia TNCT
                .GetText .ColLetterToNumber("U"), .MaxRows - 6, vTongTyLe
                'check tong TNCT phai = 100
                If Val(vTongTyLe) <> 100 Then
                    .Col = .ColLetterToNumber("U")
                    .Row = .MaxRows - 6
                    .BackColor = mErrorColor
                    .CellNote = vErrorTongTyLe
                    flagTongTyLe = 1
                Else
                    .Col = .ColLetterToNumber("U")
                    .Row = .MaxRows - 6
                    .BackColor = colorCell
                    .CellNote = ""
                End If
                
                ' kiem tra cot (5) = [09]
                If Val(vCt09) <> Val(vCt09) * Val(vTongTyLe) / 100 Then
                    .Col = .ColLetterToNumber("AB")
                    .Row = .MaxRows - 6
                    .BackColor = mErrorColor
                    .CellNote = vErrorCt09
                    flagCt09 = 1
                Else
                    .Col = .ColLetterToNumber("AB")
                    .Row = .MaxRows - 6
                    .BackColor = colorCell
                    .CellNote = ""
                End If
                
            End If
        End If
        
        'validate sheet phu luc san luong
        .Sheet = SHEET_PHU_LUC_SAN_LUONG
'        If .SheetVisible = True Then
'            .Row = 27
'            iFlagDateCompare = 0
'            Do
'                .GetText .ColLetterToNumber("K"), .Row, vDate1
'                .GetText .ColLetterToNumber("U"), .Row, vDate2
'                arrDate1 = Split(CStr(vDate1), "/")
'                arrDate2 = Split(CStr(vDate2), "/")
'                If (Trim$(vDate1) <> vbNullString And Trim$(vDate2) <> vbNullString) Then
'                    If DateSerial(CInt(arrDate1(2)), CInt(arrDate1(1)), CInt(arrDate1(0))) >= DateSerial(CInt(arrDate2(2)), CInt(arrDate2(1)), CInt(arrDate2(0))) Then
'                        iFlagDateCompare = 1
'                        .Col = .ColLetterToNumber("U")
'                        .BackColor = mErrorColor
'                        .CellNote = errDateCompare
'                    Else
'                        .Col = .ColLetterToNumber("U")
'                        .CellNote = ""  'static
'                        .BackColor = mNonErrorColor
'                    End If
'                End If
'                .Col = .ColLetterToNumber("B")
'                .Row = .Row + 1
'            Loop Until .Text = "aa"
'        End If
        
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 19, IIf(iFlagEmty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 13, IIf(iFlagCheckMaST = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 28, IIf(flagTongTyLe = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 22, IIf(iFlagDateCompare = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 25, IIf(flagDVUyQuyen = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 26, IIf(flagDVUyQuyenEmpty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 29, IIf(flagCt09 = 1, "0", "1")
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

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

Public Sub SetActiveSheet()
    Dim xmlNodeValid As MSXML.IXMLDOMNode, xmlCellNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long, lCol As Long, lRow As Long
    Dim blnNullValue As Boolean
    Dim strValue As String
    Dim strValue1 As String
    Dim strValue2 As String
    Dim strValue3 As String
    Dim strValue4 As Variant
    Dim strValue5 As Variant
    Dim strValue6 As Variant
    Dim strValue7 As Variant
    Dim i As Long
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "01-1"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(2).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(2).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                fps.Sheet = SHEET_PHU_LUC_01
                fps.GetText fps.ColLetterToNumber("D"), 27, strValue4
                    If Val(strValue1) <> 0 Or (Not IsNullValue(strValue4)) Then
                            blnNullValue = False
                    End If

                    If blnNullValue Then
                        SetAttribute xmlNodeValid, "Active", "0"
                    End If
             Case "01-2"
                blnNullValue = True
                'Check value of last node in Data
                fps.Sheet = SHEET_PHU_LUC_SAN_LUONG
                For i = 1 To TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length
                    Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")(i)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                    If Val(strValue1) <> 0 Or (Not IsNullValue(strValue1)) Then
                            blnNullValue = False
                            Exit For
                    End If
                Next
                
                If blnNullValue Then
                    SetAttribute xmlNodeValid, "Active", "0"
                End If
             Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell").length - 4)
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

Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean
    With fps
        .EventEnabled(EventAllEvents) = False
        If isImportXML = True Then
            .Sheet = 1
            .Row = 49
            .Col = .ColLetterToNumber("P")
            .value = "0"
    
            .Col = .ColLetterToNumber("AA")
            .value = "0"
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Function

Public Sub SetData()
    Dim varTemp         As Variant
    Dim ssSheet         As Integer
    Dim lCol            As Long, lRow As Long
    Dim xmlCellNode     As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode

    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet

        .Sheet = .ActiveSheet
        .EventEnabled(EventAllEvents) = True
    End With

    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

Private Sub UpdateSheets()
    Dim varTemp         As Variant
    Dim ssSheet         As Integer
    Dim lCol            As Long, lRow As Long
    Dim xmlCellNode     As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode

    Dim blCheckS2 As Boolean

    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet

        If (GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "Active") <> "0") Then

            For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               
                '.GetText lCol, lRow, varTemp
                'UpdateCell fps, lCol, lRow, varTemp
                If (lRow = 49 And lCol = .ColLetterToNumber("P")) Or (lRow = 49 And lCol = .ColLetterToNumber("AA")) Then
                    .GetText lCol, lRow, varTemp

                    If varTemp = "1" Or UCase(CStr(varTemp)) = "X" Then
                        UpdateCell fps, lCol, lRow, "x"
                    Else
                        UpdateCell fps, lCol, lRow, ""
                    End If

                Else
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
                End If
               
            Next

        End If
        'updated phu luc 01
        .Sheet = SHEET_PHU_LUC_01
        mCurrentSheet = .Sheet
        blCheckS2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheckS2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(2).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
        'end pl 01
        'updated phu luc 02
        .Sheet = 2
        mCurrentSheet = .Sheet
        blCheckS2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheckS2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
            Next
        End If
        'end pl 01
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
    tuRowDL = 53
    denRowDL = 63
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "25_27_2~25_27_3~20_30_4"
    ' end
End Sub

' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("AV"), 72, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function

' Lay ve chi tieu o to khai lan truoc
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
            Set xmlNode = .nodeFromID("AV_" & 72)
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_02TAIN_DK", "getValueCTDC", Err.number, Err.Description
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
            '[39], [40]
            Set xmlNode = .nodeFromID("AV_" & 73)
            result = "CT39_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("AV_" & 74)
            result = result & "~" & "CT40_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_02TAIN_DK", "getValueTK", Err.number, Err.Description
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
    SaveErrorLog "cls_02TAIN_DK", "getMapCT", Err.number, Err.Description
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
                '[39], [40]
                .GetText .ColLetterToNumber("AV"), 73, varTemp
                strCurrentValue = "CT39_" & Trim(varTemp)
                .GetText .ColLetterToNumber("AV"), 74, varTemp
                strCurrentValue = strCurrentValue & "~CT40_" & Trim(varTemp)

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
                
                    If Left(strMaCT, 2) = "35" Then
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



'When cell's data is changed or a control at a cell is handled, This sub is called
'parameter Col: the column of cell that rise event
'parameter Row: the row of cell that rise event
'parameter f: the flag to mark some special option
'**: when this sub is on executing, all events of fpspread will be disabled until the end of the sub
Sub CellChange(ByVal Col As Long, ByVal Row As Long, Optional ByVal f As Integer)
            
    With fps
        'Because of unclearly-rised events, this function is here for turning it off
        'at the end of this sub, this event will be turned on
        .EventEnabled(EventAllEvents) = False
        .Sheet = .ActiveSheet
        mCurrentSheet = .ActiveSheet
        If .ActiveSheet = SHEET_PHU_LUC_SAN_LUONG Then
                mCurrentSheet = .Sheet
                If Col = .ColLetterToNumber("AJ") Then
                    .Sheet = 1
                    ' tong hop chi tieu cot 8 len chi tieu 31 cua to khai
                    .Col = .ColLetterToNumber("AV")
                    .Row = 64
                    .Formula = "BG64"
                End If
        End If
        
        .EventEnabled(EventAllEvents) = True
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
    Dim sheetKhbs As Integer
    sheetKhbs = 4
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
                Case "KHBS"
                    blnNullValue = True
                    'Check value of last node in Data
                    Set xmlCellNode = TAX_Utilities_v1.Data(sheetKhbs - 1).getElementsByTagName("Cell") _
                        (TAX_Utilities_v1.Data(sheetKhbs - 1).getElementsByTagName("Cell").length - 1)
                       strValue1 = GetAttribute(xmlCellNode, "Value")
                    Set xmlCellNode = TAX_Utilities_v1.Data(sheetKhbs - 1).getElementsByTagName("Cell")(1)
                       strValue2 = GetAttribute(xmlCellNode, "Value")
                    Set xmlCellNode = TAX_Utilities_v1.Data(sheetKhbs - 1).getElementsByTagName("Cell")(7)
                       strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v1.Data(sheetKhbs - 1).getElementsByTagName("Cell").length > 31 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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

Private Function getCT29() As String
    Dim xmlDomLastData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim strDataFileName As String
    Dim strM41Target As String, strP41Target As String, strS41Target As String
    Dim iDom As Integer
    Dim MY As Date
    Dim tuThang As String
    Dim denThang As String
    Dim arrTuThang() As String
    Dim i As Integer
    Dim arrFileName() As String
    Dim dataFileKyTruoc As String
    Dim tempFile As String
    Dim arr1() As String
    Dim arr2() As String
    Dim dataFileKyTruocBs As String
    
    
    Dim isFirtFile As Boolean
    
    iDom = 0
    With xmlDomLastData
        .resolveExternals = True
        .validateOnParse = True
        .async = False
        tuThang = TAX_Utilities_v1.FirstDay
        arrTuThang = Split(tuThang, "/")
        
        MY = DateSerial(Val(arrTuThang(1)), Val(arrTuThang(0)), 1)
        MY = DateAdd("M", -1, MY)
        
        If Len(Month(MY)) = 2 Then
            denThang = Month(MY) & Year(MY)
        Else
            denThang = "0" & Month(MY) & Year(MY)
        End If
        
        ReDim arrStrXMLFileNames(0)
        LoadXMLFileNames
        
        For i = 0 To UBound(arrStrXMLFileNames)
            arrFileName = Split(arrStrXMLFileNames(i), "_")
            If UBound(arrFileName) <> -1 Then
                If arrFileName(UBound(arrFileName)) = denThang Then
                    If arrFileName(1) <> "KHBS" And InStr(arrFileName(0), "bs") <= 0 Then
                        ' to khai chinh thuc
                        dataFileKyTruoc = arrStrXMLFileNames(i)
                    ElseIf arrFileName(1) <> "KHBS" And InStr(arrFileName(0), "bs") > 0 Then
                    ' to khai bo sung
                        If isFirtFile = False Then
                            tempFile = arrStrXMLFileNames(i)
                            isFirtFile = True
                        End If
                        arr1 = Split(tempFile, "_")
                        arr2 = Split(arrStrXMLFileNames(i), "_")
                        If Val(Replace(arr1(0), "bs", "")) < Val(Replace(arr2(0), "bs", "")) Then
                            tempFile = arrStrXMLFileNames(i)
                        End If
                        dataFileKyTruocBs = tempFile
                    End If
                End If
            End If
        Next
        
         If dataFileKyTruocBs <> "" Then
            strDataFileName = TAX_Utilities_v1.DataFolder & dataFileKyTruocBs & ".xml"
        ElseIf dataFileKyTruoc <> "" Then
            strDataFileName = TAX_Utilities_v1.DataFolder & dataFileKyTruoc & ".xml"
        Else
            strDataFileName = ""
        End If
        
        If .Load(strDataFileName) = True Then
            Set xmlNode = .nodeFromID("AV_62")
            getCT29 = GetAttribute(xmlNode, "Value")
            Set xmlNode = .nodeFromID("AV_70")
            getCT29 = getCT29 & "~" & GetAttribute(xmlNode, "Value")
        Else
            getCT29 = "0~0"
        End If
        
    End With
    
   
    Set xmlDomLastData = Nothing
    Set xmlNode = Nothing
    Exit Function
    
ErrorHandle:
    SaveErrorLog "cls_02TAIN_DK", "GetMP36Previous", Err.number, Err.Description
End Function



Private Sub LoadXMLFileNames()
    Dim lngIndex As Long
    Dim fso As New FileSystemObject
    Dim fle As File
    
    For Each fle In fso.GetFolder(GetAbsolutePath(TAX_Utilities_v1.DataFolder)).Files
        If Right$(fle.Name, 4) = ".xml" And InStr(fle.Name, "02_TAIN_DK_" & strLoaiToKhaiDK) > 0 Then
            ReDim Preserve arrStrXMLFileNames(lngIndex)
            arrStrXMLFileNames(lngIndex) = Mid$(fle.Name, 1, Len(fle.Name) - 4)
            lngIndex = lngIndex + 1
        End If
    Next
End Sub

