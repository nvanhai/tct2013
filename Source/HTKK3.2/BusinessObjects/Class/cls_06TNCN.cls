VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_06TNCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "To khai quyet toan thu nhap ca nhan mau 04-TNCN" interface sheets
'this Class is belong to TAX_Business_v1 project which will be compline to DLL

Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrSolanBosung As String
Public flgLoadToKhai As Boolean

Private flgErr6B As Boolean
Private arrMST6B() As String

Private arrCMT6B() As String

Private totalRow6B As Long

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
        
        SetDateFormat fps, 1, 62, .ColLetterToNumber("M"), DDMMYYYY
        
        .Sheet = 1
        .Row = 62
        .Col = .ColLetterToNumber("M")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
    End With
    
    
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim viTriMSTDL As String
    
     With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        
        If Strloaitk = "TKCT" And StrSolanBosung = "" Then
            .Col = .ColLetterToNumber("C")
            .Row = 64
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 64
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKCT" And Val(StrSolanBosung) > 0 Then
            .Col = .ColLetterToNumber("C")
            .Row = 64
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 64
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
               ' Neu la truong hop bo sung thi hien thi cot C trong bang ke 6B (Có dieu chinh so lieu)
            .Sheet = 2
            .Col = .ColLetterToNumber("C")
            .ColHidden = False
        End If
        
        
        ' set ky ke khai tu thang den thang
        .Row = 64
        .Col = .ColLetterToNumber("K")
        .Text = TAX_Utilities_v1.FirstDay
        UpdateCell fps, .Col, .Row, .Text
        
        .Row = 64
        .Col = .ColLetterToNumber("L")
        .Text = TAX_Utilities_v1.LastDay
        UpdateCell fps, .Col, .Row, .Text
        
        .Sheet = 2
        .ActiveSheet = .Sheet
        .Row = 1
        .Col = .ColLetterToNumber("D")
        .CellType = CellTypeButton
        .TypeButtonTextColor = RGB(0, 0, 255)
        .TypeButtonColor = &H8000000F
        .TypeButtonAlign = TypeButtonAlignRight
        .Font = "Arial"
        .FontSize = 8
        .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
        
         'net check chon
        .Row = 20
        .Col = .ColLetterToNumber("C")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
        .Font = "Arial"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\Check.bmp"))
        .TypeButtonText = ""
        
        ' Sap xep theo cot Ho va ten
        .Row = 19
        .Col = .ColLetterToNumber("D")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = 8421504
        .Font = "Arial"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = GetAttribute(GetMessageCellById("0154"), "Msg")
        
        ' Sap xep theo cot thue khau tru
'        .Row = 19
'        .Col = .ColLetterToNumber("G")
'        .CellType = CellTypeButton
'        .TypeButtonAlign = TypeButtonAlignLeft
'        .TypeButtonTextColor = vbBlack
'        .TypeButtonColor = &H8000000F
'        .Font = "Arial"
'        .FontSize = 8
'        '.TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
'        .TypeButtonText = GetAttribute(GetMessageCellById("0183"), "Msg")
        
    End With
    
    'check thong tin ma so thue dai ly
    'set vi tri theo thu tu MSTDL~TENNVDLT~ChungChiHN
    viTriMSTDL = "D_22~D_60~D_62"
    updateMSTDL fps, viTriMSTDL
    
        ' Kiem tra xem phien ban 3.1.3 hay phien ban 3.1.2 tro xuong
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim isDFOld As Boolean
    ' end
    
    
    Dim i As Integer
    With fps
       ' test
'        Set xmlCellNode = TAX_Utilities_v1.Data(1).nodeFromID(GetCellID(fps, .ColLetterToNumber("D"), 22))
'        Set xmlCellNode = xmlCellNode.parentNode
'        Set xmlCellNode = xmlCellNode.lastChild
'        If Left$(GetAttribute(xmlCellNode, "CellID"), 1) = "C" Then
'            isDFOld = False
'        Else
'            isDFOld = True
'        End If
        ' end test
    
        .Sheet = 2
        .EventEnabled(EventAllEvents) = False
        i = 1
        .Col = .ColLetterToNumber("B")
        .Row = 22
        Do
'            If isDFOld = True Then
'                 .Col = .ColLetterToNumber("C")
'                 .Text = "0"
'            End If
             .Col = .ColLetterToNumber("B")
             .Row = i + 22
             totalRow6B = i
             i = i + 1
        Loop Until .Text = "aa"
        
        .EventEnabled(EventAllEvents) = True
    End With
    
End Sub

Private Function GetValidatedCellIndex(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long) As Integer
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim lCtrl As Long
    
    Set xmlCellNode = TAX_Utilities_v1.Data(lSheet - 1).nodeFromID(GetCellID(fps, lAnchorCol, lAnchorRow))
    Set xmlCellsNode = xmlCellNode.parentNode
    
    'Get Index of anchor cell
    For lCtrl = 1 To xmlCellsNode.childNodes.length
        If GetAttribute(xmlCellsNode.childNodes(lCtrl - 1), "CellID") = GetAttribute(xmlCellNode, "CellID") Then _
            Exit For
    Next
    GetValidatedCellIndex = lCtrl
    
End Function

Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim Y As Long
    Static so As Long
    If so = 0 Then
        so = 1
    End If
    With fps
        .EventEnabled(EventAllEvents) = False
            If .ActiveSheet = 2 Then
               .Sheet = 2
                
                If Col = .ColLetterToNumber("D") And Row = 19 Then
                    Y = Col
                            
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("N"), (21 + totalRow6B), SortByRow, Y, so
                    .Refresh
                    
                    ' Chieu sap xep
                    If so = 1 Then
                        so = 2 ' Sap xep giam dan
                    Else
                        so = 1 ' Sap xep tang dan
                    End If
                End If
                
                If Col = .ColLetterToNumber("G") And Row = 19 Then
                    Y = Col
                            
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("N"), (21 + totalRow6B), SortByRow, Y, so
                    .Refresh
                    
                    ' Chieu sap xep
                    If so = 1 Then
                        so = 2 ' Sap xep giam dan
                    Else
                        so = 1 ' Sap xep tang dan
                    End If
                End If
                
 
            End If
        .EventEnabled(EventAllEvents) = True
    End With

End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    'UpdateSheets
   
End Sub


Private Sub fps_KeyPress(KeyAscii As Integer)
    With fps
        If .Col = .ColLetterToNumber("F") And .ActiveSheet = 2 Then
            If (KeyAscii = 32) Then
                KeyAscii = 0
            End If
        End If
    End With
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
                         totalRow6B = i
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
    Dim varMST As Variant, varMST5A As Variant, varMST5B As Variant
    Dim sumThuNhap5A As Variant, sumThuNhap5B As Variant
    
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            'check date
            If Col = .ColLetterToNumber("M") And Row = 62 Then
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
        .EventEnabled(EventAllEvents) = True
    End With
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
        
        SetData
        
        ' Lay cac mang MST trong cac bang ke va phu luc 06B
        'prepareMSTCMT
'
        'kiemTraDuLieuImport
        
        .Sheet = 2
        InsertNode 23, .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone) - 1, 2
        
        
        UpdateSheets
        CheckDynamicError 'Set Exception Error on cells of interface
            
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
        
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()
    'chuoi vitriheader tren to khai, chuoi vi tri tren sheet header de check lai
    Dim viTriHeaderTk As String, viTriSheetHeader As String, mangMessage As String, xmlNode As MSXML.IXMLDOMNode
    Dim MSTDTNT As Variant, tSheet As Integer
    
    With fps
    
        'get MST cua DTNT
        tSheet = .Sheet
        .Sheet = 1
        .GetText .ColLetterToNumber("D"), 10, MSTDTNT
        MSTDTNT = Trim(MSTDTNT)
        
        'check thong tin header khong dc trong
        'lay vi tri header tren tk theo thu tu MSTDL, TENDLT, DIACHIDLT,QUAN/HUYEN,TP,SO HDDLT,NGAYDH,Ten NV DLT, CC Hanh Nghe

        viTriHeaderTk = "D_22~D_20~D_24~D_26~M_26~D_32~M_32~D_60~D_62"
        viTriSheetHeader = "B_30~B_31~B_32~B_33~B_34~B_35~B_36~B_37~B_38"
        mangMessage = "E_30~E_31~E_32~E_33~E_34~E_35~E_36~E_37~E_38"
        'check header
        checkErrorHeader fps, viTriHeaderTk, viTriSheetHeader, mangMessage, MSTDTNT
        
        .Sheet = tSheet
    
        CheckErrorMST
        ' Neu ton tai file data bang ke 06B thi kiem tra tren bang ke 06B
        If IIf(TAX_Utilities_v1.NodeValidity.childNodes(.SheetCount - 3).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            checkSheet1
        Else
            clearCheckSheet1
        End If
    End With
End Sub

Private Sub checkSheet1()
Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    Dim strCheck As String, iCurrentSheet As Integer
    
    Dim vMST As Variant, vErrMsgMST As Variant, vErrMsgMSTReq, varTam As Variant
    Dim vCMT As Variant, vErrMsgCMT, vErrMsgCMT1 As Variant
    Dim vHoVaTen As Variant, vErrMsgHoVaTen As Variant
    Dim ChiTieuCT09 As Variant, ChiTieuCT13 As Variant, ChiTieuCT15 As Variant, ChiTieu09 As Variant, ChiTieu13 As Variant, ChiTieu15 As Variant
    Dim ChiTieu17 As Variant, ChiTieu18 As Variant, ThueKhauTru As Variant
    Dim vThueTNDaKhauTru As Variant, vErrMsgThueTNDaKhauTru As Variant
    Dim vTongTNChiuThue As Variant, vErrMsg09 As Variant, vErrMsg13 As Variant, vErrMsg15 As Variant, vErrMsgThueKhauTru As Variant
    Dim vLastRow As Variant
    Dim iFlagvHoVaTen As Integer, iFlagvMst, iFlagvCmt As Integer, iFlagvThueDaKhauTru As Variant, iFocusFlag As Integer
    Dim iFlagvChiTieu09 As Variant, iFlagvChiTieu13 As Variant, iFlagvChiTieu15 As Variant, iFlagvThueKhauTru As Variant
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String
    Dim i, j As Integer
    Dim blCheck_S1, blCheck_S2, blCheck_S3 As Boolean
    Dim vErrMsgTrungMST, vErrMsgTrungCMT As Variant, iFlagTrungMST As Integer, iFlagTrungCMT As Integer, rowDes As Long, countRow As Long
    Dim mErrMSTTrungMSTCQT, vMSTCQ As Variant, iFlagErrTrungMSTCQT As Integer
    Dim mErrCheckCT11 As Variant, iFlagCheckCT11 As Integer
    
    iFlagvChiTieu13 = 0
    iFlagvChiTieu15 = 0
    
    ' check chi tieu 31
    Dim vCT31
    Dim vErrorCt31
    Dim iFlagCheckCT31 As Integer
    
    With fps
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 25, vErrMsgTrungMST
       .GetText .ColLetterToNumber("E"), 26, vErrMsgTrungCMT
        ' Lay thong bao sai ve ngay thanh toan trong sheet header
        .GetText .ColLetterToNumber("E"), 15, vErrMsgHoVaTen
        ' Lay thong bao ve ko nhap ma so thue trong sheet header
        .GetText .ColLetterToNumber("E"), 16, vErrMsgMSTReq
        ' Lay thong bao ve ko nhap so CMT/Ho chieu trong sheet header
        .GetText .ColLetterToNumber("E"), 17, vErrMsgCMT
        ' Lay thong bao ve sai ma so thue trong sheet header
        .GetText .ColLetterToNumber("E"), 18, vErrMsgMST
        ' Lay thong bao ve sai ve thue da khau tru trong sheet header
        .GetText .ColLetterToNumber("E"), 19, vErrMsgThueTNDaKhauTru
        .GetText .ColLetterToNumber("E"), 27, vErrMsgCMT1
        'lay thong bao ve TH CT10 < CT11
       .GetText .ColLetterToNumber("E"), 29, mErrCheckCT11
       ' check giang buoc
       .GetText .ColLetterToNumber("E"), 20, vErrMsg09
       .GetText .ColLetterToNumber("E"), 21, vErrMsg13
       .GetText .ColLetterToNumber("E"), 22, vErrMsg15
       .GetText .ColLetterToNumber("E"), 23, vErrMsgThueKhauTru
       .GetText .ColLetterToNumber("E"), 28, mErrMSTTrungMSTCQT
       
       
       .Sheet = 1
       
       .GetText .ColLetterToNumber("I"), 42, ChiTieu09
       .GetText .ColLetterToNumber("I"), 48, ChiTieu13
       .GetText .ColLetterToNumber("I"), 51, ChiTieu15
       
       .GetText .ColLetterToNumber("P"), 42, ChiTieuCT09
       .GetText .ColLetterToNumber("P"), 48, ChiTieuCT13
       .GetText .ColLetterToNumber("P"), 51, ChiTieuCT15
       
       .GetText .ColLetterToNumber("I"), 54, ChiTieu17
       .GetText .ColLetterToNumber("I"), 55, ChiTieu18
       'check chi tieu (18) <= chi tieu (17)
'       ThueKhauTru = Trim(ChiTieu17) * 0.05 - Trim(ChiTieu18)
'       If ThueKhauTru < 0 Then
'            .Col = .ColLetterToNumber("I")
'            .Row = 55
'            .CellNote = vErrMsgThueKhauTru
'            .BackColor = mAlertColor
'            iFlagvThueKhauTru = 1
'        Else
'            .Col = .ColLetterToNumber("I")
'            .Row = 55
'            .CellNote = ""
'            .BackColor = mNonErrorColor
'        End If
       
       'check chi tieu 09
'       If Val(Trim(ChiTieu09)) <> Round(Trim(ChiTieuCT09), 0) And Val(Trim(ChiTieu09)) <> (Round(Trim(ChiTieuCT09), 0) + 1) Then
'            .Col = .ColLetterToNumber("I")
'            .Row = 42
'            .CellNote = vErrMsg09
'            .BackColor = mAlertColor
'            iFlagvChiTieu09 = 0
'        Else
'            .Col = .ColLetterToNumber("I")
'            .Row = 42
'            .CellNote = ""
'            .BackColor = mNonErrorColor
'        End If
'
        'check chi tieu 13
'        If Val(Trim(ChiTieu13)) <> Round(Trim(ChiTieuCT13), 0) And Val(Trim(ChiTieu13)) <> (Round(Trim(ChiTieuCT13), 0) + 1) Then
'            .Col = .ColLetterToNumber("I")
'            .Row = 48
'            .CellNote = vErrMsg13
'            .BackColor = mAlertColor
'
'        Else
'            .Col = .ColLetterToNumber("I")
'            .Row = 48
'            .CellNote = ""
'            .BackColor = mNonErrorColor
'        End If
        
        'check chi tieu 15
'        If (Val(Trim(ChiTieu15)) <> Round(Trim(ChiTieuCT15), 0)) And (Val(Trim(ChiTieu15)) <> (Round(Trim(ChiTieuCT15), 0) + 1)) Then
'            .Col = .ColLetterToNumber("I")
'            .Row = 51
'            .CellNote = vErrMsg15
'            .BackColor = mAlertColor
'            iFlagvChiTieu15 = 0
'        Else
'            .Col = .ColLetterToNumber("I")
'            .Row = 51
'            .CellNote = ""
'            .BackColor = mNonErrorColor
'        End If
        
        .Sheet = .SheetCount
        
        .SetText .ColLetterToNumber("B"), 20, IIf(iFlagvChiTieu09 = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 21, IIf(iFlagvChiTieu13 = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 22, IIf(iFlagvChiTieu15 = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 23, IIf(iFlagvThueKhauTru = 1, "0", "1")
        
        
        .Sheet = 2 'Check tinh xac thuc cua cac thong tin tren bang ke 06B TNCN
        
        'get MST cua co quan chi tra
        .GetText .ColLetterToNumber("E"), 9, vMSTCQ
        
        .Row = 22
            Do
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

                    
                     .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
        
        blCheck_S1 = True
        blCheck_S2 = True
        blCheck_S3 = True
        iFlagvCmt = 0
        iFlagErrTrungMSTCQT = 0
        iFlagCheckCT11 = 0
        For j = 22 To .MaxRows
            .Row = j
            
            .GetText .ColLetterToNumber("D"), .Row, vHoVaTen
            .GetText .ColLetterToNumber("E"), .Row, vMST
            .GetText .ColLetterToNumber("F"), .Row, vCMT
            .GetText .ColLetterToNumber("G"), .Row, vTongTNChiuThue
            .GetText .ColLetterToNumber("H"), .Row, vThueTNDaKhauTru
            .GetText .ColLetterToNumber("B"), .Row, vLastRow
            If UCase(vLastRow) = "AA" Then Exit For ' Check cho den het cac dong co du lieu thi thoi
            
            If (Trim(vHoVaTen) = vbNullString And Trim(vMST) = vbNullString And Trim(vCMT) = vbNullString) And vTongTNChiuThue = 0 And vThueTNDaKhauTru = 0 Then
            
            Else
                'check phai CT11 < CT 10
                If vThueTNDaKhauTru >= vTongTNChiuThue And Val(vThueTNDaKhauTru) > 0 Then
                    .Col = .ColLetterToNumber("H")
                    .CellNote = mErrCheckCT11
                    .BackColor = mErrorColor
                    iFlagCheckCT11 = 1
                Else
                    .Col = .ColLetterToNumber("H")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                'Check Ho va ten tren banh ke 06B/TNCN
                If Trim(vHoVaTen) = vbNullString Then
                    .Col = .ColLetterToNumber("D")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgHoVaTen    'static
                    .BackColor = mErrorColor
                    iFlagvHoVaTen = 1
                    blCheck_S1 = False
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

                'Check ma so thue tren to khai 06B/TNCN
                If Trim(vMST) <> vbNullString Then
    '               vMST = Replace(vMST, " ", "")
    '               vMST = Replace(vMST, "-", "")
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
                        blCheck_S2 = False
                        'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
                    Else

                       'check MST ko dc trung voi MSTCQCT
                        vMSTCQ = Replace(vMSTCQ, " ", "")
                        vMSTCQ = Replace(vMSTCQ, "-", "")
                        If Trim(vMST) = Trim(vMSTCQ) Then
                            .Col = .ColLetterToNumber("E")
                            .BackColor = mErrorColor
                            .CellNote = mErrMSTTrungMSTCQT
                            iFlagErrTrungMSTCQT = 1
                        Else
                            'check MST trung nhau
                            countRow = .SearchCol(.ColLetterToNumber("E"), -1, -1, vMST, SearchFlagsNone)
                            rowDes = .SearchCol(.ColLetterToNumber("E"), countRow, -1, vMST, SearchFlagsNone)
                            If rowDes > -1 Then
                               .Col = .ColLetterToNumber("E")
                               .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                               .CellNote = vErrMsgTrungMST      'static
                               .BackColor = mErrorColor
                               iFlagTrungMST = 1
                            End If
                        End If
                        
                    End If
                    
                    

                    
                ElseIf (Trim(vMST) = vbNullString And Trim(vCMT) = vbNullString) Then
                    .Col = .ColLetterToNumber("E")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMSTReq    'static
                    .BackColor = mErrorColor
                    
                    .Col = .ColLetterToNumber("F")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgCMT    'static
                    .BackColor = mErrorColor
                    
                    
                    iFlagvMst = 1
                    blCheck_S2 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
'                    .Col = .ColLetterToNumber("E")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
                    
'                    .Col = .ColLetterToNumber("E")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
                End If

                If Trim(vCMT) <> vbNullString Then
                    'check CMT trung nhau
                     countRow = .SearchCol(.ColLetterToNumber("F"), -1, -1, vCMT, SearchFlagsNone)
                     rowDes = .SearchCol(.ColLetterToNumber("F"), countRow, -1, vCMT, SearchFlagsNone)
                    If rowDes > -1 Then
                        .Col = .ColLetterToNumber("F")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgTrungCMT      'static
                        .BackColor = mAlertColor
                        'iFlagTrungCMT = 1
                    Else
                        If InStr(1, Trim(vCMT), " ", vbTextCompare) <> 0 Then
                                .Col = .ColLetterToNumber("F")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrMsgCMT1   'static
                                .BackColor = mErrorColor
                                iFlagvCmt = 1
                        End If
                    End If
                End If
                
            End If
            
        Next
        
        
'        .Row = 22
'        Do
'            i = .Row
'            .GetText .ColLetterToNumber("K"), .Row, varTam
'            If varTam = 1 Then
'                    .Col = .ColLetterToNumber("E")
'                    .CellNote = vErrMsgTrungMST
'                    .BackColor = mErrorColor
'            End If
'            .GetText .ColLetterToNumber("L"), .Row, varTam
'            If varTam = 1 Then
'                    .Col = .ColLetterToNumber("F")
'                    .CellNote = vErrMsgTrungCMT
'                    .BackColor = mErrorColor
'            End If
'            ifl
'            .Row = .Row + 1
'            .Col = .ColLetterToNumber("B")
'        Loop Until .Text = "aa"
        ' Trong truong hop blCheck_S1 khong co Ho va ten nao bi loi thi set lai gia tri cua Header la 1
        If iFlagvCmt = 1 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 27, "0"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 27, "1"
        End If
        
        If iFlagErrTrungMSTCQT = 1 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 28, "0"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 28, "1"
        End If
        
        If blCheck_S1 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "0"
        End If
        ' Trong truong hop blCheck_S2 khong co MST nao bi loi thi set lai gia tri cua Header la 1
        If blCheck_S2 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 16, "1"
            .SetText .ColLetterToNumber("B"), 17, "1"
            .SetText .ColLetterToNumber("B"), 18, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 16, "0"
            .SetText .ColLetterToNumber("B"), 17, "0"
            .SetText .ColLetterToNumber("B"), 18, "0"
            
        End If
        'trong TH CT [11] > CT [10] set lai gia tri cua header
        If iFlagCheckCT11 = 1 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 29, "0"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 29, "1"
        End If
        ' Trong truong hop blCheck_S3 khong co so thu da khau tru nao bi loi thi set lai gia tri cua Header la 1
        If blCheck_S3 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 19, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 19, "0"
        End If
        
        'trong TH trung MST set lai gia tri cua header
        If iFlagTrungMST = 1 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 25, "0"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 25, "1"
        End If
        
        'trong TH trung CMT set lai gia tri cua header
'        If iFlagTrungCMT = 1 Then
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 26, "0"
'        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 26, "1"
'        End If
        
        .Sheet = iCurrentSheet
    End With
End Sub

Private Sub clearCheckSheet1()
    With fps
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 15, "1"
        .SetText .ColLetterToNumber("B"), 16, "1"
        .SetText .ColLetterToNumber("B"), 17, "1"
        .SetText .ColLetterToNumber("B"), 18, "1"
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
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "06B"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in E column and F column
                    If lCol = fps.ColLetterToNumber("D") Then
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
Dim isheet As Integer
'    With fps
'        .EventEnabled(EventAllEvents) = False
'         isheet = mCurrentSheet
'         mCurrentSheet = 1
'        .Sheet = 1
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 15
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 16
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 17
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 17
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 18
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 19
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 20
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 21
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 22
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'
'        .Col = fps.ColLetterToNumber("P")
'        .Row = 23
'        .Lock = True
'        UpdateCell fps, .Col, .Row, .Value
'        .EventEnabled(EventAllEvents) = True
'        mCurrentSheet = isheet
'        .ActiveSheet = mCurrentSheet
'    End With
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
    Dim xmlNodeData As MSXML.IXMLDOMNode
    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet
        
        If (GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(0), "Active") <> "0") Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
        If (GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(1), "Active") <> "0") Then
        .Sheet = 2
        mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               If lCol = .ColLetterToNumber("C") Then
                 If Trim(varTemp) = "1" Then
                    varTemp = "x"
                 ElseIf Trim(varTemp) = "0" Then
                    varTemp = ""
                 End If
               End If
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
    Dim i As Integer, iCol As Long, iRow As Long
    Dim varMST, varCMT As Variant
With fps
fps.EventEnabled(EventAllEvents) = False
    ' set node xml
    .Sheet = 2
    InsertNode 23, .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone) - 1, 2
    ' lay ve so row du lieu
    totalRow6B = .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone) - 22
    UpdateSheets
    
'    If .ActiveSheet = 2 Then
'                    i = 1
'                    .Col = .ColLetterToNumber("B")
'                    .Row = 22
'                    Do
'                         .Text = str(i)
'                         ' MST
'                        .GetText .ColLetterToNumber("E"), .Row, varMST
'                         varMST = Left(varMST, IIf(InStr(1, varMST, ".") <> 0, InStr(1, varMST, ".") - 1, Len(varMST)))
'                         .Col = .ColLetterToNumber("E")
'                         .Text = varMST
'                        ' CMT
'                        .GetText .ColLetterToNumber("F"), .Row, varCMT
'                        varCMT = Left(varCMT, IIf(InStr(1, varCMT, ".") <> 0, InStr(1, varCMT, ".") - 1, Len(varCMT)))
'                        .Col = .ColLetterToNumber("F")
'                        .Text = varCMT
'                        ' end
'                         .Col = .ColLetterToNumber("B")
'                         .Row = i + 22
'                         totalRow6B = i
'                         i = i + 1
'                    Loop Until .Text = "aa"
'                    .Row = iRow
'                    .Col = iCol
'                    .SetActiveCell iCol, iRow
'            End If
        fps.EventEnabled(EventAllEvents) = True
        End With
End Sub

Private Sub fps_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
    Dim sumLaoDong As Variant, sumThuNhap As Variant, sumSoThueTNDaKhauTru As Variant
    Dim varHoTen As Variant, varMST As Variant, varCMT As Variant
    Dim varThuNhap As Variant, varSoThueTNDaKhauTru As Variant
    
    If NewSheet = 1 Then
        With fps
            'Because of unclearly-rised events, this function is here for turning it off
            'at the end of this sub, this event will be turned on
            .EventEnabled(EventAllEvents) = False
            .Sheet = 2 'Bang ke 6D
            sumLaoDong = 0
            sumThuNhap = 0
            sumSoThueTNDaKhauTru = 0
            .Row = 22
            Do
                 .GetText .ColLetterToNumber("C"), .Row, varHoTen
                 .GetText .ColLetterToNumber("D"), .Row, varMST
                 .GetText .ColLetterToNumber("F"), .Row, varCMT
                 .GetText .ColLetterToNumber("G"), .Row, varThuNhap
                 .GetText .ColLetterToNumber("H"), .Row, varSoThueTNDaKhauTru
                                  
                 If Trim(varHoTen) <> vbNullString And (Trim(varMST) <> vbNullString Or Trim(varCMT) <> vbNullString) Then
                    sumLaoDong = sumLaoDong + 1
                    sumThuNhap = sumThuNhap + IIf(varThuNhap = vbNullString, 0, Replace(varThuNhap, " ", ""))
                    sumSoThueTNDaKhauTru = sumSoThueTNDaKhauTru + IIf(varSoThueTNDaKhauTru = vbNullString, 0, varSoThueTNDaKhauTru)
                 End If
                 .Row = .Row + 1
                 .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
             ' Cap nhat vao to quyet toan 02A.
            .Sheet = 1
            mCurrentSheet = .Sheet
'            .Col = .ColLetterToNumber("I")
'
'             'Chi tieu [17]
'             .Row = 37
'             .Value = sumLaoDong
'             UpdateCell fps, .Col, .Row, .Value
'
'             'Chi tieu [18]
'             .Row = 38
'             .Value = sumThuNhap
'             UpdateCell fps, .Col, .Row, .Value
'
'             'Chi tieu [19]
'             .Row = 39
'             .Value = sumSoThueTNDaKhauTru
'             UpdateCell fps, .Col, .Row, .Value
                        
            .EventEnabled(EventAllEvents) = True
        End With
    End If
End Sub


Private Sub kiemtra6B()
    Dim varTemp As Variant
    Dim varHoVaTen, varMST, varCMT  As Variant
    Dim varThuNhap, varTam, varBaoHiem, varSoNguoiGiamTru, varSoThangGiamTru, varThuNhapGiamThue, varSoThueDuocGiam, varSoThueKhauTru, vErrMsgTrungMST, vErrMsgTrungCMT As Variant
    Dim strCheck As String
    
    Dim varMSTCQCT As Variant
    
    Dim startRowMST As Long
    Dim rowCountMST As Long
    Dim totalRowMST As Long
    Dim oldRowMST As Long
    Dim ret2, tam As Long
    Dim flgErr6BMST As Variant
    Dim flgErr6BCMT As Variant
    
    flgErr6BMST = True
    flgErr6BCMT = True
    totalRowMST = UBound(arrMST6B)
    Dim i As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 2
             .Row = 22
            Do
                    .Col = .ColLetterToNumber("K")
                    .Text = "0"
                    .Col = .ColLetterToNumber("L")
                    .Text = "0"
                .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
        
         'Bang ke 6B
        .Row = 22
        Do
            i = .Row
            ' Cot Ho va ten
            .GetText .ColLetterToNumber("D"), .Row, varHoVaTen
            ' Cot MST
            
            .GetText .ColLetterToNumber("E"), .Row, varMST
            ' Ma la null
                'dhdang check trung MST
                If Trim(varMST) <> "" Then
                    ret2 = .SearchCol(.ColLetterToNumber("E"), 22, (22 + .MaxRows), varMST, SearchFlagsNone)
                    If ret2 > -1 And ret2 <> .Row Then
                         flgErr6BMST = False
                        .Col = .ColLetterToNumber("K")
                        tam = .Row
                        .Row = ret2
                        .Text = "1"
    '                    .CellNote = vErrMsgTrungMST
    '                    .BackColor = mErrorColor
                        .Row = tam
                        .Text = "1"
                    End If
                 End If
                 'dhdang check trung CMT
                 .GetText .ColLetterToNumber("F"), .Row, varCMT
                If Trim(varCMT) <> "" Then
                    ret2 = .SearchCol(.ColLetterToNumber("F"), 22, (22 + .MaxRows), varCMT, SearchFlagsNone)
                    If ret2 > -1 And ret2 <> .Row Then
                         flgErr6BCMT = False
                        .Col = .ColLetterToNumber("L")
                        tam = .Row
                        .Row = ret2
                        .Text = "1"
                        .Row = tam
                        .Text = "1"
                    End If
                End If
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 25, IIf(flgErr6BMST = False, "0", "1")
        .SetText .ColLetterToNumber("B"), 26, IIf(flgErr6BCMT = False, "0", "1")
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub prepareMSTCMT()
    Dim countMST As Long, i As Integer
    Dim varMST, varCMT As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 2 'Bang ke 6B
        .Row = 22
        countMST = 0
        Do
            countMST = countMST + 1
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        If countMST = 1 Then
            ReDim arrMST6B(1) As String
            ReDim arrCMT6B(1) As String
            arrMST6B(0) = ""
            arrCMT6B(0) = ""
        Else
            .Row = 22
            ReDim arrMST6B(22 + countMST - 1) As String
            ReDim arrCMT6B(22 + countMST - 1) As String
            For i = 0 To 21
                arrMST6B(i) = ""
                arrCMT6B(i) = ""
            Next
            Do
                
                .GetText .ColLetterToNumber("E"), .Row, varMST
                varMST = Left(varMST, IIf(InStr(1, varMST, ".") <> 0, InStr(1, varMST, ".") - 1, Len(varMST)))
                .Col = .ColLetterToNumber("E")
                .Text = varMST
                
                If Trim(varMST) <> vbNullString Then
                    arrMST6B(.Row) = varMST
                Else
                    arrMST6B(.Row) = ""
                End If
                
                .GetText .ColLetterToNumber("F"), .Row, varCMT
                varCMT = Left(varCMT, IIf(InStr(1, varCMT, ".") <> 0, InStr(1, varCMT, ".") - 1, Len(varCMT)))
                .Col = .ColLetterToNumber("F")
                .Text = varCMT
                
                If Trim(varCMT) <> vbNullString Then
                    arrCMT6B(.Row) = varCMT
                Else
                    arrCMT6B(.Row) = ""
                End If
                .Row = .Row + 1
                .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
        End If
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub


Public Sub kiemTraDuLieuImport()
    'kiemtra6B
    'prepareMSTCMT
End Sub


Public Sub InsertNode(ByVal fRow As Long, ByVal lRow As Long, ByVal mSheet As Byte)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCell As MSXML.IXMLDOMNode
    
    
    Dim ctlRow As Long
    Dim cTemplate, cReport As String
    
    Dim sumRowDel As Long
    sumRowDel = TAX_Utilities_v1.Data(mSheet - 1).getElementsByTagName("Cells").length
    For ctlRow = 1 To sumRowDel - 1
        If mSheet = 2 Then
            Set xmlNodeCells = TAX_Utilities_v1.Data(mSheet - 1).nodeFromID("E" & "_" & 22 + ctlRow & "").parentNode
            xmlNodeCells.parentNode.removeChild xmlNodeCells
        End If
    Next
    
    If mSheet = 2 Then
        Set xmlNodeCells = TAX_Utilities_v1.Data(mSheet - 1).nodeFromID("E_22").parentNode
    End If
    
    For ctlRow = fRow To lRow
        
        Set xmlNodeNewCells = xmlNodeCells.cloneNode(True)
        
        For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
            cTemplate = Left$(GetAttribute(xmlNodeNewCell, "CellID"), 2)
            cReport = Left$(GetAttribute(xmlNodeNewCell, "CellID2"), 2)
            
            SetAttribute xmlNodeNewCell, "CellID", cTemplate & ctlRow
                    
            ' Set first cell = 1
            SetAttribute xmlNodeNewCell, "FirstCell", "1"
            If mSheet = 2 Then
                If Trim(cTemplate) = "D_" Or Trim(cTemplate) = "E_" Or Trim(cTemplate) = "F_" Then
                    SetAttribute xmlNodeNewCell, "Value", vbNullString
                Else
                    SetAttribute xmlNodeNewCell, "Value", "0"
                End If
                SetAttribute xmlNodeNewCell, "CellID2", cReport & (ctlRow - 7)
            End If
            
        Next
        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, Null
        
    
    Next
'   **********************************
'    added
'   Date: 12/04/06
    'mAdjustData = True
    TAX_Utilities_v1.AdjustData(mSheet - 1) = True
'   **********************************
EXIT_SUB:
    Set xmlNodeNewCell = Nothing
    Set xmlNodeNewCells = Nothing
    Set xmlNodeCells = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_06TNCN", "InsertNode 06_TNCN", Err.number, Err.Description
End Sub

' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 14
    denRowDL = 19
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "7_8_2"
    ' end
End Sub


