VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_02XSTNCN"
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
Public flgLoadToKhai As Boolean
Private totalRow2AXS As Long

'sheet1
Private Const tTLanDau_C As String = "C"
Private Const tTLanDau_R As Long = 53
Private Const tTLanThu_C As String = "I"
Private Const tTLanThu_R As Long = 53
Private Const ngayKy_C As String = "M"
Private Const ngayKy_R As Long = 50
'sheet2

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
    Dim mRow As Long, mCol As Long
    With fps
        
        SetDateFormat fps, 1, ngayKy_R, .ColLetterToNumber(ngayKy_C), DDMMYYYY
        
        .Sheet = 1 'To khai 02BH/TNCN
        .Row = ngayKy_R
        .Col = .ColLetterToNumber(ngayKy_C)
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
            .Col = .ColLetterToNumber(tTLanDau_C)
            .Row = tTLanDau_R
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber(tTLanThu_C)
            .Row = tTLanThu_R
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKCT" And Val(StrSolanBosung) > 0 Then
            .Col = .ColLetterToNumber(tTLanDau_C)
            .Row = tTLanDau_R
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber(tTLanThu_C)
            .Row = tTLanThu_R
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
            
            ' Neu la truong hop bo sung thi hien thi cot C trong bang ke 2A (Có dieu chinh so lieu)
            .Sheet = 2
            .Col = .ColLetterToNumber("C")
            .ColHidden = False
            
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
        
        ' Nut check chon
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
        .TypeButtonText = GetAttribute(GetMessageCellById("0244"), "Msg")
        
        ' Sap xep theo cot thu nhap chiu thue
        .Row = 19
        .Col = .ColLetterToNumber("G")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = 8421504
        .Font = "Arial"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = GetAttribute(GetMessageCellById("0245"), "Msg")
        
        ' Sap xep theo cot thue thu nhap ca nhan da khau tru
        .Row = 19
        .Col = .ColLetterToNumber("J")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = 8421504
        .Font = "Arial"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = GetAttribute(GetMessageCellById("0246"), "Msg")
        
        .EventEnabled(EventAllEvents) = True
        
    End With
    
    'check thong tin ma so thue dai ly
    'set vi tri theo thu tu MSTDL~TENNVDLT~ChungChiHN
    viTriMSTDL = "D_22~D_48~D_50"
    updateMSTDL fps, viTriMSTDL
    
            ' Kiem tra xem phien ban 3.1.3 hay phien ban 3.1.2 tro xuong
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim isDFOld As Boolean
    ' end
    
    Dim i As Integer
    With fps
        Set xmlCellNode = TAX_Utilities_New.Data(1).nodeFromID(GetCellID(fps, .ColLetterToNumber("D"), 22))
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.lastChild
        If Left$(GetAttribute(xmlCellNode, "CellID"), 1) = "C" Then
            isDFOld = False
        Else
            isDFOld = True
        End If
        ' end test
        .Sheet = 2
        .EventEnabled(EventAllEvents) = False
        i = 1
        .Col = .ColLetterToNumber("B")
        .Row = 22
        Do
             If isDFOld = True Then
                 .Col = .ColLetterToNumber("C")
                 .Text = "0"
             End If
             
             .Col = .ColLetterToNumber("B")
             .Row = i + 22
             totalRow2AXS = i
             i = i + 1
        Loop Until .Text = "aa"
        
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
                            
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("M"), (21 + totalRow2AXS), SortByRow, Y, so
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
                            
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("M"), (21 + totalRow2AXS), SortByRow, Y, so
                    .Refresh
                    
                    ' Chieu sap xep
                    If so = 1 Then
                        so = 2 ' Sap xep giam dan
                    Else
                        so = 1 ' Sap xep tang dan
                    End If
                End If
                
                If Col = .ColLetterToNumber("J") And Row = 19 Then
                    Y = Col
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("M"), (21 + totalRow2AXS), SortByRow, Y, so
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

Public Sub kiemTraDuLieuImport()
    Dim varTemp As Variant
    Dim sumThueDaKhauTru As Variant
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 2 'Bang ke 5A
        .Row = 22
        Do
            .GetText .ColLetterToNumber("J"), .Row, sumThueDaKhauTru
            
            .GetText .ColLetterToNumber("D"), .Row, varTemp
            ' Cu them mot lao dong thi tinh la mot nguoi
            If Trim(varTemp) <> vbNullString Then
                .Col = .ColLetterToNumber("K")
                .Text = "1"
            Else
                .Col = .ColLetterToNumber("K")
                .Text = "0"
            End If
            
            If Val(sumThueDaKhauTru) > 0 Then
                ' Va tinh la lao dong co thue khau tru
                .Col = .ColLetterToNumber("L")
                .Text = "1"
                ' Tinh tong thu nhap cua rieng lao dong co thue khau tru nay
                .GetText .ColLetterToNumber("G"), .Row, varTemp
                .Col = .ColLetterToNumber("M")
                .Text = varTemp
            Else
                ' Va tinh la lao dong ko co thue khau tru
                .Col = .ColLetterToNumber("L")
                .Text = "0"
                ' Khong tinh tong thu nhap cua lao dong nay
                .Col = .ColLetterToNumber("M")
                .Text = "0"
            End If
            
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        .EventEnabled(EventAllEvents) = True
    End With
    checkSheet1
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
                         totalRow2AXS = i
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
            If Col = .ColLetterToNumber(ngayKy_C) And Row = ngayKy_R Then
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
            
            If flgLoadToKhai = True Then
                .EventEnabled(EventAllEvents) = True
                Exit Sub
            End If
            
            If Col = .ColLetterToNumber("D") Then
                .GetText .ColLetterToNumber("D"), .Row, varTemp
                ' Cu them mot lao dong thi tinh la mot nguoi
                If Trim(varTemp) <> vbNullString Then
                    .Col = .ColLetterToNumber("K")
                    .Text = "1"
                Else ' Neu ko thi ko tinh nguoi do
                    .Col = .ColLetterToNumber("K")
                    .Text = "0"
                End If
            End If
            ' Lay so thue da khau tru cua lao dong do
            If Col = .ColLetterToNumber("J") Then
                UpdateCell fps, Col, Row, .value
'                .GetText .ColLetterToNumber("J"), .Row, varTemp
'                If Val(varTemp) > 0 Then
'                    ' Va tinh la lao dong co thue khau tru
'                    .Col = .ColLetterToNumber("L")
'                    .Text = "1"
'                    ' Tinh tong thu nhap cua rieng lao dong co thue khau tru nay
'                    .GetText .ColLetterToNumber("G"), .Row, varTemp
'                    .Col = .ColLetterToNumber("M")
'                    .Text = varTemp
'                Else
'                    ' Va tinh la lao dong ko co thue khau tru
'                    .Col = .ColLetterToNumber("L")
'                    .Text = "0"
'                    ' Khong tinh tong thu nhap cua lao dong nay
'                    .Col = .ColLetterToNumber("M")
'                    .Text = "0"
'                End If
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
    'If flgLoadToKhai = True Then Exit Sub
    With fps
        .EventEnabled(EventAllEvents) = False
        ASheet = .ActiveSheet
        SSheet = .Sheet
        

        SetData

        ' insert them node xml
        .Sheet = 2
        InsertNode 23, .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone) - 1
        
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
        'If flgLoadToKhai = True Then Exit Sub
        'get MST cua DTNT
        tSheet = .Sheet
        .Sheet = 1
        .GetText .ColLetterToNumber("D"), 10, MSTDTNT
        MSTDTNT = Trim(MSTDTNT)
        
        'check thong tin header khong dc trong
        'lay vi tri header tren tk theo thu tu MSTDL, TENDLT, DIACHIDLT,QUAN/HUYEN,TP,SO HDDLT,NGAYDH,Ten NV DLT, CC Hanh Nghe

        viTriHeaderTk = "D_22~D_20~D_24~D_26~M_26~D_32~M_32~D_46~D_48"
        viTriSheetHeader = "B_37~B_38~B_39~B_40~B_41~B_42~B_43~B_44~B_45"
        mangMessage = "E_37~E_38~E_39~E_40~E_41~E_42~E_43~E_44~E_45"
        'check header
        checkErrorHeader fps, viTriHeaderTk, viTriSheetHeader, mangMessage, MSTDTNT
        
        .Sheet = tSheet
        
        CheckErrorMST
        ' Neu ton tai file data bang ke 05A thi kiem tra tren bang ke 5A
        If IIf(TAX_Utilities_New.NodeValidity.childNodes(.SheetCount - 3).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            checkSheet1
        Else
            clearCheckSheet1
        End If
        'kiemTraDuLieuImport
    End With
End Sub

Private Sub checkSheet1()
Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    Dim strCheck As String, iCurrentSheet As Integer
    
    Dim vMST As Variant, vErrMsgMST As Variant, vErrMsgMSTReq As Variant
    Dim vErrMsgMSTTrungCQCT As Variant, vErrMsgMSTTrungNhau As Variant
    Dim vHoVaTen As Variant, vErrMsgHoVaTen As Variant
    Dim vCMT, vCMT1 As Variant
    Dim vErrMsgCMT, vErrMsgCMTTrungNhau As Variant
    
    Dim vTongThuNhap As Variant
    Dim vThuNhapGiamThue As Variant, vErrMsgThuNhapGiamThue As Variant
    Dim vSoThueDuocGiam As Variant, vErrMsgSoThueDuocGiam As Variant
    Dim vSoThueDaKhauTru As Variant, vErrMsgSoThueDaKhauTru As Variant
    Dim vLastRow As Variant
    Dim vMSTCQCT As Variant
    Dim vMST1 As Variant
    Dim iFlagvHoVaTen As Integer, iFlagvMst As Integer, iFocusFlag As Integer, iFlagMsgMSTReq As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String
    Dim i, j As Integer
    Dim blCheck_S1, blCheck_S2, blCheck_S3, blCheck_S4, blCheck_S5, blCheck_S6, blCheck_S7 As Boolean
    Dim blCheck_S8, blCheck_S9 As Boolean, iFlagCMT As Integer
    Dim vErrMsgRow37 As Variant, iFlagErrRow37 As Integer, CountRow As Integer, rowDes As Long
    
    iFlagErrRow37 = 0
    
    With fps
        .EventEnabled(EventAllEvents) = False
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        
        ' Lay thong bao sai ve ngay thanh toan trong sheet header
        .GetText .ColLetterToNumber("E"), 18, vErrMsgHoVaTen
        ' Lay thong bao ve ko nhap ma so thue trong sheet header
        .GetText .ColLetterToNumber("E"), 19, vErrMsgMSTReq
        ' Lay thong bao ve sai ma so thue trong sheet header
        .GetText .ColLetterToNumber("E"), 20, vErrMsgMST
        ' Lay thong bao ve ma so thue bi trung voi co quan chi tra trong sheet header
        .GetText .ColLetterToNumber("E"), 25, vErrMsgMSTTrungCQCT
        ' Lay thong bao ve ma so thue bi trung voi nhau trong sheet header
        .GetText .ColLetterToNumber("E"), 26, vErrMsgMSTTrungNhau
        
        .GetText .ColLetterToNumber("E"), 27, vErrMsgThuNhapGiamThue
        .GetText .ColLetterToNumber("E"), 28, vErrMsgSoThueDuocGiam
        .GetText .ColLetterToNumber("E"), 29, vErrMsgSoThueDaKhauTru
        
        ' Lay thong bao ve sai CMT trong sheet header
        .GetText .ColLetterToNumber("E"), 31, vErrMsgCMT
        ' Lay thong bao ve CMT trung nhau
        .GetText .ColLetterToNumber("E"), 32, vErrMsgCMTTrungNhau
        'lay thong bao CT 12 > CT 10
        .GetText .ColLetterToNumber("E"), 36, vErrMsgRow37
        
        .Sheet = 2 'Check tinh xac thuc cua cac thong tin tren bang ke 02A/TNCN
        
        ' Lay MST cua co quan chi tra
        .GetText .ColLetterToNumber("E"), 9, vMSTCQCT
        

        iFlagCMT = 1
        .Row = 22
        'set lai cac bien check rang buoc
        blCheck_S1 = True
        blCheck_S2 = True
        blCheck_S3 = True
        blCheck_S4 = True
        blCheck_S5 = True
        blCheck_S6 = True
        blCheck_S7 = True
        blCheck_S8 = True
        blCheck_S9 = True
        Do
            'luu so dong hien tai
            j = .Row

            'lay gia tri cac chi tieu
            .GetText .ColLetterToNumber("D"), .Row, vHoVaTen
            .GetText .ColLetterToNumber("E"), .Row, vMST
            .GetText .ColLetterToNumber("F"), .Row, vCMT
            .GetText .ColLetterToNumber("G"), .Row, vTongThuNhap
            .GetText .ColLetterToNumber("H"), .Row, vThuNhapGiamThue
            .GetText .ColLetterToNumber("I"), .Row, vSoThueDuocGiam
            .GetText .ColLetterToNumber("J"), .Row, vSoThueDaKhauTru
            
            If Trim(vHoVaTen) <> vbNullString Or Trim(vMST) <> vbNullString Or Trim(vCMT) <> vbNullString Or vThuNhapGiamThue > 0 Or vSoThueDuocGiam > 0 Or vSoThueDaKhauTru > 0 Or vTongThuNhap > 0 Then
                'Check Ho va ten tren banh ke 02A/TNCN
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
                
                'set not error on vCMT
                .Col = .ColLetterToNumber("F")
                .BackColor = mNonErrorColor
                .CellNote = ""
                'set not error on vMST
                .Col = .ColLetterToNumber("E")
                .BackColor = mNonErrorColor
                .CellNote = ""
                
                If Trim(vMST) = vbNullString And Trim(vCMT) = vbNullString Then
                    .Col = .ColLetterToNumber("E")
                    .BackColor = mErrorColor
                    .CellNote = vErrMsgMSTReq
                    .Col = .ColLetterToNumber("F")
                    .BackColor = mErrorColor
                    .CellNote = vErrMsgMSTReq
                    iFlagMsgMSTReq = 1
                End If
                
                'Check ma so thue tren to khai 02A/TNCN
                If Trim(vMST) <> vbNullString Then
                   ' Set lai co cot CMT (F)
                   .Col = .ColLetterToNumber("F")
                   .CellNote = ""
                   .BackColor = mNonErrorColor
'                   vMST = Left(vMST, 10)
                   strCheck = ""
                   If Len(vMST) = 10 Or Len(vMST) = 13 Or Len(vMST) = 14 Then
                        If Len(Trim(vMST)) = 14 And InStr(1, vMST, "-") = 11 Then
                        
                        ElseIf InStr(1, vMST, "-") > 0 Or Len(Trim(vMST)) = 14 Then
                             .Col = .ColLetterToNumber("E")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgMST   'static
                            .BackColor = mErrorColor
                            iFlagvMst = 1
                            blCheck_S2 = False
                        End If
                        strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                   End If
                   ' Kiem tra dinh dang theo cau truc cua co quan thue
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
                       ' Kiem tra trung voi co quan chi tra
                       vMST = Replace(vMST, " ", "")
                       vMST = Replace(vMST, "-", "")
                       vMSTCQCT = Replace(vMSTCQCT, " ", "")
                       vMSTCQCT = Replace(vMSTCQCT, "-", "")
        
                        If Trim(vMST) = Trim(vMSTCQCT) Then
                            .Col = .ColLetterToNumber("E")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgMSTTrungCQCT    'static
                            .BackColor = mErrorColor
                            iFlagvMst = 1
                            blCheck_S3 = False
    '                        capture the first error cell
                            If iFocusFlag = 0 Then
                                strFocusSheetName = "'" & .SheetName & "'"
                                strFocusCol = str(.Col)
                                strFocusRow = str(.Row)
                                iFocusFlag = 1
                            End If
                        Else
'                           kiem tra MST trung nhau
                            
                            CountRow = 0
                            CountRow = .SearchCol(.ColLetterToNumber("E"), -1, -1, vMST, SearchFlagsNone)
                            rowDes = .SearchCol(.ColLetterToNumber("E"), CountRow, -1, vMST, SearchFlagsNone)
                            If rowDes > -1 Then
                                blCheck_S4 = False
                                .Row = j
                                .Col = .ColLetterToNumber("E")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrMsgMSTTrungNhau     'static
                                .BackColor = mErrorColor

                            End If
                        End If
                    End If
                  
                    'set ve dong hien tai
                   .Row = j
                    
'                    If blCheck_S4 = False Then
'
'                    End If
                    
                End If
                
                If Trim(vCMT) <> vbNullString Then
                    
                    If InStr(1, vCMT, " ", vbTextCompare) <> 0 Then
                        .Col = .ColLetterToNumber("F")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgCMT    'static
                        .BackColor = mErrorColor
                        
                        
                        iFlagvMst = 1
                        blCheck_S8 = False
                        'capture the first error cell
                        If iFocusFlag = 0 Then
                            strFocusSheetName = "'" & .SheetName & "'"
                            strFocusCol = str(.Col)
                            strFocusRow = str(.Row)
                            iFocusFlag = 1
                        End If
                    Else
                        ' Kiem tra cmt trung nhau
                        blCheck_S9 = True
                        CountRow = 0
                        CountRow = .SearchCol(.ColLetterToNumber("F"), -1, -1, vCMT, SearchFlagsNone)
                        rowDes = .SearchCol(.ColLetterToNumber("F"), CountRow, -1, vCMT, SearchFlagsNone)
                        If rowDes > -1 Then
'                            blCheck_S9 = False
                            .Row = j
                            .Col = .ColLetterToNumber("F")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vErrMsgCMTTrungNhau     'static
                            .BackColor = mErrorColor
                            iFlagCMT = 0
                        End If
                        
'                        If blCheck_S9 = False Then
'
'                        End If
                        
                    End If
                End If
                
                'Check thu nhap lam can cu tinh giam thue voi thu nhap chiu thue tren banh ke 02A/TNCN
                If Val(vThuNhapGiamThue) > Val(vTongThuNhap) And Val(vThuNhapGiamThue) > 0 Then
                    .Col = .ColLetterToNumber("H")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgThuNhapGiamThue    'static
                    .BackColor = mErrorColor
                    iFlagvHoVaTen = 1
                    blCheck_S5 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("H")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
'                Check so thue duoc giam voi thu nhap lam can cu tinh giam thue tren banh ke 02A/TNCN
                If Val(vThuNhapGiamThue) <= Val(vSoThueDuocGiam) And Val(vSoThueDuocGiam) > 0 Then
                    .Col = .ColLetterToNumber("I")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgSoThueDuocGiam     'static
                    .BackColor = mErrorColor
                    iFlagvHoVaTen = 1
                    blCheck_S6 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("I")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                'check CT 12 phai < CT 10
'                If Val(vSoThueDuocGiam) >= Val(vTongThuNhap) Then
'                    .Col = .ColLetterToNumber("I")
'                    .CellNote = vErrMsgRow37
'                    .BackColor = mErrorColor
'                    iFlagErrRow37 = 1
'                Else
'                    .Col = .ColLetterToNumber("I")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                End If
                'Check so thue da khau tru voi so thu nhap chiu thue tren banh ke 02A/TNCN
                If Val(vTongThuNhap) <= Val(vSoThueDaKhauTru) And Val(vSoThueDaKhauTru) > 0 Then
                    .Col = .ColLetterToNumber("J")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgSoThueDaKhauTru    'static
                    .BackColor = mErrorColor
                    iFlagvHoVaTen = 1
                    blCheck_S7 = False
                    'capture the first error cell
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        iFocusFlag = 1
                    End If
                Else
                    .Col = .ColLetterToNumber("J")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            End If

            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until UCase(.Text) = "AA"

        ' Trong truong hop blCheck_S1 khong co Ho va ten nao bi loi thi set lai gia tri cua Header la 1
        If blCheck_S1 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 18, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 18, "0"
        End If
        ' Trong truong hop blCheck_S2 khong co MST nao bi loi thi set lai gia tri cua Header la 1
        If blCheck_S2 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 19, "1"
            .SetText .ColLetterToNumber("B"), 20, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 19, "1"
            .SetText .ColLetterToNumber("B"), 20, "0"
            
        End If
        ' Trong truong hop blCheck_S3 khong co MST nao bi trung voi MST cua co quan chi tra thi set lai gia tri cua Header la 1
        If blCheck_S3 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 25, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 25, "0"
        End If
        
        ' Trong truong hop blCheck_S3 khong co MST nao bi trung voi nhau thi set lai gia tri cua Header la 1
        If blCheck_S4 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 26, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 26, "0"
        End If
        
        If blCheck_S5 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 27, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 27, "0"
        End If
        
        If blCheck_S6 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 28, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 28, "0"
        End If
        
        If blCheck_S7 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 29, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 29, "0"
        End If
        
        If blCheck_S8 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 31, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 31, "0"
        End If
        
        If iFlagCMT = 1 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 32, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 32, "0"
        End If
        If iFlagErrRow37 = 1 Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 36, "0"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 36, "1"
        End If
        'set lai header
        'truong hop neu ca MST va CMT deu de trong
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 19, IIf(iFlagMsgMSTReq = 1, "0", "1")
        .Sheet = iCurrentSheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub clearCheckSheet1()
    With fps
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 18, "1"
        .SetText .ColLetterToNumber("B"), 19, "1"
        .SetText .ColLetterToNumber("B"), 20, "1"
        .SetText .ColLetterToNumber("B"), 25, "1"
        .SetText .ColLetterToNumber("B"), 26, "1"
        .SetText .ColLetterToNumber("B"), 27, "1"
        .SetText .ColLetterToNumber("B"), 28, "1"
        .SetText .ColLetterToNumber("B"), 29, "1"
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
            Case "02A"
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
        End Select
    Next
    
    Set xmlCellNode = Nothing
    Set xmlNodeValid = Nothing
End Sub
Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean
'Dim xmlNode As MSXML.IXMLDOMNode
'Dim lRow As Long, lCol As Long
'    With fps
'        For Each xmlNode In TAX_Utilities_New.Data(1).getElementsByTagName("Cell")
'            ParserCellID fps, GetAttribute(xmlNode, "CellID"), lCol, lRow
'            .Col = lCol
'            .Row = lRow
''            If .Col = .ColLetterToNumber("D") And .Row = 22 Then
''                GoTo continue
''            End If
'            .Sheet = 2
'            Select Case .CellType
'                    Case CellTypeCheckBox
'                        fps.Text = vbNullString
'                        UpdateCell fps, lCol, lRow, vbNullString
'                    Case CellTypeComboBox
'                        fps.Text = vbNullString
'                        UpdateCell fps, lCol, lRow, vbNullString
'                    Case CellTypeNumber
'                        fps.value = 0
'                        UpdateCell fps, lCol, lRow, "0"
'                    Case Else
'                        fps.value = vbNullString
'                        UpdateCell fps, lCol, lRow, vbNullString
'            End Select
''continue:
'        Next
'    End With
End Function

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()
    ' Set lai formula cho cot STT
'    With fps
'        If .ActiveSheet = 2 Then
'            .Sheet = 2
'            .Col = .ColLetterToNumber("B")
'            .Row = 22
'            .Formula = "B21+1"
'        End If
'    End With
'Dim isheet As Integer
'Dim i      As Integer
'    With fps
''        If .ActiveSheet = 2 Then
'            .Sheet = 2
'            .Col = .ColLetterToNumber("B")
'            .Row = 22
'            i = 1
'            Do
'              .Text = CStr(i)
'              .Row = 22 + i
'              i = i + 1
'              .Col = .ColLetterToNumber("B")
'            Loop Until UCase(.Text) = "AA"
''        End If
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
               If lCol = .ColLetterToNumber("F") Then
                    varTemp = Trim(varTemp)
               End If
               If lCol = .ColLetterToNumber("C") Then
                    varTemp = Trim(varTemp)
                    If varTemp = "1" Then
                        varTemp = "x"
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
    Dim i As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
            ' insert them node xml
            .Sheet = 2
            InsertNode 23, .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone) - 1
            ' Tinh tong so dong du lieu tren BK
            totalRow2AXS = .SearchCol(.ColLetterToNumber("B"), 22, -1, "aa", SearchFlagsNone) - 22
            UpdateSheets
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub myParseCellID(ByRef mCol As Long, ByRef mRow As Long, cellid As String)
    With fps
        ParserCellID fps, cellid, mCol, mRow
    End With
End Sub


' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 13
    denRowDL = 18
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "12_13_2"
    ' end
End Sub



Public Sub InsertNode(ByVal fRow As Long, ByVal lRow As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCell As MSXML.IXMLDOMNode
    Dim arrTemp() As String
    Dim mSheet As Byte
    
    
    Dim ctlRow As Long
    Dim cTemplate, cReport As String
    
    Dim sumRowDel As Long
    
    mSheet = 2
    
    sumRowDel = TAX_Utilities_New.Data(mSheet - 1).getElementsByTagName("Cells").length
    For ctlRow = 1 To sumRowDel - 1
        Set xmlNodeCells = TAX_Utilities_New.Data(mSheet - 1).nodeFromID("D" & "_" & 22 + ctlRow & "").parentNode
        xmlNodeCells.parentNode.removeChild xmlNodeCells
    Next
    

   Set xmlNodeCells = TAX_Utilities_New.Data(mSheet - 1).nodeFromID("D_22").parentNode
    
    For ctlRow = fRow To lRow
        
        Set xmlNodeNewCells = xmlNodeCells.cloneNode(True)
        
        For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
            cTemplate = Left$(GetAttribute(xmlNodeNewCell, "CellID"), 2)
            arrTemp = Split(GetAttribute(xmlNodeNewCell, "CellID2"), "_")
            'cReport = Left$(GetAttribute(xmlNodeNewCell, "CellID2"), 2)
            cReport = Trim(arrTemp(0)) & "_"
            
            SetAttribute xmlNodeNewCell, "CellID", cTemplate & ctlRow
                    
            ' Set first cell = 1
            SetAttribute xmlNodeNewCell, "FirstCell", "1"
            If mSheet = 2 Then
                If Trim(cTemplate) = "D_" Or Trim(cTemplate) = "E_" Or Trim(cTemplate) = "F_" Then
                    SetAttribute xmlNodeNewCell, "Value", vbNullString
                Else
                    SetAttribute xmlNodeNewCell, "Value", "0"
                End If
                SetAttribute xmlNodeNewCell, "CellID2", cReport & (ctlRow + 17)
            End If
            
        Next
        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, Null
        
    
    Next
'   **********************************
'    added
'   Date: 12/04/06
    'mAdjustData = True
    TAX_Utilities_New.AdjustData(mSheet - 1) = True
'   **********************************
EXIT_SUB:
    Set xmlNodeNewCell = Nothing
    Set xmlNodeNewCells = Nothing
    Set xmlNodeCells = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_02XSTNCN", "InsertNode 02XS_TNCN", Err.number, Err.Description
End Sub
