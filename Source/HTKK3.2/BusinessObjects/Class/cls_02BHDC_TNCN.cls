VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_02BHDC_TNCN"
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
Private totalRow2ABH As Long

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
'        SetDateFormat fps, 1, 34, .ColLetterToNumber("H"), DDMMYYYY
'
'        .Sheet = 1 'To khai 02BH/TNCN
'        .Row = 34
'        .Col = .ColLetterToNumber("H")
'        .Text = Format(Date, "dd/mm/yyyy")
'        .TypeHAlign = TypeHAlignLeft
        'sua loi lech cot (tam thoi)
         SetDateFormat fps, 1, 63, .ColLetterToNumber("M"), DDMMYYYY
        
        .Sheet = 1 'To khai 02BH/TNCN
        .Row = 63
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

            ' Neu la truong hop bo sung thi hien thi cot C trong bang ke 2A (Có dieu chinh so lieu)
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
        .TypeButtonColor = &H8000000F
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
        .TypeButtonColor = &H8000000F
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
        .TypeButtonColor = &H8000000F
        .Font = "Arial"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = GetAttribute(GetMessageCellById("0246"), "Msg")
        
        'Phu luc 02-2/BK-BHDC
        
        .Sheet = 3
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
        .TypeButtonColor = &H8000000F
        .Font = "Arial"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = GetAttribute(GetMessageCellById("0343"), "Msg")
        
        ' Sap xep theo cot thu nhap chiu thue
        .Row = 19
        .Col = .ColLetterToNumber("G")
        .CellType = CellTypeButton
        .TypeButtonAlign = TypeButtonAlignLeft
        .TypeButtonTextColor = vbBlack
        .TypeButtonColor = &H8000000F
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
        .TypeButtonColor = &H8000000F
        .Font = "Arial"
        .FontSize = 8
        .TypeButtonPicture = LoadPicture(GetAbsolutePath("..\Pictures\sort2.bmp"))
        .TypeButtonText = GetAttribute(GetMessageCellById("0246"), "Msg")
        
        .EventEnabled(EventAllEvents) = True
    End With
    
    'check thong tin ma so thue dai ly
    'set vi tri theo thu tu MSTDL~TENNVDLT~ChungChiHN
    viTriMSTDL = "D_22~D_61~D_63"
    updateMSTDL fps, viTriMSTDL
    
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
                            
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("L"), (21 + totalRow2ABH), SortByRow, Y, so
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
                            
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("L"), (21 + totalRow2ABH), SortByRow, Y, so
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
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("L"), (21 + totalRow2ABH), SortByRow, Y, so
                    .Refresh
                    
                    ' Chieu sap xep
                    If so = 1 Then
                        so = 2 ' Sap xep giam dan
                    Else
                        so = 1 ' Sap xep tang dan
                    End If
                End If
            ElseIf .ActiveSheet = 3 Then
               .Sheet = 2
                
                If Col = .ColLetterToNumber("D") And Row = 19 Then
                    Y = Col
                            
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("L"), (21 + totalRow2ABH), SortByRow, Y, so
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
                            
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("L"), (21 + totalRow2ABH), SortByRow, Y, so
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
                    .Sort .ColLetterToNumber("D"), 22, .ColLetterToNumber("L"), (21 + totalRow2ABH), SortByRow, Y, so
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
'                ' Va tinh la lao dong co thue khau tru
'                .Col = .ColLetterToNumber("L")
'                .Text = "1"
'                ' Tinh tong thu nhap cua rieng lao dong co thue khau tru nay
'                .GetText .ColLetterToNumber("F"), .Row, varTemp
'                .Col = .ColLetterToNumber("M")
'                .Text = varTemp
'            Else
'                ' Va tinh la lao dong ko co thue khau tru
'                .Col = .ColLetterToNumber("L")
'                .Text = "0"
'                ' Khong tinh tong thu nhap cua lao dong nay
'                .Col = .ColLetterToNumber("M")
'                .Text = "0"
            End If
            
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        .EventEnabled(EventAllEvents) = True
    End With
    checkSheet1
End Sub


Private Sub fps_EditChange(ByVal Col As Long, ByVal Row As Long)
    Dim a As String
    a = ""
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Integer, iCol As Long, iRow As Long
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim tCol As Long, tRow As Long
    With fps
        .EventEnabled(EventAllEvents) = False
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 2 Then
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 22
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 22
                         totalRow2ABH = i
                         i = i + 1
                    Loop Until .Text = "aa"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    
                    .SetActiveCell iCol, iRow
            ElseIf .ActiveSheet = 3 Then
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 22
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 22
                         totalRow2ABH = i
                         i = i + 1
                    Loop Until .Text = "aa"
                    
                    .SetActiveCell iCol, iRow
                    .Row = iRow
                    .Col = iCol
                    
                    .SetActiveCell iCol, iRow
              End If
        End If
  
       .EventEnabled(EventAllEvents) = True
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
            If Col = .ColLetterToNumber("M") And Row = 63 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    Else
'                        .SetFocus
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
        ElseIf .ActiveSheet = 3 Then
            
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
        tSheet = .Sheet
        'If flgLoadToKhai = True Then Exit Sub

        'get MST cua DTNT
        .Sheet = 1
        .GetText .ColLetterToNumber("D"), 10, MSTDTNT
        MSTDTNT = Trim(MSTDTNT)

        'check thong tin header khong dc trong
        'lay vi tri header tren tk theo thu tu MSTDL, TENDLT, DIACHIDLT,QUAN/HUYEN,TP,SO HDDLT,NGAYDH,Ten NV DLT, CC Hanh Nghe

        viTriHeaderTk = "D_22~D_20~D_24~D_26~M_26~D_32~M_32~D_61~D_63"
        viTriSheetHeader = "B_39~B_40~B_41~B_42~B_43~B_44~B_45~B_46~B_47"
        mangMessage = "E_39~E_40~E_41~E_42~E_43~E_44~E_45~E_46~E_47"
        'check header
        checkErrorHeader fps, viTriHeaderTk, viTriSheetHeader, mangMessage, MSTDTNT
        
        .Sheet = tSheet

        CheckErrorMST
        ' Neu ton tai file data bang ke 02-1/BK-BH thi kiem tra tren bang ke 02-1/BK-BH
        If IIf(TAX_Utilities_v1.NodeValidity.childNodes(1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            checkSheet1
        Else
            clearCheckSheet1
        End If
        
        If IIf(TAX_Utilities_v1.NodeValidity.childNodes(2).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            checkSheet2
        Else
            clearCheckSheet2
        End If
        kiemTraDuLieuImport
    End With
End Sub

Private Sub checkSheet1()
Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    Dim strCheck As String, iCurrentSheet As Integer
    
    Dim vMST As Variant, vErrMsgMST As Variant, vErrMsgMSTReq As Variant
    Dim vErrMsgMSTTrungCQCT As Variant, vErrMsgMSTTrungNhau As Variant
    Dim vHoVaTen As Variant, vErrMsgHoVaTen As Variant
    Dim vTongThuNhap As Variant
    Dim vThuNhapGiamThue As Variant, vErrMsgThuNhapGiamThue As Variant
    Dim vSoThueDuocGiam As Variant, vErrMsgSoThueDuocGiam As Variant, vErrMsgThueDuocGiam2 As Variant
    Dim vSoThueDaKhauTru As Variant, vErrMsgSoThueDaKhauTru As Variant
    Dim vSoCMND As Variant, vSoCMND1 As Variant, vErrMsgSoCMNDTrung As Variant, vErrMsgCheckDauCachCMND As Variant, iFlagCMNDTrung As Integer, iFlagCMNDCoCach As Integer
    Dim vErrMsgEmpty As Variant, iFlagEmty As Integer
    
    Dim vLastRow As Variant
    Dim vMSTCQCT As Variant
    Dim vMST1 As Variant
    
    Dim iFlagvHoVaTen As Integer, iFlagvMst As Integer, iFocusFlag As Integer, iFlagErrCT12CT10 As Integer
    Dim iFlagCheckMSTTrungTemp As Variant, iFlagCheckCMNDTrungTemp As Variant
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String
    Dim i, j As Integer
    Dim blCheck_S1, blCheck_S2, blCheck_S3, blCheck_S4, blCheck_S5, blCheck_S6, blCheck_S7 As Boolean

    iFlagErrCT12CT10 = 0

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
        .GetText .ColLetterToNumber("E"), 34, vErrMsgThueDuocGiam2
        .GetText .ColLetterToNumber("E"), 38, vErrMsgEmpty
        .GetText .ColLetterToNumber("E"), 36, vErrMsgSoCMNDTrung
        .GetText .ColLetterToNumber("E"), 37, vErrMsgCheckDauCachCMND
        
        .Sheet = 2 'Check tinh xac thuc cua cac thong tin tren bang ke 02A/TNCN
        
        ' Lay MST cua co quan chi tra
        .GetText .ColLetterToNumber("E"), 9, vMSTCQCT
        
        blCheck_S1 = True
        blCheck_S2 = True
        blCheck_S3 = True
        blCheck_S4 = True
        blCheck_S5 = True
        blCheck_S6 = True
        blCheck_S7 = True
        
        .Row = 22
        Do
            'luu lai dong hien tai de set lai sau
            j = .Row
            
            .GetText .ColLetterToNumber("D"), .Row, vHoVaTen
            .GetText .ColLetterToNumber("E"), .Row, vMST
            .GetText .ColLetterToNumber("F"), .Row, vSoCMND
            .GetText .ColLetterToNumber("G"), .Row, vTongThuNhap
            .GetText .ColLetterToNumber("H"), .Row, vThuNhapGiamThue
            .GetText .ColLetterToNumber("I"), .Row, vSoThueDuocGiam
            .GetText .ColLetterToNumber("J"), .Row, vSoThueDaKhauTru
            
            If Trim(vHoVaTen) <> vbNullString Or Trim(vMST) <> vbNullString Or vTongThuNhap > 0 Or Trim(vSoCMND) <> vbNullString Or vThuNhapGiamThue <> 0 Or vSoThueDuocGiam <> 0 Or vSoThueDaKhauTru <> 0 Then

                'set mNonErrorColor cho CMDND va MST
                .Col = .ColLetterToNumber("E")
                .BackColor = mNonErrorColor
                .CellNote = ""
                .Col = .ColLetterToNumber("F")
                .BackColor = mNonErrorColor
                .CellNote = ""
                
                'Check Ho va ten tren bang ke 02A/TNCN
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
                
                'check neu MST trong thi phai nhap so CMND
                If Trim(vMST) = vbNullString And Trim(vSoCMND) = vbNullString Then
                    .Col = .ColLetterToNumber("E")
                    .BackColor = mErrorColor
                    .CellNote = vErrMsgEmpty
                    .Col = .ColLetterToNumber("F")
                    .BackColor = mErrorColor
                    .CellNote = vErrMsgEmpty

                    iFlagEmty = 1
                End If

                'check so CMND/Ho chieu khong chua dau cach
                If Trim(vSoCMND) <> vbNullString Then
                    If InStr(1, vSoCMND, " ") > 0 Then
                        .Col = .ColLetterToNumber("F")
                        .BackColor = mErrorColor
                        .CellNote = vErrMsgCheckDauCachCMND

                        iFlagCMNDCoCach = 1
                    End If
                End If
                
                'Check ma so thue tren to khai 02A/TNCN
                If Trim(vMST) <> vbNullString Then
                    '               vMST = Replace(vMST, " ", "")
                    '               vMST = Replace(vMST, "-", "")
                    '                    vMST = Left(vMST, 10)
                    strCheck = ""

                    If Len(vMST) = 10 Or Len(vMST) = 13 Then
                        If InStr(1, vMST, "-") > 0 Then
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

                    End If

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
                    End If
                Else
                    
                    iFlagvMst = 1
                    blCheck_S2 = True

                End If
                
                'check trung
                'set flag de kiem tra neu trung thi thoat vong lap
                iFlagCheckCMNDTrungTemp = 0
                iFlagCheckMSTTrungTemp = 0
                If Trim(vMST) <> vbNullString Or Trim(vSoCMND) <> vbNullString Then

                    ' Kiem tra trung nhau
                    i = 22

                    Do
                        .Row = i

                        If i = j Then
                            GoTo out
                        End If

                        ' lay MST
                        .GetText .ColLetterToNumber("E"), i, vMST1
                        'lay so CMTND/Ho chieu
                        .GetText .ColLetterToNumber("F"), i, vSoCMND1
                        
                        vMST = Replace(vMST, " ", "")
                        vMST = Replace(vMST, "-", "")
                        vMST1 = Replace(vMST1, " ", "")
                        vMST1 = Replace(vMST1, "-", "")
                        
                        'check MST trung
                        If Trim(vMST) <> vbNullString And Trim(vMST) = Trim(vMST1) Then
                            iFlagCheckMSTTrungTemp = 1
                            iFlagvMst = 1
                            blCheck_S4 = False
            
                            'capture the first error cell
                            If iFocusFlag = 0 Then
                                strFocusSheetName = "'" & .SheetName & "'"
                                strFocusCol = str(.Col)
                                strFocusRow = str(.Row)

                                iFocusFlag = 1
                            End If
                        End If

                        If Trim(vSoCMND) <> vbNullString And Trim(UCase(vSoCMND)) = Trim(UCase(vSoCMND1)) Then
                            iFlagCMNDTrung = 1
                            iFlagCheckCMNDTrungTemp = 1
                        End If

                        If Val(iFlagCheckCMNDTrungTemp) = 1 Or Val(iFlagCheckMSTTrungTemp) = 1 Then
                            Exit Do
                        End If

out:
                        i = i + 1
                        .Col = .ColLetterToNumber("B")
                        ' Check cho den het cac dong co du lieu thi thoi
                    Loop Until UCase(.Text) = "AA"

                End If
                
                .Row = j
                'neu trung CMND thi set color va cellnote cho cell
                If iFlagCheckCMNDTrungTemp = 1 Then
                        .Col = .ColLetterToNumber("F")
                        .BackColor = mAlertColor
                        .CellNote = vErrMsgSoCMNDTrung
                End If
                'neu trung MST thi set color va cellnote cho cell
                If iFlagCheckMSTTrungTemp = 1 Then
                        .Col = .ColLetterToNumber("E")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgMSTTrungNhau     'static
                        .BackColor = mErrorColor
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
                'check CT12 < CT10
'                If Val(vSoThueDuocGiam) >= Val(vTongThuNhap) Then
'                    .Col = .ColLetterToNumber("I")
'                    .BackColor = mErrorColor
'                    .CellNote = vErrMsgThueDuocGiam2
'
'                    iFlagErrCT12CT10 = 1
'                Else
'                    .Col = .ColLetterToNumber("I")
'                    .BackColor = mNonErrorColor
'                    .CellNote = ""
'                End If

                'Check so thue da khau tru voi so thu nhap chiu thue tren banh ke 02A/TNCN
                If (vSoThueDaKhauTru) >= Val(vTongThuNhap) And Val(vSoThueDaKhauTru) > 0 Then
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
'loitieptheo:
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
            ' Check cho den het cac dong co du lieu thi thoi
        Loop Until UCase(.Text) = "AA"
        ' Trong truong hop blCheck_S1 khong co Ho va ten nao bi loi thi set lai gia tri cua Header la 1
'loitieptheo1:
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
            .SetText .ColLetterToNumber("B"), 19, "0"
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
        'update lai header
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 34, IIf(iFlagErrCT12CT10 = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 38, IIf(iFlagEmty = 1, "0", "1")
        '.SetText .ColLetterToNumber("B"), 36, IIf(iFlagCMNDTrung = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 37, IIf(iFlagCMNDCoCach = 1, "0", "1")
        .Sheet = iCurrentSheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub


Private Sub checkSheet2()
Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    Dim strCheck As String, iCurrentSheet As Integer
    
    Dim vMST As Variant, vErrMsgMST As Variant, vErrMsgMSTReq As Variant
    Dim vErrMsgMSTTrungCQCT As Variant, vErrMsgMSTTrungNhau As Variant
    Dim vHoVaTen As Variant, vErrMsgHoVaTen As Variant
    Dim vTongThuNhap As Variant
    Dim vThuNhapGiamThue As Variant, vErrMsgThuNhapGiamThue As Variant
    Dim vSoThueDuocGiam As Variant, vErrMsgSoThueDuocGiam As Variant, vErrMsgThueDuocGiam2 As Variant
    Dim vSoThueDaKhauTru As Variant, vErrMsgSoThueDaKhauTru As Variant
    Dim vSoCMND As Variant, vSoCMND1 As Variant, vErrMsgSoCMNDTrung As Variant, vErrMsgCheckDauCachCMND As Variant, iFlagCMNDTrung As Integer, iFlagCMNDCoCach As Integer
    Dim vErrMsgEmpty As Variant, iFlagEmty As Integer
    
    Dim vLastRow As Variant
    Dim vMSTCQCT As Variant
    Dim vMST1 As Variant
    
    Dim iFlagvHoVaTen As Integer, iFlagvMst As Integer, iFocusFlag As Integer, iFlagErrCT12CT10 As Integer
    Dim iFlagCheckMSTTrungTemp As Variant, iFlagCheckCMNDTrungTemp As Variant
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String
    Dim i, j As Integer
    Dim blCheck_S1, blCheck_S2, blCheck_S3, blCheck_S4, blCheck_S5, blCheck_S6, blCheck_S7 As Boolean

    iFlagErrCT12CT10 = 0

    With fps
        .EventEnabled(EventAllEvents) = False
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        
        ' Lay thong bao sai ve ngay thanh toan trong sheet header
        .GetText .ColLetterToNumber("E"), 18, vErrMsgHoVaTen
        ' Lay thong bao ve ko nhap ma so thue trong sheet header
        .GetText .ColLetterToNumber("E"), 50, vErrMsgMSTReq
        ' Lay thong bao ve sai ma so thue trong sheet header
        .GetText .ColLetterToNumber("E"), 51, vErrMsgMST
        ' Lay thong bao ve ma so thue bi trung voi co quan chi tra trong sheet header
        .GetText .ColLetterToNumber("E"), 53, vErrMsgMSTTrungCQCT
        ' Lay thong bao ve ma so thue bi trung voi nhau trong sheet header
        .GetText .ColLetterToNumber("E"), 54, vErrMsgMSTTrungNhau
        
        .GetText .ColLetterToNumber("E"), 55, vErrMsgThuNhapGiamThue
        .GetText .ColLetterToNumber("E"), 56, vErrMsgSoThueDuocGiam
        .GetText .ColLetterToNumber("E"), 57, vErrMsgSoThueDaKhauTru
        .GetText .ColLetterToNumber("E"), 58, vErrMsgThueDuocGiam2
        .GetText .ColLetterToNumber("E"), 61, vErrMsgEmpty
        .GetText .ColLetterToNumber("E"), 59, vErrMsgSoCMNDTrung
        .GetText .ColLetterToNumber("E"), 60, vErrMsgCheckDauCachCMND
        
        .Sheet = 3 'Check tinh xac thuc cua cac thong tin tren bang ke 02-2/BK-DC
        
        ' Lay MST cua co quan chi tra
        .GetText .ColLetterToNumber("E"), 9, vMSTCQCT
        
        blCheck_S1 = True
        blCheck_S2 = True
        blCheck_S3 = True
        blCheck_S4 = True
        blCheck_S5 = True
        blCheck_S6 = True
        blCheck_S7 = True
        
        .Row = 22
        Do
            'luu lai dong hien tai de set lai sau
            j = .Row
            
            .GetText .ColLetterToNumber("D"), .Row, vHoVaTen
            .GetText .ColLetterToNumber("E"), .Row, vMST
            .GetText .ColLetterToNumber("F"), .Row, vSoCMND
            .GetText .ColLetterToNumber("G"), .Row, vTongThuNhap
            .GetText .ColLetterToNumber("H"), .Row, vThuNhapGiamThue
            .GetText .ColLetterToNumber("I"), .Row, vSoThueDuocGiam
            .GetText .ColLetterToNumber("J"), .Row, vSoThueDaKhauTru
            
            If Trim(vHoVaTen) <> vbNullString Or Trim(vMST) <> vbNullString Or vTongThuNhap > 0 Or Trim(vSoCMND) <> vbNullString Or vThuNhapGiamThue <> 0 Or vSoThueDuocGiam <> 0 Or vSoThueDaKhauTru <> 0 Then

                'set mNonErrorColor cho CMDND va MST
                .Col = .ColLetterToNumber("E")
                .BackColor = mNonErrorColor
                .CellNote = ""
                .Col = .ColLetterToNumber("F")
                .BackColor = mNonErrorColor
                .CellNote = ""
                
                'Check Ho va ten tren bang ke 02-2/BK-DC
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
                
                'check neu MST trong thi phai nhap so CMND
                If Trim(vMST) = vbNullString And Trim(vSoCMND) = vbNullString Then
                    .Col = .ColLetterToNumber("E")
                    .BackColor = mErrorColor
                    .CellNote = vErrMsgEmpty
                    .Col = .ColLetterToNumber("F")
                    .BackColor = mErrorColor
                    .CellNote = vErrMsgEmpty

                    iFlagEmty = 1
                End If

                'check so CMND/Ho chieu khong chua dau cach
                If Trim(vSoCMND) <> vbNullString Then
                    If InStr(1, vSoCMND, " ") > 0 Then
                        .Col = .ColLetterToNumber("F")
                        .BackColor = mErrorColor
                        .CellNote = vErrMsgCheckDauCachCMND

                        iFlagCMNDCoCach = 1
                    End If
                End If
                
                'Check ma so thue tren to khai 02A/TNCN
                If Trim(vMST) <> vbNullString Then
                    '               vMST = Replace(vMST, " ", "")
                    '               vMST = Replace(vMST, "-", "")
                    '                    vMST = Left(vMST, 10)
                    strCheck = ""

                    If Len(vMST) = 10 Or Len(vMST) = 13 Then
                        If InStr(1, vMST, "-") > 0 Then
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

                    End If

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
                    End If
                Else
                    
                    iFlagvMst = 1
                    blCheck_S2 = True

                End If
                
                'check trung
                'set flag de kiem tra neu trung thi thoat vong lap
                iFlagCheckCMNDTrungTemp = 0
                iFlagCheckMSTTrungTemp = 0
                If Trim(vMST) <> vbNullString Or Trim(vSoCMND) <> vbNullString Then

                    ' Kiem tra trung nhau
                    i = 22

                    Do
                        .Row = i

                        If i = j Then
                            GoTo out
                        End If

                        ' lay MST
                        .GetText .ColLetterToNumber("E"), i, vMST1
                        'lay so CMTND/Ho chieu
                        .GetText .ColLetterToNumber("F"), i, vSoCMND1
                        
                        vMST = Replace(vMST, " ", "")
                        vMST = Replace(vMST, "-", "")
                        vMST1 = Replace(vMST1, " ", "")
                        vMST1 = Replace(vMST1, "-", "")
                        
                        'check MST trung
                        If Trim(vMST) <> vbNullString And Trim(vMST) = Trim(vMST1) Then
                            iFlagCheckMSTTrungTemp = 1
                            iFlagvMst = 1
                            blCheck_S4 = False
            
                            'capture the first error cell
                            If iFocusFlag = 0 Then
                                strFocusSheetName = "'" & .SheetName & "'"
                                strFocusCol = str(.Col)
                                strFocusRow = str(.Row)

                                iFocusFlag = 1
                            End If
                        End If

                        If Trim(vSoCMND) <> vbNullString And Trim(UCase(vSoCMND)) = Trim(UCase(vSoCMND1)) Then
                            iFlagCMNDTrung = 1
                            iFlagCheckCMNDTrungTemp = 1
                        End If

                        If Val(iFlagCheckCMNDTrungTemp) = 1 Or Val(iFlagCheckMSTTrungTemp) = 1 Then
                            Exit Do
                        End If

out:
                        i = i + 1
                        .Col = .ColLetterToNumber("B")
                        ' Check cho den het cac dong co du lieu thi thoi
                    Loop Until UCase(.Text) = "AA"

                End If
                
                .Row = j
                'neu trung CMND thi set color va cellnote cho cell
                If iFlagCheckCMNDTrungTemp = 1 Then
                        .Col = .ColLetterToNumber("F")
                        .BackColor = mAlertColor
                        .CellNote = vErrMsgSoCMNDTrung
                End If
                'neu trung MST thi set color va cellnote cho cell
                If iFlagCheckMSTTrungTemp = 1 Then
                        .Col = .ColLetterToNumber("E")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = vErrMsgMSTTrungNhau     'static
                        .BackColor = mErrorColor
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
                'check CT12 < CT10
'                If Val(vSoThueDuocGiam) >= Val(vTongThuNhap) Then
'                    .Col = .ColLetterToNumber("I")
'                    .BackColor = mErrorColor
'                    .CellNote = vErrMsgThueDuocGiam2
'
'                    iFlagErrCT12CT10 = 1
'                Else
'                    .Col = .ColLetterToNumber("I")
'                    .BackColor = mNonErrorColor
'                    .CellNote = ""
'                End If

                'Check so thue da khau tru voi so thu nhap chiu thue tren banh ke 02A/TNCN
                If (vSoThueDaKhauTru) >= Val(vTongThuNhap) And Val(vSoThueDaKhauTru) > 0 Then
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
'loitieptheo:
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
            ' Check cho den het cac dong co du lieu thi thoi
        Loop Until UCase(.Text) = "AA"
        ' Trong truong hop blCheck_S1 khong co Ho va ten nao bi loi thi set lai gia tri cua Header la 1
'loitieptheo1:
        If blCheck_S1 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 50, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 50, "0"
        End If
        ' Trong truong hop blCheck_S2 khong co MST nao bi loi thi set lai gia tri cua Header la 1
        If blCheck_S2 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 51, "1"
            .SetText .ColLetterToNumber("B"), 52, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 51, "0"
            .SetText .ColLetterToNumber("B"), 52, "0"
            
        End If
        ' Trong truong hop blCheck_S3 khong co MST nao bi trung voi MST cua co quan chi tra thi set lai gia tri cua Header la 1
        If blCheck_S3 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 53, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 53, "0"
        End If
        
        ' Trong truong hop blCheck_S3 khong co MST nao bi trung voi nhau thi set lai gia tri cua Header la 1
        If blCheck_S4 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 54, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 54, "0"
        End If
        
        If blCheck_S5 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 55, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 55, "0"
        End If
        
        If blCheck_S6 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 56, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 56, "0"
        End If
        
        If blCheck_S7 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 57, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 57, "0"
        End If
        'update lai header
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 58, IIf(iFlagErrCT12CT10 = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 59, IIf(iFlagEmty = 1, "0", "1")
        '.SetText .ColLetterToNumber("B"), 36, IIf(iFlagCMNDTrung = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 60, IIf(iFlagCMNDCoCach = 1, "0", "1")
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


Private Sub clearCheckSheet2()
    With fps
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 50, "1"
        .SetText .ColLetterToNumber("B"), 51, "1"
        .SetText .ColLetterToNumber("B"), 52, "1"
        .SetText .ColLetterToNumber("B"), 53, "1"
        .SetText .ColLetterToNumber("B"), 54, "1"
        .SetText .ColLetterToNumber("B"), 55, "1"
        .SetText .ColLetterToNumber("B"), 56, "1"
        .SetText .ColLetterToNumber("B"), 57, "1"
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
            Case "02-1"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
                    If lCol = fps.ColLetterToNumber("D") Or _
                       lCol = fps.ColLetterToNumber("E") Or lCol = fps.ColLetterToNumber("G") Then
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
           Case "02-2"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_v1.Data(2).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
                    If lCol = fps.ColLetterToNumber("D") Or _
                       lCol = fps.ColLetterToNumber("E") Or lCol = fps.ColLetterToNumber("G") Then
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
Dim xmlNode As MSXML.IXMLDOMNode
Dim lRow As Long, lCol As Long
    With fps
        If .ActiveSheet = 2 Or .ActiveSheet = 3 Then
            For Each xmlNode In TAX_Utilities_v1.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlNode, "CellID"), lCol, lRow
                .Col = lCol
                .Row = lRow
                .Sheet = mCurrentSheet
                Select Case .CellType
                        Case CellTypeCheckBox
                            fps.Text = vbNullString
                            UpdateCell fps, lCol, lRow, vbNullString
                        Case CellTypeComboBox
                            fps.Text = vbNullString
                            UpdateCell fps, lCol, lRow, vbNullString
                        Case CellTypeNumber
                            fps.value = 0
                            UpdateCell fps, lCol, lRow, "0"
                        Case Else
                            fps.value = vbNullString
                            UpdateCell fps, lCol, lRow, vbNullString
                End Select
            Next
        End If
    End With
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
               If lCol = .ColLetterToNumber("F") Then
                    varTemp = Trim(varTemp)
               End If
               If lCol = .ColLetterToNumber("C") Then
                    varTemp = Trim(varTemp)
                    If varTemp = "1" Then
                        varTemp = "x"
                    Else
                        varTemp = ""
                    End If
               End If
               UpdateCell fps, lCol, lRow, varTemp
             Next
        End If
        
        If (GetAttribute(TAX_Utilities_v1.NodeValidity.childNodes(2), "Active") <> "0") Then
            .Sheet = 3
            mCurrentSheet = .Sheet
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(2).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               If lCol = .ColLetterToNumber("F") Then
                    varTemp = Trim(varTemp)
               End If
               If lCol = .ColLetterToNumber("C") Then
                    varTemp = Trim(varTemp)
                    If varTemp = "1" Then
                        varTemp = "x"
                    ElseIf varTemp = "0" Then
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
    Dim i As Integer
    With fps
        .EventEnabled(EventAllEvents) = False
            If .ActiveSheet = 2 Then
                .Col = .ColLetterToNumber("B")
                .Row = 22
                i = 1
                Do
                    .Text = str(i)
                    .Row = i + 22
                    i = i + 1
                Loop Until .Text = "aa"
            ElseIf .ActiveSheet = 3 Then
                .Col = .ColLetterToNumber("B")
                .Row = 22
                i = 1
                Do
                    .Text = str(i)
                    .Row = i + 22
                    i = i + 1
                Loop Until .Text = "aa"
            End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

'Private Sub fps_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
'    Dim sumLaoDong As Variant, sumThuNhap As Variant, sumLDKhauTru As Variant, sumTNKhauTru As Variant
'    Dim sumSoThueTNDaKhauTru As Variant
'    Dim varHoTen As Variant, varMST As Variant
'    Dim varThuNhap As Variant, varSoThueTNDaKhauTru As Variant
'
'    'If flgLoadToKhai = True Then Exit Sub
'
'    If NewSheet = 1 Then
'        kiemTraDuLieuImport
'        With fps
'            'Because of unclearly-rised events, this function is here for turning it off
'            'at the end of this sub, this event will be turned on
'            .EventEnabled(EventAllEvents) = False
'
'            .Sheet = 2 'Bang ke 2A
'
'            .Row = 17
'
'            sumLaoDong = 0
'            sumThuNhap = 0
'            sumLDKhauTru = 0
'            sumTNKhauTru = 0
'            sumSoThueTNDaKhauTru = 0
'
'            ' Lay cac thong tin sau khi da tong hop tren bang ke
'            .GetText .ColLetterToNumber("K"), .Row, sumLaoDong
'            .GetText .ColLetterToNumber("G"), .Row, sumThuNhap
'            .GetText .ColLetterToNumber("L"), .Row, sumLDKhauTru
'            .GetText .ColLetterToNumber("M"), .Row, sumTNKhauTru
'            .GetText .ColLetterToNumber("J"), .Row, sumSoThueTNDaKhauTru
'
'
'             ' Cap nhat vao to quyet toan 02A.
'            .Sheet = 1
'            mCurrentSheet = .Sheet
'            .Col = .ColLetterToNumber("I")
'
'             'Chi tieu [21]
'             .Row = 36
'             .value = sumLaoDong
'             UpdateCell fps, .Col, .Row, .value
'
'             'Chi tieu [22]
'             .Row = 37
'             .value = sumThuNhap
'             UpdateCell fps, .Col, .Row, .value
'
'             'Chi tieu [23]
'             .Row = 38
'             .value = sumLDKhauTru
'             UpdateCell fps, .Col, .Row, .value
'
'             'Chi tieu [24]
''             .Row = 39
''             .value = sumTNKhauTru
''             UpdateCell fps, .Col, .Row, .value
'
'             'Chi tieu [25]
'             .Row = 40
'             .value = sumSoThueTNDaKhauTru
'             UpdateCell fps, .Col, .Row, .value
'
'            .EventEnabled(EventAllEvents) = True
'        End With
'    End If
'End Sub


' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 13
    denRowDL = 18
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "12_13_2~12_13_3"
    ' end
End Sub

