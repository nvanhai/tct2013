VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_08BTNCN"
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
Public flgLoadToKhai As Boolean

' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

Public Sub Prepare1()
    With fps
        SetDateFormat fps, 1, 63, .ColLetterToNumber("M"), DDMMYYYY
        .Sheet = 1 '08B_TNCN
        .Row = 63
        .Col = .ColLetterToNumber("M")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        
'        .Row = .MaxRows - 1
'            .Col = .ColLetterToNumber("M")
'            If .value = "" Then
'                SetDateFormat fps, 1, .MaxRows - 1, .ColLetterToNumber("M"), DDMMYYYY
'                .Text = Format(Date, "dd/mm/yyyy")
'                .TypeHAlign = TypeHAlignLeft
'                UpdateCell fps, .Col, .Row, .Text
'            End If
      End With
End Sub

Public Sub Prepare2()
    Dim a2 As Variant, b2 As Variant, c2 As Variant, oldValue As Variant
    Dim count1 As Long
    Dim iCol As Long, iRow As Long
    Dim xmlNodeCell As MSXML.IXMLDOMNode
    Dim soThangMG As Integer
    Dim viTriMSTDL As String, strTenNVDDL As String, strCCHN As String, xmlNode As MSXML.IXMLDOMNode
    
    With fps
            'tinh so dong` them vao`
            mCurrentSheet = .ActiveSheet
            .Sheet = 1
            .EventEnabled(EventAllEvents) = False
            .Sheet = 1
            mCurrentSheet = .Sheet
        
            If Strloaitk = "TKCT" And StrSolanBosung = "" Then
                .Col = .ColLetterToNumber("C")
                .Row = 38
                .Text = 1
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("I")
                .Row = 38
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                
            ElseIf Strloaitk = "TKCT" And Val(StrSolanBosung) > 0 Then
                        
                .Col = .ColLetterToNumber("C")
                .Row = 38
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("I")
                .Row = 38
                .Text = StrSolanBosung
                UpdateCell fps, .Col, .Row, .Text
            End If
            
             ' set ky ke khai tu thang den thang
            .Row = 38
            .Col = .ColLetterToNumber("K")
            .Text = TAX_Utilities_v2.FirstDay
            UpdateCell fps, .Col, .Row, .Text
            
            .Row = 38
            .Col = .ColLetterToNumber("L")
            .Text = TAX_Utilities_v2.LastDay
            UpdateCell fps, .Col, .Row, .Text
            
            
            
            
         'set thang mien giam thue cho ky ke khai nam 2011
         If TAX_Utilities_v2.Year = "2011" Then
            soThangMG = 7
            .SetText .ColLetterToNumber("AB"), 6, soThangMG
         ElseIf TAX_Utilities_v2.Year = "2009" Then
            soThangMG = 6
            .SetText .ColLetterToNumber("AB"), 6, soThangMG
         End If
         
         'set thong tin nguoi ky, ngay ky, thong tin nhan vien dai ly thue
         Set xmlNodeCell = TAX_Utilities_v2.Data(0).documentElement.childNodes(TAX_Utilities_v2.Data(0).childNodes.length)
         Set xmlNodeCell = xmlNodeCell.firstChild
         Set xmlNodeCell = xmlNodeCell.firstChild
         ParserCellID fps, GetAttribute(xmlNodeCell, "CellID"), iCol, iRow
         .Row = iRow
         .Col = .Col
         .Formula = ""
         Set xmlNodeCell = xmlNodeCell.nextSibling
         ParserCellID fps, GetAttribute(xmlNodeCell, "CellID"), iCol, iRow
         .Row = iRow
         .Col = .Col
         .Formula = ""

         Set xmlNodeCell = xmlNodeCell.nextSibling
         ParserCellID fps, GetAttribute(xmlNodeCell, "CellID"), iCol, iRow
         .Row = iRow
         .Col = .Col
         .Formula = ""

         Set xmlNodeCell = xmlNodeCell.nextSibling
         ParserCellID fps, GetAttribute(xmlNodeCell, "CellID"), iCol, iRow
         .Row = iRow
         .Col = .Col
         .Formula = ""
        .EventEnabled(EventAllEvents) = True
        
        'check thong tin ma so thue dai ly
        'set vi tri theo thu tu MSTDL~TENNVDLT~ChungChiHN
        'lay vi tri cell ten nhan vien dai ly thue
        Set xmlNode = TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").Item(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length - 6)
        strTenNVDDL = GetAttribute(xmlNode, "CellID")
        'lay vi tri cell chung chi hanh nghe so
        Set xmlNode = TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").Item(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length - 5)
        strCCHN = GetAttribute(xmlNode, "CellID")
        
        
        For Each xmlNode In TAX_Utilities_v2.Data(0).childNodes(2).childNodes(2).childNodes
            ParserCellID fps, GetAttribute(xmlNode.childNodes(7), "CellID"), iCol, iRow
            
            'set lai gia tri ct 41
            If iCol = .ColLetterToNumber("T") Then
                oldValue = GetAttribute(xmlNode.childNodes(7), "Value")
                .Col = iCol
                .Row = iRow
                .Text = oldValue
'                .Formula = ""
            End If
            
            'set lai gia tri ct 42
            ParserCellID fps, GetAttribute(xmlNode.childNodes(8), "CellID"), iCol, iRow

            If iCol = .ColLetterToNumber("W") Then
                oldValue = GetAttribute(xmlNode.childNodes(8), "Value")
                .Col = iCol
                .Row = iRow
                .Text = oldValue
'                .Formula = ""
            End If
            
        Next
        
        viTriMSTDL = "D_22~" & Trim(strTenNVDDL) & "~" & Trim(strCCHN)
        updateMSTDL fps, viTriMSTDL
        
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
'
'            If .ActiveSheet = 1 Then
'                   fps.EventEnabled(EventAllEvents) = False
'                    i = 1
'                    .Col = .ColLetterToNumber("B")
'                    .Row = 56
'                    Do
'                         .Text = str(i)
'                         UpdateCell fps, .Col, .Row, .Text
'                         .GetText .ColLetterToNumber("B"), .Row, strtext
'
'                             If i = 1 Then
'
'                                .Col = .ColLetterToNumber("J")
'                                .Row = 56
'                                .value = .Text
'                                UpdateCell fps, .Col, .Row, .value
'                                lValue = .value
'                                .Col = .ColLetterToNumber("R")
'                                .Row = 50
'                                UpdateCell fps, .Col, .Row, .value
'                                lrow43 = .value
'                                .Col = .ColLetterToNumber("L")
'                                .Row = 56
'                                If lrow43 <> 0 And lValue <> 0 Then
'                                .value = lrow43 * lValue / 100
'                                Else
'                                .value = 0
'                                End If
'                                 UpdateCell fps, .Col, .Row, .value
''                                .Col = .ColLetterToNumber("P")
''                                .Row = 56
''
''                                UpdateCell fps, .Col, .Row, .Value
''                                .Col = .ColLetterToNumber("R")
''                                .Row = 56
''
''                                UpdateCell fps, .Col, .Row, .Value
''                                .Col = .ColLetterToNumber("W")
''                                .Row = 50
''
''                                UpdateCell fps, .Col, .Row, .Value
''
''
'
'                           Else
'
'                               .Col = .ColLetterToNumber("J")
'                                .Row = 56 + i - 1
'                                .value = .Text
'                                UpdateCell fps, .Col, .Row, .value
'                                lValue = .value
'                                .Col = .ColLetterToNumber("R")
'                                .Row = 50
'                                UpdateCell fps, .Col, .Row, .value
'                                lrow43 = .value
'                                .Col = .ColLetterToNumber("L")
'                                .Row = 56 + i - 1
'                                If lrow43 <> 0 And lValue <> 0 Then
'                                .value = lrow43 * lValue / 100
'                                Else
'                                .value = 0
'                                End If
'                                 UpdateCell fps, .Col, .Row, .value
''                                .Col = .ColLetterToNumber("P")
''                                .Row = 56 + i - 1
''
''                                UpdateCell fps, .Col, .Row, .Value
''                                .Col = .ColLetterToNumber("R")
''                                .Row = 56 + i - 1
''
''                                UpdateCell fps, .Col, .Row, .Value
''                                .Col = .ColLetterToNumber("W")
''                                .Row = 56 + i - 1
''
''                                UpdateCell fps, .Col, .Row, .Value
'                           End If
'
'
'                         .Col = .ColLetterToNumber("B")
'                         .Row = i + 56
'                         i = i + 1
'                    Loop Until .Text = "aa"
'                            .Col = .ColLetterToNumber("J")
'                           .Row = 56 + i
'                           UpdateCell fps, .Col, .Row, .value
'                           .Col = .ColLetterToNumber("L")
'                           .Row = 56 + i
'                           UpdateCell fps, .Col, .Row, .value
'                           .Col = .ColLetterToNumber("N")
'                           .Row = 56 + i
'                           UpdateCell fps, .Col, .Row, .value
'                           .Col = .ColLetterToNumber("P")
'                           .Row = 56 + i
'                           UpdateCell fps, .Col, .Row, .value
'                           .Col = .ColLetterToNumber("R")
'                           .Row = 56 + i
'                           UpdateCell fps, .Col, .Row, .value
'                           .Col = .ColLetterToNumber("T")
'                           .Row = 56 + i
'                           UpdateCell fps, .Col, .Row, .value
'                           .Col = .ColLetterToNumber("W")
'                           .Row = 56 + i
'                           UpdateCell fps, .Col, .Row, .value
'                          .Col = .ColLetterToNumber("Z")
'                          .Row = 56 + i
'                          UpdateCell fps, .Col, .Row, .value
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
End Sub

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
                    .Row = 57
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 57
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
    Dim varCT11 As Variant
    Dim varCT12 As Variant
    Dim count1 As Long
    Dim CellNgayKy As String
    Dim strValue As String
    Dim varThangKT As Variant
'    CellNgayKy = GetAttribute(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length - 6), "CellID")
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            ' check ngay ky
            If Col = .ColLetterToNumber("M") And Row = .MaxRows - 1 Then
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
      
'            CheckDynamicError
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
    Dim colC, colH, colU, TiLeTNCT, errorTyLe, TongTyLe As Variant
    Dim vErrMsgMST As Variant
   
    Dim iFlagvName As Integer, iFlagvDate As Integer, iFlagvMst As Integer, iFlagvThang As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer
    Dim strFormula As String
    Dim vKhoanGT, vThueTNCN, vThueTNCNgiam, vNgayKy As Variant
    Dim errEmpty As Variant, iFlagEmpty As Integer, iFlagTyLe As Integer, checkExistData As Boolean
    Dim countRow As Long, rowDes As Long, iFlagMSTTrung As Integer, mErrMSTTrung As Variant
    Dim mErrMSTTonTai As Variant, vCT5MST As Variant, iFlagMSTTonTai As Integer, rowMST As Integer
    
    'chuoi vitriheader tren to khai, chuoi vi tri tren sheet header de check lai
    Dim viTriHeaderTk As String, viTriSheetHeader As String, mangMessage As String, xmlNode As MSXML.IXMLDOMNode, strTenNVDDL, strCCHN As String
    
    iFlagvName = 0
    iFlagvDate = 0
    iFlagvMst = 0
    iFlagvThang = 0
    iFocusFlag = 0
    iFlagEmpty = 0
    iFlagTyLe = 0
    iFlagMSTTrung = 0
    
    Dim iWhiteRow As Integer
    With fps

        CheckErrorMST
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 15, errorName
        .GetText .ColLetterToNumber("E"), 16, vErrMsgMST
        .GetText .ColLetterToNumber("E"), 17, errorDate
        .GetText .ColLetterToNumber("E"), 19, error1
        .GetText .ColLetterToNumber("E"), 20, errorThang
        .GetText .ColLetterToNumber("E"), 12, errorPhuongXa
        .GetText .ColLetterToNumber("E"), 21, errorTyLe
        .GetText .ColLetterToNumber("E"), 23, errEmpty
        .GetText .ColLetterToNumber("E"), 24, mErrMSTTrung
        .GetText .ColLetterToNumber("E"), 25, mErrMSTTonTai
        .Sheet = 1
     
        'getCT5
        .GetText .ColLetterToNumber("D"), 10, vCT5MST
    
        'check thong tin header khong dc trong
        'lay vi tri header tren tk theo thu tu MSTDL, TENDLT, DIACHIDLT,QUAN/HUYEN,TP,SO HDDLT,NGAYDH,Ten NV DLT, CC Hanh Nghe
        'lay vi tri cell ten nhan vien dai ly thue
        Set xmlNode = TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").Item(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length - 6)
        strTenNVDDL = GetAttribute(xmlNode, "CellID")
        'lay vi tri cell chung chi hanh nghe so
        Set xmlNode = TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").Item(TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length - 5)
        strCCHN = GetAttribute(xmlNode, "CellID")
        viTriHeaderTk = "D_22~D_20~D_24~D_26~M_26~D_32~M_32~" & Trim(strTenNVDDL) & "~" & Trim(strCCHN)
        viTriSheetHeader = "B_26~B_27~B_28~B_29~B_30~B_31~B_32~B_33~B_34"
        mangMessage = "E_26~E_27~E_28~E_29~E_30~E_31~E_32~E_33~E_34"
        'check header
        checkErrorHeader fps, viTriHeaderTk, viTriSheetHeader, mangMessage, vCT5MST
        
        .Sheet = 1
        iFlagMSTTonTai = 0
        'check Null
                
        .Row = 57
        Do
            .GetText .ColLetterToNumber("C"), .Row, colC
            .GetText .ColLetterToNumber("G"), .Row, vMST
            .GetText .ColLetterToNumber("K"), .Row, TiLeTNCT
            .GetText .ColLetterToNumber("P"), .Row, vKhoanGT
            .GetText .ColLetterToNumber("T"), .Row, vThueTNCN
            .GetText .ColLetterToNumber("W"), .Row, vThueTNCNgiam
            
            If colC <> vbNullString Or vMST <> vbNullString Or TiLeTNCT <> 0 Or vKhoanGT <> 0 Or vThueTNCN <> 0 Or vThueTNCNgiam <> 0 Then
                 'tim MST cua DNTN co trong danh sach chi tiet ke khai khong
                 rowMST = .SearchCol(.ColLetterToNumber("G"), 56, -1, vCT5MST, SearchFlagsNone)
                 If rowMST = -1 Then
                    iFlagMSTTonTai = 1
                 End If
                 
                  If colC = vbNullString Then
                    .Col = .ColLetterToNumber("C")
                    .BackColor = mErrorColor
                    .CellNote = errorName
                    iFlagvName = 1
                  Else
                    .Col = .ColLetterToNumber("C")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                  End If
                  If vMST = vbNullString Then
                    .Col = .ColLetterToNumber("G")
                    .BackColor = mErrorColor
                    .CellNote = errEmpty
                    iFlagEmpty = 1
                  ElseIf Trim(vMST) <> vbNullString Then
        '                vMST = Replace(vMST, " ", "")
        '                vMST = Replace(vMST, "-", "")
                         vMST = Left(vMST, 10)
                       strCheck = ""
                       If Len(vMST) = 10 Or Len(vMST) = 13 Then
                            strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                       End If
                       If Mid(strCheck, 2, 1) <> "0" Then
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
                        Else
                        .Col = .ColLetterToNumber("G")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        End If
                        
                        'check MST trung nhau
                         countRow = .SearchCol(.ColLetterToNumber("G"), -1, -1, vMST, SearchFlagsNone)
                         rowDes = .SearchCol(.ColLetterToNumber("G"), countRow, -1, vMST, SearchFlagsNone)
                        If rowDes > -1 Then
                            .Col = .ColLetterToNumber("G")
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = mErrMSTTrung      'static
                            .BackColor = mErrorColor
                            iFlagMSTTrung = 1
                        End If
                End If
                checkExistData = True
            End If
             count1 = count1 + 1
            .Row = count1 + 57
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        ' ThuyNT: Kiem tra tong ty le phai bang 100% 29/4/2010
        .GetText .ColLetterToNumber("K"), .Row + 1, TongTyLe
        If checkExistData Then
            If Round(CDbl(TongTyLe), 2) <> 100 Then
                .Col = .ColLetterToNumber("K")
                .Row = .Row + 1
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                .CellNote = errorTyLe  'static
                .BackColor = mErrorColor
                iFlagTyLe = 1
            Else
                .Col = .ColLetterToNumber("K")
                .Row = .Row + 1
                .BackColor = 16709097
                .CellNote = ""
            End If
        End If
        'check ngay ky khong dc de trong
        .GetText .ColLetterToNumber("M"), .MaxRows - 1, vNgayKy
        If Trim(vNgayKy) = vbNullString Then
            .Col = .ColLetterToNumber("M")
            .Row = .MaxRows - 1
            .BackColor = mErrorColor
            .CellNote = error1
            iFlagvDate = 1
        Else
            .Col = .ColLetterToNumber("M")
            .Row = .MaxRows - 1
            .BackColor = mNonErrorColor
            .CellNote = ""
        End If
        'set neu MST CT5 khong ton tai trong danh sach chi tiet phan bo thi set color error
        If iFlagMSTTonTai = 1 Then
            .Col = .ColLetterToNumber("D")
            .Row = 10
            .CellNote = mErrMSTTonTai
            .BackColor = mErrorColor
        Else
            .Col = .ColLetterToNumber("D")
            .Row = 10
            .CellNote = ""
            .BackColor = -2147483633
        End If
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 15, IIf(iFlagvName = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 16, IIf(iFlagvMst = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 21, IIf(iFlagTyLe = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 23, IIf(iFlagEmpty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 19, IIf(iFlagvDate = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 24, IIf(iFlagMSTTrung = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 25, IIf(iFlagMSTTonTai = 1, "0", "1")
         If iFocusFlag = 1 Then
            ' setcell for focus
            strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
            .Col = .ColLetterToNumber("C")
            .Row = 15
            .Formula = strFormula
'            strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
            .Col = .ColLetterToNumber("H")
            .Row = 16
            .Formula = strFormula
        End If
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
                        If Not TAX_Utilities_v2.Data(mCurrentSheet - 1).nodeFromID( _
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
                        If Not TAX_Utilities_v2.Data(mCurrentSheet - 1).nodeFromID(GetCellID(fpSpread1, i, pRow - lRowCtrl)) Is Nothing Then
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
        For Each xmlCellNodeData In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
           ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
'               If (lRow = 5 And lCol = .ColLetterToNumber("R")) Or (lRow = 7 And lCol = .ColLetterToNumber("R")) Then
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
Public Sub FinishImport()

End Sub


' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 13
    denRowDL = 18
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "2_3_2"
    ' end
End Sub


