VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01NTNN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrTKThang_PS As String
Public StrSolanBosung As String

Public loaiKHBS As String

Private Const Row_MSTDN = 21
Private Const Col_MSTDN = "Y"

' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

Public FlagThaydoi As String

  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    With fps
        SetDateFormat fps, 1, 69, .ColLetterToNumber("AX"), DDMMYYYY
        SetDateFormat fps, 1, 55, .ColLetterToNumber("AD"), DDMMYYYY
        
        
'        SetDateFormat fps, 2, 23, .ColLetterToNumber("BG"), DDMMYYYY
        SetDateFormat fps, 2, 46, .ColLetterToNumber("BG"), DDMMYYYY
        SetDateFormat fps, 2, 31, .ColLetterToNumber("BE"), DDMMYYYY

        
        .Sheet = 1 'To khai 02/TNCN
        .Row = 69
        .Col = .ColLetterToNumber("AX")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
'        .Sheet = 2 'To khai NTNN
'        .Row = 23
'        .Col = .ColLetterToNumber("BG")
'        .Text = Format(Date, "dd/mm/yyyy")
'        .TypeHAlign = TypeHAlignLeft
        ' set CQT ra quyet dinh hoan
        setCQTQuanLyHoanThue fps
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
Dim ngay As Variant
Dim kythang As Variant
Dim kyngay As Variant
'lay gia tri tong cua chi tieu 14
Dim tongNTNN As Variant, tienDoiThanhChu As String
Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
Dim xmlNode As MSXML.IXMLDOMNode
Dim iCol As Long, iRow As Long

Dim vChitieu10 As Variant
Dim i As Integer

    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        ' vttoan: check to khai thang hoac lan phat sinh
        ngay = TAX_Utilities_v2.Day
        .SetText .ColLetterToNumber("O"), 32, ngay
        .GetText .ColLetterToNumber("W"), 31, kythang
        .GetText .ColLetterToNumber("AS"), 31, kyngay
        .GetText .ColLetterToNumber("BF"), 33, tongNTNN
         'xu ly ngay phat sinh
        If Trim(ngay) = "" Then
        .SetText .ColLetterToNumber("BM"), 31, ""
        UpdateCell fps, .ColLetterToNumber("BM"), 31, ""
        Else
        .SetText .ColLetterToNumber("BM"), 31, ngay & "/" & TAX_Utilities_v2.Month & "/" & TAX_Utilities_v2.Year
        UpdateCell fps, .ColLetterToNumber("BM"), 31, ngay & "/" & TAX_Utilities_v2.Month & "/" & TAX_Utilities_v2.Year
        End If
        'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
        If StrTKThang_PS = "TK_THANG" Then
            .Col = .ColLetterToNumber("L")
            .Row = 31
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .SetText .ColLetterToNumber("P"), 5, kythang
            .Col = .ColLetterToNumber("M")
            .Row = 31
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf StrTKThang_PS = "TK_LANPS" Then
            .Col = .ColLetterToNumber("M")
            .Row = 31
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .SetText .ColLetterToNumber("P"), 5, kyngay
            .Col = .ColLetterToNumber("L")
            .Row = 31
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            
            
        End If
      
        'end
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("C")
            .Row = 31
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 31
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 31
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("C")
            .Row = 31
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 31
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 31
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
       
            
        If str(tongNTNN) <> "" Then
            tongNTNN = Replace(tongNTNN, ".", "")
            tongNTNN = Replace(tongNTNN, ",", "")
            tienDoiThanhChu = doiTien(CStr(tongNTNN))
            .SetText .ColLetterToNumber("S"), 33, tienDoiThanhChu
            UpdateCell fps, .ColLetterToNumber("S"), 33, tienDoiThanhChu
        End If
        
            'Xoa cong thuc cua o Chu ky
            Set xmlCellNode = TAX_Utilities_v2.Data(0).nodeFromID(GetCellID(fps, .ColLetterToNumber("L"), 55))
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode.lastChild
            Set xmlCellNode = xmlCellNode.firstChild
            Set xmlCellNode = xmlCellNode.firstChild
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            .Row = iRow
            .Col = iCol
            .Formula = ""
            ' Ho va ten nhan vien dai ly
            Set xmlCellNode = xmlCellNode.nextSibling
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            .Row = iRow
            .Col = iCol
            .Formula = ""
            ' So chung chi hanh nghe
            Set xmlCellNode = xmlCellNode.nextSibling
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            .Row = iRow
            .Col = iCol
            .Formula = ""
            
            ' Set nut tai BK
            .Sheet = 1
            .ActiveSheet = .Sheet
            .Row = 3
            .Col = .ColLetterToNumber("D")
            .CellType = CellTypeButton
            .TypeButtonTextColor = RGB(0, 0, 255)
            .TypeButtonColor = &H8000000F
            .TypeButtonAlign = TypeButtonAlignRight
            .Font = "Tahoma"
            .FontSize = 8
            .TypeButtonText = GetAttribute(GetMessageCellById("0104"), "Msg")
            
            ' set lock chi tieu 12
            .Sheet = 1
            .Col = .ColLetterToNumber("B")
            .Row = 55
            i = 1
            Do
                .GetText .ColLetterToNumber("AY"), .Row, vChitieu10
                If Val(vChitieu10) < 0 Then
                    .Col = .ColLetterToNumber("BG")
                    .Lock = True
                End If
                
                .Row = 55 + i
                i = i + 1
                .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    Dim varTemp
    Dim rowDT As Long
    Dim iRow As Long, iCol As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
    Dim xmlNode As MSXML.IXMLDOMNode
    
    Dim tongNTNN As Variant, tienDoiThanhChu As String
    
    With fps
        .EventEnabled(EventAllEvents) = False
            If .ActiveSheet = 1 Then
                rowDT = TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length - 14
                Set xmlCellNode = TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").Item(rowDT)
                ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
                
                If (Col = .ColLetterToNumber("BQ")) And Row = iRow Then
                    .GetText .ColLetterToNumber("BQ"), Row, tongNTNN
                    tongNTNN = Replace(tongNTNN, ".", "")
                    tongNTNN = Replace(tongNTNN, ",", "")
                    tienDoiThanhChu = doiTien(CStr(tongNTNN))
                    .SetText .ColLetterToNumber("S"), 33, tienDoiThanhChu
                    UpdateCell fps, .ColLetterToNumber("S"), 33, tienDoiThanhChu
                End If
            End If
            ' Set lai chi tieu 12 =0 neu chi tieu 10 < 0
            If Col = .ColLetterToNumber("AY") And Row >= 55 And Row <= .MaxRows - 17 Then
                .GetText Col, Row, varTemp
                If IsNumeric(varTemp) Then
                    If Val(varTemp) < 0 Then
                        .Col = .ColLetterToNumber("BG")
                        .Row = Row
                        .Text = "0"
                        .Lock = True
                        UpdateCell fps, Col, Row, .Text
                    Else
                        .Col = .ColLetterToNumber("BG")
                        .Row = Row
                        .Lock = False
                    End If
                End If
            End If
            ' Set lai chi tieu 12 =0 neu chi tieu 10 < 0
            If Col = .ColLetterToNumber("BG") And Row >= 55 And Row <= .MaxRows - 17 Then
                .GetText .ColLetterToNumber("AY"), Row, varTemp
                If IsNumeric(varTemp) Then
                    If Val(varTemp) < 0 Then
                        .Col = .ColLetterToNumber("BG")
                        .Row = Row
                        .Text = "0"
                        .Lock = False
                        UpdateCell fps, Col, Row, .Text
                    Else
                        .Col = .ColLetterToNumber("BG")
                        .Row = Row
                        .Lock = False
                    End If
                End If
            End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Integer, iCol As Long, iRow As Long
    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 1 Then
                   fps.EventEnabled(EventAllEvents) = True
                    i = 1
                    .Col = .ColLetterToNumber("B")
                    .Row = 55
                    Do
                         .Text = str(i)
                         .Col = .ColLetterToNumber("B")
                         .Row = i + 55
                         i = i + 1
                    Loop Until .Text = "aa"
                    .SetActiveCell iCol, iRow

                    .Row = iRow
                    .Col = iCol
                     fps.EventEnabled(EventAllEvents) = False
                     .SetActiveCell iCol, iRow
            End If
         End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    'lay gia tri tong cua chi tieu 14
    Dim tongNTNN As Variant, tienDoiThanhChu As String
    
    Dim rowDT As Long
    Dim iRow As Long, iCol As Long
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
    Dim xmlNode As MSXML.IXMLDOMNode
    
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            rowDT = TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length - 16
            Set xmlCellNode = TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").Item(rowDT)
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            
            If (Col = .ColLetterToNumber("AU") Or Col = .ColLetterToNumber("BM")) And Row = iRow Then
                
                .GetText .ColLetterToNumber("BQ"), Row, tongNTNN
                tongNTNN = Replace(tongNTNN, ".", "")
                tongNTNN = Replace(tongNTNN, ",", "")
                tienDoiThanhChu = doiTien(CStr(tongNTNN))
                .SetText .ColLetterToNumber("S"), 33, tienDoiThanhChu
                UpdateCell fps, .ColLetterToNumber("S"), 33, tienDoiThanhChu
            End If
        'check date
            If Col = .ColLetterToNumber("AX") And Row = .MaxRows - 3 Then
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
            If Col = .ColLetterToNumber("AD") Then
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
            
            ' Set lai chi tieu 12 =0 neu chi tieu 10 < 0
            If Col = .ColLetterToNumber("AY") And Row >= 55 And Row <= .MaxRows - 17 Then
                .GetText Col, Row, varTemp
                If IsNumeric(varTemp) Then
                    If Val(varTemp) < 0 Then
                        .Col = .ColLetterToNumber("BG")
                        .Row = Row
                        .Text = "0"
                        .Lock = True
                        UpdateCell fps, Col, Row, .Text
                    Else
                        .Col = .ColLetterToNumber("BG")
                        .Row = Row
                        .Lock = False
                    End If
                End If
            End If
            ' Set lai chi tieu 12 =0 neu chi tieu 10 < 0
            If Col = .ColLetterToNumber("BG") And Row >= 55 And Row <= .MaxRows - 17 Then
                .GetText .ColLetterToNumber("AY"), Row, varTemp
                If IsNumeric(varTemp) Then
                    If Val(varTemp) < 0 Then
                        .Col = .ColLetterToNumber("BG")
                        .Row = Row
                        .Text = "0"
                        .Lock = True
                        UpdateCell fps, Col, Row, .Text
                    Else
                        .Col = .ColLetterToNumber("BG")
                        .Row = Row
                        .Lock = False
                    End If
                End If
            End If
            
            
        End If
        .EventEnabled(EventAllEvents) = True
    End With
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

    Dim iFlagNgayTT As Integer

    Dim MSTDN, strNoiDung As Variant
    Dim iFlagEmty, iFlagEmtyND As Variant
    Dim errMST    As Variant
    Dim errEmty   As Variant
    Dim errCt08   As Variant
    Dim errNgayTT As Variant
    Dim ct07      As Variant
    Dim ct08      As Variant
    Dim ct11      As Variant
    Dim i         As Variant
    Dim iWhiteRow As Boolean
    Dim strNgayHD, strSoTienTT, strNgayTT, errNgayTT2, errThueSuatGTGT, strDoanhThu, strTyLeGT, strThueSuat, strDoanhThuTT, strTyLeTNDN, strSoThueMG As Variant
    Dim vMSTDN, vMST, strCheckMST, thueSuatGTGT As Variant
    
    Dim iFlagCheckMaST, iFlagCheckMaSTDN, iFlagCheckMaSTDN10 As Integer
    Dim errMSTDN As Variant
    
    Dim ct13, ct14 As Variant, iFlagCheckCT13CT14 As Integer, errCheckCT13CT14 As Variant
    
    Dim tongCt9, tongCt13 As Variant
    Dim vErrorTongCt9, vErrorTongCt13 As Variant
    Dim iFlagCheckTongCt9 As Integer, iFlagCheckTongCt13 As Integer
    Dim isCheckTong9 As Boolean, isCheckTong13 As Boolean
    Dim countRowDyn As Long
    Dim ct9 As Variant

    With fps
          
        'vttoan
        ' bat cau truc mst
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 12, errMST
        .GetText .ColLetterToNumber("E"), 13, errEmty
        .GetText .ColLetterToNumber("E"), 14, errCt08
        .GetText .ColLetterToNumber("E"), 15, errNgayTT
        .GetText .ColLetterToNumber("E"), 20, errNgayTT2
        .GetText .ColLetterToNumber("E"), 17, errMSTDN
        
        .GetText .ColLetterToNumber("E"), 19, errMST
        .GetText .ColLetterToNumber("E"), 21, errThueSuatGTGT
        .GetText .ColLetterToNumber("E"), 23, errCheckCT13CT14
        
        .GetText .ColLetterToNumber("E"), 24, vErrorTongCt9
        .GetText .ColLetterToNumber("E"), 25, vErrorTongCt13
        
        Dim kythang As Variant
        Dim nam     As Variant
        kythang = TAX_Utilities_v2.Month
        nam = TAX_Utilities_v2.Year
        
        .Sheet = 1
        .GetText .ColLetterToNumber("Y"), 13, vMSTDN

        iFlagNgayTT = 0
        iFlagCheckMaST = 0
        iFlagCheckMaSTDN = 0
        iFlagCheckMaSTDN10 = 0
        
        .Col = .ColLetterToNumber("B")
        .Row = 55
        i = 1

        Do
            .GetText .ColLetterToNumber("C"), .Row, strNoiDung
            .GetText .ColLetterToNumber("R"), .Row, strNgayHD
            .GetText .ColLetterToNumber("X"), .Row, strSoTienTT
            .GetText .ColLetterToNumber("AD"), .Row, strNgayTT
            .GetText .ColLetterToNumber("AI"), .Row, strDoanhThu
            .GetText .ColLetterToNumber("AM"), .Row, strTyLeGT
            .GetText .ColLetterToNumber("AQ"), .Row, strThueSuat
            .GetText .ColLetterToNumber("AY"), .Row, strDoanhThuTT
            .GetText .ColLetterToNumber("BC"), .Row, strTyLeTNDN
            .GetText .ColLetterToNumber("BG"), .Row, strSoThueMG
            .GetText .ColLetterToNumber("L"), .Row, MSTDN
            .GetText .ColLetterToNumber("AQ"), .Row, thueSuatGTGT
            .GetText .ColLetterToNumber("BM"), .Row, ct13
            .GetText .ColLetterToNumber("BQ"), .Row, ct14
            .GetText .ColLetterToNumber("AU"), .Row, ct9
            

            If Trim(MSTDN) <> "" Or Trim(strNoiDung) <> "" Or Trim(strNgayHD) <> "" Or Val(strSoTienTT) <> 0 Or Trim(strNgayTT) <> "" Or Val(strDoanhThu) <> 0 Or Val(strTyLeGT) <> 0 Or Val(strThueSuat) <> 0 Or Val(strDoanhThuTT) <> 0 Or Val(strTyLeTNDN) <> 0 Or Val(strSoThueMG) <> 0 Then
                iWhiteRow = False
                If Val(ct9) <> 0 Then
                    isCheckTong9 = True
                End If
                
                If Val(ct13) <> 0 Then
                    isCheckTong13 = True
                End If
            Else
                iWhiteRow = True
            End If
        
            If (strNoiDung = vbNullString Or Trim(strNoiDung) = "") And iWhiteRow = False Then
                .Col = .ColLetterToNumber("C")
                .CellNote = errEmty
                .BackColor = mErrorColor

                iFlagEmtyND = 1
            Else
                .Col = .ColLetterToNumber("C")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            
            'canh bao: Thue suat GTGT cot 8
            If Trim(thueSuatGTGT) <> "" Then
                If thueSuatGTGT <> "0" And thueSuatGTGT <> "5" And thueSuatGTGT <> "10" Then
                    .Col = .ColLetterToNumber("AQ")
                    .CellNote = errThueSuatGTGT
                    .BackColor = mAlertColor
                Else
                    .Col = .ColLetterToNumber("AQ")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            End If
        
            'dntai check ngay thanh toan < ky ke khai
            If Trim(strNgayTT) <> "" Then
                If CInt(Right(strNgayTT, 4)) > TAX_Utilities_v2.Year Then
                    .Col = .ColLetterToNumber("AD")
                    .CellNote = GetAttribute(GetMessageCellById("0113"), "Msg")
                    .BackColor = mErrorColor

                    iFlagNgayTT = 1
                ElseIf CInt(Right(strNgayTT, 4)) = CInt(TAX_Utilities_v2.Year) And CInt(Right("00" & Mid(strNgayTT, 4, 2), 2)) > CInt(TAX_Utilities_v2.Month) Then
                    .Col = .ColLetterToNumber("AD")
                    .CellNote = errNgayTT
                    .BackColor = mErrorColor

                    iFlagNgayTT = 1
                Else
                    If Mid$(strNgayTT, 4, 2) <> Trim$(kythang) Then
                         .Col = .ColLetterToNumber("AD")
                        .CellNote = errNgayTT2
                        .BackColor = mAlertColor
                    Else
                        .Col = .ColLetterToNumber("AD")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                End If
                
                ' nkhoan:
'                If Mid$(strNgayTT, 4, 2) <> Trim$(kythang) Or Right$(strNgayTT, 4) <> Trim$(nam) Then
'                    .Col = .ColLetterToNumber("AD")
'                    .CellNote = errNgayTT2
'                    .BackColor = mAlertColor
'                Else
'                    .Col = .ColLetterToNumber("AD")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                End If
               
            End If

            ' bat dinh dang chi tieu [07],[08],[11]
            
            Dim ct10 As Variant
            Dim ct12 As Variant
          
            Dim ct06 As Variant
            .GetText .ColLetterToNumber("AI"), .Row, ct06
            '         .GetText .ColLetterToNumber("AU"), .Row, ct09
            .GetText .ColLetterToNumber("AY"), .Row, ct10
            .GetText .ColLetterToNumber("BG"), .Row, ct12
            '         .GetText .ColLetterToNumber("BM"), .Row, ct13
            .GetText .ColLetterToNumber("AM"), .Row, ct07
            .GetText .ColLetterToNumber("AQ"), .Row, ct08
            .GetText .ColLetterToNumber("BC"), .Row, ct11
         
            'nkhoan: check ma so thue
            .GetText .ColLetterToNumber("L"), .Row, vMST

            If Trim(MSTDN) <> "" Or Trim(strNoiDung) <> "" Or Trim(strNgayHD) <> "" Or Val(strSoTienTT) <> 0 Or Trim(strNgayTT) <> "" Or Val(strDoanhThu) <> 0 Or Val(strTyLeGT) <> 0 Or Val(strThueSuat) <> 0 Or Val(strDoanhThuTT) <> 0 Or Val(strTyLeTNDN) <> 0 Or Val(strSoThueMG) <> 0 Then
                If vMST <> vbNullString Then
                    vMST = Replace(vMST, " ", "")

                    'check MST 14 ky tu
                    If Len(vMST) = 14 And Mid(vMST, 11, 1) = "-" Then
                        vMST = Replace(vMST, "-", "")
                    End If

                    strCheckMST = ""

                    If Len(vMST) = 10 Or Len(vMST) = 13 Then
                        strCheckMST = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))

                        If Mid(strCheckMST, 2, 1) <> "0" Then
                            .Col = .ColLetterToNumber("L")
                            .BackColor = mErrorColor
                            .CellNote = errMST

                            iFlagCheckMaST = 1
                        Else
                            .Col = .ColLetterToNumber("L")
                            .BackColor = mNonErrorColor
                            .CellNote = ""
                            'iFlagCheckMaST = 0
                        End If

                    ElseIf Len(vMST) <> 10 Or Len(vMST) <> 13 Then
                        .Col = .ColLetterToNumber("L")
                        .BackColor = mErrorColor
                        .CellNote = errMST

                        iFlagCheckMaST = 1
              
                    End If

                Else
                    .Col = .ColLetterToNumber("L")
                    .BackColor = mNonErrorColor
                    .CellNote = ""

                    iFlagEmty = 0
                End If
            End If

            '16709097
            'dntai 09/04/2012
            'check ct 13, 14 khong dc de am theo yeu cau moi cua chi Thuy tren ban 3.1.2
'            If Val(Trim(ct13)) < 0 Then
'                .Col = .ColLetterToNumber("BM")
'                .BackColor = mErrorColor
'                .CellNote = errCheckCT13CT14
'                iFlagCheckCT13CT14 = 1
'            Else
'                .Col = .ColLetterToNumber("BM")
'                .BackColor = 16709097
'                .CellNote = ""
'            End If
'
'            If Val(Trim(ct14)) < 0 Then
'                .Col = .ColLetterToNumber("BQ")
'                .BackColor = mErrorColor
'                .CellNote = errCheckCT13CT14
'                iFlagCheckCT13CT14 = 1
'            Else
'                .Col = .ColLetterToNumber("BQ")
'                .BackColor = 16709097
'                .CellNote = ""
'            End If
            countRowDyn = countRowDyn + 1
            
            .Row = 55 + i
            i = i + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        
        ' check dong tong
        .Row = .Row + 1
        .GetText .ColLetterToNumber("AU"), .Row, tongCt9
        .GetText .ColLetterToNumber("BM"), .Row, tongCt13
        ' check tong 9
        If isCheckTong9 = True Then
              If Val(tongCt9) = 0 Then
                    .Col = .ColLetterToNumber("AU")
                    .BackColor = mErrorColor
                    .CellNote = vErrorTongCt9
                    iFlagCheckTongCt9 = 1
              Else
                    .Col = .ColLetterToNumber("AU")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
              End If
        Else
             .Col = .ColLetterToNumber("AU")
            .BackColor = mNonErrorColor
            .CellNote = ""
        End If
        
        ' check tong 13
        If isCheckTong13 = True Then
              If Val(tongCt13) = 0 Then
                    .Col = .ColLetterToNumber("BM")
                    .BackColor = mErrorColor
                    .CellNote = vErrorTongCt13
                    iFlagCheckTongCt13 = 1
              Else
                    .Col = .ColLetterToNumber("BM")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
              End If
        Else
             .Col = .ColLetterToNumber("BM")
            .BackColor = mNonErrorColor
            .CellNote = ""
        End If
    
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 12, IIf(iFlagEmty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 13, IIf(iFlagEmtyND = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 15, IIf(iFlagNgayTT = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 17, IIf(iFlagCheckMaSTDN = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 19, IIf(iFlagCheckMaST = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 23, IIf(iFlagCheckCT13CT14 = 1, "0", "1")
        
        .SetText .ColLetterToNumber("B"), 24, IIf(iFlagCheckTongCt9 = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 25, IIf(iFlagCheckTongCt13 = 1, "0", "1")
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

End Sub

Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean

End Function

Public Sub SetData()

End Sub

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
        If (GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "Active") <> "0") Then
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
                'If Not (lRow = 5 And (lCol = .ColLetterToNumber("AT") Or lCol = .ColLetterToNumber("BC"))) Then
                    ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                    .GetText lCol, lRow, varTemp
                    UpdateCell fps, lCol, lRow, varTemp
                'End If
             Next
        End If
      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing

End Sub

'doc vao mot so va tra ve chu cua so do
Private Function doiSoThanhChu(ByVal so As String) As String
    Dim result As String
    result = ""
    Select Case so
            Case "0": result = ""
            Case "1": result = GetAttribute(GetMessageCellById("0204"), "Msg")
            Case "2": result = GetAttribute(GetMessageCellById("0205"), "Msg")
            Case "3": result = GetAttribute(GetMessageCellById("0206"), "Msg")
            Case "4": result = GetAttribute(GetMessageCellById("0207"), "Msg")
            Case "5": result = GetAttribute(GetMessageCellById("0208"), "Msg")
            Case "6": result = GetAttribute(GetMessageCellById("0209"), "Msg")
            Case "7": result = GetAttribute(GetMessageCellById("0210"), "Msg")
            Case "8": result = GetAttribute(GetMessageCellById("0211"), "Msg")
            Case "9": result = GetAttribute(GetMessageCellById("0212"), "Msg")
    End Select
    doiSoThanhChu = result
End Function
'ham doi tien tu so thanh chu
Private Function doiTien(ByVal so As String) As String
    Dim GTtien As String
    Dim value As String
    Dim i As Integer
    'luu lai gia tri so ban dau de xu ly
    Dim soTam2 As String
    Dim soTam As String
    soTam = so
    soTam2 = "1"
    For i = 0 To Len(so) - 1 Step 1

            If Len(so) Mod 3 = 2 And Mid(so, 1, 1) = "1" Then
                value = GetAttribute(GetMessageCellById("0213"), "Msg")
            Else
                value = doiSoThanhChu(Mid(so, 1, 1))
            End If
            doiTien = doiTien + value
            If (Len(so) Mod 3 = 0) Then
                soTam2 = so
            End If
            If (Len(so) Mod 3 = 0) And Mid(so, 1, 1) <> "0" Then
                doiTien = doiTien + GetAttribute(GetMessageCellById("0214"), "Msg")
            ElseIf Len(so) Mod 3 = 2 And Mid(so, 1, 1) <> "0" And Mid(so, 1, 1) <> "1" Then
                doiTien = doiTien + GetAttribute(GetMessageCellById("0219"), "Msg")
            ElseIf Len(so) Mod 3 = 2 And Mid(so, 1, 1) = "0" And Mid(so, 2, 1) <> "0" Then
                doiTien = doiTien + GetAttribute(GetMessageCellById("0220"), "Msg")
            ElseIf Len(so) Mod 3 = 1 And Len(so) <> 1 And Val(Mid(soTam2, 1, 3)) <> 0 Then
                doiTien = doiTien + giaTriTien(so)
            End If
            so = Mid(so, 2, Len(so) - 1)
    Next
    doiTien = Trim(doiTien) + GetAttribute(GetMessageCellById("0221"), "Msg")
    doiTien = UCase(Left(doiTien, 1)) + Mid(doiTien, 2)
End Function
'xac dinh menh gia tien
Private Function giaTriTien(ByVal tien As String) As String
    If Len(tien) < 4 Then
        giaTriTien = GetAttribute(GetMessageCellById("0214"), "Msg")
    ElseIf Len(tien) >= 4 And Len(tien) < 7 Then
        giaTriTien = GetAttribute(GetMessageCellById("0215"), "Msg")
    ElseIf Len(tien) >= 7 And Len(tien) < 10 Then
        giaTriTien = GetAttribute(GetMessageCellById("0216"), "Msg")
    ElseIf Len(tien) >= 10 And Len(tien) < 13 Then
        giaTriTien = GetAttribute(GetMessageCellById("0217"), "Msg")
    ElseIf Len(tien) = "13" Then
        giaTriTien = GetAttribute(GetMessageCellById("0218"), "Msg")
    End If
End Function


' Lay ve danh sach cac chi tieu dieu chinh KHBS
Function getMapCT() As String
    On Error GoTo ErrHandle
    
    Dim tempValue As Variant
    Dim strPath As String
    Dim i As Integer
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT_TT28.xml"
    
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "ID") = GetAttribute(TAX_Utilities_v2.NodeValidity.childNodes(0), "DataFile") Then
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
    SaveErrorLog "cls_01NTNN", "getMapCT", Err.number, Err.Description
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
            Set xmlNode = .nodeFromID("AU_48")
            result = "CT9_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = .nodeFromID("BM_48")
            result = result & "~" & "CT13_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01NTNN", "getValueTK", Err.number, Err.Description
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
                .GetText .ColLetterToNumber("AU"), 48, varTemp
                strCurrentValue = "CT9_" & Trim(varTemp)
                .GetText .ColLetterToNumber("BM"), 48, varTemp
                strCurrentValue = strCurrentValue & "~CT13_" & Trim(varTemp)
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

' Lay ve chi tieu D o to khai lan truoc
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
            Set xmlNode = .nodeFromID("BQ_48")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_01NTNN", "getValueCTDC", Err.number, Err.Description
End Function


' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("BQ"), 48, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function

' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 17
    denRowDL = 25
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "20_30_2"
    ' end
End Sub

' insert node
Public Sub InsertNode(ByVal fRow As Long, ByVal lRow As Long, ByVal mSheet As Byte)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCell As MSXML.IXMLDOMNode
    
    Dim xmlNodeCellsTmp As MSXML.IXMLDOMNode
    Dim xmlNodeCellTmp As MSXML.IXMLDOMNode
    Dim arrRow() As String
    Dim tmpValue As String
    Dim arrRow2() As String
    Dim tmpValue2 As String
    Dim lCol1 As Long, lRow1 As Long
    
    Dim i As Integer
    
    
    Dim arrTemp() As String
    
    Dim ctlRow As Long
    Dim cTemplate, cReport As String
    
    Dim sumRowDel As Long
    sumRowDel = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Cells").length - 4
    For ctlRow = 1 To sumRowDel - 1
        Set xmlNodeCells = TAX_Utilities_v2.Data(mSheet - 1).nodeFromID("C" & "_" & 55 + ctlRow & "").parentNode
        xmlNodeCells.parentNode.removeChild xmlNodeCells
    Next
    
    ' set lai row cho cac dong tinh
    ' group 5
    
    Set xmlNodeCellsTmp = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Cells").Item(4)
    tmpValue = "67~69~67~30~31~31~31~31"
    tmpValue2 = "65~66~71~30~31~31~31~31"
    arrRow = Split(tmpValue, "~")
    arrRow2 = Split(tmpValue2, "~")
    i = 0
    For Each xmlNodeCellTmp In xmlNodeCellsTmp.childNodes
        If CDbl(arrRow(i)) > 55 Then
            ' Increase value of row attribute + 1 (CellID)
            ParserCellID fps, GetAttribute(xmlNodeCellTmp, "CellID"), lCol1, lRow1
            SetAttribute xmlNodeCellTmp, "CellID", GetCellID(fps, lCol1, CDbl(arrRow(i)) + lRow - fRow + 1)
            
            ' Increase value of row attribute + 1 (CellID2)
            ParserCellID fps, GetAttribute(xmlNodeCellTmp, "CellID2"), lCol1, lRow1
            SetAttribute xmlNodeCellTmp, "CellID2", GetCellID(fps, lCol1, CDbl(arrRow2(i)) + lRow - fRow + 1)
        
        End If
        i = i + 1
    Next
    
    ' group 2
    Set xmlNodeCellsTmp = TAX_Utilities_v2.Data(mSheet - 1).getElementsByTagName("Cells").Item(2)
    tmpValue = "57~48~57~57~57~57~48~57~48"
    tmpValue2 = "55~48~55~55~55~55~48~55~48"
    arrRow = Split(tmpValue, "~")
    arrRow2 = Split(tmpValue2, "~")
    i = 0
    For Each xmlNodeCellTmp In xmlNodeCellsTmp.childNodes
        If CDbl(arrRow(i)) > 55 Then
            ' Increase value of row attribute + 1 (CellID)
            ParserCellID fps, GetAttribute(xmlNodeCellTmp, "CellID"), lCol1, lRow1
            SetAttribute xmlNodeCellTmp, "CellID", GetCellID(fps, lCol1, CDbl(arrRow(i)) + lRow - fRow + 1)
            
            ' Increase value of row attribute + 1 (CellID2)
            ParserCellID fps, GetAttribute(xmlNodeCellTmp, "CellID2"), lCol1, lRow1
            SetAttribute xmlNodeCellTmp, "CellID2", GetCellID(fps, lCol1, CDbl(arrRow2(i)) + lRow - fRow + 1)
        
        End If
        i = i + 1
    Next
 
    
    Set xmlNodeCells = TAX_Utilities_v2.Data(mSheet - 1).nodeFromID("C_55").parentNode

    
    
    For ctlRow = fRow To lRow

        'IncreaseRowInDOM_01NTNN fps, TAX_Utilities_v2.Data(mSheet - 1), 56, 1, 1, sumRowDel

        Set xmlNodeNewCells = xmlNodeCells.cloneNode(True)
        
        For Each xmlNodeNewCell In xmlNodeNewCells.childNodes
            arrTemp = Split(GetAttribute(xmlNodeNewCell, "CellID"), "_")
            cTemplate = Trim(arrTemp(0)) & "_"
            arrTemp = Split(GetAttribute(xmlNodeNewCell, "CellID2"), "_")
            cReport = Trim(arrTemp(0)) & "_"
            'cTemplate = Left$(GetAttribute(xmlNodeNewCell, "CellID"), 2)
            'cReport = Left$(GetAttribute(xmlNodeNewCell, "CellID2"), 2)
            
            SetAttribute xmlNodeNewCell, "CellID", cTemplate & ctlRow
                    
            ' Set first cell = 1
            SetAttribute xmlNodeNewCell, "FirstCell", "1"
            If Trim(cTemplate) = "C_" Or Trim(cTemplate) = "L_" Or Trim(cTemplate) = "R_" Then
                SetAttribute xmlNodeNewCell, "Value", vbNullString
            Else
                SetAttribute xmlNodeNewCell, "Value", "0"
            End If
            SetAttribute xmlNodeNewCell, "CellID2", cReport & (ctlRow - 2)

            
        Next
        
        ' Insert new node to DOM object
        xmlNodeCells.parentNode.insertBefore xmlNodeNewCells, Null
  
        
    
    Next
    
    
'   **********************************
'    added
'   Date: 12/04/06
    'mAdjustData = True
    TAX_Utilities_v2.AdjustData(mSheet - 1) = True
'   **********************************
EXIT_SUB:
    Set xmlNodeNewCell = Nothing
    Set xmlNodeNewCells = Nothing
    Set xmlNodeCells = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "cls_01NTNN", "InsertNode 01_NTNN", Err.number, Err.Description
End Sub


'
Public Sub FinishImport()
    With fps
        fps.EventEnabled(EventAllEvents) = False
            ' set node xml
        .Sheet = 1
        InsertNode 56, .SearchCol(.ColLetterToNumber("B"), 55, -1, "aa", SearchFlagsNone) - 1, 1
        UpdateSheets
        fps.EventEnabled(EventAllEvents) = True
    End With
End Sub


' Tinh lai row trong cac section tinh
Private Sub IncreaseRowInDOM_01NTNN(fpSpread1 As fpSpread, xmlDomData As MSXML.DOMDocument, ByVal pRow As Long, ByVal lRows As Long, ByVal lRow2s As Long, rowDel As Long)
    On Error GoTo ErrorHandle
    
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim lCol As Long, lRow As Long, i As Long
    Dim count As Integer
        
    If xmlDomData Is Nothing Then Exit Sub
    Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
    
    For i = xmlNodeListCell.length - 1 To 0 Step -1
        count = count + 1
        If count > 21 Then
            Exit For
        End If
        
        ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID"), lCol, lRow
        If lRow >= pRow Then
            ' Increase value of row attribute + 1 (CellID)
            SetAttribute xmlNodeListCell(i), "CellID", GetCellID(fpSpread1, lCol, lRow + lRows - rowDel)
            
            ' Increase value of row attribute + 1 (CellID2)
            ParserCellID fpSpread1, GetAttribute(xmlNodeListCell(i), "CellID2"), lCol, lRow
            SetAttribute xmlNodeListCell(i), "CellID2", GetCellID(fpSpread1, lCol, lRow + lRow2s - rowDel)
        End If
        
    Next
        
    Set xmlNodeListCell = Nothing
    
    Exit Sub
    
ErrorHandle:
    SaveErrorLog "mdlFunctions", "IncreaseRowInDOM_01NTNN", Err.number, Err.Description
End Sub


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
    For Each xmlNodeValid In TAX_Utilities_v2.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
                Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v2.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v2.Data(1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v2.Data(1).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_v2.Data(1).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_v2.Data(1).getElementsByTagName("Cell").length > 31 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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

