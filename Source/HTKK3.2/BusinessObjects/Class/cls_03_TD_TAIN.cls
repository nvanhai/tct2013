VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_03_TD_TAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Thu Nhap Doanh Nghiep" interface sheets
'this Class is belong to TAX_Business project which will be compline to DLL
'Coder: PhuocHK
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Private Const S1Txt1Row = 9
Private Const S1Txt1Col = "I"
Private Const S1Txt1aRow = 10
Private Const S1Txt1aCol = "I"
Private Const S1Txt1bRow = 11
Private Const S1Txt1bCol = "I"
Private Const S1Txt1cRow = 12
Private Const S1Txt1cCol = "I"
Dim vt As String, va As String, vb As String, vc As String
Dim strComboHien As String, strCombo As String
Dim arrTaiNguyen As New Scripting.Dictionary, mblKhoiDau As Boolean
Dim arrTaiNguyenGoc As New Scripting.Dictionary
Dim mblKeyF5 As Boolean
Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
Dim arrData() As String
Public loaiKHBS As String
Public FlagThaydoi As String
Dim arrDanhSachKHBS() As String

Public StrSolanBosung As String
Public Strloaitk As String

Public strTkThangQuy As String

' Kiem tra thong tin dai ly thue
' Neu khong co DL thue se an cac row nay di
Public FlagDLThue As Boolean
Public tuRowDL As Long
Public denRowDL As Long
Public strTuRowDenRowPL As String

Private Const Row_MSTDN = 14
Private Const Col_MSTDN = "F"
Private Const rowStatic = 42
Private Const sheetPl01_KHBS_GTGT = 3
Private Const sheetPl03_1TD_TAIN = 2


'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim strDataFileName As String
    Dim strDDKTTN, strDonvitinh As String
    Dim arrDanhsach() As String, blChon As Boolean
    
    Dim valueDefault As String, textDefault As String, thuesuatDefault As String, temp As String
    
    blChon = False
    strDataFileName = GetCatalogueFileName(blChon)
    
    arrTaiNguyen.RemoveAll
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Cell")
        For Each xmlNode In xmlNodeListCell
            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                If (blChon = True And arrDanhsach(6) = "1") Or blChon = False Then
                    If (Left$(Trim(arrDanhsach(1)), 2) = "09") Then
                        strCombo = strCombo + CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 10) + CPab(arrDanhsach(5), 1) + Chr$(9)
                        strComboHien = strComboHien + arrDanhsach(2) + Chr$(9)
                        If Trim(arrDanhsach(1)) = "0902" Then
                               valueDefault = Trim(arrDanhsach(1))
                               textDefault = Trim(arrDanhsach(2))
                               thuesuatDefault = Trim(arrDanhsach(4))
                               temp = CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 10) + CPab(arrDanhsach(5), 1)
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    With fps
        SetDateFormat fps, 1, 63, fps.ColLetterToNumber("R"), DDMMYYYY
        
        SetDateFormat fps, 3, 46, .ColLetterToNumber("BG"), DDMMYYYY
        SetDateFormat fps, 3, 31, .ColLetterToNumber("BE"), DDMMYYYY
        
        .Sheet = 1
        .Row = 63
        .Col = fps.ColLetterToNumber("R")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        'SetDateFormat fps, 1, 19, .ColLetterToNumber("AJ"), DDMMYYYY
    End With
    
     With fps
        'them danh muc hang vao combobox danh muc 1
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        .Col = .ColLetterToNumber("Z")
        .Row = 24
        .TypeComboBoxList = strCombo 'CPab("Maple", 100) + Chr$(9) + CPab("Oak", 100) + Chr$(9) + CPab("Ash", 100)
        .Text = temp
        'Combo Hien
        .Col = .ColLetterToNumber("H")
        .Row = 24
        .TypeComboBoxList = strComboHien
        .Text = textDefault
        
        .Col = .ColLetterToNumber("AA")
        .Row = 24
        .Text = valueDefault
        
        .Col = .ColLetterToNumber("F")
        .Row = 25
        .Text = thuesuatDefault
       
    End With
    
    setCQTQuanLy
    
    setCQTQuanLyHoanThue fps
End Sub
Private Sub setCQTQuanLy()
        Dim xmlDomData As New MSXML.DOMDocument
        Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
        Dim xmlNode As MSXML.IXMLDOMNode
        Dim arrDanhsach() As String
        Dim tenCucThue As String
        Dim maCucThue As String
        Dim tenChiCucThue As String
        Dim maChiCucThue As String
        
        If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    
                    If arrDanhsach(0) = "1" Then
                        tenCucThue = tenCucThue + Right$(arrDanhsach(3), Len(arrDanhsach(3)) - 6) + Chr$(9)
                        maCucThue = maCucThue + arrDanhsach(1) + Chr$(9)
                    End If
                    'If arrDanhsach(0) = "0" Then
                        tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
                        maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
                    'End If
                End If
            Next
            Set xmlDomData = Nothing
            Set xmlNodeListCell = Nothing
            Set xmlNode = Nothing
        End If
        
        With fps
            'phu luc 01_2TD/GTGT
            .Sheet = sheetPl03_1TD_TAIN 'set thong tin CQL
            .Row = rowStatic
            .Col = .ColLetterToNumber("P")
            .TypeComboBoxList = tenCucThue
            
            .Row = rowStatic
            .Col = .ColLetterToNumber("AH")
            .TypeComboBoxList = maCucThue
            
            
            
            '.EventEnabled(EventAllEvents) = True
        End With
End Sub
Private Sub setChangeComb(ByVal Col As Long, Row As Long)
    Dim varTemp As Variant
    Dim intIndexCuc As Integer
    Dim intIndexChiCuc As Integer
    Dim varIdCucThue As Variant
    
    Dim xmlDomData As New MSXML.DOMDocument
    Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim arrDanhsach() As String
    Dim tenChiCucThue As String
    Dim maChiCucThue As String
    
  '  Dim checkDL As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        If .ActiveSheet = sheetPl03_1TD_TAIN Then
            .Sheet = sheetPl03_1TD_TAIN
            mCurrentSheet = sheetPl03_1TD_TAIN
            .Col = Col
            .Row = Row
            
'            If Col = .ColLetterToNumber("L") Then 'Khi chon Combo
'                'Chuyen sang du lieu tuong ung Lay gia tri index cot L va gia tri Index cot AG
'                intIndexCuc = .typeComboboxCurSel
'                .Col = .ColLetterToNumber("AG")
'                .Row = Row
'                .typeComboboxCurSel = intIndexCuc
'                'Nhap lai Combo cot AG
'                .Col = .ColLetterToNumber("AG")
'                UpdateCell fps, .Col, .Row, .Text
'                ' Lay Id cua cuc thue vua chon o tren
'                fps.GetText .Col, .Row, varIdCucThue
'
'                If varIdCucThue <> "" Or varIdCucThue <> vbNullString Then
'                    ' Lay thong tin cho chi cuc thue
'                    varIdCucThue = Left$(varIdCucThue, 3)
'
'                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
'                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
'                        For Each xmlNode In xmlNodeListCell
'                            If GetAttribute(xmlNode, "Value") <> "" Then
'                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
'
'                                If arrDanhsach(0) = "0" And arrDanhsach(2) = varIdCucThue Then
'                                    tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
'                                    maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
'                                End If
'                            End If
'                        Next
'                        Set xmlDomData = Nothing
'                        Set xmlNodeListCell = Nothing
'                        Set xmlNode = Nothing
'                    End If
'
'                    .Row = Row
'                    .Col = .ColLetterToNumber("P")
'                    .TypeComboBoxList = tenChiCucThue
'                    .Text = ""
'
'                    .Row = Row
'                    .Col = .ColLetterToNumber("AH")
'                    .TypeComboBoxList = maChiCucThue
'                    .Text = ""
'                End If
'            End If
                        
            If Col = .ColLetterToNumber("P") Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
                intIndexChiCuc = .typeComboboxCurSel
                .Col = .ColLetterToNumber("AH")
                .Row = Row
                .typeComboboxCurSel = intIndexChiCuc
                'Nhap lai Combo cot Q
                .Col = .ColLetterToNumber("AH")
                UpdateCell fps, .Col, .Row, .Text
            End If
            
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim strColC As String
    Dim i As Long
    Dim intCtrl As Integer
    Dim strId As Variant
    Dim xmlNode As MSXML.IXMLDOMNode
    With fps
        .Sheet = 1
        mCurrentSheet = .Sheet

        'update MSTDN luôn bang MST tren phan TT chung
        UpdateMST fps, Col_MSTDN, Row_MSTDN

        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("C")
            .Row = 27
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 27
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 27
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("C")
            .Row = 27
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 27
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 27
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        ' set loai tk thang/ quy
         If strTkThangQuy = "TK_THANG" Then
            .Col = .ColLetterToNumber("O")
            .Row = 27
            .Text = "0"
            UpdateCell fps, .Col, .Row, .Text
        ElseIf strTkThangQuy = "TK_QUY" Then
            .Col = .ColLetterToNumber("O")
            .Row = 27
            .Text = "1"
            UpdateCell fps, .Col, .Row, .Text
        End If
            'end
        ' set lock CQT phan bo
        .Sheet = 2
        Dim vChon As Variant
        If .SheetVisible = True Then
        .Row = 42
         Do
                .GetText .ColLetterToNumber("G"), .Row, vChon
                If vChon = 1 Then
                    .Col = .ColLetterToNumber("P")
                    .Lock = True
                End If
                
                .Row = .Row + 1
                .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
         End If
        mCurrentSheet = .Sheet
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
        
        CheckDynamicError 'Set Exception Error on cells of interface
            
        UpdateSheet
            
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()

    Dim iCurrentSheet As Integer, iFlagEmpty As Integer, iWhiteRow As Integer, iFlagCheckMaST   As Integer
    Dim vTenCoQuan    As Variant, vCapCuc As Variant, vCapDiaPhuong As Variant, vMessRequire As Variant
    Dim row_require   As Integer
    Dim vSTT As Variant, vTenNhaMay As Variant, vCQTCapCuc As Variant, vCQTCapDiaPhuong As Variant, vTyLePhanBo As Variant, vSoThuePhaiNop As Variant
    Dim vMST As Variant
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String, iFocusFlag As Integer, errMST As Variant, strCheck As String
    
    Dim vNhaMayTD, vSanLuong, vDonGia, vThueMG, vThuePS As Variant
    Dim errorThueMG As Variant
    Dim flagThueMG As Integer
    
    iWhiteRow = 1
    
    row_require = 15
    
    With fps
        iCurrentSheet = .Sheet
        'get message
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), row_require, vMessRequire
        .GetText .ColLetterToNumber("E"), 13, errMST
        
        .GetText .ColLetterToNumber("E"), 17, errorThueMG
        
        .Sheet = iCurrentSheet
        
        'check for to khai
        .Sheet = 1
        .Row = 44
        iFlagCheckMaST = 0
        If .SheetVisible = True Then
            mCurrentSheet = .Sheet
            Do
                .GetText .ColLetterToNumber("C"), .Row, vNhaMayTD
                
                .GetText .ColLetterToNumber("F"), .Row, vMST
                .GetText .ColLetterToNumber("I"), .Row, vSanLuong
                .GetText .ColLetterToNumber("K"), .Row, vDonGia
                .GetText .ColLetterToNumber("O"), .Row, vThuePS
                .GetText .ColLetterToNumber("R"), .Row, vThueMG
                
                If Trim(vNhaMayTD) <> "" Or Trim(vMST) <> "" Or Val(vSanLuong) <> 0 Or Val(vDonGia) <> 0 Or Val(vThueMG) <> 0 Then iWhiteRow = 0
                
                If iWhiteRow = 0 Then
                
                        .Col = .ColLetterToNumber("C")
                        If (Trim(.Text) = vbNullString) Then
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                            .CellNote = vMessRequire      'static
                            .BackColor = mErrorColor
            
                            iFlagEmpty = 1
                        Else
                            .BackColor = mNonErrorColor
                            .CellNote = ""
                            .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        End If
                        
                        'validate mst
                        .Col = .ColLetterToNumber("F")
                        If Trim$(vMST) <> vbNullString Then
                            vMST = Replace(vMST, " ", "")
                            vMST = Replace(vMST, "-", "")
                            'vMST = Left(vMST, 10)
        
                            strCheck = ""
        
                            If Len(vMST) = 10 Or Len(vMST) = 13 Then
                                strCheck = CheckTaxCode(Mid$(vMST, 1, 1), Mid$(vMST, 2, 1), Mid$(vMST, 3, 1), Mid$(vMST, 4, 1), Mid$(vMST, 5, 1), Mid$(vMST, 6, 1), Mid$(vMST, 7, 1), Mid$(vMST, 8, 1), Mid$(vMST, 9, 1), Mid$(vMST, 10, 1), Mid$(vMST, 11, 1), Mid$(vMST, 12, 1), Mid$(vMST, 13, 1))
                            End If
                               
                            If Mid$(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                                .Col = .ColLetterToNumber("F")
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
                                .Col = .ColLetterToNumber("F")
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                            End If
        
                        Else
                            If (Trim$(vMST) = vbNullString) Then
                                .Col = .ColLetterToNumber("F")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = ""
                                .BackColor = mNonErrorColor
                                
                                'iFlagEmty = 1
                            End If
                        End If
                        
                        ' kiem tra chi tieu 28 <= 27
                        If Val(vThueMG) > Val(vThuePS) Then
                            .Col = .ColLetterToNumber("R")
                            .CellNote = errorThueMG   'static
                            .BackColor = mErrorColor
    
                            flagThueMG = 1
                        Else
                            .Col = .ColLetterToNumber("R")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        End If
                End If
                .Row = .Row + 1
                .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
        End If
        
        .Sheet = 2
        Dim vChon As Variant
        If .SheetVisible = True Then
            mCurrentSheet = .Sheet
            'check phu luc 01_2
            .Row = 42
            iWhiteRow = 1
            
            Do
             
                
                .GetText .ColLetterToNumber("C"), .Row, vTenNhaMay
                .GetText .ColLetterToNumber("G"), .Row, vChon
                .GetText .ColLetterToNumber("H"), .Row, vMST
                .GetText .ColLetterToNumber("P"), .Row, vCQTCapDiaPhuong
                .GetText .ColLetterToNumber("R"), .Row, vTyLePhanBo
                .GetText .ColLetterToNumber("U"), .Row, vSoThuePhaiNop
                
                
                If Trim$(vTenNhaMay) <> "" Or Trim$(vMST) <> "" Or Trim$(vCQTCapDiaPhuong) <> "" Or Trim$(vTyLePhanBo) <> "0" Or Trim$(vSoThuePhaiNop) <> "0" Then
                    iWhiteRow = 0
                End If
            
                .Col = .ColLetterToNumber("C")
                If (Trim(.Text) = vbNullString And Trim(vChon) = "1" And iWhiteRow = 0) Then
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vMessRequire      'static
                    .BackColor = mErrorColor
    
                    iFlagEmpty = 1
                Else
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                End If
    
                
                .Col = .ColLetterToNumber("P")
                If (Trim(.Text) = vbNullString And Trim(vChon) = "0" And iWhiteRow = 0) Then
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vMessRequire      'static
                    .BackColor = mErrorColor
    
                    iFlagEmpty = 1
                Else
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                End If
                
                'check mst
                If Trim$(vMST) <> vbNullString Then
                    vMST = Replace(vMST, " ", "")
                    vMST = Replace(vMST, "-", "")
                    'vMST = Left(vMST, 10)

                    strCheck = ""

                    If Len(vMST) = 10 Or Len(vMST) = 13 Then
                        strCheck = CheckTaxCode(Mid$(vMST, 1, 1), Mid$(vMST, 2, 1), Mid$(vMST, 3, 1), Mid$(vMST, 4, 1), Mid$(vMST, 5, 1), Mid$(vMST, 6, 1), Mid$(vMST, 7, 1), Mid$(vMST, 8, 1), Mid$(vMST, 9, 1), Mid$(vMST, 10, 1), Mid$(vMST, 11, 1), Mid$(vMST, 12, 1), Mid$(vMST, 13, 1))
                    End If
                       
                    If Mid$(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                        .Col = .ColLetterToNumber("H")
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
                        .Col = .ColLetterToNumber("H")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If

                Else
                    If (Trim$(vMST) = vbNullString And iWhiteRow = 0) Then
                        .Col = .ColLetterToNumber("H")
                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                        
                        'iFlagEmty = 1
                    End If
                End If
                
                .Row = .Row + 1
                .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
        End If
        
        'set flag
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), row_require, IIf(iFlagEmpty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 13, IIf(iFlagCheckMaST = 1, "0", "1")
        
        .SetText .ColLetterToNumber("B"), 17, IIf(flagThueMG = 1, "0", "1")
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
    For Each xmlNodeValid In TAX_Utilities_v1.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "03-1"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_v1.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_v1.Data(1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                fps.Sheet = sheetPl03_1TD_TAIN
                fps.GetText fps.ColLetterToNumber("C"), 42, strValue4
                    If Val(strValue1) <> 0 Or (Not IsNullValue(strValue4)) Then
                            blnNullValue = False
                    End If

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

Public Sub Prepare_KHBS()
    Dim xmlDocument As New MSXML.DOMDocument
    Dim xmlNodeListItem As MSXML.IXMLDOMNodeList
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim xmlNodeItem As MSXML.IXMLDOMNode
    Dim strPath As String
    Dim strComboName As String
    Dim strComboMCT As String
    Dim strCombo As String
    Dim blChon As Boolean
    Dim arrDanhsach() As String
    
    fps.Sheet = fps.SheetCount - 1
    fps.ActiveSheet = fps.Sheet
    
    
    blChon = False
    strPath = GetCatalogueFileName(blChon)
    
    'combo box khbs
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Cell")
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                If (blChon = True And arrDanhsach(6) = "1") Or blChon = False Then
                    strCombo = strCombo + CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(3), 10) + CPab(arrDanhsach(4), 10) + CPab(arrDanhsach(5), 1) + Chr$(9)
                  '  strCombo = strCombo + CPab(arrDanhsach(2), 200) + CPab(arrDanhsach(1), 10) + Chr$(9)
                    strComboName = strComboName + arrDanhsach(2) + Chr$(9)
                    strComboMCT = strComboMCT + "10" + Chr$(9)
                End If
            End If
        Next
    End If
    
    
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
        .Text = "10"
        .Lock = True
        .Row = 13
        .TypeComboBoxList = strComboMCT
        .Text = "10"
        .Lock = True
    
'Cho phep nhap am cot 4, 5, 7, 9 to khai bo sung
        .Sheet = 1
        .Row = 13
        .Col = .ColLetterToNumber("S")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("V")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("AD")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
        .Col = .ColLetterToNumber("AJ")
        .TypeNumberMin = Val("-99999999999999")
        .TypeNumberMax = Val("99999999999999")
    
    
     End With
    
End Sub

Public Sub UpdateChangeKHBS()
    Dim varTemp As Variant
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    Dim xmlCell As MSXML.IXMLDOMNode
    Dim xmlCellUpdate As MSXML.IXMLDOMNode
    Dim xmlCellGoc As MSXML.IXMLDOMNode
    Dim strId As String
    With fps
       If loaiKHBS = "frmKHBS_TT" Then
            .Sheet = 1
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateLastKHBSCell fps, lCol, lRow, varTemp
            Next
       Else
            For Each xmlCell In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
               If Left(GetAttribute(xmlCell, "CellID"), 2) = "BC" Then
                    strId = Trim(Mid(GetAttribute(xmlCell, "Value"), 200, 10))
                        
                    For Each xmlCellGoc In TAX_Utilities_v1.DataKHBS.getElementsByTagName("Cell")
                        If Left(GetAttribute(xmlCellGoc, "CellID"), 1) = "E" Then
                            If GetAttribute(xmlCellGoc, "Value") = strId Then
                            'update secsion
                               Set xmlCellUpdate = seachCellNode(TAX_Utilities_v1.Data(0).getElementsByTagName("Cell"), GetAttribute(xmlCellGoc, "CellID"))
                               updateSecsion xmlCellUpdate.parentNode
                                'updateSecsion (xmlCellGoc.parentNode)
                            Else
                            'insert secsion
                            
                            End If
                            
                        End If
                    Next
               
               End If
            Next
       End If
        
        
        .Sheet = mCurrentSheet
    End With
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

Private Sub updateSecsion(xmlCellNode As MSXML.IXMLDOMNode)
Dim xmlCellNodeData As MSXML.IXMLDOMNode
Dim varTemp As Variant
Dim lCol As Long, lRow As Long
With fps
    For Each xmlCellNodeData In xmlCellNode.childNodes
     If Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "S" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "V" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 2) = "AD" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "AG" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 2) = "AJ" Or Left(GetAttribute(xmlCellNodeData, "CellID"), 2) = "AN" Then
        ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
        .Sheet = 1
        .GetText lCol, lRow, varTemp
        UpdateLastKHBSCell fps, lCol, lRow, varTemp
     End If
    Next
End With
End Sub

Private Sub fillData()
    Dim varTemp         As Variant
    Dim lCol            As Long, lRow As Long
    Dim xmlCellNode     As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    Dim arrDanhsach()   As String
    Dim xmlCell         As MSXML.IXMLDOMNode

    With fps
        Dim i           As Integer
        Dim strFormular As String
        ReDim arrDanhSachKHBS(TAX_Utilities_v1.Data(0).getElementsByTagName("Section")(1).childNodes.length - 1)
        i = 0

        For Each xmlCellNodeData In TAX_Utilities_v1.DataKHBS.getElementsByTagName("Cell")

            If Left(GetAttribute(xmlCellNodeData, "CellID"), 1) = "F" Then
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                arrDanhSachKHBS(i) = varTemp & "~" & GetAttribute(xmlCellNodeData, "CellID")
                i = i + 1
            End If

        Next
        
        .Sheet = .SheetCount - 1
    End With

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

Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
With fps
        If Col = .ColLetterToNumber("G") And Row >= 42 And .ActiveSheet = 2 Then
            fps.EventEnabled(EventAllEvents) = False
            .Sheet = 2
            mCurrentSheet = .Sheet
            .Row = Row
            .Col = .ColLetterToNumber("G")
            If (.value = "1") Then
                .Row = Row
                ' chi tieu STT
                .Col = .ColLetterToNumber("P")
                .Lock = True
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                
                .Col = .ColLetterToNumber("AH")
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
            Else
                .Row = Row
                ' chi tieu STT
                .Col = .ColLetterToNumber("P")
                .Lock = False
                
            End If
            fps.EventEnabled(EventAllEvents) = True
        End If

    End With
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    setChangeComb Col, Row
    
    Dim intIndexCombo As Integer
    Dim strId As Variant
    
    
    With fps
        .EventEnabled(EventAllEvents) = False
       If .ActiveSheet = 1 Then
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("H") And Row = 24 And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("Z")
            .typeComboboxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber("Z")
            UpdateCell fps, .Col, .Row, .Text
             
            'Chon lai combo cua colF
            .Col = .ColLetterToNumber("AA")
            .GetText .ColLetterToNumber("Z"), Row, strId
            .Text = Trim(Mid(strId, 200, 10))
             UpdateCell fps, .Col, .Row, .Text

             .Col = .ColLetterToNumber("F")
             .Row = 25
             .Text = Trim(Mid(strId, 221, 10))
             UpdateCell fps, .Col, .Row, .Text
        End If
     End If
     .EventEnabled(EventAllEvents) = True
   End With
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i          As Integer, iCol As Long, iRow As Long
    Dim tCol       As Long, tRow As Long
    Dim strId      As Variant, Col7 As Variant, intIndexCombo As Integer
    Dim NguoiKyCol As Long, NguoiKyRow As Long

    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow

        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            'Ke khai bo sung
'            If .ActiveSheet = .SheetCount - 1 Then
'                'fps_Change iCol, iRow
'                fps.EventEnabled(EventAllEvents) = False
'                i = 1
'                .Col = .ColLetterToNumber("B")
'                .Row = 9
'
'                Do
'                    .Text = str(i)
'                    .Col = .ColLetterToNumber("B")
'                    .Row = i + 9
'                    i = i + 1
'                Loop Until .Text = "aa"
'
'                i = 1
'                .Row = .Row + 3
'
'                Do
'                    .Text = str(i)
'                    .Col = .ColLetterToNumber("B")
'                    .Row = 1 + .Row
'                    i = i + 1
'                Loop Until .Text = "bb"
'
'                .SetActiveCell iCol, iRow
'                .Row = iRow
'                .Col = iCol
'                .SetActiveCell iCol, iRow
'                fps.EventEnabled(EventAllEvents) = True
'            End If
            'To khai 01
            If .ActiveSheet = 1 Then
                'fps_Change iCol, iRow
                fps.EventEnabled(EventAllEvents) = False
                i = 1
                .Col = .ColLetterToNumber("B")
                .Row = 44

                Do
                    .Text = str(i)
                    .Col = .ColLetterToNumber("B")
                    .Row = i + 44
                    i = i + 1
                Loop Until .Text = "aa"

                .SetActiveCell iCol, iRow
                .Row = iRow
                .Col = iCol
                .SetActiveCell iCol, iRow
                fps.EventEnabled(EventAllEvents) = True
            End If
            
            ' set lock CQT phan bo
            If .ActiveSheet = 2 Then
                .Sheet = 2
                Dim vChon As Variant
                If .SheetVisible = True Then
                .Row = 42
                 Do
                        .GetText .ColLetterToNumber("G"), .Row, vChon
                        If vChon = 1 Then
                            .Col = .ColLetterToNumber("P")
                            .Lock = True
                        Else
                            .Col = .ColLetterToNumber("P")
                            .Lock = False
                        End If
                        
                        .Row = .Row + 1
                        .Col = .ColLetterToNumber("B")
                    Loop Until .Text = "aa"
                 End If
            End If
        End If
    End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, _
                          ByVal Row As Long, _
                          ByVal NewCol As Long, _
                          ByVal NewRow As Long, _
                          Cancel As Boolean)
    Dim varTemp   As Variant
    Dim strCellID As String
    Dim tCol      As Long, tRow As Long
    Dim cCol      As Long, cRow As Long
    Dim strId     As Variant
    Dim xmlNode   As MSXML.IXMLDOMNode
    Dim count     As Long

    With fps
        mCurrentSheet = .Sheet
        .Sheet = .ActiveSheet
        .EventEnabled(EventAllEvents) = False

        If .ActiveSheet = 1 Then

            'tinh so dong
            count = 0
            .Col = .ColLetterToNumber("B")
            .Row = 44
            Do
                count = count + 1
                .Row = .Row + 1
            Loop Until .value = "aa"

            If Col = .ColLetterToNumber("R") And Row = 63 + count - 1 Then
                .GetText Col, Row, varTemp

                If varTemp <> "" And varTemp <> "../../...." Then
                    If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                        .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                        .TypeHAlign = TypeHAlignLeft
                    Else
                        '.SetFocus
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

        'KHBS
        If .ActiveSheet = .SheetCount - 1 Then
            Dim count1 As Integer
            .Row = 9
            .Col = .ColLetterToNumber("B")

            Do While .value <> "bb"
                count1 = count1 + 1
                .Row = .Row + 1
            Loop

            If Col = .ColLetterToNumber("BG") And Row = 20 + count1 - 5 Then
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

                UpdateCell fps, .Col, .Row, .Text
            End If

            .Sheet = .SheetCount - 1
        End If

        .EventEnabled(EventAllEvents) = True
    End With

    '    CellChange Col, Row
End Sub

Sub fps_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Long, TipWidth As Long, TipText As String, ShowTip As Boolean)
    ' Set tip to display and set tip's content
      If Col = fps.ColLetterToNumber("G") Then
        ShowTip = True
        TipText = TipText
      End If
End Sub
' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 13
    denRowDL = 18
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "11_12_2~20_30_3"
    ' end
End Sub


' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("V"), 37, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function

' Lay ve chi tieu 29 o to khai lan truoc
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
            Set xmlNode = .nodeFromID("V_" & 37)
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_03_TD_TAIN", "getValueCTDC", Err.number, Err.Description
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
            '[28], [29]
            Set xmlNode = .nodeFromID("R_" & 37)
            result = "CT28_" & Val(GetAttribute(xmlNode, "Value"))
'            Set xmlNode = .nodeFromID("V_" & 37)
'            result = result & "~" & "CT29_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_03_TD_TAIN", "getValueTK", Err.number, Err.Description
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
    SaveErrorLog "03_TD_TAIN", "getMapCT", Err.number, Err.Description
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
                '[28], [29]
                .GetText .ColLetterToNumber("R"), 37, varTemp
                strCurrentValue = "CT28_" & Trim(varTemp)
'                .GetText .ColLetterToNumber("V"), 37, varTemp
'                strCurrentValue = strCurrentValue & "~CT29_" & Trim(varTemp)

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
                
                If Left(strMaCT, 2) = "28" Then
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



Public Sub UpdateSheet()
    Dim varTemp As Variant
    Dim temp As Integer
    Dim blCheck_S2 As Boolean
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    With fps
        .Sheet = 1
        temp = mCurrentSheet
        mCurrentSheet = 1
        For Each xmlCellNodeData In TAX_Utilities_v1.Data(0).getElementsByTagName("Cell")
            ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
            .GetText lCol, lRow, varTemp
            UpdateCell fps, lCol, lRow, varTemp
        Next
        
        ' Phu luc 1_2
        .Sheet = 2
        mCurrentSheet = 2
        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(1).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                 If (lRow >= 42 And lCol = .ColLetterToNumber("G")) Or (lRow >= 42 And lCol = .ColLetterToNumber("AI")) Then
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
        
        ' Update lai sheet KHBS
        .Sheet = .SheetCount - 1
        mCurrentSheet = .Sheet
        blCheck_S2 = IIf(TAX_Utilities_v1.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v1.Data(.Sheet - 1).getElementsByTagName("Cell")
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



' ham set thong tin sau khi tai XML
Public Sub finishImportXML()
    Dim i As Integer
    Dim CQT As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
         .Sheet = 2
         If .SheetVisible = True Then
                .Row = 42
                .Col = .ColLetterToNumber("B")
                i = 1
                Do
                     .GetText .ColLetterToNumber("P"), .Row, CQT
                     If Trim$(CStr(CQT)) = "" Or CQT = vbNullString Then
                        .Text = i
                        .Col = .ColLetterToNumber("G")
                        .value = 1
                        .Col = .ColLetterToNumber("P")
                        .Lock = True
                        i = i + 1
                     End If
                     .Col = .ColLetterToNumber("B")
                    .Row = .Row + 1
                Loop Until .Text = "aa"
                
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
    sheetKhbs = 3
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
