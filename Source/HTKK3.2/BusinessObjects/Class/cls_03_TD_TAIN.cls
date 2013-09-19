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
    
    With fps
     SetDateFormat fps, 1, 63, fps.ColLetterToNumber("O"), DDMMYYYY
    With fps
        .Sheet = 1
        .Row = 63
        .Col = fps.ColLetterToNumber("O")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
    End With
        'SetDateFormat fps, 1, 19, .ColLetterToNumber("AJ"), DDMMYYYY
    End With
    setCQTQuanLy
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
                        tenCucThue = tenCucThue + arrDanhsach(3) + Chr$(9)
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
            .Col = .ColLetterToNumber("L")
            .TypeComboBoxList = tenCucThue
            
            .Row = rowStatic
            .Col = .ColLetterToNumber("AG")
            .TypeComboBoxList = maCucThue
            
            .Row = rowStatic
            .Col = .ColLetterToNumber("P")
            .TypeComboBoxList = tenChiCucThue
    
            .Row = rowStatic
            .Col = .ColLetterToNumber("AH")
            .TypeComboBoxList = maChiCucThue
            
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
            
            If Col = .ColLetterToNumber("L") Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot L va gia tri Index cot AG
                intIndexCuc = .TypeComboBoxCurSel
                .Col = .ColLetterToNumber("AG")
                .Row = Row
                .TypeComboBoxCurSel = intIndexCuc
                'Nhap lai Combo cot AG
                .Col = .ColLetterToNumber("AG")
                UpdateCell fps, .Col, .Row, .Text
                ' Lay Id cua cuc thue vua chon o tren
                fps.GetText .Col, .Row, varIdCucThue
                
                If varIdCucThue <> "" Or varIdCucThue <> vbNullString Then
                    ' Lay thong tin cho chi cuc thue
                    varIdCucThue = Left$(varIdCucThue, 3)
                    
                    If xmlDomData.Load(GetAbsolutePath("..\InterfaceIni\Catalogue_Tinh_Thanh.xml")) Then
                        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
                        For Each xmlNode In xmlNodeListCell
                            If GetAttribute(xmlNode, "Value") <> "" Then
                                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                                
                                If arrDanhsach(0) = "0" And arrDanhsach(2) = varIdCucThue Then
                                    tenChiCucThue = tenChiCucThue + arrDanhsach(3) + Chr$(9)
                                    maChiCucThue = maChiCucThue + arrDanhsach(1) + Chr$(9)
                                End If
                            End If
                        Next
                        Set xmlDomData = Nothing
                        Set xmlNodeListCell = Nothing
                        Set xmlNode = Nothing
                    End If
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("P")
                    .TypeComboBoxList = tenChiCucThue
                    .Text = ""
                    
                    .Row = Row
                    .Col = .ColLetterToNumber("AH")
                    .TypeComboBoxList = maChiCucThue
                    .Text = ""
                End If
            End If
                        
            If Col = .ColLetterToNumber("P") Then 'Khi chon Combo
                'Chuyen sang du lieu tuong ung Lay gia tri index cot C va gia tri Index cot Q
                intIndexChiCuc = .TypeComboBoxCurSel
                .Col = .ColLetterToNumber("AH")
                .Row = Row
                .TypeComboBoxCurSel = intIndexChiCuc
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
            
        .ActiveSheet = ASheet
        .Sheet = SSheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

'Cause this interface sheets have dynamic rows, this function will set cellnote of error for rising-error cell
'no parameter
Private Sub CheckDynamicError()

    Dim iCurrentSheet As Integer, iFlagEmpty As Integer, iWhiteRow As Integer
    Dim vTenCoQuan    As Variant, vCapCuc As Variant, vCapDiaPhuong As Variant, vMessRequire As Variant
    Dim row_require   As Integer
    Dim vSTT As Variant, vTenNhaMay As Variant, vCQTCapCuc As Variant, vCQTCapDiaPhuong As Variant, vTyLePhanBo As Variant, vSoThuePhaiNop As Variant
    
    row_require = 15
    
    With fps
        iCurrentSheet = .Sheet
        'get message
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), row_require, vMessRequire
        
        .Sheet = iCurrentSheet
        .Sheet = 2
        If .SheetVisible = True Then
            mCurrentSheet = .Sheet
            'check phu luc 01_2
            .Row = 42
            iWhiteRow = 1
            
            Do
                .GetText .ColLetterToNumber("B"), .Row, vSTT
                .GetText .ColLetterToNumber("C"), .Row, vTenNhaMay
                .GetText .ColLetterToNumber("L"), .Row, vCQTCapCuc
                .GetText .ColLetterToNumber("P"), .Row, vCQTCapDiaPhuong
                .GetText .ColLetterToNumber("R"), .Row, vTyLePhanBo
                .GetText .ColLetterToNumber("U"), .Row, vSoThuePhaiNop
                If Trim$(vSTT) <> "" Or Trim$(vTenNhaMay) <> "" Or Trim$(vCQTCapCuc) <> "" Or Trim$(vCQTCapDiaPhuong) <> "" Or Trim$(vTyLePhanBo) <> "0" Or Trim$(vSoThuePhaiNop) <> "0" Then
                    iWhiteRow = 0
                End If
            
                .Col = .ColLetterToNumber("C")
                If (Trim(.Text) = vbNullString And iWhiteRow = 0) Then
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vMessRequire      'static
                    .BackColor = mErrorColor
    
                    iFlagEmpty = 1
                Else
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                End If
    
                .Col = .ColLetterToNumber("L")
                If (Trim(.Text) = vbNullString And iWhiteRow = 0) Then
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
                If (Trim(.Text) = vbNullString And iWhiteRow = 0) Then
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vMessRequire      'static
                    .BackColor = mErrorColor
    
                    iFlagEmpty = 1
                Else
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                End If
    
                .Row = .Row + 1
                .Col = .ColLetterToNumber("B")
            Loop Until .Text = "aa"
        End If
        
        'set flag
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), row_require, IIf(iFlagEmpty = 1, "0", "1")
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
    For Each xmlNodeValid In TAX_Utilities_New.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
            Case "03-1"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 1)
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
                Set xmlCellNode = TAX_Utilities_New.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                Set xmlCellNode = TAX_Utilities_New.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(sheetPl01_KHBS_GTGT - 1).getElementsByTagName("Cell").length - 4)
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
            For Each xmlCellNodeData In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
                .GetText lCol, lRow, varTemp
                UpdateLastKHBSCell fps, lCol, lRow, varTemp
            Next
       Else
            For Each xmlCell In TAX_Utilities_New.Data(1).getElementsByTagName("Cell")
               If Left(GetAttribute(xmlCell, "CellID"), 2) = "BC" Then
                    strId = Trim(Mid(GetAttribute(xmlCell, "Value"), 200, 10))
                        
                    For Each xmlCellGoc In TAX_Utilities_New.DataKHBS.getElementsByTagName("Cell")
                        If Left(GetAttribute(xmlCellGoc, "CellID"), 1) = "E" Then
                            If GetAttribute(xmlCellGoc, "Value") = strId Then
                            'update secsion
                               Set xmlCellUpdate = seachCellNode(TAX_Utilities_New.Data(0).getElementsByTagName("Cell"), GetAttribute(xmlCellGoc, "CellID"))
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

Private Sub filldata()
    Dim varTemp         As Variant
    Dim lCol            As Long, lRow As Long
    Dim xmlCellNode     As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    Dim arrDanhsach()   As String
    Dim xmlCell         As MSXML.IXMLDOMNode

    With fps
        Dim i           As Integer
        Dim strFormular As String
        ReDim arrDanhSachKHBS(TAX_Utilities_New.Data(0).getElementsByTagName("Section")(1).childNodes.length - 1)
        i = 0

        For Each xmlCellNodeData In TAX_Utilities_New.DataKHBS.getElementsByTagName("Cell")

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

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    setChangeComb Col, Row
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
            If .ActiveSheet = .SheetCount - 1 Then
                'fps_Change iCol, iRow
                fps.EventEnabled(EventAllEvents) = False
                i = 1
                .Col = .ColLetterToNumber("B")
                .Row = 9

                Do
                    .Text = str(i)
                    .Col = .ColLetterToNumber("B")
                    .Row = i + 9
                    i = i + 1
                Loop Until .Text = "aa"

                i = 1
                .Row = .Row + 3

                Do
                    .Text = str(i)
                    .Col = .ColLetterToNumber("B")
                    .Row = 1 + .Row
                    i = i + 1
                Loop Until .Text = "bb"

                .SetActiveCell iCol, iRow
                .Row = iRow
                .Col = iCol
                .SetActiveCell iCol, iRow
                fps.EventEnabled(EventAllEvents) = True
            End If
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
        End If

        fps.EventEnabled(EventAllEvents) = True
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

            If Col = .ColLetterToNumber("O") And Row = 63 + count - 1 Then
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
