VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_04TBAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrSolanBosung As String
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
Dim strDataFileName As String
Dim arrDanhsach() As String
Dim strComboHien As String
Dim strCombo As String
Dim MSTDN As String
Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
Dim xmlNode As MSXML.IXMLDOMNode

strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_loai_ThongTinThayDoi.xml")
    ' testttt svn
    ' Lay danh muc loai hoa don
    ' 15/11/2010
    If xmlDomData.Load(strDataFileName) Then
        Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
        For Each xmlNode In xmlNodeListCell
            If GetAttribute(xmlNode, "Value") <> "" Then
                arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                'strCombo = strCombo + CPab(arrDanhsach(0), 10) + CPab(arrDanhsach(1), 10)  + CPab(arrDanhsach(2), 200) + Chr$(9)
                strCombo = strCombo + arrDanhsach(0) + Chr$(9)
                strComboHien = strComboHien + arrDanhsach(1) + Chr$(9)
            End If
        Next
    End If
    
     With fps
        SetDateFormat fps, 1, 30, .ColLetterToNumber("F"), DDMMYYYY
        
'        SetDateFormat fps, 1, 35, .ColLetterToNumber("AE"), DDMMYYYY
        'SetKyHieuHDFormat fps, 1, 35, .ColLetterToNumber("AA")
        .Sheet = 1

        ' Set gia tri ten loai hoa don
        .Col = .ColLetterToNumber("C")
        .Row = 35
        .TypeComboBoxList = strComboHien
        ' Set gia tri cho combo an
        .Col = .ColLetterToNumber("AA")
        .Row = 35
        .TypeComboBoxList = strCombo
        'set value ngay dieu chinh
        
        .Row = 47
        .Col = .ColLetterToNumber("Q")
        .Text = TAX_Utilities_v2.Day & "/" & TAX_Utilities_v2.Month & "/" & TAX_Utilities_v2.Year
        SetDateFormat fps, 1, 47, .ColLetterToNumber("Q"), DDMMYYYY

    End With
    
    
    
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    'bien ten dvcq
    Dim strDvcq As String
    Dim strTenDvcq As String
    Dim countRowDynamic As Variant
    Dim nguoiky As Variant
    Dim strId As Variant
    Dim indexCombo As Integer
    
    Dim xmlCellNode As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim txmlCellNode As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
    Dim xmlNode As MSXML.IXMLDOMNode
    Dim iCol As Long, iRow As Long
    
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
         mCurrentSheet = .Sheet
        .Row = 35
        Do
            countRowDynamic = countRowDynamic + 1
            .Col = .ColLetterToNumber("B")
            .Row = .Row + 1
        Loop Until .Text = "aa"
        'lay ten va MST DVCQ tren Header
            strDvcq = GetDVCQ()
            strTenDvcq = GetTenDVCQ()
            .SetText .ColLetterToNumber("F"), countRowDynamic + 41, strDvcq
            .SetText .ColLetterToNumber("F"), countRowDynamic + 39, strTenDvcq
        'lay ten nguoi ky tren header
        .Row = countRowDynamic + 48
        .Col = .ColLetterToNumber("Q")
        If Trim(.Text) = "" Then
            nguoiky = GetNguoiKy()
            .SetText .ColLetterToNumber("Q"), countRowDynamic + 48, nguoiky
        End If
        
         'Xoa cong thuc cua o Chu ky
        Set xmlCellNode = TAX_Utilities_v2.Data(0).nodeFromID(GetCellID(fps, .ColLetterToNumber("D"), 35))
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode
        Set xmlCellNode = xmlCellNode.parentNode.lastChild
        Set xmlCellNode = xmlCellNode.firstChild
        Set xmlCellNode = xmlCellNode.lastChild
        ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
        .Row = iRow
        .Col = iCol
        .Formula = ""
        
        ' Set format ngay ky
        .Row = countRowDynamic + 46
        .Col = .ColLetterToNumber("Q")
        '.Text = Format(Date, "dd/mm/yyyy")
        .Text = TAX_Utilities_v2.Day & "/" & TAX_Utilities_v2.Month & "/" & TAX_Utilities_v2.Year
        .TypeHAlign = TypeHAlignLeft
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    Dim strId                    As Variant
    Dim intIndexCombo            As Integer

    Dim iFlagExist               As Integer, iFlagNgayCoThongBaoPhatHanh As Integer, iFlagEmpty As Integer ',iFlagNgayThongBaoEmpty As Integer
    Dim i                        As Integer, j As Integer, iCol As Long, iRow As Long
    Dim vTarget22                As Variant, vTarget20 As Variant, vTarget21 As Variant, vPhatHanh As Variant, vDieuChinh As Variant, arrDate() As String
    Dim countRow                 As Long, rowDes As Long, vErrExist As Variant
    Dim vGiaTriThayDoiCheckTrung As Variant, vGiaTriThayDoi As Variant
    Dim tRow As Long, tCol As Long, mblLoi As Boolean

    fps.EventEnabled(EventAllEvents) = False

    With fps
        
        .Col = Col
        .Row = Row

        If Col = .ColLetterToNumber("C") And .Text <> "" Then
            intIndexCombo = .typeComboboxCurSel
            .Col = .ColLetterToNumber("AA")
            .typeComboboxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber("AA")
            UpdateCell fps, .Col, .Row, .Text
            
            .Sheet = .SheetCount
            .GetText .ColLetterToNumber("E"), 20, vTarget20 ' du lieu bat buoc phai nhap
            .GetText .ColLetterToNumber("E"), 21, vErrExist
            .GetText .ColLetterToNumber("E"), 22, vTarget22
            
            '2. Check Exist
            .Sheet = 1
            '.Row = 35
'            Do
'                .Col = .ColLetterToNumber("C")
'                vGiaTriThayDoi = .Text
'                vGiaTriThayDoiCheckTrung = vGiaTriThayDoi
'
'                If vGiaTriThayDoi <> vbNullString Then
'                    countRow = .SearchCol(.ColLetterToNumber("C"), -1, -1, vGiaTriThayDoiCheckTrung, SearchFlagsNone)
'                    rowDes = .SearchCol(.ColLetterToNumber("C"), countRow, -1, vGiaTriThayDoiCheckTrung, SearchFlagsNone)
'
'                    If rowDes > -1 Then
'                        .Col = .ColLetterToNumber("C")
'                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                        .CellNote = vErrExist      'static
'                        .BackColor = mErrorColor
'
'                        iFlagExist = 1
'                    Else
'                        .Col = .ColLetterToNumber("C")
'                        .BackColor = mNonErrorColor
'                        .CellNote = ""
'                        .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    End If
'
'                Else
'                    .Col = .ColLetterToNumber("C")
'                    .BackColor = mNonErrorColor
'                    .CellNote = ""
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    'iFlagEmpty = 1
'                End If
'
'                .Row = .Row + 1
'                .Col = .ColLetterToNumber("B")
'            Loop Until .Text = "aa"
'
'            .SetText .ColLetterToNumber("B"), 21, IIf(iFlagExist = 1, "0", "1")
                
             If blGiaTriCu And Col = .ColLetterToNumber("C") Then 'And Not mblKeyF5 Then
                If Not .Visible Then
                    .Visible = True
                    DisplayMessage "0270", msOKOnly, miWarning
                    .Visible = False
                Else
                    DisplayMessage "0270", msOKOnly, miWarning
                End If
                
                .Sheet = 1
                mCurrentSheet = .Sheet
                .Row = RowGiaTriCu
                .Col = ColGiaTriCu
                .Text = strGiaTriCu
                .Sheet = 1
                mCurrentSheet = .Sheet
                .Row = RowGiaTriCu
                .Col = ColGiaTriCu
                UpdateCell fps, .Col, .Row, .Text
                blGiaTriCu = False
                mblLoi = True
               
            End If
                
            .EventEnabled(EventChange) = True
        End If

    End With

    fps.EventEnabled(EventAllEvents) = True
End Sub

Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    Dim temp As Variant, itemp As Variant, iRow As Long, iCol As Long, iColRow As String
    Dim i    As Integer, aSection As Integer
    
    With fps
        .Sheet = 1
        .Row = 35

        If Col = fps.ColLetterToNumber("C") Then
            blGiaTriCu = False
            fps.Col = Col
            fps.Row = Row
            temp = fps.Text

            Do
                .Col = .ColLetterToNumber("C")
                itemp = .Text

                If .Row <> Row Then
                    If itemp = temp Then 'Neu co thi dat co blGiaTriCu = True
                        RowGiaTriCu = Row
                        ColGiaTriCu = .ColLetterToNumber("C")
                        .Row = Row
                        .Col = .ColLetterToNumber("C")
                        blGiaTriCu = True
                        Exit Sub
                    End If
                End If

                .Col = .ColLetterToNumber("B")
                .Row = i + iRow
                i = i + 1
            Loop Until .Text = "aa"

        End If

    End With

End Sub

Private Sub fps_KeyPress(KeyAscii As Integer)
     With fps
        If .Col = .ColLetterToNumber("P") Or .Col = .ColLetterToNumber("R") Then
            If (KeyAscii < 48 And KeyAscii <> 3 And KeyAscii <> 22 And KeyAscii <> 13 And KeyAscii <> 8) Or KeyAscii > 57 Then
                KeyAscii = 0
            End If
        End If
    End With
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Integer, j As Integer, iCol As Long, iRow As Long
    'Dim curRow As String
    'Dim varTemp As Variant
    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
            If .ActiveSheet = 1 Then
                fps.EventEnabled(EventAllEvents) = True
                i = 1
                .Col = .ColLetterToNumber("B")
                .Row = 35

                Do
                    .Text = str(i)
                    .Col = .ColLetterToNumber("B")
                    .Row = i + 35
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
    Dim tuSo As Variant
    Dim denso As Variant
    Dim strId As Variant
    Dim strLoai As Variant
    Dim strTemp As Variant
    Dim countRowDynamic As Integer
    Dim strSoHieu As Variant

    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("F") And Row >= 30 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    'If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                              .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    'End If
                Else
                 .SetText Col, Row, ""
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

Private Sub CheckDynamicError()
    Dim iFlagExist               As Integer, iFlagNgayCoThongBaoPhatHanh As Integer, iFlagEmpty As Integer ', iFlagDaiDien As Integer  ',iFlagNgayThongBaoEmpty As Integer
    Dim vTarget22                As Variant, vTarget20 As Variant, vTarget21 As Variant, vPhatHanh As Variant, vDieuChinh As Variant, arrDate() As String, vDaiDienPL As Variant
    Dim countRow                 As Long, rowDes As Long, vErrExist As Variant, iCountChange As Integer
    Dim vGiaTriThayDoiCheckTrung As Variant, vGiaTriThayDoi As Variant

    With fps
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 20, vTarget20 ' du lieu bat buoc phai nhap
        .GetText .ColLetterToNumber("E"), 21, vErrExist
        .GetText .ColLetterToNumber("E"), 22, vTarget22
         
        '1. Check Ngay dieu chinh
        .Sheet = 1
        .GetText .ColLetterToNumber("F"), 30, vPhatHanh

        If (Len(vPhatHanh) > 0) Then
            arrDate = Split(CStr(vPhatHanh), "/")
            If DateSerial(CInt(arrDate(2)), CInt(arrDate(1)), CInt(arrDate(0))) >= DateSerial(CInt(TAX_Utilities_v2.Year), CInt(TAX_Utilities_v2.Month), CInt(TAX_Utilities_v2.Day)) Then
                iFlagNgayCoThongBaoPhatHanh = 1
                .Col = .ColLetterToNumber("F")
                .Row = 30
                .BackColor = mErrorColor
                .CellNote = vTarget22
            Else
                .Col = .ColLetterToNumber("F")
                .Row = 30
                .CellNote = ""  'static
                .BackColor = mNonErrorColor
            End If

        Else
            iFlagEmpty = 1
            .Col = .ColLetterToNumber("F")
            .Row = 30
            .BackColor = mErrorColor
            .CellNote = vTarget20
        End If

        '2. Check Exist
        .Sheet = 1
        .Row = 35
        iCountChange = 0
        Do
            .Col = .ColLetterToNumber("C")
            vGiaTriThayDoi = .Text
            vGiaTriThayDoiCheckTrung = vGiaTriThayDoi

            If vGiaTriThayDoi <> vbNullString Then
                countRow = .SearchCol(.ColLetterToNumber("C"), -1, -1, vGiaTriThayDoiCheckTrung, SearchFlagsNone)
                rowDes = .SearchCol(.ColLetterToNumber("C"), countRow, -1, vGiaTriThayDoiCheckTrung, SearchFlagsNone)

                If rowDes > -1 Then
                    .Col = .ColLetterToNumber("C")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrExist      'static
                    .BackColor = mErrorColor

                    iFlagExist = 1
                Else
                    .Col = .ColLetterToNumber("C")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                End If

            Else
                .Col = .ColLetterToNumber("C")
                .BackColor = mNonErrorColor
                .CellNote = ""
                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                'iFlagEmpty = 1
            End If
            
            iCountChange = iCountChange + 1
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        
        '3.validate
        .Sheet = 1
        .Row = 35
        Do
            .Col = .ColLetterToNumber("C")
            If (.Text <> vbNullString) Then
                .Col = .ColLetterToNumber("D")
                If (Trim(.Text) = vbNullString) Then
                    .Col = .ColLetterToNumber("D")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vTarget20      'static
                    .BackColor = mErrorColor
                    iFlagEmpty = 1
                Else
                    .Col = .ColLetterToNumber("D")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                End If
            End If
            .Col = .ColLetterToNumber("C")
            If (.Text <> vbNullString) Then
                .Col = .ColLetterToNumber("M")
                If (Trim(.Text) = vbNullString) Then
                    .Col = .ColLetterToNumber("M")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vTarget20      'static
                    .BackColor = mErrorColor
                    iFlagEmpty = 1
                Else
                    .Col = .ColLetterToNumber("M")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                End If
            End If
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        
        '4 nguoi dai dien phap luat
        .Sheet = 1
        .GetText .ColLetterToNumber("Q"), 49 + iCountChange - 1, vDaiDienPL

        If vDaiDienPL = vbNullString Or vDaiDienPL = "" Then
            iFlagEmpty = 1
            .Col = .ColLetterToNumber("Q")
            .Row = 49 + iCountChange - 1
            .BackColor = mErrorColor
            .CellNote = vTarget20
        Else
            'iFlagEmpty = 0
            .Col = .ColLetterToNumber("Q")
            .Row = 49 + iCountChange - 1
            .CellNote = ""  'static
            .BackColor = mNonErrorColor
        End If
        
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 20, IIf(iFlagEmpty = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 21, IIf(iFlagExist = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 22, IIf(iFlagNgayCoThongBaoPhatHanh = 1, "0", "1")
    End With

End Sub
Sub CheckErrorMST()
    Dim vError1       As Variant, vError2 As Variant, vError3 As Variant
    Dim MST1          As Variant, MST2 As Variant, MST3 As Variant, MST4 As Variant
    Dim MST5          As Variant, MST6 As Variant, MST7 As Variant, MST8 As Variant
    Dim MST9          As Variant, MST10 As Variant, MST11 As Variant, MST12 As Variant, MST13 As Variant
    Dim iCurrentSheet As Integer, strCheck As String
    Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer

    With fps
        
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 15, vError1
        .GetText .ColLetterToNumber("E"), 16, vError2
        .GetText .ColLetterToNumber("E"), 17, vError3
        
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
            .Row = 15
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 15
            .Text = "1"
        End If
        
        If iFlagTaxCode2 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 16
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 16
            .Text = "1"
        End If
        
        If iFlagTaxCode3 = 1 Then
            .Col = .ColLetterToNumber("B")
            .Row = 17
            .Text = "0"
        Else
            .Col = .ColLetterToNumber("B")
            .Row = 17
            .Text = "1"
        End If

        .Sheet = 1
        .Col = .ColLetterToNumber("E")
        .Row = 4
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

Private Sub kiemTraDuLieuImport()

End Sub

Public Sub ResetErrorCells()

End Sub

Public Function ResetData() As Boolean
    Dim totalCell  As Long
    Dim countRowReset As Long
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol As Long, lRow As Long
    totalCell = TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length
    countRowReset = (totalCell - 5) / 14
    With fps
        For Each xmlNodeReset In TAX_Utilities_v2.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            .Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            If (lRow < 35 Or lRow > 35 + countRowReset) And mCurrentSheet = 1 Then
            Else
                .Col = lCol
                .Row = lRow
                
                Select Case .CellType
                    Case CellTypeCheckBox
                        .Text = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                    Case CellTypeComboBox
                        .Text = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                    Case CellTypeNumber
                        .value = 0
                        UpdateCell fps, lCol, lRow, "0"
                    Case Else
                        .value = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                End Select
                'mAdjustData = IIf(IsUpdate = True, IsUpdate, mAdjustData)
                TAX_Utilities_v2.AdjustData(mCurrentSheet - 1) = True
            End If
        Next
    End With
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


Public Function GetDVCQ() As String
'    Dim xmlDomHeader As New MSXML.DOMDocument
'
'    xmlDomHeader.Load GetAbsolutePath(TAX_Utilities_v2.DataFolder & "Header_01.xml")
'    GetDVCQ = GetAttribute(xmlDomHeader.nodeFromID("C_34"), "Value")
'
'    Set xmlDomHeader = Nothing
    Dim xmlDomHeader As New MSXML.DOMDocument
    Dim strMST As String
    
    xmlDomHeader.Load GetAbsolutePath(TAX_Utilities_v2.DataFolder & "Header_01.xml")
    strMST = GetAttribute(xmlDomHeader.nodeFromID("C_34"), "Value")
    If Len(strMST) = 13 Then
        strMST = Left$(strMST, 10) & "-" & Right$(strMST, 3)
    End If
    'GetDVCQ = GetAttribute(xmlDomHeader.nodeFromID("C_34"), "Value")
    GetDVCQ = strMST
    
    Set xmlDomHeader = Nothing
End Function
Public Function GetTenDVCQ() As String
    Dim xmlDomHeader As New MSXML.DOMDocument
    
    xmlDomHeader.Load GetAbsolutePath(TAX_Utilities_v2.DataFolder & "Header_01.xml")
    GetTenDVCQ = GetAttribute(xmlDomHeader.nodeFromID("C_32"), "Value")
    
    Set xmlDomHeader = Nothing
End Function
Public Function GetNguoiKy() As String
    Dim xmlDomHeader As New MSXML.DOMDocument
    
    xmlDomHeader.Load GetAbsolutePath(TAX_Utilities_v2.DataFolder & "Header_01.xml")
    GetNguoiKy = GetAttribute(xmlDomHeader.nodeFromID("C_26"), "Value")
    
    Set xmlDomHeader = Nothing
End Function
