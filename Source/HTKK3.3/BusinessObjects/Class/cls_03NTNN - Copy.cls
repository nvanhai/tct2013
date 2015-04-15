VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_03NTNN"
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
Dim mblKeyF5 As Boolean
Dim mblKeyF6 As Boolean
Public loaiKHBS As String

Public FlagThaydoi As String


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
    Dim arrNganhKD    As String
    Dim arrNganhKD_TL As String
    Dim vMaNganhKD As String
    Dim nganhKD()     As String
    Dim xmlDocument   As New MSXML.DOMDocument
    Dim xmlNode1      As MSXML.IXMLDOMNode

    xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_TyLeThue_TNDN.xml"))
    Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
        
    Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")

    For Each xmlNode1 In xmlNodeListMap

        nganhKD = Split(GetAttribute(xmlNode1, "Value"), "###")
        arrNganhKD = arrNganhKD + nganhKD(1) + Chr$(9)
        arrNganhKD_TL = arrNganhKD_TL + nganhKD(2) + Chr$(9)
        vMaNganhKD = vMaNganhKD + nganhKD(0) + Chr$(9)
    Next

    With fps
        .Sheet = 1
'        'day du lieu vao combobox
'        .Col = .ColLetterToNumber("C")
'        .Row = 34
'        .TypeComboBoxList = arrNganhKD
'        .Col = .ColLetterToNumber("CR")
'        .Row = 34
'        .TypeComboBoxList = arrNganhKD_TL
        
'        .Row = 34
'        .Col = .ColLetterToNumber("CQ")
'        .TypeComboBoxList = vMaNganhKD
        
        'format ngay thang
        SetDateFormat fps, 1, 50, .ColLetterToNumber("BG"), DDMMYYYY
        SetDateFormat fps, 1, 34, .ColLetterToNumber("BE"), DDMMYYYY
        
        .Row = 26
        .Col = .ColLetterToNumber("F")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        .Row = 50
        .Col = .ColLetterToNumber("BG")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        .Sheet = .SheetCount - 1
        .Row = 23
        .Col = .ColLetterToNumber("BG")
        .Text = Format(Date, "dd/mm/yyyy")
         SetDateFormat fps, .Sheet, 23, .ColLetterToNumber("BG"), DDMMYYYY
         .TypeHAlign = TypeHAlignLeft
    End With

End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim ngay            As Variant
    Dim kythang         As Variant
    Dim kyngay          As Variant
    Dim xmlCellNode     As MSXML.IXMLDOMNode, xmlCellsNode As MSXML.IXMLDOMNode
    Dim txmlCellNode    As MSXML.IXMLDOMNode, txmlCellsNode As MSXML.IXMLDOMNode
    Dim tongNTNN        As Variant
    Dim tienDoiThanhChu As String
    Dim i As Integer
    Dim vChitieu6 As Variant
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        ngay = TAX_Utilities_New.Day
        
        
        If Trim(ngay) = "" Then
            .SetText .ColLetterToNumber("S"), 26, ""
            UpdateCell fps, .ColLetterToNumber("S"), 26, ""
        Else
            .SetText .ColLetterToNumber("S"), 26, ngay & "/" & TAX_Utilities_New.Month & "/" & TAX_Utilities_New.Year
            UpdateCell fps, .ColLetterToNumber("S"), 26, ngay & "/" & TAX_Utilities_New.Month & "/" & TAX_Utilities_New.Year
        End If
        
        ' doi tien sang chu
        .GetText .ColLetterToNumber("U"), 26, tongNTNN

        If CStr(tongNTNN) <> "" Then
            tongNTNN = Replace(tongNTNN, ".", "")
            tongNTNN = Replace(tongNTNN, ",", "")
            tienDoiThanhChu = doiTien(CStr(tongNTNN))
            .SetText .ColLetterToNumber("AD"), 38, tienDoiThanhChu
            'UpdateCell fps, .ColLetterToNumber("U"), 26, tienDoiThanhChu
            UpdateCell fps, .ColLetterToNumber("AD"), 38, tienDoiThanhChu
        End If
        
        '- set To khai thang

        If StrTKThang_PS = "TK_THANG" Then
            .Col = .ColLetterToNumber("P")
            .Row = 26
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("Q")
            .Row = 26
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text

            ' set lan dau hoac lan bo sung
            If Strloaitk = "TKCT" Then
                .Col = .ColLetterToNumber("M")
                .Row = 26
                .Text = 1
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("O")
                .Row = 26
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                
                
            ElseIf Strloaitk = "TKBS" Then
                .Col = .ColLetterToNumber("M")
                .Row = 26
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("O")
                .Row = 26
                .Text = StrSolanBosung
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            
            
            '- Set to khai lan bo sung
        ElseIf StrTKThang_PS = "TK_LANPS" Then
            .Col = .ColLetterToNumber("P")
            .Row = 26
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("Q")
            .Row = 26
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("M")
            .Row = 26
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("O")
            .Row = 26
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            

        End If
      
        'end
        ' Kiem tra neu chi tieu 6 <0 thi 8, 9 khong dc sua
        .Sheet = 1
        .Row = 34
        .Col = .ColLetterToNumber("B")
        i = 1
        Do
            .GetText .ColLetterToNumber("BK"), .Row, vChitieu6
            If Val(vChitieu6) < 0 Then
                ' Lock chi tieu 8
                .Col = .ColLetterToNumber("BX")
                .Lock = True
                ' Lock chi tieu 9
                .Col = .ColLetterToNumber("CG")
                .Lock = True
            End If
            .Row = 34 + i
            i = i + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"
        
        .EventEnabled(EventAllEvents) = True
    End With

End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    Dim strId As Variant
    Dim strMa As Variant
    Dim intIndexCombo As Integer
    
    Dim strValue As Variant
    
    fps.EventEnabled(EventAllEvents) = False

    With fps
        .Sheet = 1
            .Col = .ColLetterToNumber("C")
            intIndexCombo = .TypeComboBoxCurSel
            .Col = .ColLetterToNumber("CR")
            .TypeComboBoxCurSel = intIndexCombo
            .GetText .ColLetterToNumber("CR"), Row, strId
            
            .Col = .ColLetterToNumber("CQ")
            .TypeComboBoxCurSel = intIndexCombo
            .GetText .ColLetterToNumber("CQ"), Row, strMa
            
            .Col = .ColLetterToNumber("BU")
            
            If Col = .ColLetterToNumber("C") Then
                If (Trim(strId) <> vbNullString) Then
                    .Text = Trim(strId)
                    UpdateCell fps, .Col, Row, .Text
                End If
'
'                ' Lock cot thue suat khong cho sua doi voi truong hop khac tien vay lai
'                If Trim$(strMa) = "08" Then
'                      .Col = .ColLetterToNumber("BU")
'                      .Lock = False
'                Else
'                      .Col = .ColLetterToNumber("BU")
'                      .Lock = True
'                End If
            End If
            
            ' Luu ma cua nganh nghe kinh doanh
            .Col = .ColLetterToNumber("CQ")
            .Row = Row
            .TypeComboBoxCurSel = intIndexCombo
            UpdateCell fps, .Col, Row, .Text
            
            ' Set CT khi thay doi cot 6,7,8
            If (Col = .ColLetterToNumber("BK") Or Col = .ColLetterToNumber("C") Or Col = .ColLetterToNumber("BU") Or Col = .ColLetterToNumber("BX")) And Row >= 34 And Row <= .MaxRows - 17 Then
                .Col = .ColLetterToNumber("CG")
                .Row = Row
                .Formula = "IF(ROUND((BK" & Row & "*BU" & Row & "/100),0)-BX" & Row & "<0,0,ROUND((BK" & Row & "*BU" & Row & "/100),0)-BX" & Row & ")"
                UpdateCell fps, .Col, Row, .Text
                .Formula = ""
            End If
            
            ' Neu chi tieu 6 <0 set chi tieu 8 =0 va khong cho sua
            If (Col = .ColLetterToNumber("BK")) And Row >= 34 And Row <= .MaxRows - 17 Then
                .GetText .ColLetterToNumber("BK"), Row, strValue
                If IsNumeric(strValue) Then
                    If Val(strValue) < 0 Then
                        ' Neu chi tieu 6 <0 set chi tieu 8 =0 va khong cho sua
                        .Col = .ColLetterToNumber("BX")
                        .Row = Row
                        .Text = "0"
                        .Lock = True
                        UpdateCell fps, .Col, Row, .Text
                        ' Neu chi tieu 6 <0 set chi tieu 9 =0 va khong cho sua
                        .Col = .ColLetterToNumber("CG")
                        .Row = Row
                        .Text = "0"
                        .Lock = True
                        UpdateCell fps, .Col, Row, .Text
                    Else
                        ' chi tieu 8
                        .Col = .ColLetterToNumber("BX")
                        .Row = Row
                        .Lock = False
                        ' chi tieu 9
                        .Col = .ColLetterToNumber("CG")
                        .Row = Row
                        .Lock = False
                    End If
                End If
            End If
            
            '
            
    End With

    fps.EventEnabled(EventAllEvents) = True
End Sub

Private Sub fps_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim iCol As Long, iRow As Long

    With fps
        iCol = .ActiveCol
        iRow = .ActiveRow
        GetCellSpan fps, iCol, iRow
        
        If (KeyCode = vbKeyF5) Or (KeyCode = vbKeyF6) Then
                  
            If .ActiveSheet = 1 Then
                fps.EventEnabled(EventAllEvents) = True
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

Private Sub fps_LeaveCell(ByVal Col As Long, _
                          ByVal Row As Long, _
                          ByVal NewCol As Long, _
                          ByVal NewRow As Long, _
                          Cancel As Boolean)
    Dim varTemp         As Variant
    'lay gia tri tong cua chi tieu 14
    Dim tienDoiThanhChu As String
    Dim tongNTNN        As Variant
    Dim xmlCellNode     As MSXML.IXMLDOMNode
    Dim iCol            As Long, iRow As Long
    Dim iCol1           As Long, iRow1 As Long

    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet

        If .ActiveSheet = 1 Then
        
            Set xmlCellNode = TAX_Utilities_New.Data(0).nodeFromID(GetCellID(fps, .ColLetterToNumber("C"), 34))
            If xmlCellNode Is Nothing Then Exit Sub
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode.childNodes(3)
            Set xmlCellNode = xmlCellNode.firstChild
            Set xmlCellNode = xmlCellNode.firstChild
            Set xmlCellNode = xmlCellNode.nextSibling
            Set xmlCellNode = xmlCellNode.nextSibling
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol1, iRow1
            Set xmlCellNode = xmlCellNode.nextSibling

            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
            ' doi tien sang chu
            .GetText iCol, iRow, tongNTNN
            tongNTNN = Replace(tongNTNN, ".", "")
            tongNTNN = Replace(tongNTNN, ",", "")
            tienDoiThanhChu = doiTien(CStr(tongNTNN))
       
            .SetText iCol1, iRow1, tienDoiThanhChu
            UpdateCell fps, iCol1, iRow1, tienDoiThanhChu
           
            If Col = .ColLetterToNumber("BE") Then
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
            
            Set xmlCellNode = TAX_Utilities_New.Data(0).nodeFromID(GetCellID(fps, .ColLetterToNumber("C"), 34))
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode.lastChild
            Set xmlCellNode = xmlCellNode.firstChild
            Set xmlCellNode = xmlCellNode.firstChild
            Set xmlCellNode = xmlCellNode.nextSibling
            Set xmlCellNode = xmlCellNode.nextSibling
            Set xmlCellNode = xmlCellNode.nextSibling

            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow

            If (Col = iCol And Row = iRow) Then
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
    Dim strFocusSheetName, strFocusRow, strFocusCol, strFormula As String
    Dim strNgayTT, i, vErrMsgMST, vErrMsgNgayHD, vErrMsgSoTien, noidung, sotien As Variant
    Dim vAlertThueTNDN_PN, vThueTNDN_PN As Variant
    Dim vErrThueTNDN_PN As Variant
    Dim iFlagvThueTNDN_PN As Integer
    Dim vErrCT678 As Variant
    Dim iFlagvCT678 As Integer
    
    Dim iFlagThueSuat As Integer
    Dim vErrorThueSuat As Variant
    
    Dim iFocusFlag, iFlagvMst As Integer
    Dim ngayHD As String
    
    If TAX_Utilities_New.NodeValidity.childNodes(0).Attributes.getNamedItem("Active").nodeValue <> "0" Then
        CheckErrorMST2 34, 17, "AD", 1
        compareTowDate 1, 34, "BE", 15
    End If

    With fps
       
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 19, vErrMsgNgayHD
        .GetText .ColLetterToNumber("E"), 20, vErrMsgSoTien
        .GetText .ColLetterToNumber("E"), 21, vAlertThueTNDN_PN
        .GetText .ColLetterToNumber("E"), 22, vErrThueTNDN_PN
        .GetText .ColLetterToNumber("E"), 23, vErrCT678
        .GetText .ColLetterToNumber("E"), 24, vErrorThueSuat
        
        
        .Sheet = 1

        .Col = .ColLetterToNumber("B")
        .Row = 34
        i = 1
        Dim col2 As Variant
        Dim col3 As Variant
        Dim col4 As Variant
        Dim col5 As Variant
        Dim col6 As Variant
        Dim Col7 As Variant
        Dim col8 As Variant
        Dim col9 As Variant
        
        Dim strMa As Variant
        
            
        Do
            '-- check phan noi dung
            .GetText .ColLetterToNumber("AD"), .Row, col2
            .GetText .ColLetterToNumber("AL"), .Row, col3
            .GetText .ColLetterToNumber("AU"), .Row, col4
            .GetText .ColLetterToNumber("BE"), .Row, col5
            .GetText .ColLetterToNumber("BK"), .Row, col6
            .GetText .ColLetterToNumber("BU"), .Row, Col7
            .GetText .ColLetterToNumber("BX"), .Row, col8
            .GetText .ColLetterToNumber("CG"), .Row, col9
          
            .GetText .ColLetterToNumber("C"), .Row, noidung
'            .GetText .ColLetterToNumber("CQ"), .Row, strMa

            ' reset error
            .Col = .ColLetterToNumber("BU")
            .CellNote = ""
            .BackColor = mNonErrorColor

            If Trim$(col2) <> vbNullString Or Trim$(col3) <> vbNullString Or Trim$(col4) <> "0" Or Trim$(col5) <> vbNullString Or Trim$(col6) <> "0" Or Trim$(col8) <> "0" Or Trim$(col9) <> "0" Then
                If noidung = vbNullString Then
                    .Col = .ColLetterToNumber("C")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMST   'static
                    .BackColor = mErrorColor

                    iFlagvMst = 1
                            
                    ' Forcus ve Cell co du lieu bi loi
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        .Sheet = .SheetCount
                        strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
                        .Col = .ColLetterToNumber("C")
                        .Row = 28
                        .Formula = strFormula
                    End If

                Else
                    .Sheet = 1
                    .Col = .ColLetterToNumber("C")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
                
                ' Neu chi tieu 9 < 6 thi canh bao do
                If Val(col9) <> 0 And Val(col9) > Val(col6) Then
                    .Col = .ColLetterToNumber("CG")
                    .CellNote = vErrThueTNDN_PN
                    .BackColor = mErrorColor
                    iFlagvThueTNDN_PN = 1
                Else
                ' Check neu cot 9 sua khac voi gia tri tinh dc thi se canh bao vang
                    If Val(col9) <> Round(Val(col6) * Val(Col7) / 100, 0) - Val(col8) Then
                        If Val(col9) = 0 And Round(Val(col6) * Val(Col7) / 100, 0) - Val(col8) < 0 Then
                            .Col = .ColLetterToNumber("CG")
                            .CellNote = ""
                            .BackColor = mNonErrorColor
                        Else
                            .Col = .ColLetterToNumber("CG")
                            .CellNote = vAlertThueTNDN_PN
                            .BackColor = mAlertColor
                        End If
                    Else
                        .Col = .ColLetterToNumber("CG")
                        .CellNote = ""
                        .BackColor = mNonErrorColor
                    End If
                End If
                
                ' Kiem tra chi tieu 6,7,8
                If Round(Val(col6) * Val(Col7) / 100, 0) - Val(col8) < 0 And Val(col8) <> 0 Then
                    .Col = .ColLetterToNumber("BX")
                    .CellNote = vErrCT678
                    .BackColor = mErrorColor
                    iFlagvCT678 = 1
                Else
                    .Col = .ColLetterToNumber("BX")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            End If
            
'            ' Kiem tra thue suat cua tien lai
'             If Trim$(strMa) = "08" Then
'                 If Val(Col7) <> 5 And Val(Col7) <> 10 Then
'                     .Col = .ColLetterToNumber("BU")
'                     .CellNote = vErrorThueSuat
'                     .BackColor = mErrorColor
'                     iFlagThueSuat = 1
'                 Else
'                     .Col = .ColLetterToNumber("BU")
'                     .CellNote = ""
'                     .BackColor = mNonErrorColor
'                 End If
'             End If
        
            .Sheet = 1
            .Row = 34 + i
            i = i + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"

        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 18, IIf(iFlagvMst = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 22, IIf(iFlagvThueTNDN_PN = 1, "0", "1")
        .SetText .ColLetterToNumber("B"), 23, IIf(iFlagvCT678 = 1, "0", "1")
        '.SetText .ColLetterToNumber("B"), 24, IIf(iFlagThueSuat = 1, "0", "1")
        ' ----
        
'        .Sheet = 1
'        .Col = .ColLetterToNumber("B")
'        .Row = 34
'        i = 1
'
'        Do
'            .GetText .ColLetterToNumber("CG"), .Row, sotien
'
'            If sotien < 0 Then
'                .Col = .ColLetterToNumber("BX")
'                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                .CellNote = vErrMsgSoTien   'static
'                .BackColor = mErrorColor
'
'                iFlagvMst = 1
'
'                ' Forcus ve Cell co du lieu bi loi
'                If iFocusFlag = 0 Then
'                    strFocusSheetName = "'" & .SheetName & "'"
'                    strFocusCol = str(.Col)
'                    strFocusRow = str(.Row)
'                    .Sheet = .SheetCount
'                    strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
'                    .Col = .ColLetterToNumber("C")
'                    .Row = 20
'                    .Formula = strFormula
'                End If
'
'            Else
'                .Sheet = 1
'                .Col = .ColLetterToNumber("BX")
'                .CellNote = ""
'                .BackColor = mNonErrorColor
'            End If
'
'            .Sheet = 1
'            .Row = 34 + i
'            i = i + 1
'            .Col = .ColLetterToNumber("B")
'        Loop Until .Text = "aa"

'        .Sheet = .SheetCount
'        .SetText .ColLetterToNumber("B"), 20, IIf(iFlagvMst = 1, "0", "1")
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

'<CSCM>
'--------------------------------------------------------------------------------
' Description:       check Ma so thue trong cac bang co dong dong
' Created by :       Khai Hoan
' Date-Time  :       9/16/2011-4:31:43 PM
' paramater  :       - iRow,iCol toa do cua Cell du lieu Ma so thue.
'                    - iSheet: sheet dag lam viec.
'                    - iRow2: dong Messages thong bao loi cua Sheet Header trong file Exel
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub CheckErrorMST2(iRow As Integer, _
                          iRow2 As Integer, _
                          iCol As String, _
                          isheet As Integer)
    Dim strCheck, strFocusSheetName, strFocusRow, strFocusCol, strFormula As String

    Dim strNgayHD, strNgayTT, i, vErrMsgMST, vMST As Variant

    Dim iFocusFlag, iFlagvMst As Integer
    'Dim mNonErrorColor = vbWhite
    
    With fps
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), iRow2, vErrMsgMST

        iFocusFlag = 0
        
        .Sheet = isheet
        .Col = .ColLetterToNumber("B")
        .Row = iRow
        i = 1

        Do
            '----- check ma so thue  -------
            .GetText .ColLetterToNumber(iCol), .Row, vMST

            If Trim(vMST) <> vbNullString Then
                vMST = Replace(vMST, " ", "")
                vMST = Replace(vMST, "-", "")
                strCheck = ""
            
                If Len(vMST) = 10 Or Len(vMST) = 13 Then
                    strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                End If
                   
                If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                    .Col = .ColLetterToNumber(iCol)
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMST   'static
                    .BackColor = mErrorColor

                    iFlagvMst = 1
                        
                    ' Forcus ve Cell co du lieu bi loi
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)
                        .Sheet = .SheetCount
                        strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
                        .Col = .ColLetterToNumber("C")
                        .Row = iRow2
                        .Formula = strFormula
                    End If

                Else
                    .Sheet = isheet
                    .Col = .ColLetterToNumber(iCol)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If

            Else
                .Sheet = isheet
                .Col = .ColLetterToNumber(iCol)
                .CellNote = ""
                .BackColor = mNonErrorColor
               
            End If

            .Sheet = isheet
            .Row = iRow + i
            i = i + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"

        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), iRow2, IIf(iFlagvMst = 1, "0", "1")
    End With

End Sub


'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       HTKK 3.1
' Description:       So sanh du lieu ngay kieu ngay voi ky tinh thue
' Created by :       nkhoan
' Date-Time  :       9/23/2011-11:35:48 AM
'
' Parameters :       isheet: sheet dang lam viec
'                    iRow, iCol: toa do cua o du lieu can Validate
'                    iRowErorrs: dong thong bao loi tren Sheet Header(cot "E" mac dinh).
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub compareTowDate(isheet As Integer, _
                          iRow As Integer, _
                          iCol As String, _
                          iRowErorrs As Integer)
                          
    Dim iFocusFlag, iFlagvNgayHD, i As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String
    Dim vNgayHD           As Variant
    Dim vErrMsgNgayHD     As Variant
    Dim strFormula As String
    
    iFocusFlag = 0
    iFlagvNgayHD = 0

    With fps
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), iRowErorrs, vErrMsgNgayHD
        
        .Sheet = isheet
        .Col = .ColLetterToNumber("B")
        .Row = iRow
        i = 1

        Do
            
            .GetText .ColLetterToNumber(iCol), .Row, vNgayHD
            
            If Trim(vNgayHD) = "" Then
            ElseIf Trim(vNgayHD) <> "" Then

                If CInt(Right(vNgayHD, 4)) > CInt(TAX_Utilities_New.Year) Then
                    .Col = .ColLetterToNumber(iCol)
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgNgayHD
                    .BackColor = mErrorColor

                    iFlagvNgayHD = 1
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)

                        iFocusFlag = 1
                        .Sheet = .SheetCount
                        strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
                        .Col = .ColLetterToNumber("C")
                        .Row = iRowErorrs
                        .Formula = strFormula
                    End If

                ElseIf CInt(Right(vNgayHD, 4)) = CInt(TAX_Utilities_New.Year) And CInt(Right("00" & Mid(vNgayHD, 4, 2), 2)) > CInt(TAX_Utilities_New.Month) Then
                    .Col = .ColLetterToNumber(iCol)
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgNgayHD
                    .BackColor = mErrorColor

                    iFlagvNgayHD = 1
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        strFocusCol = str(.Col)
                        strFocusRow = str(.Row)

                        iFocusFlag = 1
                        .Sheet = .SheetCount
                        strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
                        .Col = .ColLetterToNumber("C")
                        .Row = iRowErorrs
                        .Formula = strFormula
                    End If

                ElseIf (CheckValidNgayHD(CStr(vNgayHD)) = False) Then
                    '                    .Col = .ColLetterToNumber(iCol)
                    '                    .CellNote = GetAttribute(GetMessageCellById("0124"), "Msg") 'static
                    '                    .BackColor = mAlertColor
                    '
                    '                    iFlagvNgayHD = 0
                    '                    If iFocusFlag = 0 Then
                    '                        strFocusSheetName = "'" & .SheetName & "'"
                    '                        strFocusCol = str(.Col)
                    '                        strFocusRow = str(.Row)
                    '
                    '                        iFocusFlag = 1
                    '        .Sheet = .SheetCount
                    '                        strFormula = Trim(strFocusSheetName) & "!" & Trim(.ColNumberToLetter(CLng(strFocusCol))) & Trim(strFocusRow)
                    '                        .Col = .ColLetterToNumber("C")
                    '                        .Row = iRowErorrs
                    '                        .Formula = strFormula
                    '                    End If

                Else
                    .Sheet = isheet
                    .Col = .ColLetterToNumber(iCol)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If

            Else
                .Sheet = isheet
                .Col = .ColLetterToNumber(iCol)
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
        
            .Sheet = isheet
            .Row = iRow + i
            i = i + 1
            .Col = .ColLetterToNumber("B")
        Loop Until .Text = "aa"

        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), iRowErorrs, IIf(iFlagvNgayHD = 1, "0", "1")
    End With

End Sub
Private Function CheckValidNgayHD(vdate As String) As Boolean
  
   Dim dd As String, mm As String, yyyy As String, dDate As Date, dDateNow As Date
   Dim sDateNow As String
   On Error GoTo e
    dd = Right("00" & Left(vdate, InStr(vdate, "/") - 1), 2)
    mm = Right("00" & Mid(vdate, 4, 2), 2)
    yyyy = Right("0000" & vdate, 4)
    sDateNow = "01/" & TAX_Utilities_New.Month & "/" & TAX_Utilities_New.Year
    dDateNow = Format(CDate(sDateNow), "dd/mm/yyyy")

    dDate = Format(CDate(vdate), "dd/mm/yyyy")
        If CInt(yyyy) > CInt(TAX_Utilities_New.Year) Then
            CheckValidNgayHD = False
            Exit Function
        End If
        If DateAdd("m", 5, dDate) < dDateNow Then
        CheckValidNgayHD = False
            Exit Function
        End If
   CheckValidNgayHD = True
   Exit Function

e:
    CheckValidNgayHD = False





 End Function

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
    Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol         As Long, lRow As Long
    Dim IsUpdate     As Boolean
    Dim idtkhai      As Variant
     Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim iCol As Long, iRow As Long
     Dim a As Long
    Dim totalCell1 As Long, countRowReset1 As Long
   
   a = (TAX_Utilities_New.Data(0).getElementsByTagName("Cell").length - 16) / 9
    With fps
        totalCell1 = TAX_Utilities_New.Data(fps.SheetCount - 2).getElementsByTagName("Cell").length
        countRowReset1 = (totalCell1 - 8) / 6
        For Each xmlNodeReset In TAX_Utilities_New.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            fps.Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            fps.Col = lCol
            fps.Row = lRow
            If lRow >= 34 And lRow <= 34 + a And mCurrentSheet = 1 Then
             Select Case fps.CellType
                    Case CellTypeCheckBox
                        fps.Text = vbNullString
                        UpdateCell fps, lCol, .Row, vbNullString
                    Case CellTypeComboBox
                        fps.Text = vbNullString
                        UpdateCell fps, lCol, .Row, vbNullString
                    Case CellTypeNumber
                        fps.value = 0
                        UpdateCell fps, lCol, .Row, vbNullString
                    Case Else
                        fps.value = vbNullString
                        UpdateCell fps, lCol, .Row, vbNullString
                End Select
                'mAdjustData = IIf(IsUpdate = True, IsUpdate, mAdjustData)
                TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
                
            End If
           ' Reset KHBS
            If (lRow >= 20 + countRowReset1) And mCurrentSheet = (fps.SheetCount - 1) Then
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
                TAX_Utilities_New.AdjustData(mCurrentSheet - 1) = True
            End If
        Next
        Set xmlCellNode = TAX_Utilities_New.Data(1 - 1).nodeFromID(GetCellID(fps, .ColLetterToNumber("C"), 34))
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode
            Set xmlCellNode = xmlCellNode.parentNode.lastChild
            Set xmlCellNode = xmlCellNode.firstChild
            Set xmlCellNode = xmlCellNode.firstChild
            Set xmlCellNode = xmlCellNode.nextSibling
            Set xmlCellNode = xmlCellNode.nextSibling
            Set xmlCellNode = xmlCellNode.nextSibling
            Set xmlCellNode = xmlCellNode.nextSibling
            Set xmlCellNode = xmlCellNode.nextSibling
            Set xmlCellNode = xmlCellNode.nextSibling
            Set xmlCellNode = xmlCellNode.nextSibling
            Set xmlCellNode = xmlCellNode.nextSibling
            Set xmlCellNode = xmlCellNode.nextSibling
            Set xmlCellNode = xmlCellNode.nextSibling
            ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), iCol, iRow
       
            .SetText iCol, iRow, ""
            UpdateCell fps, iCol, iRow, ""
        
         
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
    Dim blCheckS2 As Boolean
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
        ' Update sheet KHBS
       .Sheet = .SheetCount - 1
        mCurrentSheet = .Sheet
        If (GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(.Sheet - 1), "Active") <> "0") Then
            For Each xmlCellNodeData In TAX_Utilities_New.Data(.Sheet - 1).getElementsByTagName("Cell")
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
    Dim value As String
    Dim i As Integer
    'luu lai gia tri so ban dau de xu ly
    Dim soTam2 As String
    Dim soTam As String
    Dim strSoAm As String
    If Val(so) < 0 Then
        so = Mid$(so, 2)
        strSoAm = GetAttribute(GetMessageCellById("0241"), "Msg")
    End If
    
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
    doiTien = strSoAm + " " + UCase(Left(doiTien, 1)) + Mid(doiTien, 2)
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
    strPath = GetAbsolutePath(GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "TemplateFolder")) & "MapMCT_TT28.xml"
    
    If xmlDocument.Load(strPath) Then
        Set xmlNodeListItem = xmlDocument.getElementsByTagName("Root").Item(0).childNodes
        For Each xmlNode In xmlNodeListItem
            If GetAttribute(xmlNode, "ID") = GetAttribute(TAX_Utilities_New.NodeValidity.childNodes(0), "DataFile") Then
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
    SaveErrorLog "cls_03NTNN", "getMapCT", Err.number, Err.Description
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
            Set xmlNode = .nodeFromID("CG_26")
            result = "CT9_" & Val(GetAttribute(xmlNode, "Value"))
            Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_03NTNN", "getValueTK", Err.number, Err.Description
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
                .GetText .ColLetterToNumber("CG"), 26, varTemp
                strCurrentValue = "CT9_" & Trim(varTemp)
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
            Set xmlNode = .nodeFromID("CG_26")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_03NTNN", "getValueCTDC", Err.number, Err.Description
End Function


' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("CG"), 26, varTemp
        strCurrentValue = Val(varTemp) - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function


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
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_New.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
                Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
                Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell") _
                    (TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length - 1)
                    strValue1 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell")(1)
                    strValue2 = GetAttribute(xmlCellNode, "Value")
                 Set xmlCellNode = TAX_Utilities_New.Data(1).getElementsByTagName("Cell")(7)
                    strValue3 = GetAttribute(xmlCellNode, "Value")
          
                    If (Not IsNullValue(strValue1)) Or TAX_Utilities_New.Data(1).getElementsByTagName("Cell").length > 20 Or (Not IsNullValue(strValue2)) Or (Not IsNullValue(strValue3)) Then
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


' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 24
    denRowDL = 32
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "20_30_2"
    ' end
End Sub


