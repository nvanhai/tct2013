VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_01AC"
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
  
'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'No parameter
Public Sub Prepare1()
    With fps

        Dim xmlDocument As New MSXML.DOMDocument
        Dim xmlNode As MSXML.IXMLDOMNode
        xmlDocument.Load (GetAbsolutePath("..\InterfaceIni\Catalogue_loai_HD.xml"))
        Dim xmlNodeListMap As MSXML.IXMLDOMNodeList
        Dim strQuanHeNNTID As String
        Dim strQuanHeNNT As String
        Dim strDataFileName As String
        Dim arrDanhsach() As String
        Dim strComboHien As String
        Dim strCombo As String
        Dim strCombo1 As String
        Dim xmlNodeListCell As MSXML.IXMLDOMNodeList
        Dim xmlDomData As New MSXML.DOMDocument, xmlDomCurrentData As New MSXML.DOMDocument
       .Sheet = 1
       strDataFileName = GetAbsolutePath("..\InterfaceIni\Catalogue_loai_HD.xml")
        ' Lay danh muc loai hoa don
        ' 15/11/2010
        If xmlDomData.Load(strDataFileName) Then
            Set xmlNodeListCell = xmlDomData.getElementsByTagName("Item")
            For Each xmlNode In xmlNodeListCell
                If GetAttribute(xmlNode, "Value") <> "" Then
                    arrDanhsach = Split(GetAttribute(xmlNode, "Value"), "###")
                    strCombo = strCombo + CPab(arrDanhsach(0), 10) + CPab(arrDanhsach(1), 10) + CPab(arrDanhsach(0), 200) + Chr$(9)
'                    strCombo = strCombo + CPab(arrDanhsach(1), 10) + Chr$(9)
                    strComboHien = strComboHien + arrDanhsach(2) + Chr$(9)
                End If
            Next
        End If
       
'        Set xmlNodeListMap = xmlDocument.getElementsByTagName("Item")
'        For Each xmlNode In xmlNodeListMap
'            strQuanHeNNT = strQuanHeNNT + GetAttribute(xmlNode, "Value") + Chr$(9)
'        Next
'        strQuanHeNNTID = "01GTKT" + Chr$(9) + "02GTTT" + Chr$(9) + "06HDXK" + Chr$(9) + "07KPTQ" + Chr$(9) + "03XKNB" + Chr$(9) + "04HGDL" + Chr$(9) + "01/" + Chr$(9) + "02/" + Chr$(9) + "01/" + Chr$(9) + "02/" + Chr$(9) + "01/" + Chr$(9) + "02/" + Chr$(9)
        .Row = 22
        .Col = .ColLetterToNumber("H")
        
        .TypeComboBoxList = strComboHien
        .Text = strQuanHeNNT
        .Col = .ColLetterToNumber("N")
        .TypeComboBoxList = strCombo
        SetDateFormat fps, 1, 22, .ColLetterToNumber("G"), DDMMYYYY
         SetKyHieuHDFormat fps, 1, 22, .ColLetterToNumber("J")
        .Row = 32
        ' Set format ngay ky
        .Col = .ColLetterToNumber("J")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim i As Integer
    With fps
        .Sheet = 1
        .EventEnabled(EventAllEvents) = False
        i = 1
        .Col = .ColLetterToNumber("B")
        .Row = 22
        Do
             .Col = .ColLetterToNumber("C")
             .Text = "0"
             
             .Col = .ColLetterToNumber("B")
             .Row = i + 22
             totalRow2ABH = i
             i = i + 1
        Loop Until .Text = "aa"
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Function GetValidatedCellIndex(ByVal lSheet As Long, lAnchorRow As Long, ByVal lAnchorCol As Long) As Integer
    
    
End Function

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    Dim strId As Variant
    Dim intIndexCombo As Integer
    fps.EventEnabled(EventAllEvents) = False
    With fps
        .Sheet = 1
        .Col = Col
        .Row = Row
         If Col = .ColLetterToNumber("H") And .Text <> "" Then
            intIndexCombo = .TypeComboBoxCurSel
            .Col = .ColLetterToNumber("N")
            .TypeComboBoxCurSel = intIndexCombo
            'Update combo F
            .Col = .ColLetterToNumber("N")
            UpdateCell fps, .Col, .Row, .Text
             
            
            .GetText .ColLetterToNumber("N"), .Row, strId
            
            
            
            .Col = .ColLetterToNumber("M")
            If Trim(Left(strId, 1)) = "0" Then
                .SetText .ColLetterToNumber("I"), .Row, Trim(Mid(strId, 11, 6))
                UpdateCell fps, .ColLetterToNumber("I"), .Row, .Text
            Else
                .SetText .ColLetterToNumber("I"), .Row, Trim(Mid(strId, 11, 3))
                UpdateCell fps, .ColLetterToNumber("I"), .Row, .Text
            End If
            .EventEnabled(EventChange) = True
         End If
    End With
   fps.EventEnabled(EventAllEvents) = True
End Sub

Public Sub kiemTraDuLieuImport()
   
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
            If .ActiveSheet = 1 Then
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
    Dim tuSo As Variant
    Dim denSo As Variant
    Dim strId As Variant
    Dim strLoai As Variant
    Dim strTemp As Variant
    
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
           
            If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("G") And Row >= 22 Then
                .GetText Col, Row, varTemp
                If varTemp <> "" And varTemp <> "../../...." Then
                    'If Format_ddmmyyyy(CStr(varTemp)) <> "" Then
                              .SetText Col, Row, Format_ddmmyyyy(CStr(varTemp))
                    'End If
                Else
                 .SetText Col, Row, ""
                End If
            End If

         ' Chuyen ky hieu HD thanh chu cai in hoa
            If Col = .ColLetterToNumber("J") Then
                 .GetText Col, Row, varTemp
                 .SetText Col, Row, UCase(CStr(varTemp))
                 .Col = Col
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
            End If
           
           ' Format dinh dang tu so den so co 7 ky tu
            If Col = .ColLetterToNumber("K") Or Col = .ColLetterToNumber("L") Then
                 .GetText Col, Row, varTemp
                 .SetText Col, Row, FormatSoHD(CStr(varTemp))
                 .Col = Col
                 .Row = Row
                UpdateCell fps, .Col, .Row, .Text
           ' Ho tro tinh so luong
                .GetText .ColLetterToNumber("K"), Row, tuSo
                .GetText .ColLetterToNumber("L"), Row, denSo
                
                .SetText .ColLetterToNumber("M"), Row, Val(denSo) - Val(tuSo)
                .Col = Col
                .Row = Row
                UpdateCell fps, .Col, .Row, .Text
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
    With fps
        'If flgLoadToKhai = True Then Exit Sub
        
        'CheckErrorMST
        ' Neu ton tai file data bang ke 05A thi kiem tra tren bang ke 5A
        If IIf(TAX_Utilities_New.NodeValidity.childNodes(.SheetCount - 3).Attributes.getNamedItem("Active").nodeValue <> "0", True, False) Then
            checkSheet1
        Else
            clearCheckSheet1
        End If
        kiemTraDuLieuImport
    End With
End Sub

Private Sub checkSheet1()

   Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    Dim strCheck As String, iCurrentSheet As Integer
    
    Dim vMST As Variant, vErrMsgMST As Variant, vErrMsgMSTReq As Variant
    Dim vErrMsgMSTTrungCQCT As Variant, vErrMsgMSTTrungNhau As Variant, vErrMsgSoHD As Variant
    Dim vHoVaTen As Variant, vErrMsgHoVaTen As Variant
    Dim vTongThuNhap As Variant
    Dim vThuNhapGiamThue As Variant, vErrMsgThuNhapGiamThue As Variant
    Dim vSoThueDuocGiam As Variant, vErrMsgSoThueDuocGiam As Variant
    Dim vSoThueDaKhauTru As Variant, vErrMsgSoThueDaKhauTru As Variant, vErrMsgTusoDenSo As Variant
    Dim errMauSoHD As Variant
    Dim errSoLienHD As Variant
    Dim vLastRow As Variant
    Dim vMSTCQCT As Variant
    Dim vMST1 As Variant
    Dim Tuso1 As Variant
    Dim Tuso2 As Variant
    Dim denso1 As Variant
    Dim denso2 As Variant
    Dim Tenloaihd1 As Variant
    Dim Tenloaihd2 As Variant
    Dim mauso1 As Variant
    Dim mauso2 As Variant
    Dim kyhieu1 As Variant
    Dim kyhieu2 As Variant
    Dim SohieuHD As Variant
     Dim iWhiteRow As Integer
    Dim strSoLuong As Variant
    Dim strId As Variant, strLoai As String, strTemp As String, result As String
    
    Dim iFlagvHoVaTen As Integer, iFlagvMst As Integer, iFocusFlag As Integer
    Dim strFocusSheetName As String, strFocusRow As String, strFocusCol As String
    Dim i, j As Integer
    Dim blCheck_S1, blCheck_S2, blCheck_S3, blCheck_S4, blCheck_S5, blCheck_S6, blCheck_S7, blCheck_S8 As Boolean
    With fps
        .EventEnabled(EventAllEvents) = False
        iCurrentSheet = .Sheet
        .Sheet = .SheetCount
        ' Lay thong bao ve sai ma so thue trong sheet header
        .GetText .ColLetterToNumber("E"), 12, vErrMsgMST
        .GetText .ColLetterToNumber("E"), 13, vErrMsgMSTReq
        .GetText .ColLetterToNumber("E"), 14, vErrMsgMSTTrungNhau
        .GetText .ColLetterToNumber("E"), 15, vErrMsgSoHD
        .GetText .ColLetterToNumber("E"), 16, vErrMsgTusoDenSo
        .GetText .ColLetterToNumber("E"), 17, errMauSoHD
        .GetText .ColLetterToNumber("E"), 18, errSoLienHD
        ' Lay thong bao ve ma so thue bi trung voi co quan chi tra trong sheet header

        
        .Sheet = 1
        
        '.GetText .ColLetterToNumber("E"), 9, vMSTCQCT
        blCheck_S1 = True
        blCheck_S2 = True
        blCheck_S3 = True
        blCheck_S4 = True
        blCheck_S5 = True
        blCheck_S6 = True
        blCheck_S7 = True
        blCheck_S8 = True
        
        .Row = 22
        Do
            .Col = .ColLetterToNumber("K")
            .CellNote = ""
            .BackColor = mNonErrorColor
            
            .Col = .ColLetterToNumber("L")
            .CellNote = ""
            .BackColor = mNonErrorColor
            
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
            ' Check cho den het cac dong co du lieu thi thoi
        Loop Until UCase(.Text) = "AA"
        
        .Row = 22
        Do
            j = .Row
            
            '.GetText .ColLetterToNumber("E"), .Row, vHoVaTen
            '.GetText .ColLetterToNumber("D"), .Row, vMST
            .GetText .ColLetterToNumber("H"), .Row, Tuso1
            .GetText .ColLetterToNumber("I"), .Row, denso1
            .GetText .ColLetterToNumber("D"), .Row, Tenloaihd1
            .GetText .ColLetterToNumber("E"), .Row, mauso1
            .GetText .ColLetterToNumber("F"), .Row, kyhieu1
            '.GetText .ColLetterToNumber("O"), .Row, SohieuHD
            
            iWhiteRow = 1
            .GetText .ColLetterToNumber("J"), .Row, strSoLuong
            .GetText .ColLetterToNumber("W"), .Row, strId
            'check mau so hoa don
            If Trim(Tenloaihd1) <> "" Or Trim(mauso1) <> "" Or Trim(kyhieu1) <> "" Or Trim(Tuso1) <> "" Or Trim(denso1) <> "" Or Trim(strSoLuong) <> "0" Then
                ' Khong phai dong trang
                iWhiteRow = 0
            End If
            
            'Ten HD
            If Trim(Tenloaihd1) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("D")
                .CellNote = vErrMsgMSTReq  'static
                .BackColor = mErrorColor
                'iFlagEmty = 1
            Else
                .Col = .ColLetterToNumber("D")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            ' Mau so HD
            If Trim(mauso1) = "" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("E")
                .CellNote = vErrMsgMSTReq  'static
                .BackColor = mErrorColor
               ' iFlagEmty = 1
            Else
                .Col = .ColLetterToNumber("E")
                .CellNote = ""
                .BackColor = mNonErrorColor
            End If
            strLoai = Trim(Left(strId, 1))
            If strLoai = "0" Then
                strTemp = Trim(Mid(strId, 11, 6))
            Else
                strTemp = Trim(Mid(strId, 11, 3))
            End If
            result = CheckMauSoHD(CStr(mauso1), strLoai, strTemp)
            If result = "1" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("E")
                .CellNote = errMauSoHD  'static
                .BackColor = mErrorColor
                'iFlagMauSoHD = 1
                blCheck_S7 = False
            ElseIf result = "2" And iWhiteRow = 0 Then
                .Col = .ColLetterToNumber("E")
                .CellNote = errSoLienHD  'static
                .BackColor = mErrorColor
                'iFlagSoLienHD = 1
                blCheck_S8 = False
            Else
                .Col = .ColLetterToNumber("X")
                .CellNote = ""
                .BackColor = mNonErrorColor
            
            End If
            '/-----------------------------------
            Tuso1 = Val(Tuso1)
            denso1 = Val(denso1)
            'kiem tra tu so <= den so
            If Tuso1 > denso1 Then
                    .Col = .ColLetterToNumber("H")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgTusoDenSo    'static
                    .BackColor = mErrorColor
                     .Col = .ColLetterToNumber("I")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgTusoDenSo    'static
                    .BackColor = mErrorColor
                    blCheck_S6 = False
            End If
            'Kiem tra cau truc cua ky hieu
            Dim varTemp As Variant
            .GetText .ColLetterToNumber("F"), .Row, varTemp
'            .SetText .ColLetterToNumber("O"), .Row, CheckSoHD(CStr(varTemp))
'            .Col = .ColLetterToNumber("O")
'            UpdateCell fps, .Col, .Row, .Text
            SohieuHD = CheckSoHD(CStr(varTemp))
            

                If Trim(SohieuHD) = "1" Then
                    .Col = .ColLetterToNumber("F")
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgSoHD    'static
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
                    .Col = .ColLetterToNumber("F")
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If
            
'            If Trim(vHoVaTen) <> vbNullString Or Trim(vMST) <> vbNullString Then
'                If Trim(vHoVaTen) = vbNullString Then
'                    .Col = .ColLetterToNumber("E")
'                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'                    .CellNote = vErrMsgMSTReq    'static
'                    .BackColor = mErrorColor
'                    iFlagvHoVaTen = 1
'                    blCheck_S1 = False
'                    'capture the first error cell
'                    If iFocusFlag = 0 Then
'                        strFocusSheetName = "'" & .SheetName & "'"
'                        strFocusCol = str(.Col)
'                        strFocusRow = str(.Row)
'                        iFocusFlag = 1
'                    End If
'                Else
'                    .Col = .ColLetterToNumber("E")
'                    .CellNote = ""
'                    .BackColor = mNonErrorColor
'                End If
'
                
                ' Kiem tra trung nhau tu so den so
                'bat dau
                    Do
                        i = .Row + 1
                        .GetText .ColLetterToNumber("H"), i, Tuso2
                        .GetText .ColLetterToNumber("I"), i, denso2
                        Tuso2 = Val(Tuso2)
                        denso2 = Val(denso2)
                        .GetText .ColLetterToNumber("D"), i, Tenloaihd2
                        .GetText .ColLetterToNumber("E"), i, mauso2
                        .GetText .ColLetterToNumber("F"), i, kyhieu2
                        Dim vMST2 As Variant
                        '.GetText .ColLetterToNumber("D"), .Row, vMST2
                        
                            If (((Tuso1 - Tuso2) * (denso1 - Tuso2)) <= 0 Or ((Tuso1 - denso2) * (denso1 - denso2)) <= 0 Or ((Tuso1 > Tuso2) And (denso1 < denso2))) And ((Tenloaihd1 = Tenloaihd2) And (mauso1 = mauso2) And (kyhieu1 = kyhieu2)) Then
                                .Row = j
                                .Col = .ColLetterToNumber("H")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrMsgMSTTrungNhau     'static
                                .BackColor = mErrorColor
                                
                                
                                .Row = j
                                .Col = .ColLetterToNumber("I")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrMsgMSTTrungNhau     'static
                                .BackColor = mErrorColor
                                
                                
                                .Row = i
                                .Col = .ColLetterToNumber("H")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrMsgMSTTrungNhau     'static
                                .BackColor = mErrorColor
                               
                                
                                .Row = i
                                .Col = .ColLetterToNumber("I")
                                .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                .CellNote = vErrMsgMSTTrungNhau     'static
                                .BackColor = mErrorColor
                                
    
                                iFlagvMst = 1
                                blCheck_S4 = False
                                'capture the first error cell
                                If iFocusFlag = 0 Then
                                    strFocusSheetName = "'" & .SheetName & "'"
                                    strFocusCol = str(.Col)
                                    strFocusRow = str(.Row)
                                    iFocusFlag = 1
                                End If
                                Else
    '                            .Row = i
    '                            .Col = .ColLetterToNumber("O")
    '                            .CellNote = ""
    '                            .BackColor = mNonErrorColor
    '                            .Text = .Text - 1
                             End If
                        .Row = i
                        .Col = .ColLetterToNumber("B")
                        ' Check cho den het cac dong co du lieu thi thoi
                    Loop Until UCase(.Text) = "AA"
                    
                    '--------------------------------------------------
            .Row = j
            .Row = .Row + 1
            .Col = .ColLetterToNumber("B")
            ' Check cho den het cac dong co du lieu thi thoi
        Loop Until UCase(.Text) = "AA"
      
'        If blCheck_S1 = True Then
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 12, "1"
'        Else
'            .Sheet = .SheetCount
'            .SetText .ColLetterToNumber("B"), 12, "0"
'        End If
        ' Trong truong hop blCheck_S2 khong co MST nao bi loi thi set lai gia tri cua Header la 1
        If blCheck_S2 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 12, "1"
            '.SetText .ColLetterToNumber("B"), 13, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 12, "0"
            '.SetText .ColLetterToNumber("B"), 13, "0"
            
        End If
                If blCheck_S3 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 13, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 13, "0"
        End If
        
        
        If blCheck_S4 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 14, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 14, "0"
        End If
        
        If blCheck_S5 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 15, "0"
        End If
        
        If blCheck_S6 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 16, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 16, "0"
        End If
        
        If blCheck_S7 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 17, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 17, "0"
        End If
        
        If blCheck_S8 = True Then
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 18, "1"
        Else
            .Sheet = .SheetCount
            .SetText .ColLetterToNumber("B"), 18, "0"
        End If
        .Sheet = iCurrentSheet
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub clearCheckSheet1()
    
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
            Case "01AC"
                blnNullValue = True
                'Check the third sheet
               For Each xmlCellNode In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
                    ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                    'Check value in D column and E column
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
            For Each xmlCellNodeData In TAX_Utilities_New.Data(0).getElementsByTagName("Cell")
               ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               .GetText lCol, lRow, varTemp
               UpdateCell fps, lCol, lRow, varTemp
             Next

      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing
End Sub

Public Sub FinishImport()

End Sub

Private Sub fps_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
    Dim sumLaoDong As Variant, sumThuNhap As Variant, sumLDKhauTru As Variant, sumTNKhauTru As Variant
    Dim sumSoThueTNDaKhauTru As Variant
    Dim varHoTen As Variant, varMST As Variant
    Dim varThuNhap As Variant, varSoThueTNDaKhauTru As Variant
    
    'If flgLoadToKhai = True Then Exit Sub
    
    If NewSheet = 1 Then
        kiemTraDuLieuImport
        With fps
            'Because of unclearly-rised events, this function is here for turning it off
            'at the end of this sub, this event will be turned on
            .EventEnabled(EventAllEvents) = False
            
            .Sheet = 1 'Bang ke 2A
            
            .Row = 17
            
            sumLaoDong = 0
            sumThuNhap = 0
            sumLDKhauTru = 0
            sumTNKhauTru = 0
            sumSoThueTNDaKhauTru = 0
                        
            ' Lay cac thong tin sau khi da tong hop tren bang ke
            .GetText .ColLetterToNumber("J"), .Row, sumLaoDong
            .GetText .ColLetterToNumber("F"), .Row, sumThuNhap
            .GetText .ColLetterToNumber("K"), .Row, sumLDKhauTru
            .GetText .ColLetterToNumber("L"), .Row, sumTNKhauTru
            .GetText .ColLetterToNumber("I"), .Row, sumSoThueTNDaKhauTru
                        
            .EventEnabled(EventAllEvents) = True
        End With
    End If
End Sub
Private Sub fps_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    Dim intIndexCombo As Integer
    Dim temp As String
'    With fps
'        If .ActiveSheet = 1 Then
'            If Col = .ColLetterToNumber("H") And .Text <> "" Then
'                intIndexCombo = .TypeComboBoxCurSel
'                 .Row = Row
'                 .Col = .ColLetterToNumber("N")
'                 .TypeComboBoxCurSel = intIndexCombo
'                UpdateCell fps, .Col, .Row, .Value
'                temp = .Text
'                .Col = .ColLetterToNumber("I")
'               .Text = temp
'            End If
'        End If
'    End With
End Sub


