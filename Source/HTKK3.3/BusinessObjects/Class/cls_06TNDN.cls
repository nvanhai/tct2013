VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_06TNDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Thu Nhap Doanh Nghiep Tam tinh" interface sheets
'this Class is belong to TAX_Business_v2 project which will be compline to DLL

Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Dim mblKeyF6 As Boolean
Dim mblKhoiDauToKhai As Boolean
Dim vThuephainop  As Variant
Dim strGiaTriCu As Variant, blGiaTriCu As Boolean, RowGiaTriCu As Long, ColGiaTriCu As Long
'Dim xNodeData0 As MSXML.IXMLDOMNode
Dim xNode As MSXML.IXMLDOMNode
Dim arrData() As String
Public loaiKHBS As String
Public Strloaitk As String
Public StrSolanBosung As String
Public FlagThaydoi As String
Public strLoaiNNKD As String
Public StrTKThang_PS As String
Public iflag As Boolean

Private Const Row_MSTDN = 47
Private Const Col_MSTDN = "J"

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
        .Sheet = 1
        
        'format ngay thang
        SetDateFormat fps, 1, 92, .ColLetterToNumber("AP"), DDMMYYYY
        SetDateFormat fps, 1, 42, .ColLetterToNumber("AC"), DDMMYYYY
        SetDateFormat fps, 1, 44, .ColLetterToNumber("AH"), DDMMYYYY
        
        
        .Row = 92
        .Col = .ColLetterToNumber("AP")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        

        
        SetDateFormat fps, .SheetCount - 1, 31, .ColLetterToNumber("BE"), DDMMYYYY
        SetDateFormat fps, .SheetCount - 1, 46, .ColLetterToNumber("BG"), DDMMYYYY
         .Sheet = .SheetCount - 1 'To khai TNDN
        .Row = 46
        .Col = .ColLetterToNumber("BG")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft

    End With
            ' set CQT ra quyet dinh hoan
    setCQTQuanLyHoanThue fps
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    Dim blCheck_S2 As Boolean
    Dim chitieu25  As Variant
    Dim chitieu26  As Variant
    Dim ngay As Variant

    Dim ct25, ct27, ct28, ct29, ct30, ct31, ct32, ct34 As Variant
    Dim ct26 As Double
    
    With fps
        .Sheet = 1
        'currentSheet2 = mCurrentSheet
        mCurrentSheet = .Sheet
        
        'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
        
        ngay = TAX_Utilities_v2.Day
         'xu ly ngay phat sinh
        If Trim(ngay) = "" Then
            .SetText .ColLetterToNumber("AI"), 54, ""
            UpdateCell fps, .ColLetterToNumber("AI"), 54, ""
        Else
            .SetText .ColLetterToNumber("AI"), 54, ngay & "/" & TAX_Utilities_v2.Month & "/" & TAX_Utilities_v2.Year
            UpdateCell fps, .ColLetterToNumber("AI"), 54, ngay & "/" & TAX_Utilities_v2.Month & "/" & TAX_Utilities_v2.Year
        End If

        
        ' set lan dau hoac lan bo sung
            If Strloaitk = "TKCT" Then
                .Col = .ColLetterToNumber("M")
                .Row = 54
                .Text = 1
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("N")
                .Row = 54
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("O")
                .Row = 54
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
            ElseIf Strloaitk = "TKBS" Then
                .Col = .ColLetterToNumber("M")
                .Row = 54
                .Text = ""
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("N")
                .Row = 54
                .Text = "1"
                UpdateCell fps, .Col, .Row, .Text
                .Col = .ColLetterToNumber("O")
                .Row = 54
                .Text = StrSolanBosung
                UpdateCell fps, .Col, .Row, .Text
            End If
            
            
             .GetText .ColLetterToNumber("AW"), 61, ct25
            .GetText .ColLetterToNumber("AW"), 64, ct27
            .GetText .ColLetterToNumber("AW"), 65, ct28
            .GetText .ColLetterToNumber("AW"), 66, ct29
            .GetText .ColLetterToNumber("AW"), 67, ct30
            .GetText .ColLetterToNumber("AW"), 68, ct31
            .GetText .ColLetterToNumber("AW"), 69, ct32
            .GetText .ColLetterToNumber("AW"), 71, ct34
            ct26 = Val(ct27) + Val(ct28) + Val(ct29) + Val(ct30) + Val(ct31) + Val(ct32)
            
            If Val(ct25) - ct26 <= 0 Then
             ' khong cho nhap chi tieu 34
                .Col = .ColLetterToNumber("AW")
                .Row = 71
                .Lock = True
             Else
                .Col = .ColLetterToNumber("AW")
                .Row = 71
                .Lock = False
             End If
            
            If Val(ct25) - ct26 - Val(ct34) <= 0 Then
             ' khong cho nhap chi tieu 36
                .Col = .ColLetterToNumber("AW")
                .Row = 73
                .Lock = True
             Else
                .Col = .ColLetterToNumber("AW")
                .Row = 73
                .Lock = False
             End If
    End With
    
End Sub

Public Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    Dim ct25, ct27, ct28, ct29, ct30, ct31, ct32, ct34 As Variant
    Dim ct26 As Double
    
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = 1
        .GetText .ColLetterToNumber("AW"), 61, ct25
        .GetText .ColLetterToNumber("AW"), 64, ct27
        .GetText .ColLetterToNumber("AW"), 65, ct28
        .GetText .ColLetterToNumber("AW"), 66, ct29
        .GetText .ColLetterToNumber("AW"), 67, ct30
        .GetText .ColLetterToNumber("AW"), 68, ct31
        .GetText .ColLetterToNumber("AW"), 69, ct32
        .GetText .ColLetterToNumber("AW"), 71, ct34
        ct26 = Val(ct27) + Val(ct28) + Val(ct29) + Val(ct30) + Val(ct31) + Val(ct32)
        
        If Col = .ColLetterToNumber("AW") And (Row = 61 Or Row = 71 Or (Row >= 64 And Row <= 69)) Then
             If Val(ct25) - ct26 <= 0 Then
             ' khong cho nhap chi tieu 34
                .Col = .ColLetterToNumber("AW")
                .Row = 71
                .Text = "0"
                .Lock = True
                UpdateCell fps, .Col, .Row, .Text
             Else
                .Col = .ColLetterToNumber("AW")
                .Row = 71
                .Lock = False
             End If
        
             If Val(ct25) - ct26 - Val(ct34) <= 0 Then
             ' khong cho nhap chi tieu 36
                .Col = .ColLetterToNumber("AW")
                .Row = 73
                .Text = "0"
                .Lock = True
                UpdateCell fps, .Col, .Row, .Text
             Else
                .Col = .ColLetterToNumber("AW")
                .Row = 73
                .Lock = False
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
    Dim varTemp   As Variant
    Dim chitieu25 As Variant
    Dim chitieu26 As Variant
    
    With fps
        mCurrentSheet = .Sheet
        .EventEnabled(EventAllEvents) = False

        ValidateDate "AP", 92, Col, Row, mCurrentSheet, fps
        ValidateDate "AC", 42, Col, Row, mCurrentSheet, fps
        ValidateDate "AH", 44, Col, Row, mCurrentSheet, fps
      
        .EventEnabled(EventAllEvents) = True
    End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       HTKK 3.1
' Description:       valida Cell du dieu kieu ngay thang
' Created by :       nkhoan
' Date-Time  :       9/29/2011-5:31:03 PM
'
' Parameters :       Col,Row: toa do cua Cell duoc tro vao
'                    iCol,iRow: toa do cua Cell chua du lieu ngay thang can Validate
'                    sheet: sheet dang lam viec
'                    fps (fpSpread)
'--------------------------------------------------------------------------------

Private Sub ValidateDate(iCol As String, _
                         ByVal iRow As Long, _
                         ByVal Col As Long, _
                         ByVal Row As Long, _
                         Sheet As Integer, _
                         fps As fpSpread)
    Dim varTemp As Variant

    With fps
        .Sheet = Sheet
    
        If Col = .ColLetterToNumber(iCol) And Row = iRow Then
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
        
    End With

End Sub
'
'End Sub


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
    Dim ct12, ct13, ct14, ct15_1, ct15_2, ct15_3 As Variant
    Dim flagEmty As Integer
    Dim vErrEmpty As Variant
    
    With fps
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 31, vErrEmpty
        
        .Sheet = 1
        .GetText .ColLetterToNumber("V"), 34, ct12
        .GetText .ColLetterToNumber("AL"), 36, ct13
        .GetText .ColLetterToNumber("I"), 40, ct14
        .GetText .ColLetterToNumber("X"), 42, ct15_1
        .GetText .ColLetterToNumber("AC"), 42, ct15_2
        .GetText .ColLetterToNumber("AH"), 44, ct15_3
        If Trim$(ct12) = "" Then
            .Row = 34
            .Col = .ColLetterToNumber("V")
            .CellNote = vErrEmpty
            .BackColor = mErrorColor
            flagEmty = 1
        Else
            .Row = 34
            .Col = .ColLetterToNumber("V")
            .CellNote = ""
            .BackColor = mNonErrorColor
        End If
        
        If Trim$(ct13) = "" Then
            .Row = 36
            .Col = .ColLetterToNumber("AL")
            .CellNote = vErrEmpty
            .BackColor = mErrorColor
            flagEmty = 1
        Else
            .Row = 36
            .Col = .ColLetterToNumber("AL")
            .CellNote = ""
            .BackColor = mNonErrorColor
        End If
    
        If Trim$(ct14) = "" Then
            .Row = 40
            .Col = .ColLetterToNumber("I")
            .CellNote = vErrEmpty
            .BackColor = mErrorColor
            flagEmty = 1
        Else
            .Row = 40
            .Col = .ColLetterToNumber("I")
            .CellNote = ""
            .BackColor = mNonErrorColor
        End If
        
        
        If Trim$(ct15_1) = "" Then
            .Row = 42
            .Col = .ColLetterToNumber("X")
            .CellNote = vErrEmpty
            .BackColor = mErrorColor
            flagEmty = 1
        Else
            .Row = 42
            .Col = .ColLetterToNumber("X")
            .CellNote = ""
            .BackColor = mNonErrorColor
        End If
        
         If Trim$(ct15_2) = "" Then
            .Row = 42
            .Col = .ColLetterToNumber("AC")
            .CellNote = vErrEmpty
            .BackColor = mErrorColor
            flagEmty = 1
        Else
            .Row = 42
            .Col = .ColLetterToNumber("AC")
            .CellNote = ""
            .BackColor = mNonErrorColor
        End If
        
         If Trim$(ct15_3) = "" Then
            .Row = 44
            .Col = .ColLetterToNumber("AH")
            .CellNote = vErrEmpty
            .BackColor = mErrorColor
            flagEmty = 1
        Else
            .Row = 44
            .Col = .ColLetterToNumber("AH")
            .CellNote = ""
            .BackColor = mNonErrorColor
        End If
        
        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), 31, IIf(flagEmty = 1, "0", "1")
    End With
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       HTKK 3.1
' Description:       C = A - B: bat loi khong cho C<0
' Created by :       nkhoan
' Date-Time  :       10/24/2011-5:11:50 PM
'
' Parameters :          iRow, iCol: toa do cua chi tieu C
'
'                       iRowErros: dong Massage thong beo loi tren sheet Header
'                       isheet: sheet muon lam viec.
'--------------------------------------------------------------------------------

'</CSCM>
Public Sub CheckErrorChiTieu(iRow As Integer, iCol As String, iRowErros As Integer, isheet As Integer)

    
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
' paramater  :       - iRow: dong thu 1 trong bang du lieu dong trong file Exel.
'                    - iCol: cot chua Cell du lieu Ma so thue.
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
' Description:       check Ma so thue trong 1 Cell thuoc dung tinh
' Created by :       Khai Hoan
' Date-Time  :       9/16/2011-4:31:43 PM
' paramater  :       - iRow,iCol toa do cua Cell du lieu Ma so thue.
'                    - iSheet: sheet dag lam viec.
'                    - iRow2: dong Messages thong bao loi cua Sheet Header trong file Exel
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub CheckErrorMST3(iRow As Integer, _
                          iRow2 As Integer, _
                          iCol As String, _
                          isheet As Integer)
    Dim strCheck, strFocusSheetName, strFocusRow, strFocusCol, strFormula As String

    Dim strNgayHD, strNgayTT, i, vErrMsgMST, vMST As Variant

    Dim iFocusFlag, iFlagvMst As Integer

    
    With fps
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), iRow2, vErrMsgMST

        iFocusFlag = 0
        
        .Sheet = isheet
       
            '----- check ma so thue  -------
            .GetText .ColLetterToNumber(iCol), iRow, vMST

            If Trim(vMST) <> vbNullString Then
                vMST = Replace(vMST, " ", "")
                vMST = Replace(vMST, "-", "")
                strCheck = ""
            
                If Len(vMST) = 10 Or Len(vMST) = 13 Then
                    strCheck = CheckTaxCode(Mid(vMST, 1, 1), Mid(vMST, 2, 1), Mid(vMST, 3, 1), Mid(vMST, 4, 1), Mid(vMST, 5, 1), Mid(vMST, 6, 1), Mid(vMST, 7, 1), Mid(vMST, 8, 1), Mid(vMST, 9, 1), Mid(vMST, 10, 1), Mid(vMST, 11, 1), Mid(vMST, 12, 1), Mid(vMST, 13, 1))
                End If
                   
                If Mid(strCheck, 2, 1) <> "0" Or Not IsNumeric(vMST) Then
                    .Col = .ColLetterToNumber(iCol)
                    .Row = iRow
                    .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                    .CellNote = vErrMsgMST   'static
                    .BackColor = mErrorColor

                    iFlagvMst = 1
                        
                    ' Forcus ve Cell co du lieu bi loi
                    If iFocusFlag = 0 Then
                        strFocusSheetName = "'" & .SheetName & "'"
                        
                        .Sheet = .SheetCount
                        strFormula = Trim(strFocusSheetName) & "!" & iCol & iRow
                        .Col = .ColLetterToNumber("C")
                        .Row = iRow2
                        .Formula = strFormula
                    End If

                Else
                    .Sheet = isheet
                    .Row = iRow
                    .Col = .ColLetterToNumber(iCol)
                    .CellNote = ""
                    .BackColor = mNonErrorColor
                End If

            Else
                .Sheet = isheet
                .Row = iRow
                .Col = .ColLetterToNumber(iCol)
                .CellNote = ""
                .BackColor = mNonErrorColor
               
            End If

        .Sheet = .SheetCount
        .SetText .ColLetterToNumber("B"), iRow2, IIf(iFlagvMst = 1, "0", "1")
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
    'Check the second sheet
    'Select Case GetAttribute(xmlNodeValid, "ID")
    For Each xmlNodeValid In TAX_Utilities_v2.NodeValidity.childNodes
        Select Case GetAttribute(xmlNodeValid, "ID")
             Case "KHBS"
                blnNullValue = True
                'Check value of last node in Data
'                Set xmlCellNode = TAX_Utilities_v2.Data(2).getElementsByTagName("Cell") _
'                    (TAX_Utilities_v2.Data(2).getElementsByTagName("Cell").length - 1)
'                    strValue1 = GetAttribute(xmlCellNode, "Value")
'
'                    If (Not IsNullValue(strValue1)) Then
'                            blnNullValue = False
'                    End If

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
 Dim xmlNodeReset As MSXML.IXMLDOMNode
    Dim lCol         As Long, lRow As Long
    Dim IsUpdate     As Boolean
    Dim idtkhai      As Variant
    Dim totalCell1 As Long, countRowReset1 As Long
    
    With fps
        totalCell1 = TAX_Utilities_v2.Data(fps.SheetCount - 2).getElementsByTagName("Cell").length
        countRowReset1 = (totalCell1 - 8) / 6


        For Each xmlNodeReset In TAX_Utilities_v2.Data(mCurrentSheet - 1).getElementsByTagName("Cell")
            fps.Sheet = mCurrentSheet
            ParserCellID fps, GetAttribute(xmlNodeReset, "CellID"), lCol, lRow
            fps.Col = lCol
            fps.Row = lRow

            If ((33 < lRow <= 44) And (58 < lRow < 82) And mCurrentSheet = 1) Or ((33 < lRow) And (lRow < 42) And (mCurrentSheet = 2)) Then

                Select Case fps.CellType

                    Case CellTypeCheckBox
                        fps.Text = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString

                    Case CellTypeComboBox
                        fps.Text = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString

                    Case CellTypeNumber
                        fps.value = 0
                        UpdateCell fps, lCol, lRow, vbNullString

                    Case Else
                        fps.value = vbNullString
                        UpdateCell fps, lCol, lRow, vbNullString
                End Select

                'mAdjustData = IIf(IsUpdate = True, IsUpdate, mAdjustData)
                TAX_Utilities_v2.AdjustData(mCurrentSheet - 1) = IIf(IsUpdate = True, IsUpdate, TAX_Utilities_v2.AdjustData(mCurrentSheet - 1))
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
                TAX_Utilities_v2.AdjustData(mCurrentSheet - 1) = True
            End If

        Next

    End With
End Function

Public Sub SetData()
    '
End Sub

Public Sub Prepare_KHBS()
   
End Sub

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function

Public Sub UpdateChangeKHBS()
   
End Sub


Private Sub UpdateSheets()
    Dim varTemp As Variant
    Dim ssSheet As Integer
    Dim lCol As Long, lRow As Long
    Dim xmlCellNode As MSXML.IXMLDOMNodeList
    Dim xmlCellNodeData As MSXML.IXMLDOMNode
    
    ' insert node
    Dim xmlNodeCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCells As MSXML.IXMLDOMNode
    Dim xmlNodeNewCell As MSXML.IXMLDOMNode
    Dim testDoc As New MSXML.DOMDocument
    ' end

    Dim blCheck_S2 As Boolean
    With fps
        .EventEnabled(EventAllEvents) = False
        ssSheet = mCurrentSheet
        .Sheet = 1
        mCurrentSheet = .Sheet

        For Each xmlCellNodeData In TAX_Utilities_v2.Data(0).getElementsByTagName("Cell")
           ParserCellID fps, GetAttribute(xmlCellNodeData, "CellID"), lCol, lRow
               If (lRow = 82 And lCol = .ColLetterToNumber("C")) Or (lRow = 16 And lCol = .ColLetterToNumber("Z")) Then
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

        ' Update lai sheet KHBS
        .Sheet = .SheetCount - 1
        mCurrentSheet = .Sheet
        blCheck_S2 = IIf(TAX_Utilities_v2.NodeValidity.childNodes(.Sheet - 1).Attributes.getNamedItem("Active").nodeValue <> "0", True, False)
        If blCheck_S2 = True Then
            For Each xmlCellNodeData In TAX_Utilities_v2.Data(.Sheet - 1).getElementsByTagName("Cell")
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
            Set xmlNode = .nodeFromID("AW_74")
            result = "CT37_" & Val(GetAttribute(xmlNode, "Value"))
           Set xmlNode = Nothing
        End If
    End With
    getValueTK = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_06TNDN", "getValueTK", Err.number, Err.Description
End Function

Public Function getDieuChinhGiam(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue, strDieuChinhGiam, strDieuChinhTang As String
    Dim tempArrOld()       As String
    Dim tempArrCurr()      As String
    Dim tempArrOldValue()  As String
    Dim tempArrCurrValue() As String
    Dim i, j As Integer
    Dim strMaCT         As String
    Dim arrMapCT()      As String ' Luu gia tri cua dong chi tieu
    Dim arrMapCTValue() As String ' Luu gia tri cua cac cell chi tieu
    strMaCT = getMapCT

    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("AW"), 74, varTemp
        strCurrentValue = "CT37_" & Trim(varTemp)
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

            If Left(strMaCT, 3) = "38_" Then
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
    SaveErrorLog "06TNDN", "getMapCT", Err.number, Err.Description
End Function



' Lay ve so tong dieu chinh
Public Function getChiTieuTongDC(strOldValue As String) As String
    Dim varTemp As Variant
    Dim strCurrentValue As String
    With fps
        .Sheet = 1
        .GetText .ColLetterToNumber("AW"), 74, varTemp
        strCurrentValue = Val(varTemp)
        strCurrentValue = strCurrentValue - Val(strOldValue)
        getChiTieuTongDC = strCurrentValue
    End With
End Function

' Lay ve chi tieu 32 o to khai lan truoc
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
            Set xmlNode = .nodeFromID("AW_74")
            result = Val(GetAttribute(xmlNode, "Value"))
        End If
    End With
    getValueCTDC = result
    Exit Function

ErrorHandle:
    SaveErrorLog "cls_06TNDN", "getValueCTDC", Err.number, Err.Description
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

' ham kiem tra thong tin dai ly
Public Sub inThongTinDL()
    ' Set tham so in thong tin dai ly thue
    FlagDLThue = isCheckTTDLT
    tuRowDL = 45
    denRowDL = 53
    'TuRow_DenRow_sheet
    strTuRowDenRowPL = "20_30_2"
    ' end
End Sub

Public Function isCheckGh() As Boolean
    If TAX_Utilities_v2.Data(0).getElementsByTagName("Cell").length = 29 Then
        isCheckGh = True
    Else
        isCheckGh = False
    End If
    Exit Function
End Function

Public Sub FinishImport()
Dim i As Integer, iCol As Long, iRow As Long
Dim maTM As Variant, strId As Variant
Dim idx As Integer
Dim varTemp As Variant
    With fps
        fps.EventEnabled(EventAllEvents) = False
'           If .ActiveSheet = 2 Then
'                   i = 1
'                   .Col = .ColLetterToNumber("B")
'                   .Row = 7
'                   Do
'                        .Text = str(i)
'
'
'                        .Col = .ColLetterToNumber("B")
'                        .Row = i + 7
'                        i = i + 1
'                   Loop Until .Text = "aa"
'
'                   .Row = iRow
'                   .Col = iCol
'                   .SetActiveCell iCol, iRow
'           End If
       fps.EventEnabled(EventAllEvents) = True
   End With
End Sub
