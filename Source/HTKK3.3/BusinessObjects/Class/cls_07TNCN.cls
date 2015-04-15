VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_07TNCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1
Public Strloaitk As String
Public StrSolanBosung As String
Public iflag As Boolean
Private Const Row_MSTDN = 8
Private Const Col_MSTDN = "I"

Public strTkThangQuy As String

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
    Dim tmp As Integer
    With fps
        SetDateFormat fps, 1, 68, .ColLetterToNumber("R"), DDMMYYYY
        .Sheet = 1 'To khai 07/TNCN
        .Row = 68
        .Col = .ColLetterToNumber("R")
        .Text = Format(Date, "dd/mm/yyyy")
        .TypeHAlign = TypeHAlignLeft
        
        ' set defalt gia tri giam tru ban than
        If strTkThangQuy = "TK_THANG" Then
            ' chi tieu 25
            If (Val(TAX_Utilities_v2.Month) >= 7 And Val(TAX_Utilities_v2.Year) = 2013) Or Val(TAX_Utilities_v2.Year) > 2013 Then
                .Col = .ColLetterToNumber("R")
                .Row = 46
                .Text = "9000000"
                
                ' chi tieu 26 label
'                .Col = .ColLetterToNumber("D")
'                .Row = 47
'                .Text = GetAttribute(GetMessageCellById("0269"), "Msg")
            Else
                .Col = .ColLetterToNumber("R")
                .Row = 46
                .Text = "4000000"
                ' chi tieu 26 label
'                .Col = .ColLetterToNumber("D")
'                .Row = 47
'                .Text = GetAttribute(GetMessageCellById("0268"), "Msg")

            End If
        ElseIf strTkThangQuy = "TK_QUY" Then
            ' chi tieu 25
            If (Val(TAX_Utilities_v2.ThreeMonths) >= 3 And Val(TAX_Utilities_v2.Year) = 2013) Or Val(TAX_Utilities_v2.Year) > 2013 Then
                .Col = .ColLetterToNumber("R")
                .Row = 46
                .Text = "27000000"
                
                 ' chi tieu 26 label
'                .Col = .ColLetterToNumber("D")
'                .Row = 47
'                .Text = GetAttribute(GetMessageCellById("0269"), "Msg")

            Else
                .Col = .ColLetterToNumber("R")
                .Row = 46
                .Text = "12000000"
                
                ' chi tieu 26 label
'                .Col = .ColLetterToNumber("D")
'                .Row = 47
'                .Text = GetAttribute(GetMessageCellById("0268"), "Msg")

            End If
        End If
        'end
    End With
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'No parameter
Public Sub Prepare2()
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        
         'update MSTDN luôn bang MST tren phan TT chung
            UpdateMST fps, Col_MSTDN, Row_MSTDN
        
        If Strloaitk = "TKCT" Then
            .Col = .ColLetterToNumber("C")
            .Row = 70
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 70
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 70
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
        ElseIf Strloaitk = "TKBS" Then
            .Col = .ColLetterToNumber("C")
            .Row = 70
            .Text = ""
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("F")
            .Row = 70
            .Text = 1
            UpdateCell fps, .Col, .Row, .Text
            .Col = .ColLetterToNumber("I")
            .Row = 70
            .Text = StrSolanBosung
            UpdateCell fps, .Col, .Row, .Text
        End If
        
        
        ' set loai tk thang/ quy
         If strTkThangQuy = "TK_THANG" Then
            .Col = .ColLetterToNumber("L")
            .Row = 70
            .Text = "0"
            UpdateCell fps, .Col, .Row, .Text
        ElseIf strTkThangQuy = "TK_QUY" Then
            .Col = .ColLetterToNumber("L")
            .Row = 70
            .Text = "1"
            UpdateCell fps, .Col, .Row, .Text
        End If
            'end
        
        ' set up cho phep sua chi tieu 31
          If TAX_Utilities_v2.Year = 2011 And TAX_Utilities_v2.Month >= 8 And TAX_Utilities_v2.Month <= 12 Then
          Else
                .Sheet = 1
                .Col = .ColLetterToNumber("R")
                .Row = 52
                .Lock = False
          End If
        
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    ' xu ly cho chi tieu [31]
    Dim strValue As Variant
    With fps
        .EventEnabled(EventAllEvents) = False
        If .ActiveSheet = 1 Then
            If Col = .ColLetterToNumber("R") And Row = 51 Then
                If TAX_Utilities_v2.Year = 2011 And TAX_Utilities_v2.Month >= 8 And TAX_Utilities_v2.Month <= 12 Then
                    .GetText .ColLetterToNumber("R"), 51, strValue
                    If Val(CStr(strValue)) <= 5000000 Then
                        .Sheet = 1
                        .Col = .ColLetterToNumber("R")
                        .Row = 52
                        .Lock = True
                        .Formula = ""
                        mCurrentSheet = 1
                        .SetText .ColLetterToNumber("R"), 52, "0"
                        UpdateCell fps, .ColLetterToNumber("R"), 52, "0"
                    Else
                        .Col = .ColLetterToNumber("R")
                        .Row = 52
                        ' to khai quy tinh /3 * bieu thue luy tien
                        If strTkThangQuy = "TK_THANG" Then
                            .Formula = "AC52"
                        Else
                            .Formula = "AD52"
                        End If
                    End If
                Else
                    .Col = .ColLetterToNumber("R")
                    .Row = 52
                    .Lock = False
                    '.Formula = "AC50"
                    ' to khai quy tinh /3 * bieu thue luy tien
                    If strTkThangQuy = "TK_THANG" Then
                        .Formula = "AC52"
                    Else
                        .Formula = "AD52"
                    End If
                End If
            End If
        End If
        .EventEnabled(EventAllEvents) = True
    End With
End Sub

Private Sub fps_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim varTemp As Variant
    
    With fps
        .EventEnabled(EventAllEvents) = False
        mCurrentSheet = .Sheet
        If .ActiveSheet = 1 Then
        'check date
            If Col = .ColLetterToNumber("R") And Row = 68 Then
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
    Dim MST1 As Variant, MST2 As Variant, MST3 As Variant, MST4 As Variant
    Dim MST5 As Variant, MST6 As Variant, MST7 As Variant, MST8 As Variant
    Dim MST9 As Variant, MST10 As Variant, MST11 As Variant, MST12 As Variant, MST13 As Variant
    Dim iFlagTaxCode1 As Integer, iFlagTaxCode2 As Integer, iFlagTaxCode3 As Integer
    Dim strCheck As String, iCurrentSheet As Integer
    Dim error2 As Variant
    Dim errCT As Variant
    Dim CT22 As Variant
    Dim ct23 As Variant
    Dim ct24 As Variant
    Dim ct25 As Variant
    Dim ct26 As Variant
    Dim ct27 As Variant
    Dim ct28 As Variant
    Dim ct29 As Variant
    Dim ct30 As Variant
    Dim ct32 As Variant
    Dim CT33 As Variant
    Dim CT35 As Variant
    Dim CT36 As Variant
    Dim CT37 As Variant
    Dim CT38 As Variant
    Dim CT39 As Variant
    
    
    Dim iFlagEmty As Variant
    Dim iFlagCt24 As Integer, iFlagCt25 As Integer
    Dim errCT24 As Variant, errCT24_1 As Variant
    Dim errCT25 As Variant, errCT25_1 As Variant
    
    With fps
        .Sheet = .SheetCount
        .GetText .ColLetterToNumber("E"), 17, errCT
        
        .GetText .ColLetterToNumber("E"), 18, errCT24
        .GetText .ColLetterToNumber("E"), 16, errCT25
        .GetText .ColLetterToNumber("E"), 25, errCT24_1
        .GetText .ColLetterToNumber("E"), 26, errCT25_1
        
        .Sheet = 1
        'sua yc canh bao rang buoc du lieu các muc I,II,III

        .GetText .ColLetterToNumber("R"), 43, CT22
        .GetText .ColLetterToNumber("R"), 44, ct23
        
'        .GetText .ColLetterToNumber("R"), 45, CT24
        
        .GetText .ColLetterToNumber("R"), 46, ct25
        .GetText .ColLetterToNumber("R"), 47, ct26
        .GetText .ColLetterToNumber("R"), 48, ct27
        .GetText .ColLetterToNumber("R"), 49, ct28
        .GetText .ColLetterToNumber("R"), 50, ct29
                
        '.GetText .ColLetterToNumber("R"), 51, CT30
        .GetText .ColLetterToNumber("R"), 53, ct32
        '.GetText .ColLetterToNumber("R"), 57, CT35
        
        .GetText .ColLetterToNumber("R"), 57, CT35
        '.GetText .ColLetterToNumber("R"), 60, CT37
        .GetText .ColLetterToNumber("R"), 59, CT39
        
        iFlagCt24 = 0
        iFlagCt25 = 0
        
'        If Trim(CT37) <> 0 And Trim(CT37) = "1" Or Trim(CT37) = "2" Or Trim(CT37) = "5" Then
'            .Col = .ColLetterToNumber("R")
'            .Row = 60
'            .CellNote = ""
'            .BackColor = mNonErrorColor
'            iFlagEmty = 0
'        ElseIf Trim(CT37) <> 0 And Trim(CT37) <> "1" Or Trim(CT37) <> "" Or Trim(CT37) <> "5" Then
'            .Col = .ColLetterToNumber("R")
'            .Row = 60
'            .CellNote = errCT
'            .BackColor = mErrorColor
'            iFlagEmty = 1
'        Else
'            .Col = .ColLetterToNumber("R")
'            .Row = 60
'            .CellNote = ""
'            .BackColor = mNonErrorColor
'            iFlagEmty = 0
'        End If
        'sua yc canh bao rang buoc du lieu các muc I,II
        If (Val(CT22) <> 0 Or Val(ct23) <> 0 Or Val(ct25) <> 0 Or Val(ct26) <> 0 Or Val(ct27) <> 0 Or Val(ct28) <> 0 Or Val(ct29) <> 0 Or Val(ct32) <> 0) And (Val(CT35) <> 0 Or Val(CT37) <> 0) Then
                'DisplayMessage "0225", msOKOnly, miInformation
                iflag = True
        Else
                iflag = False
        End If
        
        If strTkThangQuy = "TK_THANG" Then
            If (Val(TAX_Utilities_v2.Month) >= 7 And Val(TAX_Utilities_v2.Year) = 2013) Or Val(TAX_Utilities_v2.Year) > 2013 Then
                ' kiem tra chi tieu 25
                If Val(ct25) = 9000000 Or Val(ct25) = 0 Then
                    .Row = 46
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                 Else
                    .Row = 46
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mErrorColor
                    .CellNote = errCT24_1
                    iFlagCt24 = 1
                End If
                ' kiem tra chi tieu 26
                 If Val(ct26) Mod 3600000 <> 0 Then
                    .Row = 47
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mErrorColor
                    .CellNote = errCT25_1
                    iFlagCt25 = 1
                 Else
                    .Row = 47
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                 End If
            Else
                ' kiem tra chi tieu 25
                If Val(ct25) = 4000000 Or Val(ct25) = 0 Then
                    .Row = 46
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                 Else
                    .Row = 46
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mErrorColor
                    .CellNote = errCT24
                    iFlagCt24 = 1
                End If
                ' kiem tra chi tieu 26
                If Val(ct26) Mod 1600000 <> 0 Then
                    .Row = 47
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mErrorColor
                    .CellNote = errCT25
                    iFlagCt25 = 1
                 Else
                    .Row = 47
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                 End If
            End If
        ElseIf strTkThangQuy = "TK_QUY" Then
            If (Val(TAX_Utilities_v2.ThreeMonths) >= 3 And Val(TAX_Utilities_v2.Year) = 2013) Or Val(TAX_Utilities_v2.Year) > 2013 Then
                ' chi tieu 25
                If Val(ct25) <= 27000000 And (Val(ct25) Mod 9000000 = 0) Then
                    .Row = 46
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                 Else
                    .Row = 46
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mErrorColor
                    .CellNote = errCT24_1
                    iFlagCt24 = 1
                End If
                ' chi tieu 26
                If Val(ct26) Mod 3600000 <> 0 Then
                    .Row = 47
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mErrorColor
                    .CellNote = errCT25_1
                    iFlagCt25 = 1
                 Else
                    .Row = 47
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                 End If
            Else
                ' chi tieu 25
                If Val(ct25) <= 12000000 And (Val(ct25) Mod 9000000 = 0) Then
                    .Row = 46
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                 Else
                    .Row = 46
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mErrorColor
                    .CellNote = errCT24
                    iFlagCt24 = 1
                End If
                ' chi tieu 26
                If Val(ct26) Mod 1600000 <> 0 Then
                    .Row = 47
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mErrorColor
                    .CellNote = errCT25
                    iFlagCt25 = 1
                 Else
                    .Row = 47
                    .Col = .ColLetterToNumber("R")
                    .BackColor = mNonErrorColor
                    .CellNote = ""
                 End If
            End If
        End If
        
         .Sheet = .SheetCount
         .SetText .ColLetterToNumber("B"), 17, IIf(iFlagEmty = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 18, IIf(iFlagCt24 = 1, "0", "1")
         .SetText .ColLetterToNumber("B"), 16, IIf(iFlagCt25 = 1, "0", "1")
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
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = 1
        mCurrentSheet = .Sheet
        .Col = .ColLetterToNumber("R")
        .Row = 58
        .Text = 20
        UpdateCell fps, .Col, .Row, .Text
        .EventEnabled(EventAllEvents) = True
    End With
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
               If lCol = .ColLetterToNumber("V") And lRow = 41 And (varTemp = "1" Or varTemp = "x") Then
                    UpdateCell fps, lCol, lRow, "x"
               ElseIf lCol = .ColLetterToNumber("V") And lRow = 41 And (varTemp = "0" Or varTemp = "") Then
                    UpdateCell fps, lCol, lRow, ""
               Else
                    UpdateCell fps, lCol, lRow, varTemp
               End If
             Next
        End If
        
      .Sheet = .ActiveSheet
      .EventEnabled(EventAllEvents) = True
    End With
    mCurrentSheet = ssSheet
      
    Set xmlCellNodeData = Nothing
    Set xmlCellNode = Nothing

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

