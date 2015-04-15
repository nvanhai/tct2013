VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCatalogue_2_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class is used to individual features of "Gia Tri Gia Tang" interface sheets
'this Class is belong to TAX_Business_v2 project which will be compline to DLL
'Coder: PhuocHK
Option Explicit
Private Const ID_TN = "D"
Private Const ID2_TN = "E"
Private Const TEN_TN = "F"
Private Const DON_VI_TN = "G"
Private Const THUE_SUAT_TN = "H"
Private Const THAY_DOI_TN = "I"
Private Const CHECK_TN = "B"

Private Const ID_TTDB = "D"
Private Const ID2_TTDB = "E"
Private Const TEN_TTDB = "F"
Private Const DON_VI_TTDB = "G"
Private Const THUE_SUAT_TTDB = "H"
Private Const GIA_BAO_BI_TTDB = "I"
Private Const THAY_DOI_TTDB = "J"
Private Const CHECK_TTDB = "B"


Public WithEvents fps As fpSpread
Attribute fps.VB_VarHelpID = -1

'This funtion is called after an object of this class is created
'Its functions is 1st preparing for interface sheets, such as
'add control, data for the control, celltag...
'no parameter
Public Sub Prepare1()
    fps.ScrollBars = ScrollBarsVertical
End Sub

'This funtion is called after executing function "SetupData"
'Its functions is 2st preparing for interface sheets
'no parameter
Public Sub Prepare2()
    Dim xmlCellNode As MSXML.IXMLDOMNode
    Dim strValue As String, intCtrl As Integer
    Dim lCol As Long, lRow As Long
    Dim strCheckAll As String
    Dim arrStrValue() As String
    
    With fps
        .EventEnabled(EventAllEvents) = False
        For intCtrl = 1 To 2
            .Sheet = intCtrl
            strCheckAll = "1"
            
            For Each xmlCellNode In TAX_Utilities_v2.Data(intCtrl - 1).getElementsByTagName("Cell")
                ParserCellID fps, GetAttribute(xmlCellNode, "CellID"), lCol, lRow
                arrStrValue = Split(GetAttribute(xmlCellNode, "Value"), "###")
                                
                .Row = lRow
                
                'Set gia tri bao bi
                If intCtrl = 2 Then
                    .Col = .ColLetterToNumber(GIA_BAO_BI_TTDB)
                    .value = arrStrValue(5)
                End If
                
                .Col = 2
                .value = arrStrValue(UBound(arrStrValue))
                
                If .value = "0" Then
                    strCheckAll = "0"
                End If
            Next
            .Row = 2
            .Col = 2
            .value = strCheckAll
        Next
        .EventEnabled(EventAllEvents) = True
    End With
End Sub



'Calling this function whenever we don't need this class anymore.
'It will do some special task with special Error/ special Data
'no parameter
Public Sub finish()
    
End Sub

'******************************
'Description: Set active properties in validity node
'Author: ThanhDX
'Date: 24/11/2005

'******************************
Public Sub SetActiveSheet() 'ByRef xmlNodeValid As MSXML.IXMLDOMNode)
    
End Sub

'******************************
'Description: Reset all error satus in sheet contain dynamic error.
'Author: ThanhDX
'Date: 31/12/2005

'******************************
Public Sub ResetErrorCells()
    
End Sub

'*******************************************************
'Description: ResetData procedure reset specified cells
'Author:ThanhDX
'Date:23/01/2006

'*******************************************************
Public Sub ResetData()
   
End Sub

'*******************************************************
'Description: SetData procedure set specified cells
'Author:ThanhDX
'Date:04/02/2006

'*******************************************************
Public Sub SetData()

End Sub


Private Sub fps_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim intCtrl As Integer
    Dim strCheckAll As String
    Dim strValue As String, varValue As Variant
    
    With fps
        .EventEnabled(EventAllEvents) = False
        .Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        
        If mCurrentSheet = 1 Then
        'Danh muc tai nguyen
            
            .Col = Col
            Select Case Row
                Case 2
                    .Row = Row
                    strCheckAll = .value
                    For intCtrl = 3 To .MaxRows
                        .Row = intCtrl
                        
                        .GetText .ColLetterToNumber(ID_TN), .Row, varValue
                        strValue = Trim(CStr(varValue)) & "###"
                        .GetText .ColLetterToNumber(ID2_TN), .Row, varValue
                        strValue = strValue & Trim(CStr(varValue)) & "###"
                        .GetText .ColLetterToNumber(TEN_TN), .Row, varValue
                        strValue = strValue & Trim(CStr(varValue)) & "###"
                        .GetText .ColLetterToNumber(DON_VI_TN), .Row, varValue
                        strValue = strValue & Trim(CStr(varValue)) & "###"
                        .GetText .ColLetterToNumber(THUE_SUAT_TN), .Row, varValue
                        strValue = strValue & Trim(CStr(varValue)) & "###"
                        .GetText .ColLetterToNumber(THAY_DOI_TN), .Row, varValue
                        strValue = strValue & Trim(CStr(varValue)) & "###"
                        
                        .Col = Col
                        If .CellType = CellTypeCheckBox Then
                            .value = strCheckAll
                            If .value = "1" Then
                                strValue = strValue & "1"
                            Else
                                strValue = strValue & "0"
                            End If
                            UpdateCell fps, 3, intCtrl, strValue
                        End If
                    Next intCtrl
                    
                    TAX_Utilities_v2.AdjustData(.ActiveSheet - 1) = True
                Case Else
                    .Row = Row
                    
                    .GetText .ColLetterToNumber(ID_TN), .Row, varValue
                    strValue = Trim(CStr(varValue)) & "###"
                    .GetText .ColLetterToNumber(ID2_TN), .Row, varValue
                    strValue = strValue & Trim(CStr(varValue)) & "###"
                    .GetText .ColLetterToNumber(TEN_TN), .Row, varValue
                    strValue = strValue & Trim(CStr(varValue)) & "###"
                    .GetText .ColLetterToNumber(DON_VI_TN), .Row, varValue
                    strValue = strValue & Trim(CStr(varValue)) & "###"
                    .GetText .ColLetterToNumber(THUE_SUAT_TN), .Row, varValue
                    strValue = strValue & Trim(CStr(varValue)) & "###"
                    .GetText .ColLetterToNumber(THAY_DOI_TN), .Row, varValue
                    strValue = strValue & Trim(CStr(varValue)) & "###"
                    
                    If .value = "1" Then
                        strValue = strValue & "1"
                    Else
                        strValue = strValue & "0"
                    End If
                    UpdateCell fps, 3, Row, strValue
                    
                    strCheckAll = "1"
                    For intCtrl = 3 To .MaxRows
                        .Col = Col
                        .Row = intCtrl
                        If .CellType = CellTypeCheckBox And .value = "0" Then
                            strCheckAll = "0"
                            Exit For
                        End If
                    Next intCtrl
                    .Row = 2
                    .value = strCheckAll
            End Select
        
        ElseIf mCurrentSheet = 2 Then
        'Danh muc tieu thu dac biet
                        
            Select Case Row
                Case 2
                    .Col = Col
                    .Row = Row
                    strCheckAll = .value
                    For intCtrl = 3 To .MaxRows
                        .Row = intCtrl
                        
                        .GetText .ColLetterToNumber(ID_TTDB), .Row, varValue
                        strValue = Trim(CStr(varValue)) & "###"
                        .GetText .ColLetterToNumber(ID2_TTDB), .Row, varValue
                        strValue = strValue & Trim(CStr(varValue)) & "###"
                        .GetText .ColLetterToNumber(TEN_TTDB), .Row, varValue
                        strValue = strValue & Trim(CStr(varValue)) & "###"
                        .GetText .ColLetterToNumber(DON_VI_TTDB), .Row, varValue
                        strValue = strValue & Trim(CStr(varValue)) & "###"
                        .GetText .ColLetterToNumber(THUE_SUAT_TTDB), .Row, varValue
                        strValue = strValue & Trim(CStr(varValue)) & "###"
                        .GetText .ColLetterToNumber(GIA_BAO_BI_TTDB), .Row, varValue
                        strValue = strValue & Trim(CStr(varValue)) & "###"
                        .GetText .ColLetterToNumber(THAY_DOI_TTDB), .Row, varValue
                        strValue = strValue & Trim(CStr(varValue)) & "###"
                        
                        .Col = Col
                        If .CellType = CellTypeCheckBox Then
                            .value = strCheckAll
                            If .value = "1" Then
                                strValue = strValue & "1"
                            Else
                                strValue = strValue & "0"
                            End If
                            UpdateCell fps, 3, intCtrl, strValue
                        End If
                    Next intCtrl
                    
                    TAX_Utilities_v2.AdjustData(.ActiveSheet - 1) = True
                Case Else
                    .Col = Col
                    .Row = Row
                    
                    .GetText .ColLetterToNumber(ID_TTDB), .Row, varValue
                    strValue = Trim(CStr(varValue)) & "###"
                    .GetText .ColLetterToNumber(ID2_TTDB), .Row, varValue
                    strValue = strValue & Trim(CStr(varValue)) & "###"
                    .GetText .ColLetterToNumber(TEN_TTDB), .Row, varValue
                    strValue = strValue & Trim(CStr(varValue)) & "###"
                    .GetText .ColLetterToNumber(DON_VI_TTDB), .Row, varValue
                    strValue = strValue & Trim(CStr(varValue)) & "###"
                    .GetText .ColLetterToNumber(THUE_SUAT_TTDB), .Row, varValue
                    strValue = strValue & Trim(CStr(varValue)) & "###"
                    .GetText .ColLetterToNumber(GIA_BAO_BI_TTDB), .Row, varValue
                    strValue = strValue & Trim(CStr(varValue)) & "###"
                    .GetText .ColLetterToNumber(THAY_DOI_TTDB), .Row, varValue
                    strValue = strValue & Trim(CStr(varValue)) & "###"
                    
                    If .value = "1" Then
                        strValue = strValue & "1"
                    Else
                        strValue = strValue & "0"
                    End If
                    UpdateCell fps, 3, Row, strValue
                    
                    strCheckAll = "1"
                    For intCtrl = 3 To .MaxRows
                        .Col = Col
                        .Row = intCtrl
                        If .CellType = CellTypeCheckBox And .value = "0" Then
                            strCheckAll = "0"
                            Exit For
                        End If
                    Next intCtrl
                    .Row = 2
                    .value = strCheckAll
            End Select
        End If
        
        .EventEnabled(EventAllEvents) = True
    End With
        
End Sub

Private Sub fps_Change(ByVal Col As Long, ByVal Row As Long)
    Dim varValue As Variant
    Dim strValue As String
    
    With fps
        '.Sheet = .ActiveSheet
        mCurrentSheet = .Sheet
        .Col = Col
        .Row = Row
        If .CellType = CellTypeCheckBox Then
            TAX_Utilities_v2.AdjustData(.ActiveSheet - 1) = True
            Exit Sub
        End If
        
        If mCurrentSheet = 2 And Col = .ColLetterToNumber(GIA_BAO_BI_TTDB) Then
        'Danh muc tieu thu dac biet
            .Col = Col
            .Row = Row
            
            .GetText .ColLetterToNumber(ID_TTDB), .Row, varValue
            strValue = Trim(CStr(varValue)) & "###"
            .GetText .ColLetterToNumber(ID2_TTDB), .Row, varValue
            strValue = strValue & Trim(CStr(varValue)) & "###"
            .GetText .ColLetterToNumber(TEN_TTDB), .Row, varValue
            strValue = strValue & Trim(CStr(varValue)) & "###"
            .GetText .ColLetterToNumber(DON_VI_TTDB), .Row, varValue
            strValue = strValue & Trim(CStr(varValue)) & "###"
            .GetText .ColLetterToNumber(THUE_SUAT_TTDB), .Row, varValue
            strValue = strValue & Trim(CStr(varValue)) & "###"
            .GetText .ColLetterToNumber(GIA_BAO_BI_TTDB), .Row, varValue
            strValue = strValue & Trim(CStr(varValue)) & "###"
            .GetText .ColLetterToNumber(THAY_DOI_TTDB), .Row, varValue
            strValue = strValue & Trim(CStr(varValue)) & "###"
            .GetText .ColLetterToNumber(CHECK_TTDB), .Row, varValue
            strValue = strValue & Trim(CStr(varValue))
                        
            UpdateCell fps, 3, Row, strValue
            
            TAX_Utilities_v2.AdjustData(.ActiveSheet - 1) = True
            
        End If
    End With
    
End Sub

Public Function DeleteEnable(KeyCode As Integer, Shift As Integer) As Boolean
    DeleteEnable = True
End Function

Public Function InsertEnable(KeyCode As Integer, Shift As Integer) As Boolean
    InsertEnable = True
End Function
